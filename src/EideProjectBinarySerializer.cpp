#include "EideProjectBinarySerializer.h"

#include <Windows.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <cstdint>
#include <exception>
#include <filesystem>
#include <format>
#include <fstream>
#include <mutex>
#include <set>

#include "EideInternalTextBridge.h"
#include "Global.h"
#include "WindowHelper.h"

namespace e571 {
#if defined(_M_IX86)
namespace {

constexpr std::uintptr_t kImageBase = 0x400000;
constexpr std::uintptr_t kKnownDirectProjectSerializeToHandleRva = 0x45B640;
constexpr std::uintptr_t kKnownProjectCommandDispatchRva = 0x50A4F0;
constexpr std::uintptr_t kKnownProjectCommandDirect203Rva = 0x50A800;
constexpr std::uintptr_t kKnownProjectCommandDirect202Rva = 0x50A820;
constexpr size_t kActiveEditorSerializerFieldIndex = 23;
constexpr size_t kCommandObjectFieldScanCount = 96;
constexpr size_t kCommandObjectGraphMaxDepth = 2;
constexpr size_t kCommandObjectVtableScanCount = 96;
constexpr std::array<unsigned char, 8> kExpectedProjectMagic = {
	'C', 'N', 'W', 'T', 'E', 'P', 'R', 'G'
};

using FnSerializeCurrentProjectToHandle = HGLOBAL(__thiscall*)(void*);

struct SerializerAddresses {
	bool initialized = false;
	bool ok = false;
	std::uintptr_t moduleBase = 0;
	FnSerializeCurrentProjectToHandle directSerializeToHandle = nullptr;
	bool coldCommandRoutesOk = false;
	std::uintptr_t projectCommandDispatch = 0;
	std::uintptr_t projectCommandDirect203 = 0;
	std::uintptr_t projectCommandDirect202 = 0;
};

std::mutex g_serializerAddressMutex;
SerializerAddresses g_serializerAddresses;

struct SerializerContext {
	void* serializerThis = nullptr;
	std::string sourcePath;
};

enum class ColdCommandRoute {
	None,
	DispatchCommand,
	Direct2030001,
	Direct2020004,
};

struct ColdCommandCandidate {
	std::uintptr_t object = 0;
	std::uintptr_t vtable = 0;
	std::uintptr_t function = 0;
	ColdCommandRoute route = ColdCommandRoute::None;
	size_t slotIndex = 0;
	size_t depth = 0;
	std::string path;
};

std::mutex g_serializerContextMutex;
SerializerContext g_serializerContext;

std::string BuildMagicText(const std::vector<unsigned char>& bytes);

template <typename T>
T ResolveInternalAddress(std::uintptr_t moduleBase, std::uintptr_t rva)
{
	if (moduleBase == 0 || rva < kImageBase) {
		return nullptr;
	}
	return reinterpret_cast<T>(moduleBase + (rva - kImageBase));
}

bool IsReadableAddressRange(std::uintptr_t address, size_t bytes)
{
	if (address == 0 || bytes == 0) {
		return false;
	}

	MEMORY_BASIC_INFORMATION mbi = {};
	if (VirtualQuery(reinterpret_cast<LPCVOID>(address), &mbi, sizeof(mbi)) != sizeof(mbi)) {
		return false;
	}

	const DWORD protect = mbi.Protect & 0xFF;
	if (mbi.State != MEM_COMMIT || protect == PAGE_NOACCESS || protect == PAGE_GUARD) {
		return false;
	}

	const std::uintptr_t regionBase = reinterpret_cast<std::uintptr_t>(mbi.BaseAddress);
	const std::uintptr_t regionEnd = regionBase + static_cast<std::uintptr_t>(mbi.RegionSize);
	return address >= regionBase && address + bytes <= regionEnd;
}

std::string TrimAsciiCopyForSerializer(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string NormalizePathForSerializer(const std::string& pathText)
{
	if (pathText.empty()) {
		return std::string();
	}

	try {
		std::filesystem::path path(pathText);
		path = path.lexically_normal();
		if (path.is_relative()) {
			path = std::filesystem::absolute(path);
		}
		path = path.lexically_normal();
		return path.string();
	}
	catch (...) {
		return pathText;
	}
}

std::string ResolveCurrentSourcePathForSerializer()
{
	std::string sourcePath = TrimAsciiCopyForSerializer(g_nowOpenSourceFilePath);
	if (sourcePath.empty()) {
		sourcePath = TrimAsciiCopyForSerializer(GetSourceFilePath());
	}
	return NormalizePathForSerializer(sourcePath);
}

bool IsLikelyModuleCodeAddress(
	std::uintptr_t address,
	std::uintptr_t moduleBase)
{
	if (address == 0 || moduleBase == 0) {
		return false;
	}
	if (address < moduleBase || address >= (moduleBase + 0x4000000)) {
		return false;
	}
	return IsReadableAddressRange(address, 16);
}

bool IsLikelySerializerObjectAddress(
	void* serializerThis,
	std::uintptr_t moduleBase)
{
	const std::uintptr_t objectAddress = reinterpret_cast<std::uintptr_t>(serializerThis);
	if (objectAddress == 0 || moduleBase == 0) {
		return false;
	}
	if (objectAddress < moduleBase || objectAddress >= (moduleBase + 0x4000000)) {
		return false;
	}
	if (!IsReadableAddressRange(objectAddress, 0x40)) {
		return false;
	}

	__try {
		return *reinterpret_cast<const std::uintptr_t*>(objectAddress) != 0;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool TryReadPointerValue(std::uintptr_t address, std::uintptr_t* outValue)
{
	if (outValue != nullptr) {
		*outValue = 0;
	}
	if (address == 0 || !IsReadableAddressRange(address, sizeof(std::uintptr_t))) {
		return false;
	}

	__try {
		const std::uintptr_t value = *reinterpret_cast<const std::uintptr_t*>(address);
		if (outValue != nullptr) {
			*outValue = value;
		}
		return true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool IsLikelyObjectPointer(
	std::uintptr_t objectAddress,
	std::uintptr_t moduleBase,
	std::uintptr_t* outVtable = nullptr)
{
	if (outVtable != nullptr) {
		*outVtable = 0;
	}
	if (objectAddress < 0x10000 || !IsReadableAddressRange(objectAddress, sizeof(std::uintptr_t))) {
		return false;
	}

	std::uintptr_t vtable = 0;
	if (!TryReadPointerValue(objectAddress, &vtable)) {
		return false;
	}
	if (!IsLikelyModuleCodeAddress(vtable, moduleBase)) {
		return false;
	}

	if (outVtable != nullptr) {
		*outVtable = vtable;
	}
	return true;
}

std::string DescribeColdCommandRoute(ColdCommandRoute route)
{
	switch (route) {
	case ColdCommandRoute::DispatchCommand:
		return "dispatch_50A4F0";
	case ColdCommandRoute::Direct2030001:
		return "direct_50A800";
	case ColdCommandRoute::Direct2020004:
		return "direct_50A820";
	default:
		return "none";
	}
}

bool HasExpectedProjectMagic(const std::vector<unsigned char>& bytes)
{
	return
		bytes.size() >= kExpectedProjectMagic.size() &&
		std::equal(
			kExpectedProjectMagic.begin(),
			kExpectedProjectMagic.end(),
			bytes.begin());
}

void AppendTraceSummary(std::vector<std::string>& traces, const std::string& text, size_t maxEntries = 16)
{
	if (text.empty() || traces.size() >= maxEntries) {
		return;
	}
	traces.push_back(text);
}

void CollectRelatedCommandObjectCandidates(
	const SerializerAddresses& addrs,
	const ActiveEditorObjectInfo& activeInfo,
	std::vector<ColdCommandCandidate>& outCandidates,
	std::string* outTrace)
{
	outCandidates.clear();
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (addrs.moduleBase == 0) {
		if (outTrace != nullptr) {
			*outTrace = "candidate_collect_invalid_module_base";
		}
		return;
	}
	if (!addrs.coldCommandRoutesOk) {
		if (outTrace != nullptr) {
			*outTrace = "candidate_collect_command_routes_unavailable";
		}
		return;
	}

	struct PendingNode {
		std::uintptr_t object = 0;
		size_t depth = 0;
		std::string path;
	};

	std::set<std::uintptr_t> visitedObjects;
	std::set<std::uintptr_t> addedObjects;
	std::vector<PendingNode> pending;
	std::vector<std::string> traces;

	const auto enqueue = [&](std::uintptr_t object, size_t depth, const std::string& path) {
		if (object == 0 || !visitedObjects.insert(object).second) {
			return;
		}
		pending.push_back(PendingNode{ object, depth, path });
	};

	enqueue(activeInfo.rawEditorObject, 0, "raw");
	if (activeInfo.innerEditorObject != 0 && activeInfo.innerEditorObject != activeInfo.rawEditorObject) {
		enqueue(activeInfo.innerEditorObject, 0, "inner");
	}

	size_t scanCount = 0;
	for (size_t index = 0; index < pending.size(); ++index) {
		const PendingNode& node = pending[index];
		std::uintptr_t vtable = 0;
		if (!IsLikelyObjectPointer(node.object, addrs.moduleBase, &vtable)) {
			continue;
		}

		++scanCount;
		for (size_t slot = 0; slot < kCommandObjectVtableScanCount; ++slot) {
			std::uintptr_t functionAddress = 0;
			if (!TryReadPointerValue(vtable + slot * sizeof(std::uintptr_t), &functionAddress) ||
				functionAddress == 0) {
				continue;
			}

			ColdCommandRoute route = ColdCommandRoute::None;
			if (functionAddress == addrs.projectCommandDispatch) {
				route = ColdCommandRoute::DispatchCommand;
			}
			else if (functionAddress == addrs.projectCommandDirect203) {
				route = ColdCommandRoute::Direct2030001;
			}
			else if (functionAddress == addrs.projectCommandDirect202) {
				route = ColdCommandRoute::Direct2020004;
			}
			if (route == ColdCommandRoute::None || !addedObjects.insert(node.object).second) {
				continue;
			}

			outCandidates.push_back(ColdCommandCandidate{
				.object = node.object,
				.vtable = vtable,
				.function = functionAddress,
				.route = route,
				.slotIndex = slot,
				.depth = node.depth,
				.path = node.path,
			});
			AppendTraceSummary(
				traces,
				std::format(
					"candidate path={} depth={} object={} vtable={} slot={} route={}",
					node.path,
					node.depth,
					node.object,
					vtable,
					slot,
					DescribeColdCommandRoute(route)));
			break;
		}

		if (node.depth >= kCommandObjectGraphMaxDepth) {
			continue;
		}

		for (size_t fieldIndex = 1; fieldIndex < kCommandObjectFieldScanCount; ++fieldIndex) {
			std::uintptr_t childObject = 0;
			if (!TryReadPointerValue(node.object + fieldIndex * sizeof(std::uintptr_t), &childObject) ||
				childObject == 0 ||
				childObject == node.object) {
				continue;
			}
			std::uintptr_t childVtable = 0;
			if (!IsLikelyObjectPointer(childObject, addrs.moduleBase, &childVtable)) {
				continue;
			}

			enqueue(
				childObject,
				node.depth + 1,
				std::format("{}.f{}", node.path, fieldIndex));
		}
	}

	if (outTrace != nullptr) {
		std::string trace = std::format(
			"candidate_collect_ok scanned={} candidates={} raw={} inner={}",
			scanCount,
			outCandidates.size(),
			activeInfo.rawEditorObject,
			activeInfo.innerEditorObject);
		for (const auto& item : traces) {
			trace += "|" + item;
		}
		*outTrace = trace;
	}
}

bool TrySerializeCurrentProjectFromActiveEditorCommand(
	const SerializerAddresses& addrs,
	std::vector<unsigned char>& outBytes,
	std::string* outError,
	std::string* outTrace)
{
	outBytes.clear();
	if (outError != nullptr) {
		outError->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	ActiveEditorObjectInfo activeInfo{};
	if (!ResolveCurrentActiveEditorObject(addrs.moduleBase, &activeInfo) || !activeInfo.ok) {
		if (outError != nullptr) {
			*outError = "resolve active editor object failed";
		}
		if (outTrace != nullptr) {
			*outTrace = activeInfo.trace.empty()
				? "resolve_active_editor_failed"
				: activeInfo.trace;
		}
		return false;
	}

	std::vector<ColdCommandCandidate> candidates;
	std::string collectTrace;
	CollectRelatedCommandObjectCandidates(addrs, activeInfo, candidates, &collectTrace);
	if (candidates.empty()) {
		if (outError != nullptr) {
			*outError = "project command object not found from active editor";
		}
		if (outTrace != nullptr) {
			*outTrace = activeInfo.trace + "|" + collectTrace;
		}
		return false;
	}

	std::vector<std::string> attemptTraces;
	for (const auto& candidate : candidates) {
		std::vector<InternalThiscallCommandSpec> specs;
		switch (candidate.route) {
		case ColdCommandRoute::DispatchCommand:
			specs.push_back(InternalThiscallCommandSpec{
				candidate.object,
				candidate.function,
				0x2020004,
				0,
				0,
				0,
			});
			specs.push_back(InternalThiscallCommandSpec{
				candidate.object,
				candidate.function,
				0x2030001,
				0,
				0,
				0,
			});
			break;
		case ColdCommandRoute::Direct2030001:
			specs.push_back(InternalThiscallCommandSpec{
				candidate.object,
				candidate.function,
				0x2030001,
				0,
				0,
				0,
			});
			break;
		case ColdCommandRoute::Direct2020004:
			specs.push_back(InternalThiscallCommandSpec{
				candidate.object,
				candidate.function,
				0x2020004,
				0,
				0,
				0,
			});
			break;
		default:
			break;
		}

		for (const auto& spec : specs) {
			std::vector<unsigned char> capturedBytes;
			std::string captureTrace;
			if (!CaptureCustomClipboardPayloadByThiscall(spec, capturedBytes, &captureTrace)) {
				AppendTraceSummary(
					attemptTraces,
					std::format(
						"attempt_failed path={} route={} slot={} trace={}",
						candidate.path,
						DescribeColdCommandRoute(candidate.route),
						candidate.slotIndex,
						captureTrace));
				continue;
			}

			if (!HasExpectedProjectMagic(capturedBytes)) {
				AppendTraceSummary(
					attemptTraces,
					std::format(
						"attempt_bad_magic path={} route={} slot={} bytes={} magic={} trace={}",
						candidate.path,
						DescribeColdCommandRoute(candidate.route),
						candidate.slotIndex,
						capturedBytes.size(),
						BuildMagicText(capturedBytes),
						captureTrace));
				continue;
			}

			outBytes = std::move(capturedBytes);
			if (outTrace != nullptr) {
				*outTrace =
					activeInfo.trace +
					"|" +
					collectTrace +
					"|" +
					std::format(
						"cold_start_command_ok path={} depth={} object={} vtable={} slot={} route={} bytes={} magic={}",
						candidate.path,
						candidate.depth,
						candidate.object,
						candidate.vtable,
						candidate.slotIndex,
						DescribeColdCommandRoute(candidate.route),
						outBytes.size(),
						BuildMagicText(outBytes)) +
					"|" +
					captureTrace;
			}
			return true;
		}
	}

	if (outError != nullptr) {
		*outError = "active editor command serialize failed";
	}
	if (outTrace != nullptr) {
		std::string trace = activeInfo.trace + "|" + collectTrace;
		for (const auto& item : attemptTraces) {
			trace += "|" + item;
		}
		*outTrace = trace;
	}
	return false;
}

bool PopulateSerializerAddresses(SerializerAddresses& addrs)
{
	addrs = {};
	addrs.moduleBase = reinterpret_cast<std::uintptr_t>(GetModuleHandleA(nullptr));
	if (addrs.moduleBase == 0) {
		addrs.initialized = true;
		return false;
	}

	addrs.directSerializeToHandle = ResolveInternalAddress<FnSerializeCurrentProjectToHandle>(
		addrs.moduleBase,
		kKnownDirectProjectSerializeToHandleRva);
	addrs.projectCommandDispatch = reinterpret_cast<std::uintptr_t>(
		ResolveInternalAddress<void*>(addrs.moduleBase, kKnownProjectCommandDispatchRva));
	addrs.projectCommandDirect203 = reinterpret_cast<std::uintptr_t>(
		ResolveInternalAddress<void*>(addrs.moduleBase, kKnownProjectCommandDirect203Rva));
	addrs.projectCommandDirect202 = reinterpret_cast<std::uintptr_t>(
		ResolveInternalAddress<void*>(addrs.moduleBase, kKnownProjectCommandDirect202Rva));
	addrs.ok =
		addrs.directSerializeToHandle != nullptr &&
		IsLikelyModuleCodeAddress(
			reinterpret_cast<std::uintptr_t>(addrs.directSerializeToHandle),
			addrs.moduleBase);
	addrs.coldCommandRoutesOk =
		IsLikelyModuleCodeAddress(addrs.projectCommandDispatch, addrs.moduleBase) &&
		IsLikelyModuleCodeAddress(addrs.projectCommandDirect203, addrs.moduleBase) &&
		IsLikelyModuleCodeAddress(addrs.projectCommandDirect202, addrs.moduleBase);
	addrs.initialized = true;
	return addrs.ok;
}

const SerializerAddresses& GetSerializerAddresses()
{
	std::lock_guard<std::mutex> lock(g_serializerAddressMutex);
	if (!g_serializerAddresses.initialized) {
		PopulateSerializerAddresses(g_serializerAddresses);
	}
	return g_serializerAddresses;
}

HGLOBAL CallDirectSerializeProjectToHandleSafe(
	FnSerializeCurrentProjectToHandle serializeFn,
	void* serializerThis)
{
	if (serializeFn == nullptr || serializerThis == nullptr) {
		return nullptr;
	}

	__try {
		return serializeFn(serializerThis);
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return nullptr;
	}
}

bool TryGetVerifiedSerializerContext(
	std::uintptr_t moduleBase,
	void*& outSerializerThis,
	std::string& outSourcePath,
	std::string* outReason)
{
	outSerializerThis = nullptr;
	outSourcePath.clear();
	if (outReason != nullptr) {
		outReason->clear();
	}

	const std::string currentSourcePath = ResolveCurrentSourcePathForSerializer();

	std::lock_guard<std::mutex> lock(g_serializerContextMutex);
	if (g_serializerContext.serializerThis == nullptr) {
		if (outReason != nullptr) {
			*outReason = "serializer_this_missing";
		}
		return false;
	}

	const std::string recordedSourcePath = NormalizePathForSerializer(g_serializerContext.sourcePath);
	if (!currentSourcePath.empty() &&
		!recordedSourcePath.empty() &&
		_stricmp(currentSourcePath.c_str(), recordedSourcePath.c_str()) != 0) {
		if (outReason != nullptr) {
			*outReason = "serializer_source_mismatch current=" + currentSourcePath + " recorded=" + recordedSourcePath;
		}
		return false;
	}

	outSerializerThis = g_serializerContext.serializerThis;
	outSourcePath = recordedSourcePath;
	return true;
}

std::string BuildMagicText(const std::vector<unsigned char>& bytes)
{
	std::string text;
	const size_t count = (std::min)(bytes.size(), static_cast<size_t>(8));
	text.reserve(count);
	for (size_t index = 0; index < count; ++index) {
		const unsigned char ch = bytes[index];
		text.push_back((ch >= 32 && ch <= 126) ? static_cast<char>(ch) : '.');
	}
	return text;
}

bool CopyGlobalHandleBytes(
	HGLOBAL handle,
	std::vector<unsigned char>& outBytes,
	std::string* outError)
{
	outBytes.clear();
	if (outError != nullptr) {
		outError->clear();
	}
	if (handle == nullptr) {
		if (outError != nullptr) {
			*outError = "serialize returned null handle";
		}
		return false;
	}

	const SIZE_T size = GlobalSize(handle);
	if (size == 0) {
		GlobalFree(handle);
		if (outError != nullptr) {
			*outError = "serialized global handle is empty";
		}
		return false;
	}

	const void* data = GlobalLock(handle);
	if (data == nullptr) {
		GlobalFree(handle);
		if (outError != nullptr) {
			*outError = "GlobalLock failed";
		}
		return false;
	}

	const auto* byteData = static_cast<const unsigned char*>(data);
	outBytes.assign(byteData, byteData + size);
	GlobalUnlock(handle);
	GlobalFree(handle);

	if (outBytes.size() < kExpectedProjectMagic.size() ||
		!std::equal(
			kExpectedProjectMagic.begin(),
			kExpectedProjectMagic.end(),
			outBytes.begin())) {
		if (outError != nullptr) {
			*outError = "serialized bytes do not start with CNWTEPRG";
		}
		outBytes.clear();
		return false;
	}

	return true;
}

bool TrySerializeCurrentProjectFromActiveEditorSerializerField(
	const SerializerAddresses& addrs,
	std::vector<unsigned char>& outBytes,
	std::string* outError,
	std::string* outTrace)
{
	outBytes.clear();
	if (outError != nullptr) {
		outError->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	ActiveEditorObjectInfo activeInfo{};
	if (!ResolveCurrentActiveEditorObject(addrs.moduleBase, &activeInfo) || !activeInfo.ok) {
		if (outError != nullptr) {
			*outError = "resolve active editor object failed";
		}
		if (outTrace != nullptr) {
			*outTrace = activeInfo.trace.empty()
				? "resolve_active_editor_failed"
				: activeInfo.trace;
		}
		return false;
	}

	struct SerializerOwnerCandidate {
		std::uintptr_t object = 0;
		const char* name = nullptr;
	};

	std::vector<SerializerOwnerCandidate> owners;
	if (activeInfo.rawEditorObject != 0) {
		owners.push_back(SerializerOwnerCandidate{ activeInfo.rawEditorObject, "raw" });
	}
	if (activeInfo.innerEditorObject != 0 &&
		activeInfo.innerEditorObject != activeInfo.rawEditorObject) {
		owners.push_back(SerializerOwnerCandidate{ activeInfo.innerEditorObject, "inner" });
	}

	std::vector<std::string> attemptTraces;
	for (const auto& owner : owners) {
		const std::uintptr_t fieldAddress =
			owner.object + kActiveEditorSerializerFieldIndex * sizeof(std::uintptr_t);
		std::uintptr_t serializerAddress = 0;
		if (!TryReadPointerValue(fieldAddress, &serializerAddress) || serializerAddress == 0) {
			AppendTraceSummary(
				attemptTraces,
				std::format(
					"field_read_failed owner={} owner_object={} field={} field_addr={}",
					owner.name,
					owner.object,
					kActiveEditorSerializerFieldIndex,
					fieldAddress));
			continue;
		}

		std::uintptr_t serializerHeader = 0;
		(void)TryReadPointerValue(serializerAddress, &serializerHeader);
		if (!IsLikelySerializerObjectAddress(reinterpret_cast<void*>(serializerAddress), addrs.moduleBase)) {
			AppendTraceSummary(
				attemptTraces,
				std::format(
					"field_bad_object owner={} owner_object={} field={} field_addr={} serializer_this={} serializer_head={}",
					owner.name,
					owner.object,
					kActiveEditorSerializerFieldIndex,
					fieldAddress,
					serializerAddress,
					serializerHeader));
			continue;
		}

		const HGLOBAL handle = CallDirectSerializeProjectToHandleSafe(
			addrs.directSerializeToHandle,
			reinterpret_cast<void*>(serializerAddress));
		std::string directError;
		if (!CopyGlobalHandleBytes(handle, outBytes, &directError)) {
			AppendTraceSummary(
				attemptTraces,
				std::format(
					"field_direct_failed owner={} owner_object={} field={} field_addr={} serializer_this={} serializer_head={} error={}",
					owner.name,
					owner.object,
					kActiveEditorSerializerFieldIndex,
					fieldAddress,
					serializerAddress,
					serializerHeader,
					directError));
			continue;
		}

		if (outTrace != nullptr) {
			*outTrace =
				activeInfo.trace +
				"|" +
				std::format(
					"active_editor_serializer_ok owner={} owner_object={} field={} field_addr={} serializer_this={} serializer_head={} bytes={} magic={}",
					owner.name,
					owner.object,
					kActiveEditorSerializerFieldIndex,
					fieldAddress,
					serializerAddress,
					serializerHeader,
					outBytes.size(),
					BuildMagicText(outBytes));
		}
		return true;
	}

	if (outError != nullptr) {
		*outError = "active editor serializer field direct call failed";
	}
	if (outTrace != nullptr) {
		std::string trace = activeInfo.trace;
		for (const auto& item : attemptTraces) {
			trace += "|" + item;
		}
		*outTrace = trace;
	}
	return false;
}

} // namespace
#endif

ProjectBinarySerializer& ProjectBinarySerializer::Instance()
{
	static ProjectBinarySerializer instance;
	return instance;
}

void ProjectBinarySerializer::RecordVerifiedSerializerContext(
	void* serializerThis,
	const std::string& sourcePath)
{
#if defined(_M_IX86)
	if (serializerThis == nullptr) {
		return;
	}

	const std::string normalizedSourcePath = NormalizePathForSerializer(sourcePath);
	{
		std::lock_guard<std::mutex> lock(g_serializerContextMutex);
		g_serializerContext.serializerThis = serializerThis;
		g_serializerContext.sourcePath = normalizedSourcePath;
	}

	OutputStringToELog(std::format(
		"[ProjectBinarySerializer] captured serializer_this={} source={}",
		reinterpret_cast<std::uintptr_t>(serializerThis),
		normalizedSourcePath));
#else
	(void)serializerThis;
	(void)sourcePath;
#endif
}

void ProjectBinarySerializer::ClearVerifiedSerializerContext()
{
#if defined(_M_IX86)
	std::lock_guard<std::mutex> lock(g_serializerContextMutex);
	g_serializerContext = {};
#endif
}

bool ProjectBinarySerializer::SerializeCurrentProject(
	std::vector<unsigned char>& outBytes,
	std::string* outError,
	std::string* outTrace)
{
#if !defined(_M_IX86)
	outBytes.clear();
	if (outError != nullptr) {
		*outError = "current project binary serializer is only available on x86 IDE builds";
	}
	if (outTrace != nullptr) {
		*outTrace = "unsupported_platform";
	}
	return false;
#else
	outBytes.clear();
	if (outError != nullptr) {
		outError->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	const auto& addrs = GetSerializerAddresses();
	if (!addrs.ok) {
		if (outError != nullptr) {
			*outError = "project serializer addresses unavailable";
		}
		if (outTrace != nullptr) {
			*outTrace =
				"resolve_failed"
				"|module_base=" + std::to_string(addrs.moduleBase) +
				"|direct=" + std::to_string(reinterpret_cast<std::uintptr_t>(addrs.directSerializeToHandle)) +
				"|cold_routes_ok=" + std::to_string(addrs.coldCommandRoutesOk ? 1 : 0) +
				"|dispatch=" + std::to_string(addrs.projectCommandDispatch) +
				"|direct203=" + std::to_string(addrs.projectCommandDirect203) +
				"|direct202=" + std::to_string(addrs.projectCommandDirect202);
		}
		return false;
	}

	void* serializerThis = nullptr;
	std::string serializerSourcePath;
	std::string serializerContextReason;
	const bool hasVerifiedContext = TryGetVerifiedSerializerContext(
		addrs.moduleBase,
		serializerThis,
		serializerSourcePath,
		&serializerContextReason);
	if (hasVerifiedContext) {
		const HGLOBAL handle = CallDirectSerializeProjectToHandleSafe(
			addrs.directSerializeToHandle,
			serializerThis);
		std::string directError;
		if (CopyGlobalHandleBytes(handle, outBytes, &directError)) {
			if (outTrace != nullptr) {
				*outTrace =
					"serialize_ok"
					"|route=direct"
					"|serialize=" + std::to_string(reinterpret_cast<std::uintptr_t>(addrs.directSerializeToHandle)) +
					"|serializer_this=" + std::to_string(reinterpret_cast<std::uintptr_t>(serializerThis)) +
					"|source=" + serializerSourcePath +
					"|bytes=" + std::to_string(outBytes.size()) +
					"|magic=" + BuildMagicText(outBytes);
			}
			return true;
		}

		std::vector<unsigned char> activeEditorBytes;
		std::string activeEditorError;
		std::string activeEditorTrace;
		if (TrySerializeCurrentProjectFromActiveEditorSerializerField(
				addrs,
				activeEditorBytes,
				&activeEditorError,
				&activeEditorTrace)) {
			outBytes = std::move(activeEditorBytes);
			if (outTrace != nullptr) {
				*outTrace =
					"direct_route_failed_fallback_active_editor"
					"|direct=" + std::to_string(reinterpret_cast<std::uintptr_t>(addrs.directSerializeToHandle)) +
					"|serializer_this=" + std::to_string(reinterpret_cast<std::uintptr_t>(serializerThis)) +
					"|source=" + serializerSourcePath +
					"|direct_error=" + directError +
					"|" +
					activeEditorTrace;
			}
			return true;
		}

		std::vector<unsigned char> coldBytes;
		std::string coldError;
		std::string coldTrace;
		if (TrySerializeCurrentProjectFromActiveEditorCommand(addrs, coldBytes, &coldError, &coldTrace)) {
			outBytes = std::move(coldBytes);
			if (outTrace != nullptr) {
				*outTrace =
					"direct_route_failed_fallback_cold_start"
					"|direct=" + std::to_string(reinterpret_cast<std::uintptr_t>(addrs.directSerializeToHandle)) +
					"|serializer_this=" + std::to_string(reinterpret_cast<std::uintptr_t>(serializerThis)) +
					"|source=" + serializerSourcePath +
					"|direct_error=" + directError +
					"|active_editor_error=" + activeEditorError +
					"|" +
					coldTrace;
			}
			return true;
		}

		if (outError != nullptr) {
			*outError = directError.empty()
				? "serialize current project handle failed"
				: "direct serializer failed: " + directError;
			if (!activeEditorError.empty()) {
				*outError += " | active_editor: " + activeEditorError;
			}
			if (!coldError.empty()) {
				*outError += " | cold_start: " + coldError;
			}
		}
		if (outTrace != nullptr) {
			*outTrace =
				"serialize_call_failed"
				"|direct=" + std::to_string(reinterpret_cast<std::uintptr_t>(addrs.directSerializeToHandle)) +
				"|serializer_this=" + std::to_string(reinterpret_cast<std::uintptr_t>(serializerThis)) +
				"|source=" + serializerSourcePath +
				"|direct_error=" + directError +
				"|active_editor_error=" + activeEditorError +
				"|cold_error=" + coldError +
				"|" +
				activeEditorTrace +
				"|" +
				coldTrace;
		}
		return false;
	}

	std::vector<unsigned char> activeEditorBytes;
	std::string activeEditorError;
	std::string activeEditorTrace;
	if (TrySerializeCurrentProjectFromActiveEditorSerializerField(
			addrs,
			activeEditorBytes,
			&activeEditorError,
			&activeEditorTrace)) {
		outBytes = std::move(activeEditorBytes);
		if (outTrace != nullptr) {
			*outTrace =
				"serialize_ok"
				"|route=active_editor_field23_direct"
				"|context_reason=" + serializerContextReason +
				"|" +
				activeEditorTrace;
		}
		return true;
	}

	std::vector<unsigned char> coldBytes;
	std::string coldError;
	std::string coldTrace;
	if (TrySerializeCurrentProjectFromActiveEditorCommand(addrs, coldBytes, &coldError, &coldTrace)) {
		outBytes = std::move(coldBytes);
		if (outTrace != nullptr) {
			*outTrace =
				"serialize_ok"
				"|route=cold_start_active_editor_command"
				"|context_reason=" + serializerContextReason +
				"|" +
				coldTrace;
		}
		return true;
	}

	if (outError != nullptr) {
		*outError = activeEditorError.empty()
			? "project serializer context unavailable"
			: activeEditorError;
		if (!coldError.empty()) {
			*outError += " | cold_start: " + coldError;
		}
	}
	if (outTrace != nullptr) {
		*outTrace =
			"serializer_context_unavailable"
			"|direct=" + std::to_string(reinterpret_cast<std::uintptr_t>(addrs.directSerializeToHandle)) +
			"|reason=" + serializerContextReason +
			"|active_editor_error=" + activeEditorError +
			"|cold_error=" + coldError +
			"|" +
			activeEditorTrace +
			"|" +
			coldTrace;
	}
	return false;
#endif
}

bool ProjectBinarySerializer::WriteCurrentProjectToFile(
	const std::string& outputPath,
	size_t* outBytesWritten,
	std::string* outError,
	std::string* outTrace)
{
#if !defined(_M_IX86)
	if (outBytesWritten != nullptr) {
		*outBytesWritten = 0;
	}
	if (outError != nullptr) {
		*outError = "current project binary serializer is only available on x86 IDE builds";
	}
	if (outTrace != nullptr) {
		*outTrace = "unsupported_platform";
	}
	return false;
#else
	if (outBytesWritten != nullptr) {
		*outBytesWritten = 0;
	}
	if (outError != nullptr) {
		outError->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	std::vector<unsigned char> bytes;
	std::string serializeError;
	std::string serializeTrace;
	if (!SerializeCurrentProject(bytes, &serializeError, &serializeTrace)) {
		if (outError != nullptr) {
			*outError = serializeError;
		}
		if (outTrace != nullptr) {
			*outTrace = serializeTrace;
		}
		return false;
	}

	try {
		std::filesystem::path path(outputPath);
		if (path.has_parent_path()) {
			std::error_code ec;
			std::filesystem::create_directories(path.parent_path(), ec);
		}

		std::ofstream out(path, std::ios::binary | std::ios::trunc);
		if (!out.is_open()) {
			if (outError != nullptr) {
				*outError = "open output snapshot file failed";
			}
			if (outTrace != nullptr) {
				*outTrace = serializeTrace + "|open_output_failed";
			}
			return false;
		}

		out.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
		out.close();
		if (!out.good()) {
			if (outError != nullptr) {
				*outError = "write output snapshot file failed";
			}
			if (outTrace != nullptr) {
				*outTrace = serializeTrace + "|write_output_failed";
			}
			return false;
		}
	}
	catch (const std::exception& ex) {
		if (outError != nullptr) {
			*outError = std::string("write output snapshot exception: ") + ex.what();
		}
		if (outTrace != nullptr) {
			*outTrace = serializeTrace + "|write_output_exception";
		}
		return false;
	}

	if (outBytesWritten != nullptr) {
		*outBytesWritten = bytes.size();
	}
	if (outTrace != nullptr) {
		*outTrace = serializeTrace + "|write_ok path=" + outputPath;
	}
	return true;
#endif
}

} // namespace e571
