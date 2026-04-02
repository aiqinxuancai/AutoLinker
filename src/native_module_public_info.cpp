#include "native_module_public_info.hpp"
#include "EcModulePublicInfoReader.h"

#include <Windows.h>
#include <CommCtrl.h>
#include <afx.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <format>
#include <mutex>
#include <system_error>
#include <unordered_set>
#include <vector>

#include "IDEFacade.h"
#include "MemFind.h"

void OutputStringToELog(const std::string& szbuf);

namespace e571 {

namespace {

static_assert(sizeof(void*) == 4, "native_module_public_info requires a 32-bit build.");

using FnInitByteContainer = void*(__thiscall*)(void*);
using FnClearByteContainer = int(__thiscall*)(void*);
using FnPrepareModuleContext = void(__thiscall*)(void*, void*, int, int);
using FnLoadModulePublicInfo = int(__thiscall*)(void*, void*, CHAR*, CString*);
using FnHiddenModuleInfoDialogCtor = int(__thiscall*)(void*, CWnd*);
using FnCreateDialogIndirect = BOOL(__thiscall*)(void*, LPCDLGTEMPLATEA, void*, HINSTANCE);
using FnUpdateHiddenModuleInfoSelection = LRESULT(__thiscall*)(void*, int, unsigned int*);
using FnDestroyCString = void(__thiscall*)(void*);
using FnDestroyTreeCtrl = void(__thiscall*)(void*);
using FnDestroyEditCtrl = void(__thiscall*)(void*);
using FnDestroyDialogWindow = int(__thiscall*)(void*);
using FnDestroyImageList = void(__thiscall*)(void*);
using FnDestroyModuleInfoPayload = void(__thiscall*)(void*);
using FnDestroyDialogBase = void(__thiscall*)(void*);
using FnResolveModulePath = int(__stdcall*)(LPCSTR, CHAR*, LPCSTR);

constexpr std::uintptr_t kImageBase = 0x400000;
constexpr unsigned int kMaxRecordBodySize = 1024 * 1024 * 8;
constexpr ptrdiff_t kDialogObjectHwndOffset = 0x1C;
constexpr ptrdiff_t kDialogTemplateNameOffset = 0x40;
constexpr ptrdiff_t kDialogTemplateResourceOffset = 0x44;
constexpr ptrdiff_t kDialogTemplatePointerOffset = 0x48;
constexpr ptrdiff_t kHiddenModuleInfoPathOffset = 0xD4;
constexpr ptrdiff_t kHiddenModuleInfoRecordOffset = 0xD8;
constexpr ptrdiff_t kHiddenModuleInfoPayloadOffset = 0x1FC;
constexpr ptrdiff_t kHiddenModuleInfoImageListOffset = 0x53C;
constexpr size_t kHiddenModuleInfoStorageSize = 0x580;
constexpr size_t kHiddenModuleInfoRecordBytes = 0x123;
constexpr int kHiddenModuleInfoEditId = 1001;
constexpr int kHiddenModuleInfoTreeId = 1109;
constexpr DWORD kHiddenModuleInfoDialogTimeoutMs = 6000;

struct NativeByteContainer {
	std::uint32_t vftable = 0;
	std::uint32_t allocator = 0;
	unsigned char* data = nullptr;
	std::uint32_t capacityBytes = 0;
	std::uint32_t usedBytes = 0;
};

static_assert(offsetof(NativeByteContainer, data) == 8, "NativeByteContainer::data offset mismatch");
static_assert(offsetof(NativeByteContainer, usedBytes) == 16, "NativeByteContainer::usedBytes offset mismatch");

struct NativeModulePublicInfoAddresses {
	bool initialized = false;
	bool ok = false;
	bool hiddenDialogOk = false;
	std::uintptr_t moduleBase = 0;

	std::uintptr_t initByteContainer = 0;
	std::uintptr_t clearByteContainer = 0;
	std::uintptr_t prepareModuleContext = 0;
	std::uintptr_t loadModulePublicInfo = 0;
	std::uintptr_t moduleLoaderThis = 0;
	std::uintptr_t moduleLoaderArg = 0;
	std::uintptr_t recorderFlag = 0;
	std::uintptr_t recorderStateA0 = 0;
	std::uintptr_t recorderAuxContainer = 0;
	std::uintptr_t recorderBuffer = 0;
	std::uintptr_t recorderStateCC = 0;

	std::uintptr_t hiddenDialogCtor = 0;
	std::uintptr_t createDialogIndirect = 0;
	std::uintptr_t updateHiddenSelection = 0;
	std::uintptr_t destroyCString = 0;
	std::uintptr_t destroyTreeCtrl = 0;
	std::uintptr_t destroyEditCtrl = 0;
	std::uintptr_t destroyDialogWindow = 0;
	std::uintptr_t destroyImageList = 0;
	std::uintptr_t destroyModuleInfoPayload = 0;
	std::uintptr_t destroyDialogBase = 0;
	std::uintptr_t resolveModulePath = 0;
	std::uintptr_t importedModuleArrayGlobal = 0;
};

std::mutex g_nativeModuleInfoMutex;
NativeModulePublicInfoAddresses g_nativeModuleInfoAddresses;
const NativeModulePublicInfoAddresses& GetNativeModulePublicInfoAddresses(std::uintptr_t moduleBase);

struct OffsetString {
	size_t offset = 0;
	std::string text;
};

struct DecodedTypeInfo {
	std::string text;
	std::string flagsText;
	bool ok = false;
};

struct ParsedProcedureItem {
	size_t offset = 0;
	size_t bodyOffset = 0;
	size_t endOffset = 0;
	bool memberProcedure = false;
	bool isPublic = false;
	std::string name;
	std::string returnTypeText;
	std::string comment;
	std::vector<ModulePublicInfoParam> params;
};

struct HiddenModuleInfoDialogContext {
	HHOOK cbtHook = nullptr;
	HANDLE readyEvent = nullptr;
	HWND dialogHwnd = nullptr;
	std::string capturedText;
	std::string error;
};

std::mutex g_hiddenModuleInfoDialogMutex;
thread_local HiddenModuleInfoDialogContext* g_hiddenModuleInfoDialogContext = nullptr;

template <typename T>
T Bind(std::uintptr_t absoluteAddress);

template <typename T>
T* PtrAbsolute(std::uintptr_t moduleBase, std::uintptr_t normalizedAddress);

int ResolveImportedModuleIndex(const std::string& modulePath);
std::string GetFileNameOnly(const std::string& path);

constexpr unsigned char kTopLevelProcedureCore[] = { 0x01, 0x4A, 0x00, 0x01, 0x09 };
constexpr unsigned char kMemberProcedureCore[] = { 0x01, 0x03, 0x00, 0x01, 0x49 };

bool ReadFileBytes(const std::string& path, std::vector<unsigned char>& outBytes)
{
	outBytes.clear();
	std::ifstream in(path, std::ios::binary);
	if (!in.is_open()) {
		return false;
	}

	in.seekg(0, std::ios::end);
	const std::streamoff size = in.tellg();
	if (size <= 0) {
		return false;
	}
	in.seekg(0, std::ios::beg);

	outBytes.resize(static_cast<size_t>(size));
	in.read(reinterpret_cast<char*>(outBytes.data()), size);
	return in.good() || static_cast<size_t>(in.gcount()) == outBytes.size();
}

bool TryReadAnsiCString(std::uintptr_t address, size_t maxLen, std::string& outText)
{
	outText.clear();
	if (address == 0 || maxLen == 0) {
		return false;
	}

	__try {
		const char* text = reinterpret_cast<const char*>(address);
		for (size_t i = 0; i < maxLen; ++i) {
			const unsigned char ch = static_cast<unsigned char>(text[i]);
			if (ch == 0) {
				outText.assign(text, i);
				return !outText.empty();
			}
			if ((ch < 0x20 && ch != '\t') || ch == 0x7F) {
				return false;
			}
		}
	} __except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}

	return false;
}

bool TryReadMemoryCopy(std::uintptr_t address, void* outBuffer, size_t byteCount)
{
	if (address == 0 || outBuffer == nullptr || byteCount == 0) {
		return false;
	}

	__try {
		std::memcpy(outBuffer, reinterpret_cast<const void*>(address), byteCount);
		return true;
	} __except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

unsigned int ReadU32Le(const std::vector<unsigned char>& bytes, size_t offset)
{
	if (offset + 4 > bytes.size()) {
		return 0;
	}
	return
		static_cast<unsigned int>(bytes[offset]) |
		(static_cast<unsigned int>(bytes[offset + 1]) << 8) |
		(static_cast<unsigned int>(bytes[offset + 2]) << 16) |
		(static_cast<unsigned int>(bytes[offset + 3]) << 24);
}

bool MatchBytes(const std::vector<unsigned char>& bytes, size_t offset, const unsigned char* pattern, size_t patternSize)
{
	if (offset + patternSize > bytes.size()) {
		return false;
	}
	return std::memcmp(bytes.data() + offset, pattern, patternSize) == 0;
}

std::string TrimAsciiWhitespaceCopy(const std::string& value)
{
	size_t start = 0;
	while (start < value.size() && static_cast<unsigned char>(value[start]) <= 0x20) {
		++start;
	}

	size_t end = value.size();
	while (end > start && static_cast<unsigned char>(value[end - 1]) <= 0x20) {
		--end;
	}
	return value.substr(start, end - start);
}

bool ContainsUnexpectedControl(const std::string& text)
{
	for (unsigned char ch : text) {
		if (ch == '\r' || ch == '\n' || ch == '\t') {
			continue;
		}
		if (ch < 0x20 || ch == 0x7F) {
			return true;
		}
	}
	return false;
}

bool IsMetadataMarkerString(const std::string& text)
{
	if (text.empty()) {
		return true;
	}

	return
		text.find("CNWTEPRG") != std::string::npos ||
		text.rfind("_-@M<", 0) == 0 ||
		text.rfind("_-@S<", 0) == 0 ||
		text.rfind("krnln", 0) == 0;
}

bool IsLikelyReadableModuleString(const std::string& text)
{
	const std::string trimmed = TrimAsciiWhitespaceCopy(text);
	if (trimmed.empty() || ContainsUnexpectedControl(trimmed) || IsMetadataMarkerString(trimmed)) {
		return false;
	}

	size_t usefulCount = 0;
	for (unsigned char ch : trimmed) {
		if (std::isalnum(ch) != 0 || ch >= 0x80 || ch == '_' || ch == '@' || ch == '#' || ch == '.' || ch == ':' || ch == '\\' || ch == '/' || ch == '-' || ch == ' ') {
			++usefulCount;
		}
	}
	return usefulCount >= 2;
}

bool IsLikelyIdentifierName(const std::string& text)
{
	const std::string trimmed = TrimAsciiWhitespaceCopy(text);
	if (trimmed.empty() || trimmed.size() > 96 || ContainsUnexpectedControl(trimmed)) {
		return false;
	}
	if (trimmed.find('\n') != std::string::npos || trimmed.find('\r') != std::string::npos) {
		return false;
	}
	if (trimmed.find(' ') != std::string::npos || trimmed.find('：') != std::string::npos || trimmed.find('，') != std::string::npos || trimmed.find('。') != std::string::npos) {
		return false;
	}

	unsigned char first = static_cast<unsigned char>(trimmed.front());
	if (!(std::isalnum(first) != 0 || first >= 0x80 || first == '_' || first == '#')) {
		return false;
	}

	size_t usefulCount = 0;
	for (unsigned char ch : trimmed) {
		if (std::isalnum(ch) != 0 || ch >= 0x80 || ch == '_' || ch == '#') {
			++usefulCount;
			continue;
		}
		return false;
	}
	return usefulCount >= 1;
}

bool IsLikelyCommentText(const std::string& text)
{
	const std::string trimmed = TrimAsciiWhitespaceCopy(text);
	if (trimmed.empty() || ContainsUnexpectedControl(trimmed)) {
		return false;
	}
	if (trimmed.find('\n') != std::string::npos || trimmed.find('\r') != std::string::npos) {
		return true;
	}
	if (trimmed.find(' ') != std::string::npos || trimmed.find('！') != std::string::npos || trimmed.find('。') != std::string::npos || trimmed.find('，') != std::string::npos || trimmed.find('；') != std::string::npos || trimmed.find('“') != std::string::npos || trimmed.find('”') != std::string::npos) {
		return true;
	}
	return !IsLikelyIdentifierName(trimmed) && trimmed.size() >= 2;
}

std::string NormalizeLineBreaksCopy(std::string text)
{
	text.erase(std::remove(text.begin(), text.end(), '\r'), text.end());
	return text;
}

std::vector<std::string> SplitLinesCopy(const std::string& text)
{
	std::vector<std::string> out;
	size_t begin = 0;
	while (begin <= text.size()) {
		const size_t end = text.find('\n', begin);
		if (end == std::string::npos) {
			out.push_back(text.substr(begin));
			break;
		}
		out.push_back(text.substr(begin, end - begin));
		begin = end + 1;
	}
	return out;
}

std::vector<std::string> SplitCommaFields(const std::string& text)
{
	std::vector<std::string> out;
	size_t begin = 0;
	while (begin <= text.size()) {
		const size_t commaPos = text.find(',', begin);
		if (commaPos == std::string::npos) {
			out.push_back(TrimAsciiWhitespaceCopy(text.substr(begin)));
			break;
		}
		out.push_back(TrimAsciiWhitespaceCopy(text.substr(begin, commaPos - begin)));
		begin = commaPos + 1;
	}
	return out;
}

std::string JoinFieldsTail(const std::vector<std::string>& fields, size_t startIndex)
{
	std::string out;
	for (size_t i = startIndex; i < fields.size(); ++i) {
		if (fields[i].empty()) {
			continue;
		}
		if (!out.empty()) {
			out += ", ";
		}
		out += fields[i];
	}
	return out;
}

void PushRecordStrings(ModulePublicInfoRecord& record, const std::string& text)
{
	const std::string trimmed = TrimAsciiWhitespaceCopy(text);
	if (trimmed.empty()) {
		return;
	}
	if ((std::find)(record.extractedStrings.begin(), record.extractedStrings.end(), trimmed) == record.extractedStrings.end()) {
		record.extractedStrings.push_back(trimmed);
	}
}

bool ReadSizedString(const std::vector<unsigned char>& bytes, size_t offset, size_t size, std::string& outText)
{
	outText.clear();
	if (offset + size > bytes.size() || size == 0 || size > 0x4000) {
		return false;
	}

	outText.assign(reinterpret_cast<const char*>(bytes.data() + offset), size);
	outText = TrimAsciiWhitespaceCopy(outText);
	return IsLikelyReadableModuleString(outText);
}

std::vector<OffsetString> ExtractMeaningfulNullTerminatedStrings(
	const std::vector<unsigned char>& bytes,
	size_t startOffset,
	size_t endOffset)
{
	std::vector<OffsetString> out;
	if (startOffset >= endOffset || startOffset >= bytes.size()) {
		return out;
	}
	endOffset = (std::min)(endOffset, bytes.size());

	for (size_t i = startOffset; i < endOffset; ) {
		if (bytes[i] == 0) {
			++i;
			continue;
		}

		size_t j = i;
		while (j < endOffset && bytes[j] != 0) {
			++j;
		}
		if (j <= i) {
			++i;
			continue;
		}

		std::string text(reinterpret_cast<const char*>(bytes.data() + i), j - i);
		text = TrimAsciiWhitespaceCopy(text);
		if (IsLikelyReadableModuleString(text)) {
			out.push_back({ i, std::move(text) });
		}
		i = (j < endOffset) ? (j + 1) : j;
	}
	return out;
}

size_t FindNextProcedureHeaderOffset(const std::vector<unsigned char>& bytes, size_t startOffset)
{
	for (size_t i = startOffset; i + sizeof(kTopLevelProcedureCore) <= bytes.size(); ++i) {
		if (MatchBytes(bytes, i, kTopLevelProcedureCore, sizeof(kTopLevelProcedureCore)) ||
			MatchBytes(bytes, i, kMemberProcedureCore, sizeof(kMemberProcedureCore))) {
			return i;
		}
	}
	return bytes.size();
}

DecodedTypeInfo DecodeTypeCode(unsigned int typeCode, const std::string& currentAssemblyName);

std::vector<std::string> SplitModuleHeaderLines(const std::string& text)
{
	std::vector<std::string> lines;
	std::string current;
	for (char ch : text) {
		if (ch == '\r') {
			continue;
		}
		if (ch == '\n') {
			lines.push_back(TrimAsciiWhitespaceCopy(current));
			current.clear();
			continue;
		}
		current.push_back(ch);
	}
	if (!current.empty()) {
		lines.push_back(TrimAsciiWhitespaceCopy(current));
	}

	lines.erase(
		std::remove_if(
			lines.begin(),
			lines.end(),
			[](const std::string& line) {
				return line.empty();
			}),
		lines.end());
	return lines;
}

void TryParseAssemblyHeaderText(
	const std::vector<OffsetString>& allStrings,
	std::string& ioAssemblyName,
	std::string& ioAssemblyComment)
{
	for (const auto& entry : allStrings) {
		if (entry.text.find('@') == std::string::npos ||
			(entry.text.find('\n') == std::string::npos && entry.text.find('\r') == std::string::npos)) {
			continue;
		}

		const auto lines = SplitModuleHeaderLines(entry.text);
		if (lines.size() < 2 || lines[1].empty() || lines[1].front() != '@') {
			continue;
		}

		if (ioAssemblyName.empty() && IsLikelyIdentifierName(lines.front())) {
			ioAssemblyName = lines.front();
		}
		if (ioAssemblyComment.empty() && lines.size() >= 3 && IsLikelyReadableModuleString(lines[2])) {
			ioAssemblyComment = lines[2];
		}
		if (!ioAssemblyName.empty()) {
			return;
		}
	}
}

size_t FindProcedureMetadataEndOffset(
	const std::vector<unsigned char>& bytes,
	size_t startOffset,
	size_t hardEndOffset)
{
	if (startOffset >= bytes.size() || startOffset >= hardEndOffset) {
		return startOffset;
	}

	static constexpr unsigned char kBodyMarkers[][2] = {
		{ 0x36, 0x1D },
		{ 0x36, 0x1F },
		{ 0x36, 0x21 },
		{ 0x37, 0x1D },
		{ 0x37, 0x1F },
		{ 0x37, 0x21 },
	};

	const size_t scanStart = (std::min)(startOffset + 24, hardEndOffset);
	size_t best = (std::min)(hardEndOffset, startOffset + 1024);
	for (size_t i = scanStart; i + 2 <= best; ++i) {
		for (const auto& marker : kBodyMarkers) {
			if (bytes[i] == marker[0] && bytes[i + 1] == marker[1]) {
				return i;
			}
		}
	}
	return best;
}

bool TryFindNearbyTypeInfo(
	const std::vector<unsigned char>& bytes,
	size_t stringOffset,
	size_t minOffset,
	const std::string& currentAssemblyName,
	size_t maxDistance,
	DecodedTypeInfo& outTypeInfo)
{
	outTypeInfo = {};
	if (stringOffset <= minOffset || maxDistance < 4) {
		return false;
	}

	const size_t searchBegin = (stringOffset > maxDistance) ? stringOffset - maxDistance : minOffset;
	const size_t begin = (std::max)(minOffset, searchBegin);
	for (size_t candidate = stringOffset; candidate-- > begin; ) {
		if (candidate + 4 > stringOffset) {
			continue;
		}

		DecodedTypeInfo info = DecodeTypeCode(ReadU32Le(bytes, candidate), currentAssemblyName);
		if (info.ok) {
			outTypeInfo = std::move(info);
			return true;
		}
	}
	return false;
}

DecodedTypeInfo DecodeTypeCode(unsigned int typeCode, const std::string& currentAssemblyName)
{
	DecodedTypeInfo out;
	if (typeCode == 0) {
		return out;
	}

	if ((typeCode & 0x80000000u) != 0) {
		switch (typeCode & 0xFFu) {
		case 1: out.text = "整数型"; break;
		case 2: out.text = "逻辑型"; break;
		case 4: out.text = "文本型"; break;
		case 5: out.text = "字节集"; break;
		default: break;
		}
		if (!out.text.empty()) {
			const unsigned int modifiers = (typeCode >> 8) & 0xFFFFu;
			if ((modifiers & 0x2u) != 0) {
				out.flagsText = "可空";
			}
			out.ok = true;
		}
		return out;
	}

	if ((typeCode & 0xFFFFFF00u) == 0x49010000u) {
		out.text = currentAssemblyName.empty() ? "自定义类型" : currentAssemblyName;
		out.ok = true;
		return out;
	}

	return out;
}

std::vector<ModulePublicInfoParam> ParseProcedureParamsHeuristic(
	const std::vector<unsigned char>& bytes,
	size_t startOffset,
	size_t endOffset,
	const std::string& currentAssemblyName,
	const std::string& ownerName,
	const std::string& ownerComment)
{
	std::vector<ModulePublicInfoParam> out;
	const size_t scanEndOffset = FindProcedureMetadataEndOffset(bytes, startOffset, endOffset);
	const auto strings = ExtractMeaningfulNullTerminatedStrings(bytes, startOffset, scanEndOffset);
	std::unordered_set<std::string> seenNames;

	for (size_t i = 0; i < strings.size(); ++i) {
		const auto& item = strings[i];
		if (item.text == ownerName || (!ownerComment.empty() && item.text == ownerComment)) {
			continue;
		}
		if (!IsLikelyIdentifierName(item.text) || IsMetadataMarkerString(item.text)) {
			continue;
		}

		DecodedTypeInfo typeInfo;
		if (!TryFindNearbyTypeInfo(
				bytes,
				item.offset,
				startOffset,
				currentAssemblyName,
				16,
				typeInfo) ||
			!seenNames.insert(item.text).second) {
			continue;
		}

		ModulePublicInfoParam param;
		param.name = item.text;
		param.typeText = typeInfo.text;
		param.flagsText = typeInfo.flagsText;

		if (i + 1 < strings.size()) {
			const auto& next = strings[i + 1];
			DecodedTypeInfo nextTypeInfo;
			if (next.text != ownerName &&
				next.text != ownerComment &&
				!TryFindNearbyTypeInfo(
					bytes,
					next.offset,
					startOffset,
					currentAssemblyName,
					16,
					nextTypeInfo) &&
				!IsMetadataMarkerString(next.text) &&
				(IsLikelyCommentText(next.text) || next.text.size() == 1)) {
				param.comment = next.text;
				++i;
			}
		}
		out.push_back(std::move(param));
	}

	return out;
}

std::string BuildProcedureSignature(const ParsedProcedureItem& item)
{
	std::string line = ".子程序 " + item.name;
	if (!item.returnTypeText.empty()) {
		line += ", " + item.returnTypeText;
	}
	if (item.isPublic) {
		line += ", 公开";
	}
	if (!item.comment.empty()) {
		line += ", " + item.comment;
	}
	return line;
}

bool TryParseProcedureItemAt(
	const std::vector<unsigned char>& bytes,
	size_t offset,
	const std::string& currentAssemblyName,
	ParsedProcedureItem& outItem)
{
	bool memberProcedure = false;
	if (MatchBytes(bytes, offset, kTopLevelProcedureCore, sizeof(kTopLevelProcedureCore))) {
		memberProcedure = false;
	}
	else if (MatchBytes(bytes, offset, kMemberProcedureCore, sizeof(kMemberProcedureCore))) {
		memberProcedure = true;
	}
	else {
		return false;
	}

	if (offset + 17 > bytes.size()) {
		return false;
	}

	const unsigned int visibilityFlags = ReadU32Le(bytes, offset + 5);
	const unsigned int returnTypeCode = ReadU32Le(bytes, offset + 9);
	const unsigned int nameSize = ReadU32Le(bytes, offset + 13);
	if (nameSize == 0 || nameSize > 256) {
		return false;
	}

	std::string name;
	if (!ReadSizedString(bytes, offset + 17, nameSize, name) || !IsLikelyIdentifierName(name)) {
		return false;
	}

	const size_t commentLengthOffset = offset + 17 + nameSize;
	if (commentLengthOffset + 4 > bytes.size()) {
		return false;
	}
	const unsigned int commentSize = ReadU32Le(bytes, commentLengthOffset);
	if (commentSize > 0x4000 || commentLengthOffset + 4 + commentSize > bytes.size()) {
		return false;
	}

	std::string comment;
	if (commentSize > 0) {
		std::string rawComment;
		if (ReadSizedString(bytes, commentLengthOffset + 4, commentSize, rawComment) && IsLikelyCommentText(rawComment)) {
			comment = rawComment;
		}
	}

	const size_t bodyOffset = commentLengthOffset + 4 + commentSize;
	const size_t nextOffset = FindNextProcedureHeaderOffset(bytes, bodyOffset);

	outItem = {};
	outItem.offset = offset;
	outItem.bodyOffset = bodyOffset;
	outItem.endOffset = nextOffset;
	outItem.memberProcedure = memberProcedure;
	outItem.isPublic = (visibilityFlags & 0x8u) != 0;
	outItem.name = std::move(name);
	outItem.comment = std::move(comment);

	const DecodedTypeInfo returnType = DecodeTypeCode(returnTypeCode, currentAssemblyName);
	outItem.returnTypeText = returnType.text;
	outItem.params = ParseProcedureParamsHeuristic(
		bytes,
		bodyOffset,
		nextOffset,
		currentAssemblyName,
		outItem.name,
		outItem.comment);
	return true;
}

std::string JoinLines(const std::vector<std::string>& lines)
{
	std::string text;
	for (size_t i = 0; i < lines.size(); ++i) {
		text += lines[i];
		if (i + 1 < lines.size()) {
			text += "\r\n";
		}
	}
	return text;
}

bool IsLikelyConstantName(const std::string& text)
{
	if (!IsLikelyIdentifierName(text) || text.size() > 96) {
		return false;
	}

	bool hasUpper = false;
	bool hasLower = false;
	bool hasUnderscore = false;
	for (unsigned char ch : text) {
		hasUpper = hasUpper || (ch >= 'A' && ch <= 'Z');
		hasLower = hasLower || (ch >= 'a' && ch <= 'z');
		hasUnderscore = hasUnderscore || (ch == '_');
	}
	return hasUnderscore || (hasUpper && !hasLower) || (hasUpper && hasLower);
}

bool ParseModulePublicInfoFromEcFile(
	const std::string& modulePath,
	ModulePublicInfoDump& outDump,
	std::string* outError)
{
	EcModulePublicInfoReader reader;
	return reader.Load(modulePath, outDump, outError);
}

LRESULT CALLBACK HiddenModuleInfoDialogCbtProc(int code, WPARAM wParam, LPARAM lParam)
{
	HiddenModuleInfoDialogContext* ctx = g_hiddenModuleInfoDialogContext;
	if (ctx != nullptr && (code == HCBT_ACTIVATE || code == HCBT_CREATEWND)) {
		HWND hwnd = reinterpret_cast<HWND>(wParam);
		char className[64] = {};
		if (hwnd != nullptr &&
			GetClassNameA(hwnd, className, static_cast<int>(sizeof(className))) > 0 &&
			_stricmp(className, "#32770") == 0) {
			if (code == HCBT_CREATEWND) {
				auto* createInfo = reinterpret_cast<CBT_CREATEWNDA*>(lParam);
				if (createInfo != nullptr && createInfo->lpcs != nullptr) {
					createInfo->lpcs->style &= ~WS_VISIBLE;
					createInfo->lpcs->x = -32000;
					createInfo->lpcs->y = -32000;
				}
			}
			ctx->dialogHwnd = hwnd;
			ShowWindow(hwnd, SW_HIDE);
			SetWindowPos(
				hwnd,
				nullptr,
				-32000,
				-32000,
				0,
				0,
				SWP_NOSIZE | SWP_NOZORDER | SWP_HIDEWINDOW | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
			if (ctx->readyEvent != nullptr) {
				SetEvent(ctx->readyEvent);
			}
		}
	}
	return CallNextHookEx(ctx == nullptr ? nullptr : ctx->cbtHook, code, wParam, lParam);
}

std::string ReadWindowTextAString(HWND hWnd)
{
	if (hWnd == nullptr || !IsWindow(hWnd)) {
		return std::string();
	}

	const int length = GetWindowTextLengthA(hWnd);
	if (length <= 0) {
		return std::string();
	}

	std::string text(static_cast<size_t>(length) + 1, '\0');
	GetWindowTextA(hWnd, text.data(), length + 1);
	text.resize(static_cast<size_t>(length));
	return text;
}

LPCDLGTEMPLATEA ResolveDialogTemplateFromObject(void* dialogObject)
{
	if (dialogObject == nullptr) {
		return nullptr;
	}

	auto* const bytes = reinterpret_cast<unsigned char*>(dialogObject);
	const auto resourceName = *reinterpret_cast<LPCSTR*>(bytes + kDialogTemplateNameOffset);
	auto resourceHandle = *reinterpret_cast<HGLOBAL*>(bytes + kDialogTemplateResourceOffset);
	auto templatePointer = *reinterpret_cast<LPCDLGTEMPLATEA*>(bytes + kDialogTemplatePointerOffset);

	HMODULE resourceModule = GetModuleHandleA(nullptr);
	if (resourceModule == nullptr) {
		resourceModule = reinterpret_cast<HMODULE>(GetModuleHandleW(nullptr));
	}

	if (resourceName != nullptr && resourceModule != nullptr) {
		HRSRC dialogResource = FindResourceA(resourceModule, resourceName, RT_DIALOG);
		if (dialogResource != nullptr) {
			resourceHandle = LoadResource(resourceModule, dialogResource);
		}
	}

	if (resourceHandle != nullptr) {
		templatePointer = reinterpret_cast<LPCDLGTEMPLATEA>(LockResource(resourceHandle));
	}
	return templatePointer;
}

bool CaptureHiddenModuleInfoTextFromDialog(
	void* dialogObject,
	HWND dialogHwnd,
	const NativeModulePublicInfoAddresses& addrs,
	std::string& outText,
	std::string& outError)
{
	outText.clear();
	outError.clear();

	if (dialogObject == nullptr || dialogHwnd == nullptr || !IsWindow(dialogHwnd)) {
		outError = "hidden_modeless_dialog_invalid";
		return false;
	}

	HWND editHwnd = GetDlgItem(dialogHwnd, kHiddenModuleInfoEditId);
	if (editHwnd == nullptr || !IsWindow(editHwnd)) {
		outError = "hidden_modeless_dialog_edit_missing";
		return false;
	}

	const auto updateSelection = Bind<FnUpdateHiddenModuleInfoSelection>(addrs.updateHiddenSelection);
	unsigned int updateCookie = 0;
	updateSelection(dialogObject, 0, &updateCookie);

	std::string bestText = ReadWindowTextAString(editHwnd);
	if (bestText.empty()) {
		updateSelection(dialogObject, 0, &updateCookie);
		bestText = ReadWindowTextAString(editHwnd);
	}

	if (bestText.empty()) {
		outError = "hidden_modeless_dialog_text_empty";
		return false;
	}

	outText = std::move(bestText);
	return true;
}

DWORD WINAPI HiddenModuleInfoDialogWorkerProc(LPVOID parameter)
{
	auto* ctx = reinterpret_cast<HiddenModuleInfoDialogContext*>(parameter);
	if (ctx == nullptr) {
		return 0;
	}

	if (ctx->readyEvent == nullptr ||
		WaitForSingleObject(ctx->readyEvent, kHiddenModuleInfoDialogTimeoutMs) != WAIT_OBJECT_0) {
		ctx->error = "hidden_dialog_not_created";
		return 0;
	}

	const DWORD beginTick = GetTickCount();
	HWND dialogHwnd = nullptr;
	HWND treeHwnd = nullptr;
	HWND editHwnd = nullptr;
	while (GetTickCount() - beginTick < kHiddenModuleInfoDialogTimeoutMs) {
		dialogHwnd = ctx->dialogHwnd;
		if (dialogHwnd != nullptr && IsWindow(dialogHwnd)) {
			treeHwnd = GetDlgItem(dialogHwnd, kHiddenModuleInfoTreeId);
			editHwnd = GetDlgItem(dialogHwnd, kHiddenModuleInfoEditId);
			if (editHwnd != nullptr) {
				break;
			}
		}
		Sleep(20);
	}

	if (dialogHwnd == nullptr || editHwnd == nullptr) {
		ctx->error = "hidden_dialog_controls_not_ready";
		if (dialogHwnd != nullptr && IsWindow(dialogHwnd)) {
			PostMessageA(dialogHwnd, WM_CLOSE, 0, 0);
		}
		return 0;
	}

	std::string bestText;
	size_t stableRounds = 0;
	for (int i = 0; i < 150; ++i) {
		std::string currentText = ReadWindowTextAString(editHwnd);
		if (!currentText.empty()) {
			if (currentText == bestText) {
				++stableRounds;
				if (stableRounds >= 3) {
					break;
				}
			}
			else {
				bestText = std::move(currentText);
				stableRounds = 0;
			}
		}
		Sleep(20);
	}

	ctx->capturedText = std::move(bestText);
	if (ctx->capturedText.empty()) {
		ctx->error = "hidden_dialog_text_empty";
	}

	if (IsWindow(dialogHwnd)) {
		PostMessageA(dialogHwnd, WM_COMMAND, IDOK, 0);
		PostMessageA(dialogHwnd, WM_CLOSE, 0, 0);
	}
	return 0;
}

#if 0
#if 0
void ParseModulePublicInfoFromFormattedText(
	const std::string& formattedText,
	ModulePublicInfoDump& outDump)
{
	outDump.formattedText = formattedText;
	const std::string normalized = NormalizeLineBreaksCopy(formattedText);
	const auto lines = SplitLinesCopy(normalized);

	bool inCodeSection = false;
	ModulePublicInfoRecord currentRecord;
	bool hasCurrentRecord = false;
	const auto flushRecord = [&]() {
		if (!hasCurrentRecord) {
			return;
		}
		if (currentRecord.signatureText.empty() && !currentRecord.kind.empty() && !currentRecord.name.empty()) {
			currentRecord.signatureText = currentRecord.kind + " " + currentRecord.name;
		}
		outDump.records.push_back(std::move(currentRecord));
		currentRecord = {};
		hasCurrentRecord = false;
	};

	for (size_t i = 0; i < lines.size(); ++i) {
		const std::string line = TrimAsciiWhitespaceCopy(lines[i]);
		if (line.empty()) {
			continue;
		}

		if (!inCodeSection) {
			if (line.rfind("模块名称：", 0) == 0 || line.rfind("模块名称:", 0) == 0) {
				const size_t sep = line.find_first_of(":：");
				if (sep != std::string::npos) {
					outDump.moduleName = TrimAsciiWhitespaceCopy(line.substr(sep + 1));
				}
				continue;
			}
			if (line.rfind("版本：", 0) == 0 || line.rfind("版本:", 0) == 0) {
				const size_t sep = line.find_first_of(":：");
				if (sep != std::string::npos) {
					outDump.versionText = TrimAsciiWhitespaceCopy(line.substr(sep + 1));
				}
				continue;
			}
			if (line == "------------------------------" || line == ".版本 2" || line == ".版本") {
				inCodeSection = true;
				continue;
			}
			if (outDump.assemblyName.empty() &&
				line.find("@备注") == std::string::npos &&
				line.find("模块名称") == std::string::npos &&
				line.find("版本") == std::string::npos) {
				outDump.assemblyName = line;
				continue;
			}
			if (line.rfind("@备注", 0) == 0 && i + 1 < lines.size()) {
				outDump.assemblyComment = TrimAsciiWhitespaceCopy(lines[i + 1]);
				continue;
			}
			continue;
		}

		if (line.rfind(".参数 ", 0) == 0 && hasCurrentRecord) {
			const auto fields = SplitCommaFields(line.substr(4));
			if (!fields.empty()) {
				ModulePublicInfoParam param;
				param.name = fields[0];
				if (fields.size() > 1) {
					param.typeText = fields[1];
				}
				if (fields.size() > 2) {
					param.flagsText = fields[2];
				}
				if (fields.size() > 3) {
					param.comment = JoinFieldsTail(fields, 3);
				}
				currentRecord.params.push_back(std::move(param));
			}
			PushRecordStrings(currentRecord, line);
			continue;
		}

		if (line.front() != '.') {
			continue;
		}

		flushRecord();
		hasCurrentRecord = true;
		currentRecord.signatureText = line;
		PushRecordStrings(currentRecord, line);

		if (line.rfind(".程序集 ", 0) == 0) {
			currentRecord.tag = 250;
			currentRecord.kind = "程序集";
			const auto fields = SplitCommaFields(line.substr(5));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else if (line.rfind(".子程序 ", 0) == 0) {
			currentRecord.tag = 301;
			currentRecord.kind = "子程序";
			const auto fields = SplitCommaFields(line.substr(5));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else if (line.rfind(".DLL命令 ", 0) == 0) {
			currentRecord.tag = 305;
			currentRecord.kind = "DLL命令";
			const auto fields = SplitCommaFields(line.substr(6));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else if (line.rfind(".常量 ", 0) == 0) {
			currentRecord.tag = 307;
			currentRecord.kind = "常量";
			const auto fields = SplitCommaFields(line.substr(4));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else {
			currentRecord.tag = 0;
			currentRecord.kind = "声明";
			currentRecord.name = line;
		}

		PushRecordStrings(currentRecord, currentRecord.name);
		PushRecordStrings(currentRecord, currentRecord.comment);
	}

	flushRecord();
}
#endif

void ParseModulePublicInfoFromFormattedText(
	const std::string& formattedText,
	ModulePublicInfoDump& outDump)
{
	outDump.formattedText = formattedText;
	const std::string normalized = NormalizeLineBreaksCopy(formattedText);
	const auto lines = SplitLinesCopy(normalized);

	const auto startsWith = [](const std::string& text, const char* prefix) {
		return text.rfind(prefix, 0) == 0;
	};
	const auto valueAfterColon = [](const std::string& line) {
		size_t pos = line.find("：");
		if (pos != std::string::npos) {
			return TrimAsciiWhitespaceCopy(line.substr(pos + std::strlen("：")));
		}
		pos = line.find(':');
		if (pos != std::string::npos) {
			return TrimAsciiWhitespaceCopy(line.substr(pos + 1));
		}
		return std::string();
	};

	bool inCodeSection = false;
	ModulePublicInfoRecord currentRecord;
	bool hasCurrentRecord = false;
	const auto flushRecord = [&]() {
		if (!hasCurrentRecord) {
			return;
		}
		if (currentRecord.signatureText.empty() && !currentRecord.kind.empty() && !currentRecord.name.empty()) {
			currentRecord.signatureText = currentRecord.kind + " " + currentRecord.name;
		}
		outDump.records.push_back(std::move(currentRecord));
		currentRecord = {};
		hasCurrentRecord = false;
	};

	for (size_t i = 0; i < lines.size(); ++i) {
		const std::string line = TrimAsciiWhitespaceCopy(lines[i]);
		if (line.empty()) {
			continue;
		}

		if (!inCodeSection) {
			if (startsWith(line, "模块名称：") || startsWith(line, "模块名称:")) {
				outDump.moduleName = valueAfterColon(line);
				continue;
			}
			if (startsWith(line, "版本：") || startsWith(line, "版本:")) {
				outDump.versionText = valueAfterColon(line);
				continue;
			}
			if (line == "------------------------------") {
				continue;
			}
			if (startsWith(line, ".版本")) {
				inCodeSection = true;
				continue;
			}
			if (startsWith(line, "@备注")) {
				for (size_t j = i + 1; j < lines.size(); ++j) {
					const std::string nextLine = TrimAsciiWhitespaceCopy(lines[j]);
					if (!nextLine.empty()) {
						outDump.assemblyComment = nextLine;
						break;
					}
				}
				continue;
			}
			if (outDump.assemblyName.empty()) {
				outDump.assemblyName = line;
			}
			continue;
		}

		if (startsWith(line, ".参数 ") && hasCurrentRecord) {
			const auto fields = SplitCommaFields(line.substr(std::strlen(".参数 ")));
			if (!fields.empty()) {
				ModulePublicInfoParam param;
				param.name = fields[0];
				if (fields.size() > 1) {
					param.typeText = fields[1];
				}
				if (fields.size() > 2) {
					param.flagsText = fields[2];
				}
				if (fields.size() > 3) {
					param.comment = JoinFieldsTail(fields, 3);
				}
				currentRecord.params.push_back(std::move(param));
			}
			PushRecordStrings(currentRecord, line);
			continue;
		}

		if (line.front() != '.') {
			continue;
		}

		flushRecord();
		hasCurrentRecord = true;
		currentRecord.signatureText = line;
		PushRecordStrings(currentRecord, line);

		if (startsWith(line, ".程序集 ")) {
			currentRecord.tag = 250;
			currentRecord.kind = "程序集";
			const auto fields = SplitCommaFields(line.substr(std::strlen(".程序集 ")));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else if (startsWith(line, ".子程序 ")) {
			currentRecord.tag = 301;
			currentRecord.kind = "子程序";
			const auto fields = SplitCommaFields(line.substr(std::strlen(".子程序 ")));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else if (startsWith(line, ".DLL命令 ")) {
			currentRecord.tag = 305;
			currentRecord.kind = "DLL命令";
			const auto fields = SplitCommaFields(line.substr(std::strlen(".DLL命令 ")));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else if (startsWith(line, ".常量 ")) {
			currentRecord.tag = 307;
			currentRecord.kind = "常量";
			const auto fields = SplitCommaFields(line.substr(std::strlen(".常量 ")));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else {
			currentRecord.tag = 0;
			currentRecord.kind = "声明";
			currentRecord.name = line;
		}

		PushRecordStrings(currentRecord, currentRecord.name);
		PushRecordStrings(currentRecord, currentRecord.comment);
	}

	flushRecord();
}
#endif

void ParseModulePublicInfoFromFormattedText(
	const std::string& formattedText,
	ModulePublicInfoDump& outDump)
{
	outDump.formattedText = formattedText;
	const std::string normalized = NormalizeLineBreaksCopy(formattedText);
	const auto lines = SplitLinesCopy(normalized);

	const auto startsWith = [](const std::string& text, const char* prefix) {
		return text.rfind(prefix, 0) == 0;
	};
	const auto valueAfterSep = [](const std::string& line, const char* fullWidthSep) {
		size_t pos = line.find(fullWidthSep);
		if (pos != std::string::npos) {
			return TrimAsciiWhitespaceCopy(line.substr(pos + std::strlen(fullWidthSep)));
		}
		pos = line.find(':');
		if (pos != std::string::npos) {
			return TrimAsciiWhitespaceCopy(line.substr(pos + 1));
		}
		return std::string();
	};

	constexpr const char* kPrefixModuleName = "\xC4\xA3\xBF\xE9\xC3\xFB\xB3\xC6\xA3\xBA";
	constexpr const char* kPrefixVersion = "\xB0\xE6\xB1\xBE\xA3\xBA";
	constexpr const char* kPrefixNote = "\x40\xB1\xB8\xD7\xA2";
	constexpr const char* kPrefixCodeVersion = "\x2E\xB0\xE6\xB1\xBE";
	constexpr const char* kPrefixParam = "\x2E\xB2\xCE\xCA\xFD\x20";
	constexpr const char* kPrefixAssembly = "\x2E\xB3\xCC\xD0\xF2\xBC\xAF\x20";
	constexpr const char* kPrefixSub = "\x2E\xD7\xD3\xB3\xCC\xD0\xF2\x20";
	constexpr const char* kPrefixDll = "\x2E\x44\x4C\x4C\xC3\xFC\xC1\xEE\x20";
	constexpr const char* kPrefixConst = "\x2E\xB3\xA3\xC1\xBF\x20";

	bool inCodeSection = false;
	ModulePublicInfoRecord currentRecord;
	bool hasCurrentRecord = false;
	const auto flushRecord = [&]() {
		if (!hasCurrentRecord) {
			return;
		}
		if (currentRecord.signatureText.empty() && !currentRecord.kind.empty() && !currentRecord.name.empty()) {
			currentRecord.signatureText = currentRecord.kind + " " + currentRecord.name;
		}
		outDump.records.push_back(std::move(currentRecord));
		currentRecord = {};
		hasCurrentRecord = false;
	};

	for (size_t i = 0; i < lines.size(); ++i) {
		const std::string line = TrimAsciiWhitespaceCopy(lines[i]);
		if (line.empty()) {
			continue;
		}

		if (!inCodeSection) {
			if (startsWith(line, kPrefixModuleName) || startsWith(line, "module_name:")) {
				outDump.moduleName = valueAfterSep(line, kPrefixModuleName);
				continue;
			}
			if (startsWith(line, kPrefixVersion) || startsWith(line, "version:")) {
				outDump.versionText = valueAfterSep(line, kPrefixVersion);
				continue;
			}
			if (line == "------------------------------") {
				continue;
			}
			if (startsWith(line, kPrefixCodeVersion) || startsWith(line, ".version")) {
				inCodeSection = true;
				continue;
			}
			if (startsWith(line, kPrefixNote) || startsWith(line, "@note")) {
				for (size_t j = i + 1; j < lines.size(); ++j) {
					const std::string nextLine = TrimAsciiWhitespaceCopy(lines[j]);
					if (!nextLine.empty()) {
						outDump.assemblyComment = nextLine;
						break;
					}
				}
				continue;
			}
			if (outDump.assemblyName.empty()) {
				outDump.assemblyName = line;
			}
			continue;
		}

		if (startsWith(line, kPrefixParam) && hasCurrentRecord) {
			const auto fields = SplitCommaFields(line.substr(std::strlen(kPrefixParam)));
			if (!fields.empty()) {
				ModulePublicInfoParam param;
				param.name = fields[0];
				if (fields.size() > 1) {
					param.typeText = fields[1];
				}
				if (fields.size() > 2) {
					param.flagsText = fields[2];
				}
				if (fields.size() > 3) {
					param.comment = JoinFieldsTail(fields, 3);
				}
				currentRecord.params.push_back(std::move(param));
			}
			PushRecordStrings(currentRecord, line);
			continue;
		}

		if (line.front() != '.') {
			continue;
		}

		flushRecord();
		hasCurrentRecord = true;
		currentRecord.signatureText = line;
		PushRecordStrings(currentRecord, line);

		if (startsWith(line, kPrefixAssembly)) {
			currentRecord.tag = 250;
			currentRecord.kind = "assembly";
			const auto fields = SplitCommaFields(line.substr(std::strlen(kPrefixAssembly)));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else if (startsWith(line, kPrefixSub)) {
			currentRecord.tag = 301;
			currentRecord.kind = "sub";
			const auto fields = SplitCommaFields(line.substr(std::strlen(kPrefixSub)));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else if (startsWith(line, kPrefixDll)) {
			currentRecord.tag = 305;
			currentRecord.kind = "dll_command";
			const auto fields = SplitCommaFields(line.substr(std::strlen(kPrefixDll)));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else if (startsWith(line, kPrefixConst)) {
			currentRecord.tag = 307;
			currentRecord.kind = "const";
			const auto fields = SplitCommaFields(line.substr(std::strlen(kPrefixConst)));
			if (!fields.empty()) {
				currentRecord.name = fields[0];
			}
			if (fields.size() > 1) {
				currentRecord.typeText = fields[1];
			}
			if (fields.size() > 2) {
				currentRecord.flagsText = fields[2];
			}
			if (fields.size() > 3) {
				currentRecord.comment = JoinFieldsTail(fields, 3);
			}
		}
		else {
			currentRecord.tag = 0;
			currentRecord.kind = "declaration";
			currentRecord.name = line;
		}

		PushRecordStrings(currentRecord, currentRecord.name);
		PushRecordStrings(currentRecord, currentRecord.comment);
	}

	flushRecord();
}

bool LoadModulePublicInfoDumpHiddenDialog(
	const std::string& modulePath,
	std::uintptr_t moduleBase,
	ModulePublicInfoDump& outDump,
	std::string* outError)
{
	if (outError != nullptr) {
		outError->clear();
	}

	const auto& addrs = GetNativeModulePublicInfoAddresses(moduleBase);
	if (!addrs.hiddenDialogOk) {
		if (outError != nullptr) {
			*outError = "hidden_dialog_resolve_addresses_failed";
		}
		return false;
	}

	const int moduleIndex = ResolveImportedModuleIndex(modulePath);
	if (moduleIndex < 0) {
		if (outError != nullptr) {
			*outError = "imported_module_index_not_found";
		}
		return false;
	}

	auto* importedModuleArrayGlobal = PtrAbsolute<std::uintptr_t>(moduleBase, addrs.importedModuleArrayGlobal);
	if (importedModuleArrayGlobal == nullptr) {
		if (outError != nullptr) {
			*outError = "hidden_dialog_import_array_global_invalid";
		}
		return false;
	}

	std::uintptr_t importedModuleArrayAddress = 0;
	if (!TryReadMemoryCopy(
			reinterpret_cast<std::uintptr_t>(importedModuleArrayGlobal),
			&importedModuleArrayAddress,
			sizeof(importedModuleArrayAddress)) ||
		importedModuleArrayAddress == 0) {
		if (outError != nullptr) {
			*outError = std::format(
				"hidden_dialog_import_array_ptr_invalid global=0x{:X}",
				reinterpret_cast<std::uintptr_t>(importedModuleArrayGlobal));
		}
		return false;
	}

	std::uintptr_t moduleRecordAddress = 0;
	if (!TryReadMemoryCopy(
			importedModuleArrayAddress + static_cast<std::uintptr_t>(moduleIndex) * sizeof(std::uintptr_t),
			&moduleRecordAddress,
			sizeof(moduleRecordAddress))) {
		if (outError != nullptr) {
			*outError = std::format(
				"hidden_dialog_module_record_ptr_unreadable index={} array=0x{:X}",
				moduleIndex,
				importedModuleArrayAddress);
		}
		return false;
	}

	if (moduleRecordAddress == 0) {
		if (outError != nullptr) {
			*outError = "hidden_dialog_module_record_invalid";
		}
		return false;
	}

	std::uintptr_t moduleRecordNameAddress = 0;
	if (!TryReadMemoryCopy(
			moduleRecordAddress + 8,
			&moduleRecordNameAddress,
			sizeof(moduleRecordNameAddress))) {
		if (outError != nullptr) {
			*outError = std::format(
				"hidden_dialog_module_record_name_ptr_unreadable index={} record=0x{:X}",
				moduleIndex,
				moduleRecordAddress);
		}
		return false;
	}

	std::string moduleRecordName;
	if (!TryReadAnsiCString(moduleRecordNameAddress, 1024, moduleRecordName)) {
		if (outError != nullptr) {
			*outError = std::format(
				"hidden_dialog_module_name_invalid index={} record=0x{:X} name_ptr=0x{:X}",
				moduleIndex,
				moduleRecordAddress,
				moduleRecordNameAddress);
		}
		return false;
	}

	std::array<unsigned char, kHiddenModuleInfoRecordBytes> moduleRecordBytes = {};
	if (!TryReadMemoryCopy(
			moduleRecordAddress + 12,
			moduleRecordBytes.data(),
			moduleRecordBytes.size())) {
		if (outError != nullptr) {
			*outError = std::format(
				"hidden_dialog_module_record_bytes_unreadable index={} record=0x{:X}",
				moduleIndex,
				moduleRecordAddress);
		}
		return false;
	}

	const auto viewerCtor = Bind<FnHiddenModuleInfoDialogCtor>(addrs.hiddenDialogCtor);
	const auto createDialogIndirect = Bind<FnCreateDialogIndirect>(addrs.createDialogIndirect);
	const auto destroyCString = Bind<FnDestroyCString>(addrs.destroyCString);
	const auto destroyTreeCtrl = Bind<FnDestroyTreeCtrl>(addrs.destroyTreeCtrl);
	const auto destroyEditCtrl = Bind<FnDestroyEditCtrl>(addrs.destroyEditCtrl);
	const auto destroyDialogWindow = Bind<FnDestroyDialogWindow>(addrs.destroyDialogWindow);
	const auto destroyImageList = Bind<FnDestroyImageList>(addrs.destroyImageList);
	const auto destroyModulePayload = Bind<FnDestroyModuleInfoPayload>(addrs.destroyModuleInfoPayload);
	const auto destroyDialogBase = Bind<FnDestroyDialogBase>(addrs.destroyDialogBase);
	const auto resolveModulePath = Bind<FnResolveModulePath>(addrs.resolveModulePath);

	std::vector<unsigned char> dialogStorage(kHiddenModuleInfoStorageSize, 0);
	auto* dialogObject = dialogStorage.data();
	viewerCtor(dialogObject, nullptr);

	bool pathConstructed = false;
	bool payloadConstructed = false;
	bool dialogCreated = false;
	HiddenModuleInfoDialogContext ctx;

	auto cleanupDialog = [&]() {
		if (dialogCreated && IsWindow(ctx.dialogHwnd)) {
			destroyDialogWindow(dialogObject);
			ctx.dialogHwnd = nullptr;
			dialogCreated = false;
		}
		if (payloadConstructed) {
			destroyImageList(dialogObject + kHiddenModuleInfoImageListOffset);
			destroyModulePayload(dialogObject + kHiddenModuleInfoPayloadOffset);
		}
		if (pathConstructed) {
			destroyCString(dialogObject + kHiddenModuleInfoPathOffset);
		}
		destroyTreeCtrl(dialogObject + 0x98);
		destroyEditCtrl(dialogObject + 0x5C);
		destroyDialogBase(dialogObject);
	};

	if (!resolveModulePath(
			moduleRecordName.c_str(),
			reinterpret_cast<char*>(dialogObject + kHiddenModuleInfoPathOffset),
			nullptr)) {
		cleanupDialog();
		if (outError != nullptr) {
			*outError = "hidden_dialog_resolve_module_path_failed";
		}
		return false;
	}
	pathConstructed = true;

	std::memcpy(
		dialogObject + kHiddenModuleInfoRecordOffset,
		moduleRecordBytes.data(),
		moduleRecordBytes.size());
	payloadConstructed = true;

	std::lock_guard hiddenDialogLock(g_hiddenModuleInfoDialogMutex);
	g_hiddenModuleInfoDialogContext = &ctx;
	ctx.cbtHook = SetWindowsHookExA(WH_CBT, HiddenModuleInfoDialogCbtProc, nullptr, GetCurrentThreadId());
	const auto dialogTemplate = ResolveDialogTemplateFromObject(dialogObject);
	if (dialogTemplate == nullptr) {
		if (ctx.cbtHook != nullptr) {
			UnhookWindowsHookEx(ctx.cbtHook);
			ctx.cbtHook = nullptr;
		}
		g_hiddenModuleInfoDialogContext = nullptr;
		cleanupDialog();
		if (outError != nullptr) {
			*outError = "hidden_modeless_dialog_template_missing";
		}
		return false;
	}

	if (!createDialogIndirect(dialogObject, dialogTemplate, nullptr, nullptr)) {
		if (ctx.cbtHook != nullptr) {
			UnhookWindowsHookEx(ctx.cbtHook);
			ctx.cbtHook = nullptr;
		}
		g_hiddenModuleInfoDialogContext = nullptr;
		cleanupDialog();
		if (outError != nullptr) {
			*outError = "hidden_modeless_dialog_create_failed";
		}
		return false;
	}
	if (ctx.cbtHook != nullptr) {
		UnhookWindowsHookEx(ctx.cbtHook);
		ctx.cbtHook = nullptr;
	}
	g_hiddenModuleInfoDialogContext = nullptr;

	if (ctx.dialogHwnd == nullptr) {
		ctx.dialogHwnd = reinterpret_cast<HWND>(
			*reinterpret_cast<std::uintptr_t*>(dialogObject + kDialogObjectHwndOffset));
	}
	if (ctx.dialogHwnd == nullptr || !IsWindow(ctx.dialogHwnd)) {
		cleanupDialog();
		if (outError != nullptr) {
			*outError = "hidden_modeless_dialog_hwnd_missing";
		}
		return false;
	}

	dialogCreated = true;
	if (!CaptureHiddenModuleInfoTextFromDialog(dialogObject, ctx.dialogHwnd, addrs, ctx.capturedText, ctx.error)) {
		cleanupDialog();
		outDump = {};
		outDump.modulePath = modulePath;
		outDump.trace = ctx.error.empty() ? "hidden_modeless_dialog_text_empty" : ctx.error;
		if (outError != nullptr) {
			*outError = outDump.trace;
		}
		return false;
	}

	cleanupDialog();

	if (ctx.capturedText.empty()) {
		outDump = {};
		outDump.modulePath = modulePath;
		outDump.trace = ctx.error.empty() ? "hidden_dialog_text_empty" : ctx.error;
		if (outError != nullptr) {
			*outError = outDump.trace;
		}
		return false;
	}

	outDump = {};
	outDump.modulePath = modulePath;
	outDump.nativeResult = 1;
	outDump.sourceKind = "hidden_modeless_dialog";
	outDump.trace = "hidden_modeless_capture_ok";
	ParseModulePublicInfoFromFormattedText(ctx.capturedText, outDump);
	if (outDump.moduleName.empty()) {
		outDump.moduleName = GetFileNameOnly(modulePath);
	}
	return true;
}

std::uintptr_t NormalizeRuntimeAddress(std::uintptr_t runtimeAddress, std::uintptr_t moduleBase)
{
	if (runtimeAddress == 0 || moduleBase == 0 || runtimeAddress < moduleBase) {
		return 0;
	}
	return runtimeAddress - moduleBase + kImageBase;
}

std::uintptr_t ReadNormalizedImm32(std::uintptr_t instructionAddress, size_t immOffset, std::uintptr_t moduleBase)
{
	if (instructionAddress == 0) {
		return 0;
	}
	const auto runtimeValue = static_cast<std::uintptr_t>(
		*reinterpret_cast<const std::uint32_t*>(instructionAddress + immOffset));
	return NormalizeRuntimeAddress(runtimeValue, moduleBase);
}

std::uintptr_t ReadNormalizedRelativeCallTarget(
	std::uintptr_t instructionAddress,
	size_t relOffset,
	size_t instructionSize,
	std::uintptr_t moduleBase)
{
	if (instructionAddress == 0) {
		return 0;
	}
	const auto rel = *reinterpret_cast<const std::int32_t*>(instructionAddress + relOffset);
	const auto runtimeTarget = instructionAddress + instructionSize + static_cast<std::intptr_t>(rel);
	return NormalizeRuntimeAddress(runtimeTarget, moduleBase);
}

std::uintptr_t ResolveUniqueCodeAddress(
	const char* label,
	const char* pattern,
	std::uintptr_t moduleBase)
{
	const auto matches = FindSelfModelMemoryAll(pattern);
	if (matches.size() != 1) {
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] resolve {} failed, matchCount={}",
			label,
			matches.size()));
		return 0;
	}
	return NormalizeRuntimeAddress(static_cast<std::uintptr_t>(matches.front()), moduleBase);
}

std::uintptr_t ResolveUniqueImmAddress(
	const char* label,
	const char* pattern,
	size_t immOffset,
	std::uintptr_t moduleBase)
{
	const auto matches = FindSelfModelMemoryAll(pattern);
	if (matches.size() != 1) {
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] resolve {} failed, matchCount={}",
			label,
			matches.size()));
		return 0;
	}
	return ReadNormalizedImm32(static_cast<std::uintptr_t>(matches.front()), immOffset, moduleBase);
}

std::uintptr_t ResolveUniqueRelativeCallTarget(
	const char* label,
	const char* pattern,
	size_t instructionOffset,
	size_t relOffset,
	size_t instructionSize,
	std::uintptr_t moduleBase)
{
	const auto matches = FindSelfModelMemoryAll(pattern);
	if (matches.size() != 1) {
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] resolve {} failed, matchCount={}",
			label,
			matches.size()));
		return 0;
	}

	return ReadNormalizedRelativeCallTarget(
		static_cast<std::uintptr_t>(matches.front()) + instructionOffset,
		relOffset,
		instructionSize,
		moduleBase);
}

std::uintptr_t ResolveSharedImmAddress(
	const char* label,
	const char* pattern,
	size_t immOffset,
	std::uintptr_t moduleBase)
{
	const auto matches = FindSelfModelMemoryAll(pattern);
	if (matches.empty()) {
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] resolve {} failed, matchCount=0",
			label));
		return 0;
	}

	std::uintptr_t resolved = 0;
	for (const int match : matches) {
		const auto value = ReadNormalizedImm32(static_cast<std::uintptr_t>(match), immOffset, moduleBase);
		if (value == 0) {
			OutputStringToELog(std::format(
				"[NativeModulePublicInfo] resolve {} failed, zero_value_match={}",
				label,
				match));
			return 0;
		}
		if (resolved == 0) {
			resolved = value;
			continue;
		}
		if (resolved != value) {
			OutputStringToELog(std::format(
				"[NativeModulePublicInfo] resolve {} failed, mismatch existing=0x{:X} current=0x{:X} matchCount={}",
				label,
				static_cast<unsigned int>(resolved),
				static_cast<unsigned int>(value),
				matches.size()));
			return 0;
		}
	}
	return resolved;
}

std::uintptr_t ResolveSharedRelativeCallTarget(
	const char* label,
	const char* pattern,
	size_t instructionOffset,
	size_t relOffset,
	size_t instructionSize,
	std::uintptr_t moduleBase)
{
	const auto matches = FindSelfModelMemoryAll(pattern);
	if (matches.empty()) {
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] resolve {} failed, matchCount=0",
			label));
		return 0;
	}

	std::uintptr_t resolved = 0;
	for (const int match : matches) {
		const auto instructionAddress = static_cast<std::uintptr_t>(match) + instructionOffset;
		const auto value = ReadNormalizedRelativeCallTarget(
			instructionAddress,
			relOffset,
			instructionSize,
			moduleBase);
		if (value == 0) {
			OutputStringToELog(std::format(
				"[NativeModulePublicInfo] resolve {} failed, zero_target_match={}",
				label,
				match));
			return 0;
		}
		if (resolved == 0) {
			resolved = value;
			continue;
		}
		if (resolved != value) {
			OutputStringToELog(std::format(
				"[NativeModulePublicInfo] resolve {} failed, mismatch existing=0x{:X} current=0x{:X} matchCount={}",
				label,
				static_cast<unsigned int>(resolved),
				static_cast<unsigned int>(value),
				matches.size()));
			return 0;
		}
	}
	return resolved;
}

bool PopulateNativeModulePublicInfoAddresses(NativeModulePublicInfoAddresses& addrs, std::uintptr_t moduleBase)
{
	addrs = {};
	addrs.moduleBase = moduleBase;

	static constexpr char kHiddenDialogCtorPattern[] =
		"6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 51 8B 44 24 14 56 57 8B F1 50 68 5B 02 00 00 89 74 24 10 E8 ?? ?? ?? ?? 8D 7E 5C C7 44 24 14 00 00 00 00 8B CF E8 ?? ?? ?? ?? C7 07 ?? ?? ?? ?? 8D BE 98 00 00 00";
	static constexpr char kCreateDialogIndirectPattern[] =
		"B8 ?? ?? ?? ?? E8 ?? ?? ?? ?? 83 EC 34 53 56 33 DB 57 39 5D 10 8B F1 89 65 F0 89 75 DC 75 0B E8 ?? ?? ?? ?? 8B 40 08 89 45 10 E8 ?? ?? ?? ?? 8B B8 38 10 00 00 6A 10";
	static constexpr char kUpdateHiddenSelectionPattern[] =
		"64 A1 00 00 00 00 6A FF 68 ?? ?? ?? ?? 50 8B 44 24 14 64 89 25 00 00 00 00 83 EC 14 C7 00 00 00 00 00 56 8B F1 57 6A 00 8B 8E B4 00 00 00 6A 09 68 0A 11 00 00 51 FF 15 ?? ?? ?? ?? 8B F8 85 FF 74 66";
	static constexpr char kDestroyCStringPattern[] =
		"56 8B F1 8B 06 8D 48 F4 3B 0D ?? ?? ?? ?? 74 18 83 C0 F4 50 FF 15 ?? ?? ?? ?? 85 C0 7F 0A 8B 0E 83 E9 0C E8 ?? ?? ?? ?? 5E C3";
	static constexpr char kDestroyTreeCtrlPattern[] =
		"B8 ?? ?? ?? ?? E8 ?? ?? ?? ?? 51 56 8B F1 89 75 F0 C7 06 ?? ?? ?? ?? 83 65 FC 00 E8 ?? ?? ?? ?? 83 4D FC FF 8B CE E8 ?? ?? ?? ?? 8B 4D F4 5E 64 89 0D 00 00 00 00 C9 C3 B8 ?? ?? ?? ?? E8 ?? ?? ?? ?? 83 EC 2C 8B 45 0C 83 65 F0 00 53 56 57 6A 01 89 45 CC A1 ?? ?? ?? ?? 5F 8B D9 89 7D C8";
	static constexpr char kDestroyEditCtrlPattern[] =
		"B8 ?? ?? ?? ?? E8 ?? ?? ?? ?? 51 56 8B F1 89 75 F0 C7 06 ?? ?? ?? ?? 83 65 FC 00 E8 ?? ?? ?? ?? 83 4D FC FF 8B CE E8 ?? ?? ?? ?? 8B 4D F4 5E 64 89 0D 00 00 00 00 C9 C3 B8 ?? ?? ?? ?? C3 56 8B F1 E8 39 10 00 00 83 A6 BC 00 00 00 00 C7 06 ?? ?? ?? ?? 8B C6 5E C3 53 56 57 8B F9 6A 00 E8 4B";
	static constexpr char kDestroyDialogWindowPattern[] =
		"56 8B F1 83 7E 1C 00 75 04 33 C0 5E C3 53 57 6A 00 E8 ?? ?? ?? ?? FF 76 1C 8D 48 04 E8 ?? ?? ?? ?? 8B 4E 38 8B F8 85 C9 75 0B";
	static constexpr char kDestroyImageListPattern[] =
		"B8 ?? ?? ?? ?? E8 ?? ?? ?? ?? 51 89 4D F0 C7 01 ?? ?? ?? ?? 83 65 FC 00 E8 36 00 00 00 8B 4D F4 64 89 0D 00 00 00 00 C9 C3 56 8B F1 57 8B 7E 04 85 FF 74 16";
	static constexpr char kDestroyPayloadPattern[] =
		"6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 83 EC 08 53 56 8B F1 57 89 74 24 0C C7 44 24 1C 1F 00 00 00 E8 ?? ?? ?? ?? 8D 8E 1C 03 00 00";
	static constexpr char kDestroyDialogBasePattern[] =
		"B8 ?? ?? ?? ?? E8 ?? ?? ?? ?? 51 56 8B F1 89 75 F0 C7 06 ?? ?? ?? ?? 83 65 FC 00 83 7E 1C 00 74 05 E8 ?? ?? ?? ?? 83 4D FC FF 8B CE E8 ?? ?? ?? ??";
	static constexpr char kResolveModulePathPattern[] =
		"64 A1 00 00 00 00 6A FF 68 ?? ?? ?? ?? 50 64 89 25 00 00 00 00 8B 44 24 10 83 EC 08 53 56 57 8B 7C 24 28 8B F1 57 50 B9 ?? ?? ?? ?? E8 ?? ?? ?? ?? 8B 07 50 E8 ?? ?? ?? ?? BB 01 00 00 00";
	static constexpr char kImportedModuleArrayRefPattern[] =
		"A1 ?? ?? ?? ?? 8B 14 B1 8D 4C 24 10 51 B9 ?? ?? ?? ?? 8B 2C 90 8B 45 08 50 E8 ?? ?? ?? ??";

	addrs.initByteContainer = ResolveUniqueCodeAddress(
		"init_byte_container",
		"8B C1 33 C9 C7 00 ?? ?? ?? ?? C7 40 04 ?? ?? ?? ?? 89 48 08 89 48 10 89 48 0C C3",
		moduleBase);
	addrs.loadModulePublicInfo = ResolveUniqueCodeAddress(
		"load_module_public_info",
		"6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 81 EC 44 09 00 00 53 55 56 57 8B F9 33 DB 8B 8C 24 6C 09 00 00 89 7C 24 30 3B CB 74 05 E8 ?? ?? ?? ?? A1 ?? ?? ?? ?? 89 44 24 28",
		moduleBase);

	static constexpr char kPrepareAndLoadModuleCallsitePattern[] =
		"6A 01 57 68 ?? ?? ?? ?? B9 ?? ?? ?? ?? E8 ?? ?? ?? ?? 8B 4C 24 ?? 8D 44 24 ?? 50 51 68 ?? ?? ?? ?? B9 ?? ?? ?? ?? E8 ?? ?? ?? ??";
	addrs.prepareModuleContext = ResolveUniqueRelativeCallTarget(
		"prepare_module_context",
		kPrepareAndLoadModuleCallsitePattern,
		13,
		1,
		5,
		moduleBase);
	addrs.moduleLoaderArg = ResolveUniqueImmAddress(
		"module_loader_arg",
		kPrepareAndLoadModuleCallsitePattern,
		4,
		moduleBase);
	addrs.moduleLoaderThis = ResolveUniqueImmAddress(
		"module_loader_this",
		kPrepareAndLoadModuleCallsitePattern,
		9,
		moduleBase);

	static constexpr char kRecorderCleanupPattern[] =
		"B9 ?? ?? ?? ?? C7 05 ?? ?? ?? ?? FF FF FF FF C7 05 ?? ?? ?? ?? 00 00 00 00 E8 ?? ?? ?? ?? B9 ?? ?? ?? ?? E8 ?? ?? ?? ?? 8B CE C7 05 ?? ?? ?? ?? 00 00 00 00 E8";
	addrs.recorderAuxContainer = ResolveSharedImmAddress(
		"recorder_aux_container",
		kRecorderCleanupPattern,
		1,
		moduleBase);
	addrs.recorderFlag = ResolveSharedImmAddress(
		"recorder_flag",
		kRecorderCleanupPattern,
		7,
		moduleBase);
	addrs.recorderStateA0 = ResolveSharedImmAddress(
		"recorder_state_a0",
		kRecorderCleanupPattern,
		17,
		moduleBase);
	addrs.clearByteContainer = ResolveSharedRelativeCallTarget(
		"clear_byte_container",
		kRecorderCleanupPattern,
		25,
		1,
		5,
		moduleBase);
	addrs.recorderBuffer = ResolveSharedImmAddress(
		"recorder_buffer",
		kRecorderCleanupPattern,
		31,
		moduleBase);
	addrs.recorderStateCC = ResolveSharedImmAddress(
		"recorder_state_cc",
		kRecorderCleanupPattern,
		44,
		moduleBase);

	addrs.hiddenDialogCtor = ResolveUniqueCodeAddress(
		"hidden_dialog_ctor",
		kHiddenDialogCtorPattern,
		moduleBase);
	addrs.createDialogIndirect = ResolveUniqueCodeAddress(
		"create_dialog_indirect",
		kCreateDialogIndirectPattern,
		moduleBase);
	addrs.updateHiddenSelection = ResolveUniqueCodeAddress(
		"update_hidden_selection",
		kUpdateHiddenSelectionPattern,
		moduleBase);
	addrs.destroyCString = ResolveUniqueCodeAddress(
		"destroy_cstring",
		kDestroyCStringPattern,
		moduleBase);
	addrs.destroyTreeCtrl = ResolveUniqueCodeAddress(
		"destroy_tree_ctrl",
		kDestroyTreeCtrlPattern,
		moduleBase);
	addrs.destroyEditCtrl = ResolveUniqueCodeAddress(
		"destroy_edit_ctrl",
		kDestroyEditCtrlPattern,
		moduleBase);
	addrs.destroyDialogWindow = ResolveUniqueCodeAddress(
		"destroy_dialog_window",
		kDestroyDialogWindowPattern,
		moduleBase);
	addrs.destroyImageList = ResolveUniqueCodeAddress(
		"destroy_image_list",
		kDestroyImageListPattern,
		moduleBase);
	addrs.destroyModuleInfoPayload = ResolveUniqueCodeAddress(
		"destroy_module_info_payload",
		kDestroyPayloadPattern,
		moduleBase);
	addrs.destroyDialogBase = ResolveUniqueCodeAddress(
		"destroy_dialog_base",
		kDestroyDialogBasePattern,
		moduleBase);
	addrs.resolveModulePath = ResolveUniqueCodeAddress(
		"resolve_module_path",
		kResolveModulePathPattern,
		moduleBase);
	addrs.importedModuleArrayGlobal = ResolveUniqueImmAddress(
		"imported_module_array_global",
		kImportedModuleArrayRefPattern,
		1,
		moduleBase);

	addrs.ok =
		addrs.initByteContainer != 0 &&
		addrs.clearByteContainer != 0 &&
		addrs.prepareModuleContext != 0 &&
		addrs.loadModulePublicInfo != 0 &&
		addrs.moduleLoaderThis != 0 &&
		addrs.moduleLoaderArg != 0 &&
		addrs.recorderFlag != 0 &&
		addrs.recorderStateA0 != 0 &&
		addrs.recorderAuxContainer != 0 &&
		addrs.recorderBuffer != 0 &&
		addrs.recorderStateCC != 0;
	addrs.hiddenDialogOk =
		addrs.hiddenDialogCtor != 0 &&
		addrs.createDialogIndirect != 0 &&
		addrs.updateHiddenSelection != 0 &&
		addrs.destroyCString != 0 &&
		addrs.destroyTreeCtrl != 0 &&
		addrs.destroyEditCtrl != 0 &&
		addrs.destroyDialogWindow != 0 &&
		addrs.destroyImageList != 0 &&
		addrs.destroyModuleInfoPayload != 0 &&
		addrs.destroyDialogBase != 0 &&
		addrs.resolveModulePath != 0 &&
		addrs.importedModuleArrayGlobal != 0;
	addrs.initialized = true;
	return addrs.ok;
}

const NativeModulePublicInfoAddresses& GetNativeModulePublicInfoAddresses(std::uintptr_t moduleBase)
{
	std::lock_guard lock(g_nativeModuleInfoMutex);
	if (!g_nativeModuleInfoAddresses.initialized || g_nativeModuleInfoAddresses.moduleBase != moduleBase) {
		PopulateNativeModulePublicInfoAddresses(g_nativeModuleInfoAddresses, moduleBase);
	}
	return g_nativeModuleInfoAddresses;
}

template <typename T>
T Bind(std::uintptr_t absoluteAddress)
{
	return reinterpret_cast<T>(absoluteAddress - kImageBase + reinterpret_cast<std::uintptr_t>(GetModuleHandle(nullptr)));
}

template <typename T>
T* PtrAbsolute(std::uintptr_t moduleBase, std::uintptr_t normalizedAddress)
{
	if (moduleBase == 0 || normalizedAddress < kImageBase) {
		return nullptr;
	}
	return reinterpret_cast<T*>(normalizedAddress - kImageBase + moduleBase);
}

int GetRecordExtraIntCount(int tag)
{
	switch (tag) {
	case 250:
	case 252:
	case 253:
	case 301:
	case 302:
	case 303:
	case 305:
	case 307:
		return 2 + (tag == 252 ? 1 : 0);
	case 251:
	case 306:
	case 308:
		return 4;
	case 309:
	case 311:
		return 3;
	default:
		return 0;
	}
}

bool IsLikelyModulePublicInfoTag(int tag)
{
	switch (tag) {
	case 250:
	case 251:
	case 252:
	case 253:
	case 301:
	case 302:
	case 303:
	case 305:
	case 306:
	case 307:
	case 308:
	case 309:
	case 311:
		return true;
	default:
		return false;
	}
}

bool IsLikelyTextByte(unsigned char ch)
{
	return ch == '\t' || (ch >= 0x20 && ch != 0x7F) || ch >= 0x80;
}

bool ShouldKeepExtractedString(const std::string& text)
{
	if (text.size() < 3) {
		return false;
	}
	size_t usefulCount = 0;
	for (unsigned char ch : text) {
		if (std::isalnum(ch) != 0 || ch >= 0x80 || ch == '_' || ch == '.' || ch == ':' || ch == '\\' || ch == '/') {
			++usefulCount;
		}
	}
	return usefulCount >= 2;
}

std::vector<std::string> ExtractPossibleStrings(const unsigned char* data, size_t size)
{
	std::vector<std::string> out;
	if (data == nullptr || size == 0) {
		return out;
	}

	std::unordered_set<std::string> seen;
	std::string current;
	current.reserve(128);
	for (size_t i = 0; i < size; ++i) {
		const unsigned char ch = data[i];
		if (IsLikelyTextByte(ch)) {
			current.push_back(static_cast<char>(ch));
			continue;
		}

		if (ShouldKeepExtractedString(current) && seen.insert(current).second) {
			out.push_back(current);
			if (out.size() >= 16) {
				return out;
			}
		}
		current.clear();
	}
	if (ShouldKeepExtractedString(current) && seen.insert(current).second) {
		out.push_back(current);
	}
	return out;
}

std::string GetFileNameOnly(const std::string& path)
{
	const size_t pos = path.find_last_of("\\/");
	return pos == std::string::npos ? path : path.substr(pos + 1);
}

std::string BuildTempProbeModulePath(const std::string& originalPath)
{
	char tempDir[MAX_PATH] = {};
	DWORD len = GetTempPathA(static_cast<DWORD>(std::size(tempDir)), tempDir);
	std::filesystem::path dir;
	if (len > 0 && len < std::size(tempDir)) {
		dir = tempDir;
	}
	else {
		dir = std::filesystem::temp_directory_path();
	}

	const std::filesystem::path original = std::filesystem::path(originalPath);
	const std::string stem = original.stem().string();
	const std::string ext = original.extension().string();
	const DWORD pid = GetCurrentProcessId();
	const DWORD tick = GetTickCount();
	const std::string fileName = std::format(
		"{}_autolinker_probe_{}_{}{}",
		stem.empty() ? "module" : stem,
		pid,
		tick,
		ext);
	return (dir / fileName).string();
}

bool TryCopyModuleToTempPath(const std::string& sourcePath, std::string& outTempPath)
{
	outTempPath = BuildTempProbeModulePath(sourcePath);
	std::error_code ec;
	std::filesystem::copy_file(
		std::filesystem::path(sourcePath),
		std::filesystem::path(outTempPath),
		std::filesystem::copy_options::overwrite_existing,
		ec);
	return !ec;
}

bool ContainsModuleDuplicateErrorLoose(const std::string& text)
{
	if (text.empty()) {
		return false;
	}

	const bool hasModulePathHint = text.find(".ec") != std::string::npos || text.find(".EC") != std::string::npos;
	if (!hasModulePathHint) {
		return false;
	}

	return
		text.find("鍚屾枃浠跺悕") != std::string::npos ||
		text.find("宸叉湁涓") != std::string::npos ||
		text.find("妯″潡瀛樺湪") != std::string::npos ||
		text.find("閸氬本鏋冩禒璺烘倳") != std::string::npos ||
		text.find("閺勬挻膩閸") != std::string::npos;
}

bool ShouldRetryWithTempCopy(
	const std::string& modulePath,
	const ModulePublicInfoDump& dump,
	const std::string& firstError)
{
	if (ContainsModuleDuplicateErrorLoose(firstError) || ContainsModuleDuplicateErrorLoose(dump.loaderError)) {
		return true;
	}

	if (dump.nativeResult != -1 || !dump.records.empty()) {
		return false;
	}

	if (modulePath.empty()) {
		return false;
	}

	const std::string fileName = GetFileNameOnly(modulePath);
	const auto containsPathHint = [&](const std::string& text) {
		return
			(!fileName.empty() && text.find(fileName) != std::string::npos) ||
			text.find(modulePath) != std::string::npos;
	};

	return
		(!firstError.empty() && containsPathHint(firstError)) ||
		(!dump.loaderError.empty() && containsPathHint(dump.loaderError));
}

bool SafeInitByteContainer(FnInitByteContainer fn, void* container)
{
	__try {
		fn(container);
		return true;
	} __except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool SafeClearByteContainer(FnClearByteContainer fn, void* container)
{
	__try {
		fn(container);
		return true;
	} __except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool SafeReadInt(const int* value, int* outValue)
{
	if (outValue == nullptr) {
		return false;
	}
	__try {
		*outValue = *value;
		return true;
	} __except (EXCEPTION_EXECUTE_HANDLER) {
		*outValue = 0;
		return false;
	}
}

bool SafeWriteInt(int* value, int newValue)
{
	__try {
		*value = newValue;
		return true;
	} __except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool SafeSnapshotByteContainer(const NativeByteContainer* container, std::vector<unsigned char>* outBytes)
{
	if (outBytes == nullptr) {
		return false;
	}
	outBytes->clear();
	__try {
		if (container == nullptr) {
			return false;
		}
		if (container->usedBytes > 0 && container->data != nullptr) {
			outBytes->assign(container->data, container->data + container->usedBytes);
		}
		return true;
	} __except (EXCEPTION_EXECUTE_HANDLER) {
		outBytes->clear();
		return false;
	}
}

bool SafeCleanupRecorderContainers(
	FnClearByteContainer clearFn,
	int* recorderStateA0,
	NativeByteContainer* recorderAuxContainer,
	NativeByteContainer* recorderBuffer,
	int* recorderStateCC,
	bool* outClearAuxOk,
	bool* outClearBufferOk)
{
	if (outClearAuxOk != nullptr) {
		*outClearAuxOk = false;
	}
	if (outClearBufferOk != nullptr) {
		*outClearBufferOk = false;
	}

	SafeWriteInt(recorderStateA0, 0);
	const bool clearAuxOk = SafeClearByteContainer(clearFn, recorderAuxContainer);
	const bool clearBufferOk = SafeClearByteContainer(clearFn, recorderBuffer);
	SafeWriteInt(recorderStateCC, 0);

	if (outClearAuxOk != nullptr) {
		*outClearAuxOk = clearAuxOk;
	}
	if (outClearBufferOk != nullptr) {
		*outClearBufferOk = clearBufferOk;
	}
	return clearAuxOk && clearBufferOk;
}

bool SafeLoadModulePublicInfoCall(
	FnLoadModulePublicInfo fn,
	void* loaderThis,
	void* loaderArg,
	char* modulePath,
	CString* outLoaderError,
	int* outNativeResult)
{
	__try {
		*outNativeResult = fn(loaderThis, loaderArg, modulePath, outLoaderError);
		return true;
	} __except (EXCEPTION_EXECUTE_HANDLER) {
		if (outNativeResult != nullptr) {
			*outNativeResult = -1;
		}
		if (outLoaderError != nullptr) {
			outLoaderError->Empty();
		}
		return false;
	}
}

bool SafePrepareModuleContextCall(
	FnPrepareModuleContext fn,
	void* loaderThis,
	void* loaderArg,
	int moduleIndex,
	int prepareMode)
{
	__try {
		fn(loaderThis, loaderArg, moduleIndex, prepareMode);
		return true;
	} __except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

int ResolveImportedModuleIndex(const std::string& modulePath)
{
	IDEFacade& ide = IDEFacade::Instance();

	int index = ide.FindECOMIndex(modulePath);
	if (index >= 0) {
		return index;
	}

	const std::filesystem::path pathObj(modulePath);
	const std::string stem = pathObj.stem().string();
	if (!stem.empty()) {
		index = ide.FindECOMNameIndex(stem);
		if (index >= 0) {
			return index;
		}
	}

	return -1;
}

bool ParseSingleRecordedRecord(
	const unsigned char* data,
	size_t availableBytes,
	ModulePublicInfoRecord* outRecord,
	size_t* outConsumedBytes)
{
	if (outRecord == nullptr || outConsumedBytes == nullptr || data == nullptr || availableBytes < 6) {
		return false;
	}

	const unsigned short tag = *reinterpret_cast<const unsigned short*>(data);
	const unsigned int bodySize = *reinterpret_cast<const unsigned int*>(data + 2);
	if (bodySize > kMaxRecordBodySize) {
		return false;
	}

	const size_t totalRecordBytes = static_cast<size_t>(bodySize) + 6;
	if (totalRecordBytes > availableBytes) {
		return false;
	}

	ModulePublicInfoRecord record;
	record.tag = static_cast<int>(tag);
	record.bodySize = bodySize;
	record.rawBytes.resize(totalRecordBytes);
	std::memcpy(record.rawBytes.data(), data, totalRecordBytes);

	const int extraIntCount = GetRecordExtraIntCount(record.tag);
	const size_t maxHeaderBytes = static_cast<size_t>(extraIntCount) * sizeof(std::int32_t);
	const size_t availableHeaderBytes = bodySize < maxHeaderBytes ? bodySize : maxHeaderBytes;
	const size_t availableHeaderInts = availableHeaderBytes / sizeof(std::int32_t);
	record.headerInts.reserve(availableHeaderInts);
	for (size_t i = 0; i < availableHeaderInts; ++i) {
		record.headerInts.push_back(*reinterpret_cast<const std::int32_t*>(data + 6 + i * sizeof(std::int32_t)));
	}

	record.payloadOffset = 6 + static_cast<int>(availableHeaderInts * sizeof(std::int32_t));
	if (record.payloadOffset < 6) {
		record.payloadOffset = 6;
	}
	const size_t payloadOffset = static_cast<size_t>(record.payloadOffset);
	if (payloadOffset < record.rawBytes.size()) {
		record.extractedStrings = ExtractPossibleStrings(
			record.rawBytes.data() + payloadOffset,
			record.rawBytes.size() - payloadOffset);
	}

	*outRecord = std::move(record);
	*outConsumedBytes = totalRecordBytes;
	return true;
}

bool LoadModulePublicInfoDumpOnce(
	const std::string& modulePath,
	std::uintptr_t moduleBase,
	const NativeModulePublicInfoAddresses& addrs,
	ModulePublicInfoDump& outDump,
	std::string* outError)
{
	if (outError != nullptr) {
		outError->clear();
	}

	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once begin path={}",
		modulePath));

	auto* const loaderThis = reinterpret_cast<void*>(addrs.moduleLoaderThis - kImageBase + moduleBase);
	auto* const loaderArg = reinterpret_cast<void*>(addrs.moduleLoaderArg - kImageBase + moduleBase);
	auto* const recorderFlag = PtrAbsolute<int>(moduleBase, addrs.recorderFlag);
	auto* const recorderStateA0 = PtrAbsolute<int>(moduleBase, addrs.recorderStateA0);
	auto* const recorderAuxContainer = PtrAbsolute<NativeByteContainer>(moduleBase, addrs.recorderAuxContainer);
	auto* const recorderBuffer = PtrAbsolute<NativeByteContainer>(moduleBase, addrs.recorderBuffer);
	auto* const recorderStateCC = PtrAbsolute<int>(moduleBase, addrs.recorderStateCC);
	const auto initByteContainer = Bind<FnInitByteContainer>(addrs.initByteContainer);
	const auto clearByteContainer = Bind<FnClearByteContainer>(addrs.clearByteContainer);
	const auto prepareModuleContext = Bind<FnPrepareModuleContext>(addrs.prepareModuleContext);
	const auto loadModule = Bind<FnLoadModulePublicInfo>(addrs.loadModulePublicInfo);

	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once resolved path={} loaderThis=0x{:X} loaderArg=0x{:X} flag=0x{:X} stateA0=0x{:X} aux=0x{:X} buffer=0x{:X} stateCC=0x{:X} init=0x{:X} clear=0x{:X} prepare=0x{:X} load=0x{:X}",
		modulePath,
		static_cast<unsigned int>(reinterpret_cast<std::uintptr_t>(loaderThis)),
		static_cast<unsigned int>(reinterpret_cast<std::uintptr_t>(loaderArg)),
		static_cast<unsigned int>(reinterpret_cast<std::uintptr_t>(recorderFlag)),
		static_cast<unsigned int>(reinterpret_cast<std::uintptr_t>(recorderStateA0)),
		static_cast<unsigned int>(reinterpret_cast<std::uintptr_t>(recorderAuxContainer)),
		static_cast<unsigned int>(reinterpret_cast<std::uintptr_t>(recorderBuffer)),
		static_cast<unsigned int>(reinterpret_cast<std::uintptr_t>(recorderStateCC)),
		static_cast<unsigned int>(addrs.initByteContainer),
		static_cast<unsigned int>(addrs.clearByteContainer),
		static_cast<unsigned int>(addrs.prepareModuleContext),
		static_cast<unsigned int>(addrs.loadModulePublicInfo)));

	if (loaderThis == nullptr ||
		loaderArg == nullptr ||
		recorderFlag == nullptr ||
		recorderStateA0 == nullptr ||
		recorderAuxContainer == nullptr ||
		recorderBuffer == nullptr ||
		recorderStateCC == nullptr) {
		if (outError != nullptr) {
			*outError = "resolved_runtime_pointer_invalid";
		}
		return false;
	}

	int recorderFlagValue = 0;
	if (!SafeReadInt(recorderFlag, &recorderFlagValue)) {
		if (outError != nullptr) {
			*outError = "read_recorder_flag_exception";
		}
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] dump_once read_flag exception path={}",
			modulePath));
		return false;
	}
	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once before_guard path={} recorderFlag={}",
		modulePath,
		recorderFlagValue));

	if (recorderFlagValue != -1) {
		if (outError != nullptr) {
		*outError = "module_public_info_recorder_busy";
		}
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] dump_once recorder_busy path={} recorderFlag={}",
			modulePath,
			recorderFlagValue));
		return false;
	}

	bool preClearAuxOk = false;
	bool preClearBufferOk = false;
	SafeCleanupRecorderContainers(
		clearByteContainer,
		recorderStateA0,
		recorderAuxContainer,
		recorderBuffer,
		recorderStateCC,
		&preClearAuxOk,
		&preClearBufferOk);
	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once pre_cleanup path={} clearAuxOk={} clearBufferOk={}",
		modulePath,
		preClearAuxOk ? 1 : 0,
		preClearBufferOk ? 1 : 0));

	const int moduleIndex = ResolveImportedModuleIndex(modulePath);
	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once resolve_import_index path={} moduleIndex={}",
		modulePath,
		moduleIndex));
	if (moduleIndex < 0) {
		outDump = {};
		outDump.modulePath = modulePath;
		outDump.nativeResult = -1;
		outDump.trace = "imported_module_index_not_found";
		if (outError != nullptr) {
			*outError = "imported_module_index_not_found";
		}
		return false;
	}

	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once prepare_call begin path={} moduleIndex={}",
		modulePath,
		moduleIndex));
	if (!SafePrepareModuleContextCall(
			prepareModuleContext,
			loaderThis,
			loaderArg,
			moduleIndex,
			1)) {
		outDump = {};
		outDump.modulePath = modulePath;
		outDump.nativeResult = -1;
		outDump.trace = std::format("prepare_module_context_exception module_index={}", moduleIndex);
		if (outError != nullptr) {
			*outError = "prepare_module_context_exception";
		}
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] dump_once prepare_call exception path={} moduleIndex={}",
			modulePath,
			moduleIndex));
		return false;
	}
	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once prepare_call ok path={} moduleIndex={}",
		modulePath,
		moduleIndex));

	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once init_containers begin path={}",
		modulePath));
	if (!SafeInitByteContainer(initByteContainer, recorderAuxContainer) ||
		!SafeInitByteContainer(initByteContainer, recorderBuffer)) {
		if (outError != nullptr) {
			*outError = "init_byte_container_exception";
		}
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] dump_once init_container exception path={}",
			modulePath));
		SafeWriteInt(recorderFlag, -1);
		SafeWriteInt(recorderStateA0, 0);
		SafeWriteInt(recorderStateCC, 0);
		return false;
	}
	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once init_containers ok path={}",
		modulePath));

	if (!SafeWriteInt(recorderStateA0, 0) ||
		!SafeWriteInt(recorderStateCC, 0) ||
		!SafeWriteInt(recorderFlag, 0)) {
		if (outError != nullptr) {
			*outError = "write_recorder_state_exception";
		}
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] dump_once write_state exception path={}",
			modulePath));
		SafeWriteInt(recorderFlag, -1);
		SafeWriteInt(recorderStateA0, 0);
		SafeWriteInt(recorderStateCC, 0);
		SafeClearByteContainer(clearByteContainer, recorderAuxContainer);
		SafeClearByteContainer(clearByteContainer, recorderBuffer);
		return false;
	}
	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once recorder_armed path={}",
		modulePath));

	std::vector<char> modulePathBuffer(modulePath.begin(), modulePath.end());
	modulePathBuffer.push_back('\0');

	CString loaderError;
	int nativeResult = -1;
	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once load_call begin path={}",
		modulePath));
	if (!SafeLoadModulePublicInfoCall(
			loadModule,
			loaderThis,
			loaderArg,
			modulePathBuffer.data(),
			&loaderError,
			&nativeResult)) {
		SafeWriteInt(recorderFlag, -1);
		SafeWriteInt(recorderStateA0, 0);
		SafeClearByteContainer(clearByteContainer, recorderAuxContainer);
		SafeClearByteContainer(clearByteContainer, recorderBuffer);
		SafeWriteInt(recorderStateCC, 0);
		outDump = {};
		outDump.modulePath = modulePath;
		outDump.nativeResult = -1;
		outDump.trace = "load_module_public_info_exception";
		if (outError != nullptr) {
			*outError = "load_module_public_info_exception";
		}
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] dump_once load_module exception path={}",
			modulePath));
		return false;
	}
	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once load_call ok path={} nativeResult={}",
		modulePath,
		nativeResult));

	std::vector<unsigned char> recordedBytes;
	const bool snapshotOk = SafeSnapshotByteContainer(recorderBuffer, &recordedBytes);
	if (!snapshotOk) {
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] dump_once snapshot_buffer exception path={}",
			modulePath));
	}

	// Do not clear recorder buffers immediately on the success path.
	// The native loader appears to have post-return consumers in some builds; clearing
	// here can make the IDE crash after our extraction already succeeded.
	SafeWriteInt(recorderFlag, -1);
	const bool clearAuxOk = true;
	const bool clearBufferOk = true;

	OutputStringToELog(std::format(
		"[NativeModulePublicInfo] dump_once after load_module path={} nativeResult={} loaderError={} recordedBytes={} clearAuxOk={} clearBufferOk={} cleanupMode=deferred",
		modulePath,
		nativeResult,
		static_cast<LPCSTR>(loaderError),
		recordedBytes.size(),
		clearAuxOk ? 1 : 0,
		clearBufferOk ? 1 : 0));

	outDump = {};
	outDump.modulePath = modulePath;
	outDump.nativeResult = nativeResult;
	outDump.loaderError = static_cast<LPCSTR>(loaderError);

	size_t rawRecordCount = 0;
	size_t publicRecordCount = 0;
	for (size_t offset = 0; offset < recordedBytes.size();) {
		ModulePublicInfoRecord record;
		size_t consumedBytes = 0;
		if (!ParseSingleRecordedRecord(
				recordedBytes.data() + offset,
				recordedBytes.size() - offset,
				&record,
				&consumedBytes) ||
			consumedBytes == 0) {
			outDump.trace = std::format(
				"native_result={} raw_record_count={} public_record_count={} recorded_bytes={} parse_failed_at={} clear_aux_ok={} clear_buffer_ok={}",
				nativeResult,
				rawRecordCount,
				publicRecordCount,
				recordedBytes.size(),
				offset,
				clearAuxOk ? 1 : 0,
				clearBufferOk ? 1 : 0);
			if (outError != nullptr) {
				*outError = "parse_recorded_module_public_info_failed";
			}
			return false;
		}

		++rawRecordCount;
		if (IsLikelyModulePublicInfoTag(record.tag)) {
			outDump.records.push_back(std::move(record));
			++publicRecordCount;
		}
		offset += consumedBytes;
	}

	outDump.trace = std::format(
		"native_result={} raw_record_count={} public_record_count={} recorded_bytes={} clear_aux_ok={} clear_buffer_ok={}",
		nativeResult,
		rawRecordCount,
		publicRecordCount,
		recordedBytes.size(),
		clearAuxOk ? 1 : 0,
		clearBufferOk ? 1 : 0);

	if (nativeResult == -1) {
		if (outError != nullptr) {
			*outError = loaderError.IsEmpty() ? "load_module_public_info_failed" : static_cast<LPCSTR>(loaderError);
		}
		return false;
	}

	return true;
}

}  // namespace

bool LoadModulePublicInfoDumpFromSource(
	const std::string& modulePath,
	std::uintptr_t moduleBase,
	ModulePublicInfoLoadSource source,
	ModulePublicInfoDump* outDump,
	std::string* outError)
{
	if (outDump != nullptr) {
		*outDump = {};
		outDump->modulePath = modulePath;
	}
	if (outError != nullptr) {
		outError->clear();
	}
	if (modulePath.empty()) {
		if (outError != nullptr) {
			*outError = "module path is empty";
		}
		return false;
	}

	if (source == ModulePublicInfoLoadSource::kLocalEc) {
		ModulePublicInfoDump localDump;
		const bool ok = ParseModulePublicInfoFromEcFile(modulePath, localDump, outError);
		if (ok && outDump != nullptr) {
			*outDump = std::move(localDump);
		}
		return ok;
	}

	if (source == ModulePublicInfoLoadSource::kHiddenDialog) {
		ModulePublicInfoDump hiddenOnlyDump;
		const bool ok = LoadModulePublicInfoDumpHiddenDialog(modulePath, moduleBase, hiddenOnlyDump, outError);
		if (ok && outDump != nullptr) {
			*outDump = std::move(hiddenOnlyDump);
		}
		return ok;
	}

	const auto& nativeAddrs = GetNativeModulePublicInfoAddresses(moduleBase);
	if (source == ModulePublicInfoLoadSource::kNativeRecorder) {
		if (!nativeAddrs.ok) {
			if (outError != nullptr) {
				*outError = "resolve_native_addresses_failed";
			}
			return false;
		}

		ModulePublicInfoDump nativeDump;
		std::string nativeError;
		const bool ok = LoadModulePublicInfoDumpOnce(modulePath, moduleBase, nativeAddrs, nativeDump, &nativeError);
		if (ok) {
			nativeDump.sourceKind = "native_recorder";
			if (outDump != nullptr) {
				*outDump = std::move(nativeDump);
			}
			if (outError != nullptr) {
				outError->clear();
			}
			return true;
		}
		if (outDump != nullptr) {
			*outDump = std::move(nativeDump);
		}
		if (outError != nullptr) {
			*outError = nativeError.empty() ? "load_module_public_info_failed" : nativeError;
		}
		return false;
	}

	ModulePublicInfoDump hiddenDump;
	std::string hiddenError;
	if (LoadModulePublicInfoDumpHiddenDialog(modulePath, moduleBase, hiddenDump, &hiddenError)) {
		if (outDump != nullptr) {
			*outDump = std::move(hiddenDump);
		}
		if (outError != nullptr) {
			outError->clear();
		}
		return true;
	}
	if (!hiddenError.empty()) {
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] hidden_dialog failed path={} error={}",
			modulePath,
			hiddenError));
	}

	ModulePublicInfoDump parsedDump;
	std::string parseError;
	if (ParseModulePublicInfoFromEcFile(modulePath, parsedDump, &parseError)) {
		if (!hiddenError.empty()) {
			parsedDump.trace = std::format("hidden_error={} {}", hiddenError, parsedDump.trace);
		}
		if (outDump != nullptr) {
			*outDump = std::move(parsedDump);
		}
		if (outError != nullptr) {
			outError->clear();
		}
		return true;
	}

	const auto& addrs = GetNativeModulePublicInfoAddresses(moduleBase);
	if (!addrs.ok) {
		if (outError != nullptr) {
			*outError = !hiddenError.empty()
				? hiddenError
				: (parseError.empty() ? "resolve_native_addresses_failed" : parseError);
		}
		return false;
	}

	ModulePublicInfoDump dump;
	std::string firstError;
	if (LoadModulePublicInfoDumpOnce(modulePath, moduleBase, addrs, dump, &firstError)) {
		dump.sourceKind = "native_recorder";
		if (outDump != nullptr) {
			*outDump = std::move(dump);
		}
		return true;
	}

	if (ShouldRetryWithTempCopy(modulePath, dump, firstError)) {
		OutputStringToELog(std::format(
			"[NativeModulePublicInfo] retry with temp copy path={} error={} loaderError={} trace={}",
			modulePath,
			firstError,
			dump.loaderError,
			dump.trace));
		std::string tempPath;
		if (TryCopyModuleToTempPath(modulePath, tempPath)) {
			OutputStringToELog(std::format(
				"[NativeModulePublicInfo] temp copy created source={} temp={}",
				modulePath,
				tempPath));
			ModulePublicInfoDump retryDump;
			std::string retryError;
			const bool retryOk = LoadModulePublicInfoDumpOnce(tempPath, moduleBase, addrs, retryDump, &retryError);
			retryDump.trace += std::format(" retry_temp_copy=1 original_path={}", modulePath);

			std::error_code ec;
			std::filesystem::remove(std::filesystem::path(tempPath), ec);

			if (retryOk) {
				retryDump.modulePath = modulePath;
				if (outDump != nullptr) {
					*outDump = std::move(retryDump);
				}
				if (outError != nullptr) {
					outError->clear();
				}
				return true;
			}

			dump = std::move(retryDump);
			firstError = retryError;
		}
		else {
			OutputStringToELog(std::format(
				"[NativeModulePublicInfo] temp copy failed source={}",
				modulePath));
		}
	}

	if (outDump != nullptr) {
		*outDump = std::move(dump);
	}
	if (outError != nullptr) {
		*outError = !firstError.empty()
			? firstError
			: (!hiddenError.empty() ? hiddenError : parseError);
	}
	return false;
}

bool LoadModulePublicInfoDump(
	const std::string& modulePath,
	std::uintptr_t moduleBase,
	ModulePublicInfoDump* outDump,
	std::string* outError)
{
	return LoadModulePublicInfoDumpFromSource(
		modulePath,
		moduleBase,
		ModulePublicInfoLoadSource::kAuto,
		outDump,
		outError);
}

}  // namespace e571
