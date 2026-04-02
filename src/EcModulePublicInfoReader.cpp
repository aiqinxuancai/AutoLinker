#include "EcModulePublicInfoReader.h"

#include <Windows.h>

#include <algorithm>
#include <array>
#include <bit>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <format>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include <lib2.h>

#include "PathHelper.h"

namespace e571 {

namespace {

constexpr std::uint32_t kMagicEncryptedSource = 1162630231u;  // WTLE
constexpr std::uint32_t kMagicFileHeader1 = 1415007811u;      // CNWT
constexpr std::uint32_t kMagicFileHeader2 = 1196576837u;      // EPRG
constexpr std::uint32_t kMagicSection = 353465113u;
constexpr std::array<std::uint8_t, 4> kSectionNameNoKey = { 25, 115, 0, 7 };

constexpr std::int32_t kVarAttrByRef = 0x0002;
constexpr std::int32_t kVarAttrNullable = 0x0004;
constexpr std::int32_t kVarAttrArray = 0x0008;
constexpr std::int32_t kFuncAttrPublic = 0x0008;

constexpr std::int32_t kConstPageValue = 1;
constexpr std::int32_t kConstPageImage = 2;
constexpr std::int32_t kConstPageSound = 3;

constexpr std::uint8_t kConstTypeEmpty = 22;
constexpr std::uint8_t kConstTypeNumber = 23;
constexpr std::uint8_t kConstTypeBool = 24;
constexpr std::uint8_t kConstTypeDate = 25;
constexpr std::uint8_t kConstTypeText = 26;

constexpr std::int16_t kConstAttrLongText = 0x0010;

constexpr int kTagAssembly = 250;
constexpr int kTagClass = 251;
constexpr int kTagDataType = 252;
constexpr int kTagGlobalVar = 253;
constexpr int kTagSub = 301;
constexpr int kTagClassMethod = 302;
constexpr int kTagDll = 305;
constexpr int kTagConst = 307;
constexpr int kTagWindow = 309;

#pragma pack(push, 1)
struct RawSystemInfoSection {
	std::int16_t compileMajor = 0;
	std::int16_t compileMinor = 0;
	std::int32_t unknown1 = 0;
	std::int32_t unknown2 = 0;
	std::int32_t unknownType = 0;
	std::int32_t fileType = 0;
	std::int32_t unknown3 = 0;
	std::int32_t compileType = 0;
	std::int32_t unknown9[8] = {};
};

struct RawSectionHeader {
	std::uint32_t magic = 0;
	std::uint32_t infoChecksum = 0;
};

struct RawSectionInfo {
	std::uint8_t key[4] = {};
	char name[30] = {};
	std::int16_t reserveFill1 = 0;
	std::int32_t index = 0;
	std::int32_t flag1 = 0;
	std::int32_t dataChecksum = 0;
	std::int32_t dataLength = 0;
	std::int32_t reserveItems[10] = {};
};
#pragma pack(pop)

static_assert(sizeof(RawSystemInfoSection) == 60);
static_assert(sizeof(RawSectionHeader) == 8);
static_assert(sizeof(RawSectionInfo) == 92);

struct UserInfoSection {
	std::string programName;
	std::string programComment;
	std::string author;
	std::string zipCode;
	std::string address;
	std::string phone;
	std::string fax;
	std::string email;
	std::string homePage;
	std::string other;
	std::int32_t version1 = 0;
	std::int32_t version2 = 0;
};

struct BlockHeader {
	std::int32_t dwId = 0;
	std::int32_t dwUnk = 0;
};

struct VariableInfo {
	std::int32_t marker = 0;
	std::int32_t offset = 0;
	std::int32_t length = 0;
	std::int32_t dataType = 0;
	std::int16_t attr = 0;
	std::uint8_t arrayDimensions = 0;
	std::vector<std::int32_t> arrayBounds;
	std::string name;
	std::string comment;
};

struct CodePageInfo {
	BlockHeader header;
	std::int32_t unk1 = 0;
	std::int32_t baseClass = 0;
	std::string name;
	std::string comment;
	std::vector<std::int32_t> functionIds;
	std::vector<VariableInfo> pageVars;
};

struct FunctionInfo {
	BlockHeader header;
	std::int32_t type = 0;
	std::int32_t attr = 0;
	std::int32_t returnType = 0;
	std::string name;
	std::string comment;
	std::vector<VariableInfo> locals;
	std::vector<VariableInfo> params;
	std::vector<std::uint8_t> code;
};

struct DataTypeInfo {
	BlockHeader header;
	std::int32_t attr = 0;
	std::string name;
	std::string comment;
	std::vector<VariableInfo> members;
};

struct DllInfo {
	BlockHeader header;
	std::int32_t attr = 0;
	std::int32_t returnType = 0;
	std::string name;
	std::string comment;
	std::string fileName;
	std::string commandName;
	std::vector<VariableInfo> params;
};

struct ProgramHeaderInfo {
	std::int32_t versionFlag1 = 0;
	std::int32_t unk1 = 0;
	std::vector<std::uint8_t> unk2_1;
	std::vector<std::uint8_t> unk2_2;
	std::vector<std::uint8_t> unk2_3;
	std::vector<std::string> supportLibraryInfo;
	std::int32_t flag1 = 0;
	std::int32_t flag2 = 0;
	std::vector<std::uint8_t> unk3Op;
	std::vector<std::uint8_t> icon;
	std::string debugCommandLine;
};

struct ProgramSection {
	ProgramHeaderInfo header;
	std::vector<CodePageInfo> codePages;
	std::vector<FunctionInfo> functions;
	std::vector<VariableInfo> globals;
	std::vector<DataTypeInfo> dataTypes;
	std::vector<DllInfo> dlls;
};

struct FormInfo {
	BlockHeader header;
	std::int32_t unknown1 = 0;
	std::int32_t unknown2 = 0;
	std::string name;
	std::string comment;
};

struct ConstantInfo {
	std::int32_t marker = 0;
	std::int32_t offset = 0;
	std::int32_t length = 0;
	std::int16_t attr = 0;
	std::string name;
	std::string comment;
	std::string valueText;
	std::string resultText;
	bool longText = false;
};

struct ResourceSection {
	std::vector<FormInfo> forms;
	std::vector<ConstantInfo> constants;
	std::int32_t reserve = 0;
};

struct ModuleSections {
	bool hasSystemInfo = false;
	bool hasUserInfo = false;
	bool hasProgram = false;
	bool hasResources = false;
	RawSystemInfoSection systemInfo = {};
	UserInfoSection userInfo;
	ProgramSection program;
	ResourceSection resources;
};

struct PublicAssemblyInfo {
	CodePageInfo* page = nullptr;
	std::vector<FunctionInfo*> functions;
};

struct PublicClassInfo {
	CodePageInfo* page = nullptr;
	std::vector<FunctionInfo*> methods;
};

bool ReadFileBytes(const std::string& path, std::vector<std::uint8_t>& outBytes)
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

std::string TrimAsciiCopy(std::string text)
{
	size_t begin = 0;
	while (begin < text.size()) {
		const unsigned char ch = static_cast<unsigned char>(text[begin]);
		if (ch > 0x20) {
			break;
		}
		++begin;
	}

	size_t end = text.size();
	while (end > begin) {
		const unsigned char ch = static_cast<unsigned char>(text[end - 1]);
		if (ch > 0x20) {
			break;
		}
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string BytesToLocalText(const std::vector<std::uint8_t>& bytes)
{
	if (bytes.empty()) {
		return std::string();
	}
	return std::string(reinterpret_cast<const char*>(bytes.data()), bytes.size());
}

std::string RemoveTrailingNulls(std::string text)
{
	while (!text.empty() && text.back() == '\0') {
		text.pop_back();
	}
	return text;
}

std::string JoinStrings(const std::vector<std::string>& parts, const char* sep)
{
	std::string out;
	for (const auto& part : parts) {
		if (part.empty()) {
			continue;
		}
		if (!out.empty()) {
			out += sep;
		}
		out += part;
	}
	return out;
}

std::string JoinLines(const std::vector<std::string>& lines)
{
	return JoinStrings(lines, "\r\n");
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
			out.push_back(TrimAsciiCopy(text.substr(begin)));
			break;
		}
		out.push_back(TrimAsciiCopy(text.substr(begin, commaPos - begin)));
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

std::vector<std::string> SplitByCrLf(const std::string& text)
{
	std::vector<std::string> parts;
	std::string current;
	for (const char ch : text) {
		if (ch == '\r' || ch == '\n') {
			if (!current.empty()) {
				parts.push_back(current);
				current.clear();
			}
			continue;
		}
		current.push_back(ch);
	}
	if (!current.empty()) {
		parts.push_back(current);
	}
	return parts;
}

std::string GetFirstSupportLibraryToken(const std::string& rawText)
{
	const auto parts = SplitByCrLf(rawText);
	if (parts.empty()) {
		return TrimAsciiCopy(rawText);
	}
	return TrimAsciiCopy(parts.front());
}

class ByteReader {
public:
	explicit ByteReader(const std::vector<std::uint8_t>& bytes)
		: m_bytes(bytes)
	{
	}

	size_t position() const
	{
		return m_pos;
	}

	bool eof() const
	{
		return m_pos >= m_bytes.size();
	}

	bool ReadU8(std::uint8_t& value)
	{
		if (m_pos + sizeof(value) > m_bytes.size()) {
			return false;
		}
		value = m_bytes[m_pos];
		++m_pos;
		return true;
	}

	bool ReadI16(std::int16_t& value)
	{
		if (m_pos + sizeof(value) > m_bytes.size()) {
			return false;
		}
		std::memcpy(&value, m_bytes.data() + m_pos, sizeof(value));
		m_pos += sizeof(value);
		return true;
	}

	bool ReadI32(std::int32_t& value)
	{
		if (m_pos + sizeof(value) > m_bytes.size()) {
			return false;
		}
		std::memcpy(&value, m_bytes.data() + m_pos, sizeof(value));
		m_pos += sizeof(value);
		return true;
	}

	bool ReadU32(std::uint32_t& value)
	{
		if (m_pos + sizeof(value) > m_bytes.size()) {
			return false;
		}
		std::memcpy(&value, m_bytes.data() + m_pos, sizeof(value));
		m_pos += sizeof(value);
		return true;
	}

	bool ReadDouble(double& value)
	{
		if (m_pos + sizeof(value) > m_bytes.size()) {
			return false;
		}
		std::memcpy(&value, m_bytes.data() + m_pos, sizeof(value));
		m_pos += sizeof(value);
		return true;
	}

	bool ReadBool(bool& value)
	{
		std::int16_t raw = 0;
		if (!ReadI16(raw)) {
			return false;
		}
		value = raw != 0;
		return true;
	}

	bool ReadRaw(size_t size, std::vector<std::uint8_t>& out)
	{
		if (m_pos + size > m_bytes.size()) {
			return false;
		}
		out.assign(m_bytes.begin() + static_cast<std::ptrdiff_t>(m_pos),
			m_bytes.begin() + static_cast<std::ptrdiff_t>(m_pos + size));
		m_pos += size;
		return true;
	}

	bool Skip(size_t size)
	{
		if (m_pos + size > m_bytes.size()) {
			return false;
		}
		m_pos += size;
		return true;
	}

	bool ReadDynamicBytes(std::vector<std::uint8_t>& out)
	{
		std::int32_t size = 0;
		if (!ReadI32(size) || size < 0) {
			return false;
		}
		return ReadRaw(static_cast<size_t>(size), out);
	}

	bool SkipDynamicBytes()
	{
		std::int32_t size = 0;
		if (!ReadI32(size) || size < 0) {
			return false;
		}
		return Skip(static_cast<size_t>(size));
	}

	bool ReadDynamicText(std::string& out)
	{
		std::vector<std::uint8_t> bytes;
		if (!ReadDynamicBytes(bytes)) {
			return false;
		}
		out = BytesToLocalText(bytes);
		return true;
	}

	bool ReadStandardText(std::string& out)
	{
		const size_t start = m_pos;
		while (m_pos < m_bytes.size() && m_bytes[m_pos] != 0) {
			++m_pos;
		}
		if (m_pos >= m_bytes.size()) {
			return false;
		}
		out.assign(
			reinterpret_cast<const char*>(m_bytes.data() + start),
			m_pos - start);
		++m_pos;  // Skip trailing '\0'.
		return true;
	}

	bool ReadTextArray(std::vector<std::string>& out)
	{
		out.clear();
		std::int16_t count = 0;
		if (!ReadI16(count) || count < 0) {
			return false;
		}
		out.reserve(static_cast<size_t>(count));
		for (std::int16_t i = 0; i < count; ++i) {
			std::string item;
			if (!ReadDynamicText(item)) {
				return false;
			}
			out.push_back(std::move(item));
		}
		return true;
	}

private:
	const std::vector<std::uint8_t>& m_bytes;
	size_t m_pos = 0;
};

std::uint32_t GetHighType(std::uint32_t value)
{
	return (value & 0xF0000000u) >> 28;
}

std::string FormatDoubleLocal(double value)
{
	return std::format("{}", value);
}

std::string FormatDateLocal(double value)
{
	char buffer[128] = {};
	DateTimeFormat(buffer, static_cast<int>(sizeof(buffer)), value, FALSE);
	return buffer;
}

std::string BuildArraySuffix(const std::vector<std::int32_t>& bounds)
{
	if (bounds.empty()) {
		return std::string();
	}

	std::vector<std::string> parts;
	parts.reserve(bounds.size());
	for (const auto bound : bounds) {
		parts.push_back(std::to_string(bound));
	}
	return "[" + JoinStrings(parts, ",") + "]";
}

std::string BuildVarFlags(const VariableInfo& info)
{
	std::vector<std::string> parts;
	if ((info.attr & kVarAttrByRef) != 0) {
		parts.push_back("传址");
	}
	if ((info.attr & kVarAttrNullable) != 0) {
		parts.push_back("可空");
	}
	if ((info.attr & kVarAttrArray) != 0 || !info.arrayBounds.empty()) {
		parts.push_back("数组" + BuildArraySuffix(info.arrayBounds));
	}
	return JoinStrings(parts, " ");
}

std::string BuildMemberFlags(const VariableInfo& info)
{
	std::vector<std::string> parts;
	if ((info.attr & kVarAttrByRef) != 0) {
		parts.push_back("传址");
	}
	if ((info.attr & kVarAttrNullable) != 0) {
		parts.push_back("可空");
	}
	return JoinStrings(parts, " ");
}

std::string BuildMemberExtra(const VariableInfo& info)
{
	if (info.arrayBounds.empty()) {
		return std::string();
	}

	std::vector<std::string> parts;
	parts.reserve(info.arrayBounds.size());
	for (const auto bound : info.arrayBounds) {
		parts.push_back(std::to_string(bound));
	}
	return "\"" + JoinStrings(parts, ",") + "\"";
}

std::string QuoteIfNotEmpty(const std::string& text)
{
	const std::string trimmed = TrimAsciiCopy(text);
	if (trimmed.empty()) {
		return std::string();
	}
	return "\"" + trimmed + "\"";
}

void PushUniqueString(std::vector<std::string>& values, const std::string& value)
{
	if (value.empty()) {
		return;
	}
	if ((std::find)(values.begin(), values.end(), value) == values.end()) {
		values.push_back(value);
	}
}

void PushRecordStrings(ModulePublicInfoRecord& record, const std::string& text)
{
	const std::string trimmed = TrimAsciiCopy(text);
	if (trimmed.empty()) {
		return;
	}
	if ((std::find)(record.extractedStrings.begin(), record.extractedStrings.end(), trimmed) == record.extractedStrings.end()) {
		record.extractedStrings.push_back(trimmed);
	}
}

bool DecodeSectionName(const RawSectionInfo& info, std::string& outName)
{
	std::array<std::uint8_t, 30> nameBytes = {};
	std::memcpy(nameBytes.data(), info.name, nameBytes.size());

	if (!std::equal(std::begin(info.key), std::end(info.key), kSectionNameNoKey.begin())) {
		size_t keyIndex = 1;
		for (auto& ch : nameBytes) {
			ch = static_cast<std::uint8_t>(ch ^ info.key[(keyIndex % 4)]);
			++keyIndex;
		}
	}

	outName.assign(reinterpret_cast<const char*>(nameBytes.data()), nameBytes.size());
	outName = TrimAsciiCopy(RemoveTrailingNulls(std::move(outName)));
	return !outName.empty();
}

bool ReadBlockHeaders(ByteReader& reader, std::vector<BlockHeader>& outHeaders)
{
	outHeaders.clear();
	std::int32_t length = 0;
	if (!reader.ReadI32(length) || length < 0 || (length % 8) != 0) {
		return false;
	}

	const int count = length / 8;
	outHeaders.resize(static_cast<size_t>(count));
	for (int i = 0; i < count; ++i) {
		if (!reader.ReadI32(outHeaders[static_cast<size_t>(i)].dwId)) {
			return false;
		}
	}
	for (int i = 0; i < count; ++i) {
		if (!reader.ReadI32(outHeaders[static_cast<size_t>(i)].dwUnk)) {
			return false;
		}
	}
	return true;
}

bool ParseVariableList(
	std::int32_t count,
	const std::vector<std::uint8_t>& bytes,
	std::vector<VariableInfo>& outVars)
{
	outVars.clear();
	if (count < 0) {
		return false;
	}
	if (count == 0) {
		return true;
	}

	ByteReader reader(bytes);
	outVars.resize(static_cast<size_t>(count));

	for (std::int32_t i = 0; i < count; ++i) {
		if (!reader.ReadI32(outVars[static_cast<size_t>(i)].marker)) {
			return false;
		}
	}
	for (std::int32_t i = 0; i < count; ++i) {
		if (!reader.ReadI32(outVars[static_cast<size_t>(i)].offset)) {
			return false;
		}
	}

	const size_t bodyBase = reader.position();
	for (auto& item : outVars) {
		const size_t targetOffset = bodyBase + static_cast<size_t>(item.offset);
		if (targetOffset > bytes.size()) {
			return false;
		}

		ByteReader itemReader(bytes);
		if (!itemReader.Skip(targetOffset)) {
			return false;
		}

		if (!itemReader.ReadI32(item.length) ||
			!itemReader.ReadI32(item.dataType) ||
			!itemReader.ReadI16(item.attr) ||
			!itemReader.ReadU8(item.arrayDimensions)) {
			return false;
		}

		item.arrayBounds.resize(item.arrayDimensions);
		for (std::uint8_t i = 0; i < item.arrayDimensions; ++i) {
			if (!itemReader.ReadI32(item.arrayBounds[i])) {
				return false;
			}
		}

		if (!itemReader.ReadStandardText(item.name) ||
			!itemReader.ReadStandardText(item.comment)) {
			return false;
		}
	}

	return true;
}

bool ParseProgramHeader(ByteReader& reader, ProgramHeaderInfo& outHeader)
{
	if (!reader.ReadI32(outHeader.versionFlag1) ||
		!reader.ReadI32(outHeader.unk1) ||
		!reader.ReadDynamicBytes(outHeader.unk2_1) ||
		!reader.ReadDynamicBytes(outHeader.unk2_2) ||
		!reader.ReadDynamicBytes(outHeader.unk2_3) ||
		!reader.ReadTextArray(outHeader.supportLibraryInfo) ||
		!reader.ReadI32(outHeader.flag1) ||
		!reader.ReadI32(outHeader.flag2)) {
		return false;
	}

	if ((outHeader.flag1 & 0x1) != 0) {
		if (!reader.ReadRaw(16, outHeader.unk3Op)) {
			return false;
		}
	}

	if (!reader.ReadDynamicBytes(outHeader.icon) ||
		!reader.ReadDynamicText(outHeader.debugCommandLine)) {
		return false;
	}

	return true;
}

bool ParseCodePage(ByteReader& reader, CodePageInfo& outPage)
{
	if (!reader.ReadI32(outPage.unk1) ||
		!reader.ReadI32(outPage.baseClass) ||
		!reader.ReadDynamicText(outPage.name) ||
		!reader.ReadDynamicText(outPage.comment)) {
		return false;
	}

	std::int32_t length = 0;
	if (!reader.ReadI32(length) || length < 0 || (length % 4) != 0) {
		return false;
	}

	const int count = length / 4;
	outPage.functionIds.resize(static_cast<size_t>(count));
	for (int i = 0; i < count; ++i) {
		if (!reader.ReadI32(outPage.functionIds[static_cast<size_t>(i)])) {
			return false;
		}
	}

	std::int32_t varCount = 0;
	std::vector<std::uint8_t> varBytes;
	if (!reader.ReadI32(varCount) ||
		!reader.ReadDynamicBytes(varBytes) ||
		!ParseVariableList(varCount, varBytes, outPage.pageVars)) {
		return false;
	}

	return true;
}

bool ParseFunction(ByteReader& reader, FunctionInfo& outFunction)
{
	if (!reader.ReadI32(outFunction.type) ||
		!reader.ReadI32(outFunction.attr) ||
		!reader.ReadI32(outFunction.returnType) ||
		!reader.ReadDynamicText(outFunction.name) ||
		!reader.ReadDynamicText(outFunction.comment)) {
		return false;
	}

	std::int32_t localCount = 0;
	std::int32_t paramCount = 0;
	std::vector<std::uint8_t> localBytes;
	std::vector<std::uint8_t> paramBytes;
	if (!reader.ReadI32(localCount) ||
		!reader.ReadDynamicBytes(localBytes) ||
		!ParseVariableList(localCount, localBytes, outFunction.locals) ||
		!reader.ReadI32(paramCount) ||
		!reader.ReadDynamicBytes(paramBytes) ||
		!ParseVariableList(paramCount, paramBytes, outFunction.params)) {
		return false;
	}

	for (int i = 0; i < 5; ++i) {
		if (!reader.SkipDynamicBytes()) {
			return false;
		}
	}

	return reader.ReadDynamicBytes(outFunction.code);
}

bool ParseDataType(ByteReader& reader, DataTypeInfo& outDataType)
{
	if (!reader.ReadI32(outDataType.attr) ||
		!reader.ReadDynamicText(outDataType.name) ||
		!reader.ReadDynamicText(outDataType.comment)) {
		return false;
	}

	std::int32_t memberCount = 0;
	std::vector<std::uint8_t> memberBytes;
	if (!reader.ReadI32(memberCount) ||
		!reader.ReadDynamicBytes(memberBytes) ||
		!ParseVariableList(memberCount, memberBytes, outDataType.members)) {
		return false;
	}

	return true;
}

bool ParseDll(ByteReader& reader, DllInfo& outDll)
{
	if (!reader.ReadI32(outDll.attr) ||
		!reader.ReadI32(outDll.returnType) ||
		!reader.ReadDynamicText(outDll.name) ||
		!reader.ReadDynamicText(outDll.comment) ||
		!reader.ReadDynamicText(outDll.fileName) ||
		!reader.ReadDynamicText(outDll.commandName)) {
		return false;
	}

	std::int32_t paramCount = 0;
	std::vector<std::uint8_t> paramBytes;
	if (!reader.ReadI32(paramCount) ||
		!reader.ReadDynamicBytes(paramBytes) ||
		!ParseVariableList(paramCount, paramBytes, outDll.params)) {
		return false;
	}

	return true;
}

bool ParseProgramSection(const std::vector<std::uint8_t>& bytes, ProgramSection& outProgram)
{
	ByteReader reader(bytes);
	if (!ParseProgramHeader(reader, outProgram.header)) {
		return false;
	}

	std::vector<BlockHeader> headers;
	if (!ReadBlockHeaders(reader, headers)) {
		return false;
	}
	outProgram.codePages.resize(headers.size());
	for (size_t i = 0; i < headers.size(); ++i) {
		outProgram.codePages[i].header = headers[i];
		if (!ParseCodePage(reader, outProgram.codePages[i])) {
			return false;
		}
	}

	if (!ReadBlockHeaders(reader, headers)) {
		return false;
	}
	outProgram.functions.resize(headers.size());
	for (size_t i = 0; i < headers.size(); ++i) {
		outProgram.functions[i].header = headers[i];
		if (!ParseFunction(reader, outProgram.functions[i])) {
			return false;
		}
	}

	std::int32_t globalCount = 0;
	std::vector<std::uint8_t> globalBytes;
	if (!reader.ReadI32(globalCount) ||
		!reader.ReadDynamicBytes(globalBytes) ||
		!ParseVariableList(globalCount, globalBytes, outProgram.globals)) {
		return false;
	}

	if (!ReadBlockHeaders(reader, headers)) {
		return false;
	}
	outProgram.dataTypes.resize(headers.size());
	for (size_t i = 0; i < headers.size(); ++i) {
		outProgram.dataTypes[i].header = headers[i];
		if (!ParseDataType(reader, outProgram.dataTypes[i])) {
			return false;
		}
	}

	if (!ReadBlockHeaders(reader, headers)) {
		return false;
	}
	outProgram.dlls.resize(headers.size());
	for (size_t i = 0; i < headers.size(); ++i) {
		outProgram.dlls[i].header = headers[i];
		if (!ParseDll(reader, outProgram.dlls[i])) {
			return false;
		}
	}

	return true;
}

bool ParseForm(ByteReader& reader, FormInfo& outForm)
{
	if (!reader.ReadI32(outForm.unknown1) ||
		!reader.ReadI32(outForm.unknown2) ||
		!reader.ReadDynamicText(outForm.name) ||
		!reader.ReadDynamicText(outForm.comment)) {
		return false;
	}

	std::int32_t elementCount = 0;
	if (!reader.ReadI32(elementCount)) {
		return false;
	}
	return reader.SkipDynamicBytes();
}

bool ParseConstants(
	std::int32_t count,
	const std::vector<std::uint8_t>& bytes,
	std::vector<ConstantInfo>& outConstants)
{
	outConstants.clear();
	if (count < 0) {
		return false;
	}
	if (count == 0) {
		return true;
	}

	ByteReader reader(bytes);
	outConstants.resize(static_cast<size_t>(count));
	for (std::int32_t i = 0; i < count; ++i) {
		if (!reader.ReadI32(outConstants[static_cast<size_t>(i)].marker)) {
			return false;
		}
	}
	for (std::int32_t i = 0; i < count; ++i) {
		if (!reader.ReadI32(outConstants[static_cast<size_t>(i)].offset)) {
			return false;
		}
	}

	const size_t bodyBase = reader.position();
	for (auto& item : outConstants) {
		const size_t targetOffset = bodyBase + static_cast<size_t>(item.offset);
		if (targetOffset > bytes.size()) {
			return false;
		}

		ByteReader itemReader(bytes);
		if (!itemReader.Skip(targetOffset) ||
			!itemReader.ReadI32(item.length) ||
			!itemReader.ReadI16(item.attr) ||
			!itemReader.ReadStandardText(item.name) ||
			!itemReader.ReadStandardText(item.comment)) {
			return false;
		}

		const std::uint32_t pageType = GetHighType(static_cast<std::uint32_t>(item.marker));
		if (pageType == kConstPageValue) {
			std::uint8_t valueType = 0;
			if (!itemReader.ReadU8(valueType)) {
				return false;
			}
			switch (valueType) {
			case kConstTypeEmpty:
				item.valueText.clear();
				item.resultText = "文本";
				break;
			case kConstTypeNumber: {
				double value = 0.0;
				if (!itemReader.ReadDouble(value)) {
					return false;
				}
				item.valueText = FormatDoubleLocal(value);
				item.resultText = "文本";
				break;
			}
			case kConstTypeBool: {
				bool value = false;
				if (!itemReader.ReadBool(value)) {
					return false;
				}
				item.valueText = value ? "真" : "假";
				item.resultText = "文本";
				break;
			}
			case kConstTypeDate: {
				double value = 0.0;
				if (!itemReader.ReadDouble(value)) {
					return false;
				}
				item.valueText = FormatDateLocal(value);
				item.resultText = "文本";
				break;
			}
			case kConstTypeText: {
				if (!itemReader.ReadDynamicText(item.valueText)) {
					return false;
				}
				std::int16_t trailingAttr = 0;
				if (!itemReader.ReadI16(trailingAttr)) {
					return false;
				}
				item.longText = (trailingAttr & kConstAttrLongText) != 0;
				item.resultText = "文本";
				break;
			}
			default:
				item.resultText = std::format("未知常量类型({})", valueType);
				break;
			}
		}
		else if (pageType == kConstPageImage) {
			item.resultText = "图片";
			item.valueText = "<图片>";
			if (!itemReader.SkipDynamicBytes()) {
				return false;
			}
		}
		else if (pageType == kConstPageSound) {
			item.resultText = "声音";
			item.valueText = "<声音>";
			if (!itemReader.SkipDynamicBytes()) {
				return false;
			}
		}
		else {
			item.resultText = std::format("未知常量页面({})", pageType);
		}
	}

	return true;
}

bool ParseResourceSection(const std::vector<std::uint8_t>& bytes, ResourceSection& outResource)
{
	ByteReader reader(bytes);

	std::vector<BlockHeader> headers;
	if (!ReadBlockHeaders(reader, headers)) {
		return false;
	}

	outResource.forms.resize(headers.size());
	for (size_t i = 0; i < headers.size(); ++i) {
		outResource.forms[i].header = headers[i];
		if (!ParseForm(reader, outResource.forms[i])) {
			return false;
		}
	}

	std::int32_t constCount = 0;
	std::vector<std::uint8_t> constBytes;
	if (!reader.ReadI32(constCount) ||
		!reader.ReadDynamicBytes(constBytes) ||
		!ParseConstants(constCount, constBytes, outResource.constants) ||
		!reader.ReadI32(outResource.reserve)) {
		return false;
	}

	return true;
}

bool ParseUserInfoSection(const std::vector<std::uint8_t>& bytes, UserInfoSection& outUser)
{
	ByteReader reader(bytes);
	return
		reader.ReadDynamicText(outUser.programName) &&
		reader.ReadDynamicText(outUser.programComment) &&
		reader.ReadDynamicText(outUser.author) &&
		reader.ReadDynamicText(outUser.zipCode) &&
		reader.ReadDynamicText(outUser.address) &&
		reader.ReadDynamicText(outUser.phone) &&
		reader.ReadDynamicText(outUser.fax) &&
		reader.ReadDynamicText(outUser.email) &&
		reader.ReadDynamicText(outUser.homePage) &&
		reader.ReadDynamicText(outUser.other) &&
		reader.ReadI32(outUser.version1) &&
		reader.ReadI32(outUser.version2);
}

bool ParseModuleSections(const std::string& modulePath, ModuleSections& outSections, std::string* outError)
{
	std::vector<std::uint8_t> bytes;
	if (!ReadFileBytes(modulePath, bytes)) {
		if (outError != nullptr) {
			*outError = "read_module_file_failed";
		}
		return false;
	}

	if (bytes.size() < sizeof(std::uint32_t) * 2) {
		if (outError != nullptr) {
			*outError = "module_file_too_small";
		}
		return false;
	}

	const std::uint32_t magic1 = *reinterpret_cast<const std::uint32_t*>(bytes.data());
	if (magic1 == kMagicEncryptedSource) {
		if (outError != nullptr) {
			*outError = "encrypted_module_not_supported";
		}
		return false;
	}

	ByteReader reader(bytes);
	std::uint32_t header1 = 0;
	std::uint32_t header2 = 0;
	if (!reader.ReadU32(header1) ||
		!reader.ReadU32(header2) ||
		header1 != kMagicFileHeader1 ||
		header2 != kMagicFileHeader2) {
		if (outError != nullptr) {
			*outError = "module_file_header_invalid";
		}
		return false;
	}

	while (!reader.eof()) {
		RawSectionHeader header = {};
		RawSectionInfo info = {};
		std::vector<std::uint8_t> sectionBytes;
		if (!reader.ReadRaw(sizeof(header), sectionBytes)) {
			break;
		}
		std::memcpy(&header, sectionBytes.data(), sizeof(header));
		if (header.magic != kMagicSection) {
			if (outError != nullptr) {
				*outError = "section_header_invalid";
			}
			return false;
		}

		if (!reader.ReadRaw(sizeof(info), sectionBytes)) {
			if (outError != nullptr) {
				*outError = "section_info_read_failed";
			}
			return false;
		}
		std::memcpy(&info, sectionBytes.data(), sizeof(info));

		if (info.dataLength < 0) {
			if (outError != nullptr) {
				*outError = "section_length_invalid";
			}
			return false;
		}

		std::string sectionName;
		if (!DecodeSectionName(info, sectionName)) {
			if (!reader.Skip(static_cast<size_t>(info.dataLength))) {
				return false;
			}
			continue;
		}

		if (!reader.ReadRaw(static_cast<size_t>(info.dataLength), sectionBytes)) {
			if (outError != nullptr) {
				*outError = "section_body_read_failed";
			}
			return false;
		}

		if (sectionName == "系统信息段") {
			if (sectionBytes.size() < sizeof(RawSystemInfoSection)) {
				if (outError != nullptr) {
					*outError = "system_info_section_too_small";
				}
				return false;
			}
			std::memcpy(&outSections.systemInfo, sectionBytes.data(), sizeof(RawSystemInfoSection));
			outSections.hasSystemInfo = true;
		}
		else if (sectionName == "用户信息段") {
			if (!ParseUserInfoSection(sectionBytes, outSections.userInfo)) {
				if (outError != nullptr) {
					*outError = "user_info_section_parse_failed";
				}
				return false;
			}
			outSections.hasUserInfo = true;
		}
		else if (sectionName == "程序段") {
			if (!ParseProgramSection(sectionBytes, outSections.program)) {
				if (outError != nullptr) {
					*outError = "program_section_parse_failed";
				}
				return false;
			}
			outSections.hasProgram = true;
		}
		else if (sectionName == "程序资源段") {
			if (!ParseResourceSection(sectionBytes, outSections.resources)) {
				if (outError != nullptr) {
					*outError = "resource_section_parse_failed";
				}
				return false;
			}
			outSections.hasResources = true;
		}
	}

	if (!outSections.hasProgram) {
		if (outError != nullptr) {
			*outError = "program_section_missing";
		}
		return false;
	}

	return true;
}

std::string GetBuiltinTypeName(std::int32_t typeValue)
{
	switch (typeValue) {
	case 0: return "";
	case static_cast<std::int32_t>(0x80000101u): return "字节型";
	case static_cast<std::int32_t>(0x80000201u): return "短整数型";
	case static_cast<std::int32_t>(0x80000301u): return "整数型";
	case static_cast<std::int32_t>(0x80000401u): return "长整数型";
	case static_cast<std::int32_t>(0x80000501u): return "小数型";
	case static_cast<std::int32_t>(0x80000601u): return "双精度小数";
	case static_cast<std::int32_t>(0x80000002u): return "逻辑型";
	case static_cast<std::int32_t>(0x80000003u): return "日期时间型";
	case static_cast<std::int32_t>(0x80000004u): return "文本型";
	case static_cast<std::int32_t>(0x80000005u): return "字节集";
	case static_cast<std::int32_t>(0x80000006u): return "子程序指针";
	case 65537: return "窗口";
	case 65539: return "菜单";
	case 65540: return "字体";
	case 65541: return "编辑框";
	case 65542: return "图片框";
	case 65543: return "外形框";
	case 65544: return "画板";
	case 65545: return "分组框";
	case 65546: return "标签";
	case 65547: return "按钮";
	case 65548: return "选择框";
	case 65549: return "单选框";
	case 65550: return "组合框";
	case 65551: return "列表框";
	case 65552: return "选择列表框";
	case 65553: return "横向滚动条";
	case 65554: return "纵向滚动条";
	case 65555: return "进度条";
	case 65556: return "滑块条";
	case 65557: return "选择夹";
	case 65558: return "影像框";
	case 65559: return "日期框";
	case 65560: return "月历";
	case 65561: return "驱动器框";
	case 65562: return "目录框";
	case 65563: return "文件框";
	case 65564: return "颜色选择器";
	case 65565: return "超级链接框";
	case 65566: return "调节器";
	case 65567: return "通用对话框";
	case 65568: return "时钟";
	case 65569: return "打印机";
	case 65570: return "字段信息";
	case 65572: return "数据报";
	case 65573: return "客户";
	case 65574: return "服务器";
	case 65575: return "端口";
	case 65576: return "打印设置信息";
	case 65577: return "表格";
	case 65578: return "数据源";
	case 65579: return "通用提供者";
	case 65580: return "数据库提供者";
	case 65581: return "图形按钮";
	case 65582: return "外部数据库";
	case 65583: return "外部数据提供者";
	case 65584: return "对象";
	case 65585: return "变体型";
	case 65586: return "变体类型";
	case 196611: return "工具条";
	case 196612: return "超级列表框";
	case 262145: return "高级表格";
	case -1: return "";
	default: return std::string();
	}
}

class TypeNameResolver {
public:
	explicit TypeNameResolver(const ProgramSection& program)
		: m_program(program)
	{
	}

	std::string Resolve(std::int32_t typeValue)
	{
		for (const auto& dataType : m_program.dataTypes) {
			if (dataType.header.dwId == typeValue || dataType.header.dwUnk == typeValue) {
				return dataType.name;
			}
		}
		for (const auto& page : m_program.codePages) {
			if (page.header.dwId == typeValue || page.header.dwUnk == typeValue) {
				return page.name;
			}
		}

		const std::string builtin = GetBuiltinTypeName(typeValue);
		if (!builtin.empty()) {
			return builtin;
		}

		return ResolveSupportLibraryType(typeValue);
	}

private:
	bool EnsureSupportLibraryCache(std::uint16_t supportIndex)
	{
		if (supportIndex == 0 || supportIndex == 1) {
			return false;
		}

		if (const auto it = m_supportTypeCache.find(supportIndex); it != m_supportTypeCache.end()) {
			return !it->second.empty();
		}

		if (supportIndex > m_program.header.supportLibraryInfo.size()) {
			m_supportTypeCache.emplace(supportIndex, std::vector<std::string>{});
			return false;
		}

		std::string fileName = GetFirstSupportLibraryToken(
			m_program.header.supportLibraryInfo[static_cast<size_t>(supportIndex - 1)]);
		if (fileName.empty()) {
			m_supportTypeCache.emplace(supportIndex, std::vector<std::string>{});
			return false;
		}
		if (!std::filesystem::path(fileName).has_extension()) {
			fileName += ".fne";
		}

		std::vector<std::filesystem::path> candidates;
		candidates.push_back(std::filesystem::path(GetBasePath()) / fileName);
		candidates.push_back(std::filesystem::path(GetBasePath()) / "lib" / fileName);

		HMODULE module = nullptr;
		for (const auto& path : candidates) {
			if (!std::filesystem::exists(path)) {
				continue;
			}
			module = LoadLibraryA(path.string().c_str());
			if (module != nullptr) {
				break;
			}
		}
		if (module == nullptr) {
			m_supportTypeCache.emplace(supportIndex, std::vector<std::string>{});
			return false;
		}

		std::vector<std::string> names;
		const auto* getInfoProc = reinterpret_cast<PFN_GET_LIB_INFO>(GetProcAddress(module, FUNCNAME_GET_LIB_INFO));
		if (getInfoProc != nullptr) {
			const auto* libInfo = getInfoProc();
			if (libInfo != nullptr && libInfo->m_nDataTypeCount > 0 && libInfo->m_pDataType != nullptr) {
				names.reserve(static_cast<size_t>(libInfo->m_nDataTypeCount));
				for (int i = 0; i < libInfo->m_nDataTypeCount; ++i) {
					names.emplace_back(libInfo->m_pDataType[i].m_szName == nullptr ? "" : libInfo->m_pDataType[i].m_szName);
				}
			}
		}
		FreeLibrary(module);
		m_supportTypeCache.emplace(supportIndex, names);
		return !names.empty();
	}

	std::string ResolveSupportLibraryType(std::int32_t typeValue)
	{
		const auto rawValue = static_cast<std::uint32_t>(typeValue);
		if ((rawValue & 0x80000000u) != 0) {
			return std::string();
		}

		const std::uint16_t supportIndex = static_cast<std::uint16_t>(rawValue >> 16);
		const std::uint16_t typeIndex = static_cast<std::uint16_t>(rawValue & 0xFFFFu);
		if (supportIndex == 0 || typeIndex == 0) {
			return std::string();
		}
		if (!EnsureSupportLibraryCache(supportIndex)) {
			return std::format("0x{:08X}", rawValue);
		}

		const auto it = m_supportTypeCache.find(supportIndex);
		if (it == m_supportTypeCache.end() || typeIndex > it->second.size()) {
			return std::format("0x{:08X}", rawValue);
		}
		const std::string& name = it->second[static_cast<size_t>(typeIndex - 1)];
		return name.empty() ? std::format("0x{:08X}", rawValue) : name;
	}

	const ProgramSection& m_program;
	std::unordered_map<std::uint16_t, std::vector<std::string>> m_supportTypeCache;
};

ModulePublicInfoParam BuildParam(const VariableInfo& info, TypeNameResolver& resolver)
{
	ModulePublicInfoParam param;
	param.name = TrimAsciiCopy(info.name);
	param.typeText = TrimAsciiCopy(resolver.Resolve(info.dataType));
	param.flagsText = BuildVarFlags(info);
	param.comment = TrimAsciiCopy(info.comment);
	return param;
}

std::string BuildRecordSignatureKeepingFields(
	const std::string& prefix,
	const std::vector<std::string>& fields,
	size_t minFieldCount);

std::string BuildRecordSignature(const std::string& prefix, const std::vector<std::string>& fields)
{
	return BuildRecordSignatureKeepingFields(prefix, fields, 0);
}

std::string BuildRecordSignatureKeepingFields(
	const std::string& prefix,
	const std::vector<std::string>& fields,
	size_t minFieldCount)
{
	size_t fieldCount = fields.size();
	while (fieldCount > minFieldCount && fields[fieldCount - 1].empty()) {
		--fieldCount;
	}

	std::string line = prefix;
	if (fieldCount > 0) {
		line += " ";
	}
	bool first = true;
	for (size_t i = 0; i < fieldCount; ++i) {
		const auto& field = fields[i];
		if (!first) {
			line += ", ";
		}
		line += field;
		first = false;
	}
	return line;
}

ModulePublicInfoRecord BuildAssemblyRecord(const PublicAssemblyInfo& info)
{
	ModulePublicInfoRecord record;
	record.tag = kTagAssembly;
	record.kind = "assembly";
	record.name = info.page == nullptr ? std::string() : TrimAsciiCopy(info.page->name);
	record.comment = info.page == nullptr ? std::string() : TrimAsciiCopy(info.page->comment);
	record.flagsText = "公开";
	if (info.page != nullptr) {
		record.headerInts = { info.page->header.dwId, info.page->header.dwUnk };
	}
	record.signatureText = BuildRecordSignature(".程序集", { record.name, "", record.flagsText, record.comment });
	PushUniqueString(record.extractedStrings, record.name);
	PushUniqueString(record.extractedStrings, record.comment);
	return record;
}

ModulePublicInfoRecord BuildClassRecord(const PublicClassInfo& info, TypeNameResolver& resolver)
{
	ModulePublicInfoRecord record;
	record.tag = kTagClass;
	record.kind = "class";
	record.name = info.page == nullptr ? std::string() : TrimAsciiCopy(info.page->name);
	record.comment = info.page == nullptr ? std::string() : TrimAsciiCopy(info.page->comment);
	record.typeText = (info.page == nullptr) ? std::string() : TrimAsciiCopy(resolver.Resolve(info.page->baseClass));
	record.flagsText = "公开";
	if (info.page != nullptr) {
		record.headerInts = { info.page->header.dwId, info.page->header.dwUnk, info.page->baseClass };
	}
	record.signatureText = BuildRecordSignature(".类", { record.name, record.typeText, record.flagsText, record.comment });
	PushUniqueString(record.extractedStrings, record.name);
	PushUniqueString(record.extractedStrings, record.typeText);
	PushUniqueString(record.extractedStrings, record.comment);
	return record;
}

ModulePublicInfoRecord BuildFunctionRecord(
	const FunctionInfo& info,
	TypeNameResolver& resolver,
	int tag,
	const std::string& kind,
	const std::string& ownerName)
{
	ModulePublicInfoRecord record;
	record.tag = tag;
	record.kind = kind;
	record.name = TrimAsciiCopy(info.name);
	record.comment = TrimAsciiCopy(info.comment);
	record.typeText = TrimAsciiCopy(resolver.Resolve(info.returnType));
	record.flagsText = "公开";
	record.headerInts = { info.header.dwId, info.header.dwUnk, info.attr, info.type };
	record.signatureText = BuildRecordSignature(".子程序", { record.name, record.typeText, record.flagsText, record.comment });
	for (const auto& paramInfo : info.params) {
		record.params.push_back(BuildParam(paramInfo, resolver));
	}
	PushUniqueString(record.extractedStrings, ownerName);
	PushUniqueString(record.extractedStrings, record.name);
	PushUniqueString(record.extractedStrings, record.typeText);
	PushUniqueString(record.extractedStrings, record.comment);
	for (const auto& param : record.params) {
		PushUniqueString(record.extractedStrings, param.name);
		PushUniqueString(record.extractedStrings, param.typeText);
		PushUniqueString(record.extractedStrings, param.comment);
	}
	return record;
}

ModulePublicInfoRecord BuildDataTypeRecord(const DataTypeInfo& info, TypeNameResolver& resolver)
{
	ModulePublicInfoRecord record;
	record.tag = kTagDataType;
	record.kind = "data_type";
	record.name = TrimAsciiCopy(info.name);
	record.comment = TrimAsciiCopy(info.comment);
	record.flagsText = "公开";
	record.headerInts = { info.header.dwId, info.header.dwUnk, info.attr };
	record.signatureText = BuildRecordSignature(".数据类型", { record.name, record.flagsText, record.comment });
	PushUniqueString(record.extractedStrings, record.name);
	PushUniqueString(record.extractedStrings, record.comment);
	return record;
}

ModulePublicInfoRecord BuildDllRecord(const DllInfo& info, TypeNameResolver& resolver)
{
	ModulePublicInfoRecord record;
	record.tag = kTagDll;
	record.kind = "dll_command";
	record.name = TrimAsciiCopy(info.name);
	record.comment = TrimAsciiCopy(info.comment);
	record.typeText = TrimAsciiCopy(resolver.Resolve(info.returnType));
	record.flagsText = "公开";
	record.headerInts = { info.header.dwId, info.header.dwUnk, info.attr };
	record.signatureText = BuildRecordSignature(
		".DLL命令",
		{ record.name, record.typeText, QuoteIfNotEmpty(info.fileName), QuoteIfNotEmpty(info.commandName), record.flagsText, record.comment });
	for (const auto& paramInfo : info.params) {
		record.params.push_back(BuildParam(paramInfo, resolver));
	}
	PushUniqueString(record.extractedStrings, record.name);
	PushUniqueString(record.extractedStrings, info.fileName);
	PushUniqueString(record.extractedStrings, info.commandName);
	PushUniqueString(record.extractedStrings, record.typeText);
	PushUniqueString(record.extractedStrings, record.comment);
	for (const auto& param : record.params) {
		PushUniqueString(record.extractedStrings, param.name);
		PushUniqueString(record.extractedStrings, param.typeText);
		PushUniqueString(record.extractedStrings, param.comment);
	}
	return record;
}

ModulePublicInfoRecord BuildConstantRecord(const ConstantInfo& info)
{
	ModulePublicInfoRecord record;
	record.tag = kTagConst;
	record.kind = "const";
	record.name = TrimAsciiCopy(info.name);
	record.comment = TrimAsciiCopy(info.comment);
	record.typeText = "\"" + info.valueText + "\"";
	record.flagsText = "公开";
	record.signatureText = BuildRecordSignature(
		".常量",
		{ record.name, "\"" + info.valueText + "\"", record.flagsText, record.comment });
	record.headerInts = { info.marker, info.attr, info.length };
	PushUniqueString(record.extractedStrings, record.name);
	PushUniqueString(record.extractedStrings, info.valueText);
	PushUniqueString(record.extractedStrings, info.resultText);
	PushUniqueString(record.extractedStrings, record.comment);
	return record;
}

ModulePublicInfoRecord BuildGlobalVarRecord(const VariableInfo& info, TypeNameResolver& resolver)
{
	ModulePublicInfoRecord record;
	record.tag = kTagGlobalVar;
	record.kind = "global_var";
	record.name = TrimAsciiCopy(info.name);
	record.comment = TrimAsciiCopy(info.comment);
	record.typeText = TrimAsciiCopy(resolver.Resolve(info.dataType));
	record.flagsText = BuildVarFlags(info);
	record.signatureText = BuildRecordSignature(
		".全局变量",
		{ record.name, record.typeText, record.flagsText, record.comment });
	record.headerInts = { info.marker, info.attr, info.dataType };
	PushUniqueString(record.extractedStrings, record.name);
	PushUniqueString(record.extractedStrings, record.typeText);
	PushUniqueString(record.extractedStrings, record.flagsText);
	PushUniqueString(record.extractedStrings, record.comment);
	return record;
}

ModulePublicInfoRecord BuildWindowRecord(const FormInfo& info)
{
	ModulePublicInfoRecord record;
	record.tag = kTagWindow;
	record.kind = "window";
	record.name = TrimAsciiCopy(info.name);
	record.comment = TrimAsciiCopy(info.comment);
	record.flagsText = "公开";
	record.signatureText = BuildRecordSignature(".窗口", { record.name, record.flagsText, record.comment });
	record.headerInts = { info.header.dwId, info.header.dwUnk };
	PushUniqueString(record.extractedStrings, record.name);
	PushUniqueString(record.extractedStrings, record.comment);
	return record;
}

void AppendRecordLines(const ModulePublicInfoRecord& record, std::vector<std::string>& outLines, const char* childPrefix)
{
	if (!record.signatureText.empty()) {
		outLines.push_back(record.signatureText);
	}
	for (const auto& param : record.params) {
		std::string line = childPrefix;
		line += param.name;
		if (!param.typeText.empty()) {
			line += ", " + param.typeText;
		}
		else {
			line += ", ";
		}
		if (!param.flagsText.empty()) {
			line += ", " + param.flagsText;
		}
		else {
			line += ", ";
		}
		if (!param.comment.empty()) {
			line += ", " + param.comment;
		}
		outLines.push_back(std::move(line));
	}
	outLines.push_back("");
}

void AppendDataTypeLines(const DataTypeInfo& info, TypeNameResolver& resolver, std::vector<std::string>& outLines)
{
	const std::string dataTypeLine = BuildRecordSignature(
		".数据类型",
		{ TrimAsciiCopy(info.name), "公开", TrimAsciiCopy(info.comment) });
	outLines.push_back(dataTypeLine);

	for (const auto& member : info.members) {
		const std::string line = "    " + BuildRecordSignatureKeepingFields(
			".成员",
			{
				TrimAsciiCopy(member.name),
				TrimAsciiCopy(resolver.Resolve(member.dataType)),
				BuildMemberFlags(member),
				BuildMemberExtra(member),
				TrimAsciiCopy(member.comment),
			},
			2);
		outLines.push_back(line);
	}

	outLines.push_back("");
}

void ParseFormattedPublicInfoText(const std::string& formattedText, ModulePublicInfoDump& outDump)
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
			return TrimAsciiCopy(line.substr(pos + std::strlen(fullWidthSep)));
		}
		pos = line.find(':');
		if (pos != std::string::npos) {
			return TrimAsciiCopy(line.substr(pos + 1));
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
		const std::string line = TrimAsciiCopy(lines[i]);
		if (line.empty()) {
			continue;
		}

		if (!inCodeSection) {
			if (startsWith(line, "模块名称：") || startsWith(line, "module_name:")) {
				outDump.moduleName = valueAfterSep(line, "模块名称：");
				continue;
			}
			if (startsWith(line, "版本：") || startsWith(line, "version:")) {
				outDump.versionText = valueAfterSep(line, "版本：");
				continue;
			}
			if (line == "------------------------------") {
				continue;
			}
			if (startsWith(line, ".版本") || startsWith(line, ".version")) {
				inCodeSection = true;
				continue;
			}
			if (startsWith(line, "@备注") || startsWith(line, "@note")) {
				for (size_t j = i + 1; j < lines.size(); ++j) {
					const std::string nextLine = TrimAsciiCopy(lines[j]);
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
			currentRecord.kind = "assembly";
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
			currentRecord.kind = "sub";
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
			currentRecord.kind = "dll_command";
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
			currentRecord.kind = "const";
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
			currentRecord.kind = "declaration";
			currentRecord.name = line;
		}

		PushRecordStrings(currentRecord, currentRecord.name);
		PushRecordStrings(currentRecord, currentRecord.comment);
	}

	flushRecord();
}

void BuildPublicItems(const ModuleSections& sections, ModulePublicInfoDump& outDump)
{
	TypeNameResolver resolver(sections.program);

	std::vector<PublicAssemblyInfo> assemblies(sections.program.codePages.size() + 1);
	for (size_t i = 0; i < sections.program.codePages.size(); ++i) {
		assemblies[i].page = const_cast<CodePageInfo*>(&sections.program.codePages[i]);
	}

	for (const auto& function : sections.program.functions) {
		if (function.name.empty() || function.name.starts_with("_") || (function.attr & kFuncAttrPublic) == 0) {
			continue;
		}

		bool attached = false;
		for (size_t i = 0; i < sections.program.codePages.size() && !attached; ++i) {
			const auto& page = sections.program.codePages[i];
			if ((std::find)(page.functionIds.begin(), page.functionIds.end(), function.header.dwId) == page.functionIds.end()) {
				continue;
			}

			if (page.baseClass == 0) {
				assemblies[i].functions.push_back(const_cast<FunctionInfo*>(&function));
				attached = true;
			}
			else if (!page.name.empty() && (function.attr & kFuncAttrPublic) != 0) {
				assemblies[i].functions.push_back(const_cast<FunctionInfo*>(&function));
				attached = true;
			}
		}

	}

	std::vector<std::string> formattedLines;
	formattedLines.push_back(std::format("模块名称：{}", std::filesystem::path(outDump.modulePath).stem().string()));
	if (sections.hasUserInfo) {
		formattedLines.push_back(std::format("版本：{}.{}", sections.userInfo.version1, sections.userInfo.version2));
	}
	else if (sections.hasSystemInfo) {
		formattedLines.push_back(std::format("版本：{}.{}", sections.systemInfo.compileMajor, sections.systemInfo.compileMinor));
	}
	formattedLines.push_back("");

	std::string firstAssemblyName;
	std::string firstAssemblyComment;
	for (const auto& assembly : assemblies) {
		if (assembly.page == nullptr || assembly.page->name.empty()) {
			continue;
		}
		firstAssemblyName = TrimAsciiCopy(assembly.page->name);
		firstAssemblyComment = TrimAsciiCopy(assembly.page->comment);
		break;
	}
	if (!firstAssemblyName.empty()) {
		formattedLines.push_back(firstAssemblyName);
		formattedLines.push_back("@备注:");
		formattedLines.push_back(firstAssemblyComment);
		formattedLines.push_back("");
		outDump.assemblyName = firstAssemblyName;
		outDump.assemblyComment = firstAssemblyComment;
	}

	formattedLines.push_back("------------------------------");
	formattedLines.push_back("");
	formattedLines.push_back(".版本 2");
	formattedLines.push_back("");

	for (const auto& assembly : assemblies) {
		if (assembly.functions.empty()) {
			continue;
		}

		const bool hasNamedAssembly = assembly.page != nullptr && !assembly.page->name.empty();
		std::string assemblyName = hasNamedAssembly
			? TrimAsciiCopy(assembly.page->name)
			: "无主程序集";
		std::string assemblyComment = assembly.page == nullptr ? std::string() : TrimAsciiCopy(assembly.page->comment);

		if (hasNamedAssembly) {
			const auto assemblyRecord = BuildAssemblyRecord(assembly);
			AppendRecordLines(assemblyRecord, formattedLines, ".参数 ");
		}

		for (const auto* function : assembly.functions) {
			if (function == nullptr) {
				continue;
			}
			const auto functionRecord = BuildFunctionRecord(*function, resolver, kTagSub, "sub", hasNamedAssembly ? assemblyName : std::string());
			AppendRecordLines(functionRecord, formattedLines, ".参数 ");
		}
	}

	for (const auto& dataType : sections.program.dataTypes) {
		if (dataType.name.empty()) {
			continue;
		}
		AppendDataTypeLines(dataType, resolver, formattedLines);
	}

	for (const auto& dll : sections.program.dlls) {
		if (dll.name.empty()) {
			continue;
		}
		const auto dllRecord = BuildDllRecord(dll, resolver);
		AppendRecordLines(dllRecord, formattedLines, ".参数 ");
	}

	for (const auto& constant : sections.resources.constants) {
		if (constant.name.empty()) {
			continue;
		}
		const auto constantRecord = BuildConstantRecord(constant);
		AppendRecordLines(constantRecord, formattedLines, ".参数 ");
	}

	for (const auto& global : sections.program.globals) {
		if (global.name.empty()) {
			continue;
		}
		const auto globalRecord = BuildGlobalVarRecord(global, resolver);
		AppendRecordLines(globalRecord, formattedLines, ".参数 ");
	}

	for (const auto& form : sections.resources.forms) {
		if (form.name.empty()) {
			continue;
		}
		const auto windowRecord = BuildWindowRecord(form);
		AppendRecordLines(windowRecord, formattedLines, ".参数 ");
	}

	outDump.formattedText = JoinLines(formattedLines);
	outDump.records.clear();
	ParseFormattedPublicInfoText(outDump.formattedText, outDump);
	if (outDump.versionText.empty()) {
		if (sections.hasUserInfo) {
			outDump.versionText = std::format("{}.{}", sections.userInfo.version1, sections.userInfo.version2);
		}
		else if (sections.hasSystemInfo) {
			outDump.versionText = std::format("{}.{}", sections.systemInfo.compileMajor, sections.systemInfo.compileMinor);
		}
	}
}

}  // namespace

bool EcModulePublicInfoReader::Load(
	const std::string& modulePath,
	ModulePublicInfoDump& outDump,
	std::string* outError) const
{
	if (outError != nullptr) {
		outError->clear();
	}

	ModuleSections sections;
	if (!ParseModuleSections(modulePath, sections, outError)) {
		return false;
	}

	outDump = {};
	outDump.modulePath = modulePath;
	outDump.nativeResult = 1;
	outDump.sourceKind = "local_ec_structured";
	BuildPublicItems(sections, outDump);
	outDump.trace = std::format(
		"parser=local_ec_structured code_pages={} functions={} data_types={} dlls={} constants={} globals={} windows={}",
		sections.program.codePages.size(),
		sections.program.functions.size(),
		sections.program.dataTypes.size(),
		sections.program.dlls.size(),
		sections.resources.constants.size(),
		sections.program.globals.size(),
		sections.resources.forms.size());

	if (outDump.records.empty()) {
		if (outError != nullptr) {
			*outError = "local_ec_structured_no_public_items";
		}
		return false;
	}

	return true;
}

}  // namespace e571
