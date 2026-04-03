#include "e2txt.h"

#include <Windows.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <cstdlib>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <limits>
#include <memory>
#include <optional>
#include <sstream>
#include <stdexcept>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include <lib2.h>

#include "PathHelper.h"

namespace e2txt {

namespace {

constexpr std::uint32_t kMagicEncryptedSource = 1162630231u;
constexpr std::uint32_t kMagicFileHeader1 = 1415007811u;
constexpr std::uint32_t kMagicFileHeader2 = 1196576837u;
constexpr std::uint32_t kMagicSection = 353465113u;
constexpr std::array<std::uint8_t, 4> kSectionNameNoKey = { 25, 115, 0, 7 };

constexpr std::int32_t kVarAttrByRef = 0x0002;
constexpr std::int32_t kVarAttrNullable = 0x0004;
constexpr std::int32_t kVarAttrArray = 0x0008;
constexpr std::int32_t kConstPageValue = 1;
constexpr std::int32_t kConstPageImage = 2;
constexpr std::int32_t kConstPageSound = 3;
constexpr std::uint8_t kConstTypeEmpty = 22;
constexpr std::uint8_t kConstTypeNumber = 23;
constexpr std::uint8_t kConstTypeBool = 24;
constexpr std::uint8_t kConstTypeDate = 25;
constexpr std::uint8_t kConstTypeText = 26;
constexpr std::int16_t kConstAttrLongText = 0x0010;

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
	std::int32_t ownerClass = 0;
	std::int32_t attr = 0;
	std::int32_t returnType = 0;
	std::string name;
	std::string comment;
	std::vector<VariableInfo> locals;
	std::vector<VariableInfo> params;
	std::vector<std::uint8_t> lineOffset;
	std::vector<std::uint8_t> blockOffset;
	std::vector<std::uint8_t> methodReference;
	std::vector<std::uint8_t> variableReference;
	std::vector<std::uint8_t> constantReference;
	std::vector<std::uint8_t> expressionData;
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

struct ClassPublicityInfo {
	std::int32_t classId = 0;
	std::int32_t flags = 0;
};

struct FormInfo {
	BlockHeader header;
	std::int32_t unknown1 = 0;
	std::int32_t unknown2 = 0;
	std::string name;
	std::string comment;
	struct ElementInfo {
		std::int32_t id = 0;
		std::int32_t dataType = 0;
		bool isMenu = false;
		std::string name;
	};
	std::vector<ElementInfo> elements;
};

struct ConstantInfo {
	std::int32_t marker = 0;
	std::int32_t offset = 0;
	std::int32_t length = 0;
	std::int16_t attr = 0;
	std::uint32_t pageType = 0;
	std::string name;
	std::string comment;
	std::string valueText;
	bool longText = false;
	size_t textByteLength = 0;
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
	bool hasClassPublicity = false;
	RawSystemInfoSection systemInfo = {};
	UserInfoSection userInfo;
	ProgramSection program;
	ResourceSection resources;
	std::vector<ClassPublicityInfo> classPublicities;
	std::vector<std::uint8_t> ecomSectionBytes;
};

std::string FormatNumberLiteral(double value);
std::string FormatDateLiteral(double value);

struct EComDependencyRecord {
	std::string name;
	std::string path;
	bool reExport = false;
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
	if (size < 0) {
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
		if (static_cast<unsigned char>(text[begin]) > 0x20) {
			break;
		}
		++begin;
	}

	size_t end = text.size();
	while (end > begin) {
		if (static_cast<unsigned char>(text[end - 1]) > 0x20) {
			break;
		}
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string RemoveTrailingNulls(std::string text)
{
	while (!text.empty() && text.back() == '\0') {
		text.pop_back();
	}
	return text;
}

std::string BytesToLocalText(const std::vector<std::uint8_t>& bytes)
{
	if (bytes.empty()) {
		return std::string();
	}
	return std::string(reinterpret_cast<const char*>(bytes.data()), bytes.size());
}

std::string JoinStrings(const std::vector<std::string>& parts, const char* separator)
{
	std::string out;
	for (const auto& part : parts) {
		if (part.empty()) {
			continue;
		}
		if (!out.empty()) {
			out += separator;
		}
		out += part;
	}
	return out;
}

std::string Quote(const std::string& text)
{
	return "\"" + text + "\"";
}

std::string QuoteIfNotEmpty(const std::string& text)
{
	const std::string trimmed = TrimAsciiCopy(text);
	return trimmed.empty() ? std::string() : Quote(trimmed);
}

std::string BuildDefinitionLine(
	const std::string& type,
	const std::vector<std::string>& rawItems,
	int indent = 0)
{
	std::ostringstream stream;
	for (int i = 0; i < indent; ++i) {
		stream << "    ";
	}
	stream << "." << type;

	size_t count = rawItems.size();
	while (count > 0 && rawItems[count - 1].empty()) {
		--count;
	}
	if (count == 0) {
		return stream.str();
	}

	stream << " ";
	for (size_t i = 0; i < count; ++i) {
		if (i != 0) {
			stream << ", ";
		}
		stream << rawItems[i];
	}
	return stream.str();
}

std::uint32_t GetHighType(std::uint32_t value)
{
	return (value & 0xF0000000u) >> 28;
}

std::vector<std::string> SplitByCarriageLines(const std::string& text)
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
	const auto parts = SplitByCarriageLines(rawText);
	return parts.empty() ? TrimAsciiCopy(rawText) : TrimAsciiCopy(parts.front());
}

std::vector<std::string> SplitSupportLibraryTokens(const std::string& rawText)
{
	std::vector<std::string> parts = SplitByCarriageLines(rawText);
	for (auto& part : parts) {
		part = TrimAsciiCopy(std::move(part));
	}
	parts.erase(
		std::remove_if(parts.begin(), parts.end(), [](const std::string& value) { return value.empty(); }),
		parts.end());
	return parts;
}

void TraceLine(const std::string& text)
{
	std::ofstream out("e2txt_trace.log", std::ios::app | std::ios::binary);
	if (!out.is_open()) {
		return;
	}
	out << text << "\r\n";
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

	const std::vector<std::uint8_t>& bytes() const
	{
		return m_bytes;
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
		out.assign(
			m_bytes.begin() + static_cast<std::ptrdiff_t>(m_pos),
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

	bool SetPosition(size_t pos)
	{
		if (pos > m_bytes.size()) {
			return false;
		}
		m_pos = pos;
		return true;
	}

	bool PeekU8(std::uint8_t& value) const
	{
		if (m_pos + sizeof(value) > m_bytes.size()) {
			return false;
		}
		value = m_bytes[m_pos];
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

	bool ReadBStr(std::string& out, bool& isNull)
	{
		isNull = false;
		std::int32_t size = 0;
		if (!ReadI32(size) || size < 0) {
			return false;
		}
		if (size == 0) {
			out.clear();
			isNull = true;
			return true;
		}
		if (size < 1 || m_pos + static_cast<size_t>(size) > m_bytes.size()) {
			return false;
		}
		out.assign(reinterpret_cast<const char*>(m_bytes.data() + m_pos), static_cast<size_t>(size - 1));
		m_pos += static_cast<size_t>(size);
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
		out.assign(reinterpret_cast<const char*>(m_bytes.data() + start), m_pos - start);
		++m_pos;
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

std::string BuildReaderHexContext(const ByteReader& reader, size_t centerPos)
{
	const auto& bytes = reader.bytes();
	if (bytes.empty()) {
		return std::string();
	}

	const size_t begin = centerPos > 8 ? centerPos - 8 : 0;
	const size_t end = (std::min)(bytes.size(), centerPos + 8);
	std::ostringstream stream;
	stream << std::hex << std::uppercase << std::setfill('0');
	for (size_t i = begin; i < end; ++i) {
		if (i != begin) {
			stream << ' ';
		}
		stream << std::setw(2) << static_cast<unsigned int>(bytes[i]);
	}
	return stream.str();
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

bool ParseVariableList(std::int32_t count, const std::vector<std::uint8_t>& bytes, std::vector<VariableInfo>& outVars)
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

	return reader.ReadDynamicBytes(outHeader.icon) && reader.ReadDynamicText(outHeader.debugCommandLine);
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
	return reader.ReadI32(varCount) &&
		reader.ReadDynamicBytes(varBytes) &&
		ParseVariableList(varCount, varBytes, outPage.pageVars);
}

bool ParseFunction(ByteReader& reader, FunctionInfo& outFunction)
{
	if (!reader.ReadI32(outFunction.ownerClass) ||
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

	return reader.ReadDynamicBytes(outFunction.lineOffset) &&
		reader.ReadDynamicBytes(outFunction.blockOffset) &&
		reader.ReadDynamicBytes(outFunction.methodReference) &&
		reader.ReadDynamicBytes(outFunction.variableReference) &&
		reader.ReadDynamicBytes(outFunction.constantReference) &&
		reader.ReadDynamicBytes(outFunction.expressionData);
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
	return reader.ReadI32(memberCount) &&
		reader.ReadDynamicBytes(memberBytes) &&
		ParseVariableList(memberCount, memberBytes, outDataType.members);
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
	return reader.ReadI32(paramCount) &&
		reader.ReadDynamicBytes(paramBytes) &&
		ParseVariableList(paramCount, paramBytes, outDll.params);
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

bool ParseFormElements(ByteReader& reader, std::vector<FormInfo::ElementInfo>& outElements)
{
	outElements.clear();

	std::int32_t count = 0;
	std::int32_t size = 0;
	if (!reader.ReadI32(count) || !reader.ReadI32(size) || count < 0 || size < 0) {
		return false;
	}

	const size_t blockStart = reader.position();
	const size_t blockEnd = blockStart + static_cast<size_t>(size);
	if (blockEnd > reader.bytes().size()) {
		return false;
	}

	std::vector<std::int32_t> ids(static_cast<size_t>(count));
	std::vector<std::int32_t> offsets(static_cast<size_t>(count));
	for (std::int32_t i = 0; i < count; ++i) {
		if (!reader.ReadI32(ids[static_cast<size_t>(i)])) {
			return false;
		}
	}
	for (std::int32_t i = 0; i < count; ++i) {
		if (!reader.ReadI32(offsets[static_cast<size_t>(i)])) {
			return false;
		}
	}

	const size_t payloadStart = reader.position();
	if (blockEnd < payloadStart) {
		return false;
	}

	outElements.resize(static_cast<size_t>(count));
	for (std::int32_t i = 0; i < count; ++i) {
		const size_t index = static_cast<size_t>(i);
		const size_t itemPos = payloadStart + static_cast<size_t>(offsets[index]);
		if (itemPos + sizeof(std::int32_t) > blockEnd) {
			return false;
		}

		ByteReader itemReader(reader.bytes());
		if (!itemReader.SetPosition(itemPos)) {
			return false;
		}

		std::int32_t itemLength = 0;
		if (!itemReader.ReadI32(itemLength) || itemLength < 4) {
			return false;
		}
		const size_t itemEnd = itemPos + sizeof(std::int32_t) + static_cast<size_t>(itemLength);
		if (itemEnd > blockEnd) {
			return false;
		}

		auto& element = outElements[index];
		element.id = ids[index];
		if (!itemReader.ReadI32(element.dataType) || !itemReader.Skip(20) || !itemReader.ReadStandardText(element.name)) {
			return false;
		}
		element.isMenu = element.dataType == 65539;
	}

	return reader.SetPosition(blockEnd);
}

bool ParseForm(ByteReader& reader, FormInfo& outForm)
{
	return reader.ReadI32(outForm.unknown1) &&
		reader.ReadI32(outForm.unknown2) &&
		reader.ReadDynamicText(outForm.name) &&
		reader.ReadDynamicText(outForm.comment) &&
		ParseFormElements(reader, outForm.elements);
}

bool ParseConstants(std::int32_t count, const std::vector<std::uint8_t>& bytes, std::vector<ConstantInfo>& outConstants)
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

		item.pageType = GetHighType(static_cast<std::uint32_t>(item.marker));
		if (item.pageType == kConstPageValue) {
			std::uint8_t valueType = 0;
			if (!itemReader.ReadU8(valueType)) {
				return false;
			}
			item.longText = (item.attr & kConstAttrLongText) != 0;
			switch (valueType) {
			case kConstTypeEmpty:
				item.valueText.clear();
				break;
			case kConstTypeNumber: {
				double value = 0.0;
				if (!itemReader.ReadDouble(value)) {
					return false;
				}
				item.valueText = FormatNumberLiteral(value);
				break;
			}
			case kConstTypeBool: {
				bool value = false;
				if (!itemReader.ReadBool(value)) {
					return false;
				}
				item.valueText = value ? "真" : "假";
				break;
			}
			case kConstTypeDate: {
				double value = 0.0;
				if (!itemReader.ReadDouble(value)) {
					return false;
				}
				item.valueText = FormatDateLiteral(value);
				break;
			}
			case kConstTypeText: {
				bool isNull = false;
				std::string rawText;
				if (!itemReader.ReadBStr(rawText, isNull)) {
					return false;
				}
				item.textByteLength = rawText.size();
				if (isNull) {
					item.valueText.clear();
				}
				else {
					item.valueText = "“" + rawText + "”";
				}
				break;
			}
			default:
				item.valueText = "<未知常量>";
				break;
			}
		}
		else if (item.pageType == kConstPageImage) {
			item.valueText = "<图片>";
			if (!itemReader.SkipDynamicBytes()) {
				return false;
			}
		}
		else if (item.pageType == kConstPageSound) {
			item.valueText = "<声音>";
			if (!itemReader.SkipDynamicBytes()) {
				return false;
			}
		}
		else {
			item.valueText = "<未知资源>";
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
	return reader.ReadI32(constCount) &&
		reader.ReadDynamicBytes(constBytes) &&
		ParseConstants(constCount, constBytes, outResource.constants) &&
		reader.ReadI32(outResource.reserve);
}

bool ParseClassPublicitySection(const std::vector<std::uint8_t>& bytes, std::vector<ClassPublicityInfo>& outItems)
{
	outItems.clear();
	if ((bytes.size() % 8u) != 0u) {
		return false;
	}

	ByteReader reader(bytes);
	while (!reader.eof()) {
		ClassPublicityInfo item;
		if (!reader.ReadI32(item.classId) || !reader.ReadI32(item.flags)) {
			return false;
		}
		outItems.push_back(item);
	}
	return true;
}

bool ParseUserInfoSection(const std::vector<std::uint8_t>& bytes, UserInfoSection& outUser)
{
	ByteReader reader(bytes);
	return reader.ReadDynamicText(outUser.programName) &&
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
			*outError = "encrypted_source_not_supported";
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
		else if (sectionName == "辅助信息段2") {
			if (!ParseClassPublicitySection(sectionBytes, outSections.classPublicities)) {
				if (outError != nullptr) {
					*outError = "class_publicity_section_parse_failed";
				}
				return false;
			}
			outSections.hasClassPublicity = true;
		}
		else if (sectionName == "易模块记录段") {
			outSections.ecomSectionBytes = std::move(sectionBytes);
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

bool IsLikelyReadableString(const std::string& text)
{
	const std::string trimmed = TrimAsciiCopy(text);
	if (trimmed.empty()) {
		return false;
	}
	for (const unsigned char ch : trimmed) {
		if (ch == '\r' || ch == '\n' || ch == '\t') {
			continue;
		}
		if (ch < 0x20 || ch == 0x7F) {
			return false;
		}
	}
	return true;
}

std::vector<std::string> ExtractReadableNullStrings(const std::vector<std::uint8_t>& bytes)
{
	std::vector<std::string> out;
	size_t i = 0;
	while (i < bytes.size()) {
		if (bytes[i] == 0) {
			++i;
			continue;
		}
		size_t j = i;
		while (j < bytes.size() && bytes[j] != 0) {
			++j;
		}
		if (j > i) {
			std::string text(reinterpret_cast<const char*>(bytes.data() + i), j - i);
			text = TrimAsciiCopy(text);
			if (IsLikelyReadableString(text)) {
				out.push_back(std::move(text));
			}
		}
		i = (j < bytes.size()) ? (j + 1) : j;
	}
	return out;
}

bool ParseEComDependencies(const std::vector<std::uint8_t>& bytes, std::vector<EComDependencyRecord>& outRecords)
{
	outRecords.clear();
	if (bytes.empty()) {
		return true;
	}

	const auto strings = ExtractReadableNullStrings(bytes);
	for (size_t i = 0; i < strings.size(); ++i) {
		const std::string& current = strings[i];
		if (current.size() < 3) {
			continue;
		}
		std::string lower = current;
		std::transform(
			lower.begin(),
			lower.end(),
			lower.begin(),
			[](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
		if (!(lower.ends_with(".ec") || lower.ends_with(".ec\""))) {
			continue;
		}

		EComDependencyRecord record;
		record.path = current;
		if (i > 0) {
			record.name = strings[i - 1];
			if (!record.name.empty() && record.name.back() == '.') {
				record.name.pop_back();
			}
		}
		if (record.name.empty()) {
			size_t slashPos = current.find_last_of("\\/");
			size_t beginPos = (slashPos == std::string::npos) ? 0 : (slashPos + 1);
			size_t endPos = current.size();
			if (endPos >= 3 && current.substr(endPos - 3) == ".ec") {
				endPos -= 3;
			}
			record.name = current.substr(beginPos, endPos - beginPos);
			if (!record.name.empty() && record.name.front() == '$') {
				record.name.erase(record.name.begin());
			}
		}
		outRecords.push_back(std::move(record));
	}

	return true;
}

std::string BuildArraySuffix(const std::vector<std::int32_t>& bounds)
{
	if (bounds.empty()) {
		return std::string();
	}
	if (bounds.size() == 1 && bounds[0] == 0) {
		return "\"0\"";
	}
	std::vector<std::string> parts;
	parts.reserve(bounds.size());
	for (const auto bound : bounds) {
		parts.push_back(bound == 0 ? std::string() : std::to_string(bound));
	}
	return "\"" + JoinStrings(parts, ",") + "\"";
}

std::string GetBuiltinTypeName(std::int32_t typeValue)
{
	switch (typeValue) {
	case 0: return "";
	case -1: return "";
	case static_cast<std::int32_t>(0x80000000u): return "通用型";
	case static_cast<std::int32_t>(0x80000101u): return "字节型";
	case static_cast<std::int32_t>(0x80000201u): return "短整数型";
	case static_cast<std::int32_t>(0x80000301u): return "整数型";
	case static_cast<std::int32_t>(0x80000401u): return "长整数型";
	case static_cast<std::int32_t>(0x80000501u): return "小数型";
	case static_cast<std::int32_t>(0x80000601u): return "双精度小数型";
	case static_cast<std::int32_t>(0x80000002u): return "逻辑型";
	case static_cast<std::int32_t>(0x80000003u): return "日期时间型";
	case static_cast<std::int32_t>(0x80000004u): return "文本型";
	case static_cast<std::int32_t>(0x80000005u): return "字节集";
	case static_cast<std::int32_t>(0x80000006u): return "子程序指针";
	case static_cast<std::int32_t>(0x80000008u): return "条件语句型";
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
	default: return std::string();
	}
}

namespace epl_system_id {

constexpr std::int32_t kIdNaV = 0x0500FFFE;

constexpr std::int32_t kTypeMethod = 0x04000000;
constexpr std::int32_t kTypeGlobal = 0x05000000;
constexpr std::int32_t kTypeStaticClass = 0x09000000;
constexpr std::int32_t kTypeDll = 0x0A000000;
constexpr std::int32_t kTypeClassMember = 0x15000000;
constexpr std::int32_t kTypeConstant = 0x18000000;
constexpr std::int32_t kTypeFormClass = 0x19000000;
constexpr std::int32_t kTypeLocal = 0x25000000;
constexpr std::int32_t kTypeImageResource = 0x28000000;
constexpr std::int32_t kTypeStructMember = 0x35000000;
constexpr std::int32_t kTypeSoundResource = 0x38000000;
constexpr std::int32_t kTypeStruct = 0x41000000;
constexpr std::int32_t kTypeDllParameter = 0x45000000;
constexpr std::int32_t kTypeClass = 0x49000000;
constexpr std::int32_t kTypeForm = 0x52000000;
constexpr std::int32_t kTypeFormSelf = 0x06000000;
constexpr std::int32_t kTypeFormControl = 0x16000000;
constexpr std::int32_t kTypeFormMenu = 0x26000000;

constexpr std::int32_t kMaskType = static_cast<std::int32_t>(0xFF000000u);
constexpr std::int32_t kMaskNum = 0x00FFFFFF;

inline std::int32_t GetType(std::int32_t id)
{
	return id & kMaskType;
}

inline bool IsLibDataType(std::int32_t id)
{
	return (id & kMaskType) == 0 && id != 0;
}

inline void DecomposeLibDataTypeId(std::int32_t id, std::int16_t& outLib, std::int16_t& outType)
{
	outLib = static_cast<std::int16_t>((id >> 16) - 1);
	outType = static_cast<std::int16_t>((id & 0xFFFF) - 1);
}

}  // namespace epl_system_id

struct SupportTypeSymbols {
	std::string name;
	std::vector<std::string> memberNames;
};

struct SupportLibrarySymbols {
	bool attempted = false;
	std::string fileName;
	std::string filePath;
	std::vector<std::string> commandNames;
	std::vector<std::string> constantNames;
	std::vector<SupportTypeSymbols> dataTypes;
};

std::string ResolveCoreLibraryFallbackMemberName(std::int32_t typeId, std::int32_t memberId)
{
	if (memberId == 8) {
		switch (typeId) {
		case 0:
		case 9:
		case 10:
		case 13:
		case 14:
		case 15:
			return "标题";
		default:
			break;
		}
	}
	return std::string();
}

std::string GetUserProfilePath()
{
	char* raw = nullptr;
	size_t size = 0;
	if (_dupenv_s(&raw, &size, "USERPROFILE") != 0 || raw == nullptr || size == 0) {
		return std::string();
	}
	std::string value(raw);
	free(raw);
	return value;
}

void PushUniqueCandidate(std::vector<std::filesystem::path>& candidates, const std::filesystem::path& candidate)
{
	if (candidate.empty()) {
		return;
	}

	const auto normalized = candidate.lexically_normal();
	for (const auto& item : candidates) {
		if (item.lexically_normal() == normalized) {
			return;
		}
	}
	candidates.push_back(normalized);
}

std::vector<std::filesystem::path> BuildSupportLibraryCandidatePaths(
	const std::string& sourcePath,
	const std::string& libraryFileName)
{
	std::vector<std::filesystem::path> candidates;
	if (libraryFileName.empty()) {
		return candidates;
	}

	std::filesystem::path filePath = std::filesystem::path(libraryFileName);
	if (!filePath.has_extension()) {
		filePath += ".fne";
	}

	if (filePath.is_absolute()) {
		PushUniqueCandidate(candidates, filePath);
		return candidates;
	}

	auto addBaseCandidates = [&](const std::filesystem::path& baseDir) {
		if (baseDir.empty()) {
			return;
		}
		PushUniqueCandidate(candidates, baseDir / filePath);
		PushUniqueCandidate(candidates, baseDir / "lib" / filePath);

		std::filesystem::path current = baseDir;
		while (!current.empty()) {
			PushUniqueCandidate(candidates, current / "lib" / filePath);
			if (current == current.root_path()) {
				break;
			}
			current = current.parent_path();
		}
	};

	std::error_code ec;
	if (!sourcePath.empty()) {
		addBaseCandidates(std::filesystem::path(sourcePath).parent_path());
	}
	addBaseCandidates(std::filesystem::current_path(ec));
	addBaseCandidates(std::filesystem::path(GetBasePath()));

	const std::string userProfile = GetUserProfilePath();
	if (!userProfile.empty()) {
		addBaseCandidates(std::filesystem::path(userProfile) / "OneDrive" / "e5.6");
	}

	return candidates;
}

class SymbolResolver {
public:
	SymbolResolver(const ProgramSection& program, const ResourceSection& resources, const std::string& sourcePath)
		: m_program(program)
		, m_resources(resources)
		, m_sourcePath(sourcePath)
	{
		BuildUserNameCache();
	}

	std::string ResolveType(std::int32_t typeValue)
	{
		for (const auto& dataType : m_program.dataTypes) {
			if (dataType.header.dwId == typeValue || dataType.header.dwUnk == typeValue) {
				return TrimAsciiCopy(dataType.name);
			}
		}
		for (const auto& page : m_program.codePages) {
			if (page.header.dwId == typeValue || page.header.dwUnk == typeValue) {
				return TrimAsciiCopy(page.name);
			}
		}

		const std::string builtin = GetBuiltinTypeName(typeValue);
		if (!builtin.empty()) {
			return builtin;
		}
		if (epl_system_id::IsLibDataType(typeValue)) {
			return ResolveSupportLibraryType(typeValue);
		}
		if (typeValue != 0 && epl_system_id::GetType(typeValue) != 0) {
			return ResolveUserName(typeValue);
		}
		return std::string();
	}

	std::string ResolveUserName(std::int32_t id) const
	{
		if (const auto it = m_userNameCache.find(id); it != m_userNameCache.end()) {
			return it->second;
		}

		switch (epl_system_id::GetType(id)) {
		case epl_system_id::kTypeMethod: return "_Sub_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeGlobal: return "_Global_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeStaticClass: return "_Mod_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeDll: return "_Dll_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeClassMember: return "_Mem_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeFormControl: return "_Control_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeFormMenu: return "_Menu_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeConstant: return "_Const_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeFormClass: return "_FormCls_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeLocal: return "_Local_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeImageResource: return "_Img_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeSoundResource: return "_Sound_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeStructMember: return "_StructMem_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeStruct: return "_Struct_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeDllParameter: return "_DllParam_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeClass: return "_Cls_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeForm: return "_Form_0x" + ToHexSuffix(id);
		case epl_system_id::kTypeFormSelf: return std::string();
		default: break;
		}

		std::ostringstream stream;
		stream << "_User_0x" << std::hex << std::uppercase << static_cast<std::uint32_t>(id);
		return stream.str();
	}

	bool TryGetMethodOwnerName(std::int32_t methodId, std::string& outName) const
	{
		if (const auto it = m_methodOwnerNameCache.find(methodId); it != m_methodOwnerNameCache.end()) {
			outName = it->second;
			return !outName.empty();
		}
		outName.clear();
		return false;
	}

	std::string ResolveLibCmdName(std::int16_t libraryIndex, std::int32_t commandId)
	{
		if (!EnsureSupportLibraryCacheByLibraryId(libraryIndex)) {
			std::ostringstream stream;
			stream << "_Lib" << libraryIndex << "Cmd" << commandId;
			return stream.str();
		}

		const auto* symbols = FindSupportLibrarySymbolsByLibraryId(libraryIndex);
		if (symbols == nullptr || commandId < 0 || static_cast<size_t>(commandId) >= symbols->commandNames.size()) {
			std::ostringstream stream;
			stream << "_Lib" << libraryIndex << "Cmd" << commandId;
			return stream.str();
		}
		const std::string& name = symbols->commandNames[static_cast<size_t>(commandId)];
		return name.empty() ? std::string("_Lib") + std::to_string(libraryIndex) + "Cmd" + std::to_string(commandId) : name;
	}

	std::string ResolveLibConstantName(std::int16_t libraryIndex, std::int32_t constantId)
	{
		if (!EnsureSupportLibraryCacheByLibraryId(libraryIndex)) {
			std::ostringstream stream;
			stream << "_Lib" << libraryIndex << "Const" << constantId;
			return stream.str();
		}

		const auto* symbols = FindSupportLibrarySymbolsByLibraryId(libraryIndex);
		if (symbols == nullptr || constantId < 0 || static_cast<size_t>(constantId) >= symbols->constantNames.size()) {
			std::ostringstream stream;
			stream << "_Lib" << libraryIndex << "Const" << constantId;
			return stream.str();
		}
		const std::string& name = symbols->constantNames[static_cast<size_t>(constantId)];
		return name.empty() ? std::string("_Lib") + std::to_string(libraryIndex) + "Const" + std::to_string(constantId) : name;
	}

	std::string ResolveLibTypeName(std::int16_t libraryIndex, std::int32_t typeId)
	{
		if (!EnsureSupportLibraryCacheByLibraryId(libraryIndex)) {
			std::ostringstream stream;
			stream << "_Lib" << libraryIndex << "Type" << typeId;
			return stream.str();
		}

		const auto* symbols = FindSupportLibrarySymbolsByLibraryId(libraryIndex);
		if (symbols == nullptr || typeId < 0 || static_cast<size_t>(typeId) >= symbols->dataTypes.size()) {
			std::ostringstream stream;
			stream << "_Lib" << libraryIndex << "Type" << typeId;
			return stream.str();
		}

		const std::string& name = symbols->dataTypes[static_cast<size_t>(typeId)].name;
		return name.empty() ? std::string("_Lib") + std::to_string(libraryIndex) + "Type" + std::to_string(typeId) : name;
	}

	std::string ResolveLibTypeMemberName(std::int16_t libraryIndex, std::int32_t typeId, std::int32_t memberId)
	{
		if (libraryIndex == 0) {
			const std::string fallbackName = ResolveCoreLibraryFallbackMemberName(typeId, memberId);
			if (!fallbackName.empty()) {
				return fallbackName;
			}
		}

		if (!EnsureSupportLibraryCacheByLibraryId(libraryIndex)) {
			std::ostringstream stream;
			stream << "_Lib" << libraryIndex << "Type" << typeId << "Mem" << memberId;
			return stream.str();
		}

		const auto* symbols = FindSupportLibrarySymbolsByLibraryId(libraryIndex);
		if (symbols == nullptr || typeId < 0 || static_cast<size_t>(typeId) >= symbols->dataTypes.size()) {
			std::ostringstream stream;
			stream << "_Lib" << libraryIndex << "Type" << typeId << "Mem" << memberId;
			return stream.str();
		}

		const auto& members = symbols->dataTypes[static_cast<size_t>(typeId)].memberNames;
		if (memberId < 0 || static_cast<size_t>(memberId) >= members.size()) {
			std::ostringstream stream;
			stream << "_Lib" << libraryIndex << "Type" << typeId << "Mem" << memberId;
			return stream.str();
		}

		const std::string& name = members[static_cast<size_t>(memberId)];
		return name.empty() ? std::string("_Lib") + std::to_string(libraryIndex) + "Type" + std::to_string(typeId) + "Mem" + std::to_string(memberId) : name;
	}

private:
	static std::string ToHexSuffix(std::int32_t id)
	{
		std::ostringstream stream;
		stream << std::hex << std::uppercase << static_cast<std::uint32_t>(id & epl_system_id::kMaskNum);
		return stream.str();
	}

	void BuildUserNameCache()
	{
		for (const auto& page : m_program.codePages) {
			const std::string pageName = TrimAsciiCopy(page.name);
			m_userNameCache[page.header.dwId] = pageName;
			for (const auto& pageVar : page.pageVars) {
				m_userNameCache[pageVar.marker] = TrimAsciiCopy(pageVar.name);
			}
			for (const auto functionId : page.functionIds) {
				m_methodOwnerNameCache[functionId] = pageName;
			}
		}

		for (const auto& function : m_program.functions) {
			m_userNameCache[function.header.dwId] = TrimAsciiCopy(function.name);
			for (const auto& item : function.params) {
				m_userNameCache[item.marker] = TrimAsciiCopy(item.name);
			}
			for (const auto& item : function.locals) {
				m_userNameCache[item.marker] = TrimAsciiCopy(item.name);
			}
		}
		for (const auto& item : m_program.globals) {
			m_userNameCache[item.marker] = TrimAsciiCopy(item.name);
		}
		for (const auto& item : m_program.dataTypes) {
			m_userNameCache[item.header.dwId] = TrimAsciiCopy(item.name);
			for (const auto& member : item.members) {
				m_userNameCache[member.marker] = TrimAsciiCopy(member.name);
			}
		}
		for (const auto& item : m_program.dlls) {
			m_userNameCache[item.header.dwId] = TrimAsciiCopy(item.name);
			for (const auto& param : item.params) {
				m_userNameCache[param.marker] = TrimAsciiCopy(param.name);
			}
		}
		for (const auto& item : m_resources.constants) {
			m_userNameCache[item.marker] = TrimAsciiCopy(item.name);
		}
		for (const auto& item : m_resources.forms) {
			m_userNameCache[item.header.dwId] = TrimAsciiCopy(item.name);
			for (const auto& element : item.elements) {
				std::int32_t elementId = element.id;
				if (epl_system_id::GetType(elementId) == 0) {
					elementId |= element.isMenu ? epl_system_id::kTypeFormMenu : epl_system_id::kTypeFormControl;
				}
				m_userNameCache[elementId] = TrimAsciiCopy(element.name);
			}
		}
	}

	SupportLibrarySymbols* FindSupportLibrarySymbolsByLibraryId(std::int16_t libraryIndex)
	{
		if (libraryIndex < 0) {
			return nullptr;
		}
		return FindSupportLibrarySymbolsByTypeIndex(static_cast<std::uint16_t>(libraryIndex + 1));
	}

	SupportLibrarySymbols* FindSupportLibrarySymbolsByTypeIndex(std::uint16_t supportIndex)
	{
		if (supportIndex == 0) {
			return nullptr;
		}
		if (const auto it = m_supportCache.find(supportIndex); it != m_supportCache.end()) {
			return &it->second;
		}
		return nullptr;
	}

	bool EnsureSupportLibraryCacheByLibraryId(std::int16_t libraryIndex)
	{
		if (libraryIndex < 0) {
			return false;
		}
		return EnsureSupportLibraryCacheByTypeIndex(static_cast<std::uint16_t>(libraryIndex + 1));
	}

	bool EnsureSupportLibraryCacheByTypeIndex(std::uint16_t supportIndex)
	{
		if (supportIndex == 0) {
			return false;
		}
		if (const auto it = m_supportCache.find(supportIndex); it != m_supportCache.end()) {
			return !it->second.dataTypes.empty() || !it->second.commandNames.empty() || !it->second.constantNames.empty();
		}
		if (supportIndex > m_program.header.supportLibraryInfo.size()) {
			SupportLibrarySymbols emptySymbols;
			emptySymbols.attempted = true;
			m_supportCache.emplace(supportIndex, std::move(emptySymbols));
			return false;
		}

		SupportLibrarySymbols symbols;
		symbols.attempted = true;
		symbols.fileName = GetFirstSupportLibraryToken(m_program.header.supportLibraryInfo[static_cast<size_t>(supportIndex - 1)]);
		if (symbols.fileName.empty()) {
			m_supportCache.emplace(supportIndex, std::move(symbols));
			return false;
		}

		const auto candidates = BuildSupportLibraryCandidatePaths(m_sourcePath, symbols.fileName);
		HMODULE module = nullptr;
		for (const auto& path : candidates) {
			std::error_code ec;
			if (!std::filesystem::exists(path, ec)) {
				continue;
			}
			module = LoadLibraryExA(path.string().c_str(), nullptr, DONT_RESOLVE_DLL_REFERENCES);
			if (module != nullptr) {
				symbols.filePath = path.string();
				break;
			}
		}

		if (module != nullptr) {
			const auto* getInfoProc = reinterpret_cast<PFN_GET_LIB_INFO>(GetProcAddress(module, FUNCNAME_GET_LIB_INFO));
			if (getInfoProc != nullptr) {
				const auto* libInfo = getInfoProc();
				if (libInfo != nullptr) {
					if (libInfo->m_nCmdCount > 0 && libInfo->m_pBeginCmdInfo != nullptr) {
						symbols.commandNames.reserve(static_cast<size_t>(libInfo->m_nCmdCount));
						for (int i = 0; i < libInfo->m_nCmdCount; ++i) {
							symbols.commandNames.emplace_back(libInfo->m_pBeginCmdInfo[i].m_szName == nullptr ? "" : libInfo->m_pBeginCmdInfo[i].m_szName);
						}
					}
					if (libInfo->m_nLibConstCount > 0 && libInfo->m_pLibConst != nullptr) {
						symbols.constantNames.reserve(static_cast<size_t>(libInfo->m_nLibConstCount));
						for (int i = 0; i < libInfo->m_nLibConstCount; ++i) {
							symbols.constantNames.emplace_back(libInfo->m_pLibConst[i].m_szName == nullptr ? "" : libInfo->m_pLibConst[i].m_szName);
						}
					}
					if (libInfo->m_nDataTypeCount > 0 && libInfo->m_pDataType != nullptr) {
						symbols.dataTypes.reserve(static_cast<size_t>(libInfo->m_nDataTypeCount));
						for (int i = 0; i < libInfo->m_nDataTypeCount; ++i) {
							const LIB_DATA_TYPE_INFO& dataType = libInfo->m_pDataType[i];
							SupportTypeSymbols typeSymbols;
							typeSymbols.name = dataType.m_szName == nullptr ? "" : dataType.m_szName;
							if (dataType.m_nElementCount > 0 && dataType.m_pElementBegin != nullptr) {
								typeSymbols.memberNames.reserve(static_cast<size_t>(dataType.m_nElementCount));
								for (int memberIndex = 0; memberIndex < dataType.m_nElementCount; ++memberIndex) {
									typeSymbols.memberNames.emplace_back(
										dataType.m_pElementBegin[memberIndex].m_szName == nullptr
											? ""
											: dataType.m_pElementBegin[memberIndex].m_szName);
								}
							}
							symbols.dataTypes.push_back(std::move(typeSymbols));
						}
					}
				}
			}
			FreeLibrary(module);
		}

		const bool loaded = !symbols.dataTypes.empty() || !symbols.commandNames.empty() || !symbols.constantNames.empty();
		m_supportCache.emplace(supportIndex, std::move(symbols));
		return loaded;
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
		if (!EnsureSupportLibraryCacheByTypeIndex(supportIndex)) {
			std::ostringstream stream;
			stream << "0x" << std::hex << std::uppercase << rawValue;
			return stream.str();
		}

		const auto* symbols = FindSupportLibrarySymbolsByTypeIndex(supportIndex);
		if (symbols == nullptr || typeIndex > symbols->dataTypes.size()) {
			std::ostringstream stream;
			stream << "0x" << std::hex << std::uppercase << rawValue;
			return stream.str();
		}
		const std::string& name = symbols->dataTypes[static_cast<size_t>(typeIndex - 1)].name;
		return name.empty() ? std::string() : name;
	}

	const ProgramSection& m_program;
	const ResourceSection& m_resources;
	std::string m_sourcePath;
	std::unordered_map<std::int32_t, std::string> m_userNameCache;
	std::unordered_map<std::int32_t, std::string> m_methodOwnerNameCache;
	std::unordered_map<std::uint16_t, SupportLibrarySymbols> m_supportCache;
};

struct AnonymousTypeHints {
	std::unordered_map<std::int32_t, std::string> localTypeAliases;
	std::unordered_map<std::string, std::string> importedMemberAliases;
	std::unordered_map<std::string, std::string> importedDllParamAliases;
};

std::string BuildHintKey(const std::string& left, const std::string& right)
{
	return TrimAsciiCopy(left) + "\n" + TrimAsciiCopy(right);
}

bool IsAnonymousPlaceholderTypeName(const std::string& typeName)
{
	return typeName.empty() || typeName.rfind("_Struct_0x", 0) == 0;
}

void SetAnonymousTypeAlias(
	std::unordered_map<std::int32_t, std::string>& aliases,
	std::int32_t typeId,
	const std::string& alias)
{
	if (typeId == 0 || alias.empty()) {
		return;
	}
	aliases.insert_or_assign(typeId, alias);
}

bool MatchAllMemberTypes(
	const std::vector<VariableInfo>& members,
	std::int32_t expectedType,
	size_t expectedCount)
{
	if (members.size() != expectedCount) {
		return false;
	}
	return std::all_of(
		members.begin(),
		members.end(),
		[expectedType](const VariableInfo& item) { return item.dataType == expectedType; });
}

bool MatchWndClassExShape(const std::vector<VariableInfo>& members)
{
	if (members.size() != 12) {
		return false;
	}
	for (size_t i = 0; i < 9; ++i) {
		if (members[i].dataType != static_cast<std::int32_t>(0x80000301u)) {
			return false;
		}
	}
	return
		members[9].dataType == static_cast<std::int32_t>(0x80000005u) &&
		members[10].dataType == static_cast<std::int32_t>(0x80000005u) &&
		members[11].dataType == static_cast<std::int32_t>(0x80000301u);
}

std::unordered_map<std::int32_t, std::string> BuildLocalAnonymousTypeAliasMap(const ModuleSections& sections)
{
	std::unordered_map<std::int32_t, std::string> aliases;

	for (const auto& item : sections.program.dataTypes) {
		const std::string ownerName = TrimAsciiCopy(item.name);
		if (ownerName.empty()) {
			continue;
		}
		for (const auto& member : item.members) {
			const std::string memberName = TrimAsciiCopy(member.name);
			if (memberName.empty()) {
				continue;
			}
			if (ownerName == "WFSYSCTLBTN" && memberName == "Tips") {
				SetAnonymousTypeAlias(aliases, member.dataType, "WFTIPSINFO");
			}
			else if (ownerName == "WFSHADOWINFO" && memberName == "Canvas") {
				SetAnonymousTypeAlias(aliases, member.dataType, "WFCANVAS");
			}
		}
	}

	for (const auto& item : sections.program.dlls) {
		const std::string dllName = TrimAsciiCopy(item.name);
		if (dllName.empty()) {
			continue;
		}
		for (const auto& param : item.params) {
			const std::string paramName = TrimAsciiCopy(param.name);
			if (paramName.empty()) {
				continue;
			}
			if ((dllName == "GdipDrawString" || dllName == "GdipMeasureString") &&
				(paramName == "layoutRect" || paramName == "boundingBox")) {
				SetAnonymousTypeAlias(aliases, param.dataType, "RectF");
			}
			else if (dllName == "RegisterClassExW" && paramName == "pcWndClassEx") {
				SetAnonymousTypeAlias(aliases, param.dataType, "WNDCLASSEXW");
			}
		}
	}

	for (const auto& item : sections.program.dataTypes) {
		if (!TrimAsciiCopy(item.name).empty()) {
			continue;
		}
		if (aliases.find(item.header.dwId) != aliases.end()) {
			continue;
		}
		if (MatchAllMemberTypes(item.members, static_cast<std::int32_t>(0x80000501u), 4)) {
			SetAnonymousTypeAlias(aliases, item.header.dwId, "RectF");
		}
		else if (MatchWndClassExShape(item.members)) {
			SetAnonymousTypeAlias(aliases, item.header.dwId, "WNDCLASSEXW");
		}
	}

	return aliases;
}

std::vector<std::filesystem::path> BuildModuleCandidatePaths(
	const std::string& sourcePath,
	const std::string& modulePathText)
{
	std::vector<std::filesystem::path> candidates;
	std::string normalizedText = TrimAsciiCopy(modulePathText);
	if (normalizedText.empty()) {
		return candidates;
	}

	if (normalizedText.front() == '"' && normalizedText.back() == '"' && normalizedText.size() >= 2) {
		normalizedText = normalizedText.substr(1, normalizedText.size() - 2);
	}
	if (!normalizedText.empty() && normalizedText.front() == '$') {
		normalizedText.erase(normalizedText.begin());
	}

	std::filesystem::path filePath(normalizedText);
	if (filePath.extension().empty()) {
		filePath += ".ec";
	}

	if (filePath.is_absolute()) {
		PushUniqueCandidate(candidates, filePath);
		return candidates;
	}

	auto addBaseCandidates = [&](const std::filesystem::path& baseDir) {
		if (baseDir.empty()) {
			return;
		}
		PushUniqueCandidate(candidates, baseDir / filePath);
		PushUniqueCandidate(candidates, baseDir / "ecom" / filePath);

		std::filesystem::path current = baseDir;
		while (!current.empty()) {
			PushUniqueCandidate(candidates, current / "ecom" / filePath);
			if (current == current.root_path()) {
				break;
			}
			current = current.parent_path();
		}
	};

	std::error_code ec;
	if (!sourcePath.empty()) {
		addBaseCandidates(std::filesystem::path(sourcePath).parent_path());
	}
	addBaseCandidates(std::filesystem::current_path(ec));
	addBaseCandidates(std::filesystem::path(GetBasePath()));

	const std::string userProfile = GetUserProfilePath();
	if (!userProfile.empty()) {
		addBaseCandidates(std::filesystem::path(userProfile) / "OneDrive" / "e5.6");
	}

	return candidates;
}

AnonymousTypeHints BuildAnonymousTypeHints(const ModuleSections& sections, const std::string& sourcePath)
{
	AnonymousTypeHints hints;
	hints.localTypeAliases = BuildLocalAnonymousTypeAliasMap(sections);

	std::vector<EComDependencyRecord> ecomRecords;
	if (!ParseEComDependencies(sections.ecomSectionBytes, ecomRecords)) {
		return hints;
	}

	std::unordered_set<std::string> loadedPaths;
	for (const auto& record : ecomRecords) {
		const auto candidates = BuildModuleCandidatePaths(sourcePath, record.path);
		for (const auto& candidate : candidates) {
			std::error_code ec;
			if (!std::filesystem::exists(candidate, ec)) {
				continue;
			}
			const std::string normalized = candidate.lexically_normal().string();
			if (!loadedPaths.insert(normalized).second) {
				break;
			}

			ModuleSections dependencySections;
			std::string loadError;
			if (!ParseModuleSections(normalized, dependencySections, &loadError)) {
				break;
			}

			const auto dependencyAliases = BuildLocalAnonymousTypeAliasMap(dependencySections);
			for (const auto& item : dependencySections.program.dataTypes) {
				const std::string ownerName = TrimAsciiCopy(item.name);
				if (ownerName.empty()) {
					continue;
				}
				for (const auto& member : item.members) {
					const std::string memberName = TrimAsciiCopy(member.name);
					if (memberName.empty()) {
						continue;
					}
					if (const auto aliasIt = dependencyAliases.find(member.dataType); aliasIt != dependencyAliases.end()) {
						hints.importedMemberAliases[BuildHintKey(ownerName, memberName)] = aliasIt->second;
					}
				}
			}

			for (const auto& item : dependencySections.program.dlls) {
				const std::string dllName = TrimAsciiCopy(item.name);
				if (dllName.empty()) {
					continue;
				}
				for (const auto& param : item.params) {
					const std::string paramName = TrimAsciiCopy(param.name);
					if (paramName.empty()) {
						continue;
					}
					if (const auto aliasIt = dependencyAliases.find(param.dataType); aliasIt != dependencyAliases.end()) {
						hints.importedDllParamAliases[BuildHintKey(dllName, paramName)] = aliasIt->second;
					}
				}
			}
			break;
		}
	}

	return hints;
}

std::string ResolveMemberTypeWithHints(
	const VariableInfo& info,
	const std::string& ownerName,
	SymbolResolver& resolver,
	const AnonymousTypeHints& hints)
{
	std::string typeText = TrimAsciiCopy(resolver.ResolveType(info.dataType));
	if (!IsAnonymousPlaceholderTypeName(typeText)) {
		return typeText;
	}

	if (const auto it = hints.localTypeAliases.find(info.dataType); it != hints.localTypeAliases.end()) {
		return it->second;
	}
	if (const auto it = hints.importedMemberAliases.find(BuildHintKey(ownerName, info.name));
		it != hints.importedMemberAliases.end()) {
		return it->second;
	}
	return typeText;
}

std::string ResolveDllParamTypeWithHints(
	const VariableInfo& info,
	const std::string& dllName,
	SymbolResolver& resolver,
	const AnonymousTypeHints& hints)
{
	std::string typeText = TrimAsciiCopy(resolver.ResolveType(info.dataType));
	if (!IsAnonymousPlaceholderTypeName(typeText)) {
		return typeText;
	}

	if (const auto it = hints.localTypeAliases.find(info.dataType); it != hints.localTypeAliases.end()) {
		return it->second;
	}
	if (const auto it = hints.importedDllParamAliases.find(BuildHintKey(dllName, info.name));
		it != hints.importedDllParamAliases.end()) {
		return it->second;
	}
	return typeText;
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
	return JoinStrings(parts, " ");
}

std::string BuildTypeField(const VariableInfo& info, SymbolResolver& resolver)
{
	return TrimAsciiCopy(resolver.ResolveType(info.dataType));
}

std::string BuildMethodParameterLine(const VariableInfo& info, SymbolResolver& resolver)
{
	std::vector<std::string> flags;
	if ((info.attr & kVarAttrByRef) != 0) {
		flags.push_back("参考");
	}
	if ((info.attr & kVarAttrNullable) != 0) {
		flags.push_back("可空");
	}
	if ((info.attr & kVarAttrArray) != 0) {
		flags.push_back("数组");
	}
	return BuildDefinitionLine(
		"参数",
		{
			TrimAsciiCopy(info.name),
			BuildTypeField(info, resolver),
			JoinStrings(flags, " "),
			TrimAsciiCopy(info.comment),
		});
}

std::string BuildLocalVariableLine(const VariableInfo& info, SymbolResolver& resolver)
{
	return BuildDefinitionLine(
		"局部变量",
		{
			TrimAsciiCopy(info.name),
			BuildTypeField(info, resolver),
			(info.attr & 0x0001) != 0 ? "静态" : std::string(),
			BuildArraySuffix(info.arrayBounds),
			TrimAsciiCopy(info.comment),
		});
}

std::string BuildClassVariableLine(const VariableInfo& info, SymbolResolver& resolver)
{
	return BuildDefinitionLine(
		"程序集变量",
		{
			TrimAsciiCopy(info.name),
			BuildTypeField(info, resolver),
			std::string(),
			BuildArraySuffix(info.arrayBounds),
			TrimAsciiCopy(info.comment),
		});
}

std::string BuildGlobalVariableLine(const VariableInfo& info, SymbolResolver& resolver)
{
	return BuildDefinitionLine(
		"全局变量",
		{
			TrimAsciiCopy(info.name),
			BuildTypeField(info, resolver),
			(info.attr & 0x0100) != 0 ? "公开" : std::string(),
			BuildArraySuffix(info.arrayBounds),
			TrimAsciiCopy(info.comment),
		});
}

std::string BuildStructMemberLine(const VariableInfo& info, const std::string& resolvedType)
{
	return BuildDefinitionLine(
		"成员",
		{
			TrimAsciiCopy(info.name),
			resolvedType,
			(info.attr & kVarAttrByRef) != 0 ? "传址" : std::string(),
			BuildArraySuffix(info.arrayBounds),
			TrimAsciiCopy(info.comment),
		},
		1);
}

std::string BuildDllParameterLine(const VariableInfo& info, const std::string& resolvedType)
{
	std::vector<std::string> flags;
	if ((info.attr & kVarAttrByRef) != 0) {
		flags.push_back("传址");
	}
	if ((info.attr & kVarAttrArray) != 0) {
		flags.push_back("数组");
	}
	return BuildDefinitionLine(
		"参数",
		{
			TrimAsciiCopy(info.name),
			resolvedType,
			JoinStrings(flags, " "),
			TrimAsciiCopy(info.comment),
		},
		1);
}

void AppendLine(Page& page, const std::string& line)
{
	page.lines.push_back(line);
}

bool IsImportedFunction(const FunctionInfo& functionInfo)
{
	return (functionInfo.attr & 0x80) != 0;
}

bool ShouldKeepPage(const CodePageInfo& page, const std::vector<const FunctionInfo*>& functions, bool includeImportedPages)
{
	if (TrimAsciiCopy(page.name).empty() || page.name == "__HIDDEN_TEMP_MOD__") {
		return false;
	}
	if (includeImportedPages) {
		return true;
	}
	if (functions.empty()) {
		return true;
	}
	return std::any_of(
		functions.begin(),
		functions.end(),
		[](const FunctionInfo* item) { return item != nullptr && !IsImportedFunction(*item); });
}

std::string BuildVersionText(const ModuleSections& sections)
{
	if (sections.hasUserInfo) {
		return std::to_string(sections.userInfo.version1) + "." + std::to_string(sections.userInfo.version2);
	}
	if (sections.hasSystemInfo) {
		return std::to_string(sections.systemInfo.compileMajor) + "." + std::to_string(sections.systemInfo.compileMinor);
	}
	return "0.0";
}

void BuildDependencies(const ModuleSections& sections, Document& outDocument)
{
	TraceLine("BuildDependencies begin");
	for (const auto& rawText : sections.program.header.supportLibraryInfo) {
		const auto tokens = SplitSupportLibraryTokens(rawText);
		if (tokens.empty()) {
			continue;
		}

		Dependency item;
		item.kind = DependencyKind::ELib;
		item.fileName = tokens.size() > 0 ? tokens[0] : std::string();
		item.guid = tokens.size() > 1 ? tokens[1] : std::string();
		item.versionText = tokens.size() > 2 ? tokens[2] : std::string();
		const std::string majorOrMinor = tokens.size() > 3 ? tokens[3] : std::string();
		item.name = tokens.size() > 4 ? tokens[4] : majorOrMinor;
		if (item.name.empty()) {
			item.name = item.fileName;
		}
		if (!majorOrMinor.empty() && !item.versionText.empty()) {
			item.versionText += "." + majorOrMinor;
		}
		outDocument.dependencies.push_back(std::move(item));
	}
	TraceLine("BuildDependencies support_lib_done count=" + std::to_string(outDocument.dependencies.size()));

	std::vector<EComDependencyRecord> ecomRecords;
	TraceLine("BuildDependencies ecom_parse_begin bytes=" + std::to_string(sections.ecomSectionBytes.size()));
	if (ParseEComDependencies(sections.ecomSectionBytes, ecomRecords)) {
		TraceLine("BuildDependencies ecom_parse_ok count=" + std::to_string(ecomRecords.size()));
		for (const auto& record : ecomRecords) {
			Dependency item;
			item.kind = DependencyKind::ECom;
			item.name = record.name;
			item.path = record.path;
			item.reExport = record.reExport;
			outDocument.dependencies.push_back(std::move(item));
		}
	}
	TraceLine("BuildDependencies end count=" + std::to_string(outDocument.dependencies.size()));
}

std::vector<const FunctionInfo*> CollectPageFunctions(const ProgramSection& program, const CodePageInfo& page)
{
	std::vector<const FunctionInfo*> out;
	for (const auto functionId : page.functionIds) {
		for (const auto& function : program.functions) {
			if (function.header.dwId == functionId) {
				out.push_back(&function);
				break;
			}
		}
	}
	return out;
}

std::string BuildIndent(int level)
{
	return std::string(static_cast<size_t>(level > 0 ? level : 0) * 4, ' ');
}

std::string FormatNumberLiteral(double value)
{
	if (std::isfinite(value)) {
		const double rounded = std::round(value);
		if (std::fabs(value - rounded) < 1e-12 &&
			rounded >= static_cast<double>((std::numeric_limits<long long>::min)()) &&
			rounded <= static_cast<double>((std::numeric_limits<long long>::max)())) {
			return std::to_string(static_cast<long long>(rounded));
		}
	}

	std::ostringstream stream;
	stream << std::fixed << std::setprecision(15) << value;
	std::string text = stream.str();
	while (!text.empty() && text.back() == '0') {
		text.pop_back();
	}
	if (!text.empty() && text.back() == '.') {
		text.pop_back();
	}
	if (text == "-0") {
		text = "0";
	}
	return text.empty() ? "0" : text;
}

std::string FormatDateLiteral(double value)
{
	SYSTEMTIME systemTime = {};
	if (VariantTimeToSystemTime(value, &systemTime) == FALSE) {
		return "[]";
	}

	std::ostringstream stream;
	stream << "[";
	stream << std::setfill('0') << std::setw(4) << systemTime.wYear << "年";
	stream << std::setfill('0') << std::setw(2) << systemTime.wMonth << "月";
	stream << std::setfill('0') << std::setw(2) << systemTime.wDay << "日";
	if (systemTime.wHour != 0 || systemTime.wMinute != 0 || systemTime.wSecond != 0) {
		stream << std::setfill('0') << std::setw(2) << systemTime.wHour << "时";
		stream << std::setfill('0') << std::setw(2) << systemTime.wMinute << "分";
		stream << std::setfill('0') << std::setw(2) << systemTime.wSecond << "秒";
	}
	stream << "]";
	return stream.str();
}

enum class ExprKind {
	DefaultValue,
	ParamListEnd,
	ArrayLiteralEnd,
	Number,
	Bool,
	Date,
	String,
	Constant,
	EnumConstant,
	Variable,
	MethodPtr,
	ArrayLiteral,
	AccessArray,
	AccessMember,
	Call,
};

struct Expr {
	ExprKind kind = ExprKind::DefaultValue;
	double numberValue = 0.0;
	bool boolValue = false;
	std::int32_t intValue1 = 0;
	std::int32_t intValue2 = 0;
	std::int32_t intValue3 = 0;
	std::int16_t shortValue1 = 0;
	std::int16_t shortValue2 = 0;
	bool flagValue = false;
	std::string text;
	std::unique_ptr<Expr> target;
	std::unique_ptr<Expr> extra;
	std::vector<std::unique_ptr<Expr>> items;
};

enum class StatementKind {
	Expression,
	Unexamined,
	IfTrue,
	IfElse,
	WhileLoop,
	DoWhileLoop,
	CounterLoop,
	ForLoop,
	SwitchBlock,
};

struct StatementBlock;

struct SwitchCase {
	std::unique_ptr<Expr> condition;
	std::string unexaminedCode;
	std::string comment;
	bool mask = false;
	std::unique_ptr<StatementBlock> block;
};

struct Statement {
	StatementKind kind = StatementKind::Expression;
	bool mask = false;
	bool maskOnEnd = false;
	std::string comment;
	std::string commentOnEnd;
	std::string unexaminedCode;
	std::unique_ptr<Expr> expression;
	std::unique_ptr<Expr> condition;
	std::unique_ptr<Expr> expr2;
	std::unique_ptr<Expr> expr3;
	std::unique_ptr<Expr> expr4;
	std::unique_ptr<StatementBlock> block;
	std::unique_ptr<StatementBlock> elseBlock;
	std::unique_ptr<StatementBlock> defaultBlock;
	std::vector<SwitchCase> switchCases;
};

struct StatementBlock {
	std::vector<Statement> items;
};

struct OperatorInfo {
	const char* text = "";
	int precedence = 0;
	enum class Type {
		Unary,
		Binary,
		Multi,
	} type = Type::Binary;
};

std::string RenderExpr(const Expr& expr, SymbolResolver& resolver, int expectedLowestPrecedence = (std::numeric_limits<int>::max)());
bool ParseExpression(ByteReader& reader, std::unique_ptr<Expr>& outExpr, std::string* outError, bool parseMember = true);
bool ParseStatementBlock(ByteReader& reader, std::unique_ptr<StatementBlock>& outBlock, std::string* outError);
void AppendStatementLines(const StatementBlock& block, SymbolResolver& resolver, int indent, std::vector<std::string>& outLines);

bool TryGetOperatorInfo(std::int32_t methodId, OperatorInfo& outInfo)
{
	switch (methodId) {
	case 15: outInfo = { "*", 2, OperatorInfo::Type::Multi }; return true;
	case 16: outInfo = { "/", 2, OperatorInfo::Type::Multi }; return true;
	case 17: outInfo = { "\\", 3, OperatorInfo::Type::Multi }; return true;
	case 18: outInfo = { "%", 4, OperatorInfo::Type::Multi }; return true;
	case 19: outInfo = { "+", 5, OperatorInfo::Type::Multi }; return true;
	case 20: outInfo = { "-", 5, OperatorInfo::Type::Multi }; return true;
	case 21: outInfo = { "-", 1, OperatorInfo::Type::Unary }; return true;
	case 38: outInfo = { "==", 6, OperatorInfo::Type::Binary }; return true;
	case 39: outInfo = { "!=", 6, OperatorInfo::Type::Binary }; return true;
	case 40: outInfo = { "<", 6, OperatorInfo::Type::Binary }; return true;
	case 41: outInfo = { ">", 6, OperatorInfo::Type::Binary }; return true;
	case 42: outInfo = { "<=", 6, OperatorInfo::Type::Binary }; return true;
	case 43: outInfo = { ">=", 6, OperatorInfo::Type::Binary }; return true;
	case 44: outInfo = { "?=", 6, OperatorInfo::Type::Binary }; return true;
	case 45: outInfo = { "&&", 7, OperatorInfo::Type::Multi }; return true;
	case 46: outInfo = { "||", 8, OperatorInfo::Type::Multi }; return true;
	case 52: outInfo = { "=", 9, OperatorInfo::Type::Binary }; return true;
	default: return false;
	}
}

std::string RenderExprList(const std::vector<std::unique_ptr<Expr>>& items, SymbolResolver& resolver)
{
	std::string out;
	for (size_t i = 0; i < items.size(); ++i) {
		if (i > 0) {
			out += ", ";
		}
		out += items[i] == nullptr ? std::string() : RenderExpr(*items[i], resolver);
	}
	return out;
}

std::string RenderExpr(const Expr& expr, SymbolResolver& resolver, int expectedLowestPrecedence)
{
	switch (expr.kind) {
	case ExprKind::DefaultValue:
	case ExprKind::ParamListEnd:
	case ExprKind::ArrayLiteralEnd:
		return std::string();
	case ExprKind::Number:
		return FormatNumberLiteral(expr.numberValue);
	case ExprKind::Bool:
		return expr.boolValue ? "真" : "假";
	case ExprKind::Date:
		return FormatDateLiteral(expr.numberValue);
	case ExprKind::String:
		return "“" + expr.text + "”";
	case ExprKind::Constant:
		return expr.shortValue1 == -2
			? "#" + resolver.ResolveUserName(expr.intValue1)
			: "#" + resolver.ResolveLibConstantName(expr.shortValue1, expr.intValue1);
	case ExprKind::EnumConstant:
		return "#" + resolver.ResolveLibTypeName(expr.shortValue1, expr.shortValue2) + "." +
			resolver.ResolveLibTypeMemberName(expr.shortValue1, expr.shortValue2, expr.intValue1);
	case ExprKind::Variable:
		return resolver.ResolveUserName(expr.intValue1);
	case ExprKind::MethodPtr:
		return "&" + resolver.ResolveUserName(expr.intValue1);
	case ExprKind::ArrayLiteral:
		return "{" + RenderExprList(expr.items, resolver) + "}";
	case ExprKind::AccessArray:
		return (expr.target == nullptr ? std::string() : RenderExpr(*expr.target, resolver)) +
			"[" +
			(expr.extra == nullptr ? std::string() : RenderExpr(*expr.extra, resolver)) +
			"]";
	case ExprKind::AccessMember: {
		std::string prefix;
		if (expr.target != nullptr) {
			const bool omitTarget =
				expr.target->kind == ExprKind::Variable &&
				epl_system_id::GetType(expr.target->intValue1) == epl_system_id::kTypeFormSelf;
			if (!omitTarget) {
				prefix = RenderExpr(*expr.target, resolver) + ".";
			}
		}
		if (expr.shortValue1 == -2) {
			if (epl_system_id::GetType(expr.intValue2) == epl_system_id::kTypeForm &&
				(static_cast<std::uint32_t>(expr.intValue1) & 0xFF000000u) == 0) {
				return prefix + resolver.ResolveLibTypeMemberName(0, 0, expr.intValue1 - 1);
			}
			return prefix + resolver.ResolveUserName(expr.intValue1);
		}
		return prefix + resolver.ResolveLibTypeMemberName(expr.shortValue1, expr.intValue2, expr.intValue1);
	}
	case ExprKind::Call: {
		if (expr.target == nullptr && expr.shortValue1 == 0) {
			OperatorInfo operatorInfo = {};
			if (TryGetOperatorInfo(expr.intValue1, operatorInfo)) {
				const bool needBracket = operatorInfo.precedence > expectedLowestPrecedence;
				std::string text;
				if (needBracket) {
					text += "(";
				}
				if (operatorInfo.type == OperatorInfo::Type::Unary) {
					text += operatorInfo.text;
					if (!expr.items.empty() && expr.items[0] != nullptr) {
						text += RenderExpr(*expr.items[0], resolver, operatorInfo.precedence);
					}
				}
				else if (expr.items.empty()) {
					text += " ";
					text += operatorInfo.text;
					text += " ";
				}
				else {
					text += RenderExpr(*expr.items[0], resolver, operatorInfo.precedence);
					text += " ";
					text += operatorInfo.text;
					text += " ";
					if (expr.items.size() >= 2 && expr.items[1] != nullptr) {
						text += RenderExpr(*expr.items[1], resolver, operatorInfo.precedence - 1);
						for (size_t i = 2; i < expr.items.size(); ++i) {
							text += " ";
							text += operatorInfo.text;
							text += " ";
							text += expr.items[i] == nullptr ? std::string() : RenderExpr(*expr.items[i], resolver, operatorInfo.precedence - 1);
						}
					}
				}
				if (needBracket) {
					text += ")";
				}
				return text;
			}
		}

		std::string text;
		if (expr.target != nullptr) {
			text += RenderExpr(*expr.target, resolver);
			text += ".";
		}
		if (expr.flagValue && expr.shortValue1 == -2) {
			std::string ownerName;
			if (resolver.TryGetMethodOwnerName(expr.intValue1, ownerName) && !ownerName.empty()) {
				text += ownerName + ".";
			}
		}
		text += (expr.shortValue1 == -2 || expr.shortValue1 == -3)
			? resolver.ResolveUserName(expr.intValue1)
			: resolver.ResolveLibCmdName(expr.shortValue1, expr.intValue1);
		text += " (" + RenderExprList(expr.items, resolver) + ")";
		return text;
	}
	}

	return std::string();
}

bool ParseParamList(ByteReader& reader, std::vector<std::unique_ptr<Expr>>& outItems, std::string* outError)
{
	outItems.clear();
	while (true) {
		std::unique_ptr<Expr> item;
		if (!ParseExpression(reader, item, outError)) {
			return false;
		}
		if (item != nullptr && item->kind == ExprKind::ParamListEnd) {
			return true;
		}
		outItems.push_back(std::move(item));
	}
}

bool ParseCallExpressionWithoutType(
	ByteReader& reader,
	std::unique_ptr<Expr>& outExpr,
	std::string* outUnexaminedCode,
	std::string* outComment,
	bool* outMask,
	std::string* outError)
{
	outExpr = std::make_unique<Expr>();
	outExpr->kind = ExprKind::Call;

	std::int16_t flag = 0;
	std::string unexaminedCode;
	std::string comment;
	bool unexaminedIsNull = false;
	bool commentIsNull = false;
	if (!reader.ReadI32(outExpr->intValue1) ||
		!reader.ReadI16(outExpr->shortValue1) ||
		!reader.ReadI16(flag) ||
		!reader.ReadBStr(unexaminedCode, unexaminedIsNull) ||
		!reader.ReadBStr(comment, commentIsNull)) {
		if (outError != nullptr) {
			*outError = "call_expression_read_failed";
		}
		return false;
	}

	outExpr->flagValue = (flag & 0x10) != 0;
	if (outMask != nullptr) {
		*outMask = (flag & 0x20) != 0;
	}

	if (outUnexaminedCode != nullptr) {
		if (!unexaminedIsNull) {
			const size_t nullPos = unexaminedCode.find('\0');
			if (nullPos != std::string::npos) {
				unexaminedCode.erase(nullPos);
			}
			if (unexaminedCode.empty()) {
				unexaminedIsNull = true;
			}
		}
		*outUnexaminedCode = unexaminedIsNull ? std::string() : unexaminedCode;
	}
	if (outComment != nullptr) {
		if (!commentIsNull) {
			const size_t nullPos = comment.find('\0');
			if (nullPos != std::string::npos) {
				comment.erase(nullPos);
			}
			if (comment.empty()) {
				commentIsNull = true;
			}
		}
		*outComment = commentIsNull ? std::string() : comment;
	}

	if (!reader.eof()) {
		std::uint8_t marker = 0;
		if (!reader.ReadU8(marker)) {
			if (outError != nullptr) {
				*outError = "call_expression_marker_read_failed";
			}
			return false;
		}
		if (marker == 0x38) {
			if (!reader.SetPosition(reader.position() - 1) ||
				!ParseExpression(reader, outExpr->target, outError) ||
				!ParseParamList(reader, outExpr->items, outError)) {
				return false;
			}
		}
		else if (marker == 0x36) {
			if (!ParseParamList(reader, outExpr->items, outError)) {
				return false;
			}
		}
		else {
			if (outError != nullptr) {
				*outError = "call_expression_marker_invalid";
			}
			return false;
		}
	}

	return true;
}

bool ParseExpression(ByteReader& reader, std::unique_ptr<Expr>& outExpr, std::string* outError, bool parseMember)
{
	outExpr.reset();

	std::uint8_t type = 0;
	while (true) {
		if (!reader.ReadU8(type)) {
			if (outError != nullptr) {
				*outError = "expression_type_read_failed";
			}
			return false;
		}
		if (type != 0x1D && type != 0x37) {
			break;
		}
	}

	auto expr = std::make_unique<Expr>();
	switch (type) {
	case 0x01:
		expr->kind = ExprKind::ParamListEnd;
		break;
	case 0x16:
		expr->kind = ExprKind::DefaultValue;
		break;
	case 0x17:
		expr->kind = ExprKind::Number;
		if (!reader.ReadDouble(expr->numberValue)) {
			if (outError != nullptr) {
				*outError = "number_literal_read_failed";
			}
			return false;
		}
		break;
	case 0x18: {
		std::int16_t raw = 0;
		expr->kind = ExprKind::Bool;
		if (!reader.ReadI16(raw)) {
			if (outError != nullptr) {
				*outError = "bool_literal_read_failed";
			}
			return false;
		}
		expr->boolValue = raw != 0;
		break;
	}
	case 0x19:
		expr->kind = ExprKind::Date;
		if (!reader.ReadDouble(expr->numberValue)) {
			if (outError != nullptr) {
				*outError = "date_literal_read_failed";
			}
			return false;
		}
		break;
	case 0x1A: {
		bool isNull = false;
		expr->kind = ExprKind::String;
		if (!reader.ReadBStr(expr->text, isNull)) {
			if (outError != nullptr) {
				*outError = "string_literal_read_failed";
			}
			return false;
		}
		if (isNull) {
			expr->text.clear();
		}
		break;
	}
	case 0x1B:
		expr->kind = ExprKind::Constant;
		expr->shortValue1 = -2;
		if (!reader.ReadI32(expr->intValue1)) {
			if (outError != nullptr) {
				*outError = "constant_read_failed";
			}
			return false;
		}
		break;
	case 0x1C:
		expr->kind = ExprKind::Constant;
		if (!reader.ReadI16(expr->shortValue1) || !reader.ReadI16(expr->shortValue2)) {
			if (outError != nullptr) {
				*outError = "lib_constant_read_failed";
			}
			return false;
		}
		expr->shortValue1 = static_cast<std::int16_t>(expr->shortValue1 - 1);
		expr->intValue1 = static_cast<std::int32_t>(expr->shortValue2 - 1);
		break;
	case 0x1E:
		expr->kind = ExprKind::MethodPtr;
		if (!reader.ReadI32(expr->intValue1)) {
			if (outError != nullptr) {
				*outError = "method_ptr_read_failed";
			}
			return false;
		}
		break;
	case 0x1F:
		expr->kind = ExprKind::ArrayLiteral;
		while (true) {
			std::unique_ptr<Expr> item;
			if (!ParseExpression(reader, item, outError)) {
				return false;
			}
			if (item != nullptr && item->kind == ExprKind::ArrayLiteralEnd) {
				break;
			}
			expr->items.push_back(std::move(item));
		}
		break;
	case 0x20:
		expr->kind = ExprKind::ArrayLiteralEnd;
		break;
	case 0x21: {
		std::unique_ptr<Expr> callExpr;
		if (!ParseCallExpressionWithoutType(reader, callExpr, nullptr, nullptr, nullptr, outError) || callExpr == nullptr) {
			return false;
		}
		expr = std::move(callExpr);
		break;
	}
	case 0x23:
		expr->kind = ExprKind::EnumConstant;
		if (!reader.ReadI16(expr->shortValue2) ||
			!reader.ReadI16(expr->shortValue1) ||
			!reader.ReadI32(expr->intValue1)) {
			if (outError != nullptr) {
				*outError = "enum_constant_read_failed";
			}
			return false;
		}
		expr->shortValue1 = static_cast<std::int16_t>(expr->shortValue1 - 1);
		expr->shortValue2 = static_cast<std::int16_t>(expr->shortValue2 - 1);
		expr->intValue1 -= 1;
		break;
	case 0x38:
		expr->kind = ExprKind::Variable;
		if (!reader.ReadI32(expr->intValue1)) {
			if (outError != nullptr) {
				*outError = "variable_read_failed";
			}
			return false;
		}
		if (expr->intValue1 == epl_system_id::kIdNaV) {
			std::uint8_t marker = 0;
			if (!reader.ReadU8(marker) || marker != 0x3A) {
				if (outError != nullptr) {
					*outError = "nav_expression_invalid";
				}
				return false;
			}
			return ParseExpression(reader, outExpr, outError, true);
		}
		break;
	case 0x3B:
		expr->kind = ExprKind::Number;
		if (!reader.ReadI32(expr->intValue1)) {
			if (outError != nullptr) {
				*outError = "int_number_read_failed";
			}
			return false;
		}
		expr->numberValue = static_cast<double>(expr->intValue1);
		break;
	default:
		if (outError != nullptr) {
			std::ostringstream stream;
			stream << "unknown_expression_type_"
				<< std::hex << std::uppercase << static_cast<int>(type)
				<< "@0x" << reader.position() - 1
				<< "|ctx=" << BuildReaderHexContext(reader, reader.position() - 1);
			*outError = stream.str();
		}
		return false;
	}

	if (!parseMember) {
		outExpr = std::move(expr);
		return true;
	}

	const bool allowMemberChain =
		expr->kind == ExprKind::Variable ||
		expr->kind == ExprKind::Call ||
		expr->kind == ExprKind::AccessMember ||
		expr->kind == ExprKind::AccessArray;
	while (allowMemberChain && !reader.eof()) {
		std::uint8_t nextType = 0;
		if (!reader.PeekU8(nextType)) {
			break;
		}
		if (nextType == 0x39) {
			reader.ReadU8(nextType);
			auto wrapper = std::make_unique<Expr>();
			wrapper->kind = ExprKind::AccessMember;
			wrapper->target = std::move(expr);
			if (!reader.ReadI32(wrapper->intValue1) || !reader.ReadI32(wrapper->intValue2)) {
				if (outError != nullptr) {
					*outError = "member_expression_read_failed";
				}
				return false;
			}
			if (epl_system_id::IsLibDataType(wrapper->intValue2)) {
				std::int16_t libId = 0;
				std::int16_t typeId = 0;
				epl_system_id::DecomposeLibDataTypeId(wrapper->intValue2, libId, typeId);
				wrapper->shortValue1 = libId;
				wrapper->intValue2 = typeId;
				wrapper->intValue1 -= 1;
			}
			else {
				wrapper->shortValue1 = -2;
			}
			expr = std::move(wrapper);
			continue;
		}
		if (nextType == 0x3A) {
			reader.ReadU8(nextType);
			auto wrapper = std::make_unique<Expr>();
			wrapper->kind = ExprKind::AccessArray;
			wrapper->target = std::move(expr);
			if (!ParseExpression(reader, wrapper->extra, outError, true)) {
				return false;
			}
			expr = std::move(wrapper);
			continue;
		}
		if (nextType == 0x37) {
			reader.ReadU8(nextType);
		}
		break;
	}

	outExpr = std::move(expr);
	return true;
}

bool ReadExpectedMarker(ByteReader& reader, std::uint8_t expected, const char* errorCode, std::string* outError)
{
	std::uint8_t marker = 0;
	if (!reader.ReadU8(marker) || marker != expected) {
		if (outError != nullptr) {
			*outError = errorCode;
		}
		return false;
	}
	return true;
}

std::int32_t GetCallMethodId(const std::unique_ptr<Expr>& expr)
{
	return expr != nullptr && expr->kind == ExprKind::Call ? expr->intValue1 : -1;
}

bool ParseStatementBlock(ByteReader& reader, std::unique_ptr<StatementBlock>& outBlock, std::string* outError)
{
	outBlock = std::make_unique<StatementBlock>();

	static const std::array<std::uint8_t, 14> kKnownTypes = {
		0x50, 0x51, 0x52, 0x53, 0x54, 0x55, 0x6A, 0x6B, 0x6C, 0x6D, 0x6E, 0x6F, 0x70, 0x71
	};

	while (!reader.eof()) {
		std::uint8_t type = 0;
		do {
			if (reader.eof()) {
				return true;
			}
			if (!reader.ReadU8(type)) {
				if (outError != nullptr) {
					*outError = "statement_type_read_failed";
				}
				return false;
			}
		} while (std::find(kKnownTypes.begin(), kKnownTypes.end(), type) == kKnownTypes.end());

		if (type == 0x50 || type == 0x51 || type == 0x52 || type == 0x53 || type == 0x54 || type == 0x6F || type == 0x71) {
			return reader.SetPosition(reader.position() - 1);
		}
		if (type == 0x55) {
			continue;
		}

		if (type == 0x6D) {
			Statement statement;
			statement.kind = StatementKind::SwitchBlock;
			if (!ReadExpectedMarker(reader, 0x6E, "switch_first_case_missing", outError)) {
				return false;
			}

			std::uint8_t nextMarker = 0;
			while (true) {
				SwitchCase caseInfo;
				std::unique_ptr<Expr> callExpr;
				if (!ParseCallExpressionWithoutType(reader, callExpr, &caseInfo.unexaminedCode, &caseInfo.comment, &caseInfo.mask, outError)) {
					return false;
				}
				if (callExpr != nullptr && !callExpr->items.empty()) {
					caseInfo.condition = std::move(callExpr->items.front());
				}
				if (!ParseStatementBlock(reader, caseInfo.block, outError)) {
					return false;
				}
				statement.switchCases.push_back(std::move(caseInfo));
				if (!ReadExpectedMarker(reader, 0x53, "switch_case_end_missing", outError)) {
					return false;
				}
				if (!reader.ReadU8(nextMarker)) {
					if (outError != nullptr) {
						*outError = "switch_next_marker_read_failed";
					}
					return false;
				}
				if (nextMarker != 0x6E) {
					break;
				}
			}
			if (nextMarker != 0x6F) {
				if (outError != nullptr) {
					*outError = "switch_default_missing";
				}
				return false;
			}
			if (!ParseStatementBlock(reader, statement.defaultBlock, outError) ||
				!ReadExpectedMarker(reader, 0x54, "switch_end_missing", outError) ||
				!ReadExpectedMarker(reader, 0x74, "switch_tail_marker_missing", outError)) {
				return false;
			}
			outBlock->items.push_back(std::move(statement));
			continue;
		}

		std::unique_ptr<Expr> callExpr;
		std::string unexaminedCode;
		std::string comment;
		bool mask = false;
		if (!ParseCallExpressionWithoutType(reader, callExpr, &unexaminedCode, &comment, &mask, outError)) {
			return false;
		}

		if (type == 0x70) {
			std::unique_ptr<StatementBlock> loopBody;
			if (!ParseStatementBlock(reader, loopBody, outError) ||
				!ReadExpectedMarker(reader, 0x71, "loop_end_call_missing", outError)) {
				return false;
			}

			std::unique_ptr<Expr> endCallExpr;
			std::string endUnexaminedCode;
			std::string endComment;
			bool endMask = false;
			if (!ParseCallExpressionWithoutType(reader, endCallExpr, &endUnexaminedCode, &endComment, &endMask, outError)) {
				return false;
			}

			Statement statement;
			statement.mask = mask;
			statement.maskOnEnd = endMask;
			statement.comment = std::move(comment);
			statement.commentOnEnd = std::move(endComment);
			statement.unexaminedCode = std::move(unexaminedCode);
			statement.block = std::move(loopBody);

			switch (GetCallMethodId(callExpr)) {
			case 3:
				statement.kind = StatementKind::WhileLoop;
				if (!callExpr->items.empty()) {
					statement.condition = std::move(callExpr->items.front());
				}
				break;
			case 5:
				statement.kind = StatementKind::DoWhileLoop;
				statement.unexaminedCode = std::move(endUnexaminedCode);
				statement.commentOnEnd.clear();
				statement.maskOnEnd = false;
				if (endCallExpr != nullptr && !endCallExpr->items.empty()) {
					statement.condition = std::move(endCallExpr->items.front());
				}
				break;
			case 7:
				statement.kind = StatementKind::CounterLoop;
				if (!callExpr->items.empty()) {
					statement.condition = std::move(callExpr->items[0]);
				}
				if (callExpr->items.size() >= 2) {
					statement.expr2 = std::move(callExpr->items[1]);
				}
				break;
			case 9:
				statement.kind = StatementKind::ForLoop;
				if (!callExpr->items.empty()) {
					statement.condition = std::move(callExpr->items[0]);
				}
				if (callExpr->items.size() >= 2) {
					statement.expr2 = std::move(callExpr->items[1]);
				}
				if (callExpr->items.size() >= 3) {
					statement.expr3 = std::move(callExpr->items[2]);
				}
				if (callExpr->items.size() >= 4) {
					statement.expr4 = std::move(callExpr->items[3]);
				}
				break;
			default:
				if (outError != nullptr) {
					*outError = "unknown_loop_type";
				}
				return false;
			}
			outBlock->items.push_back(std::move(statement));
			continue;
		}

		if (type == 0x6C) {
			Statement statement;
			statement.kind = StatementKind::IfTrue;
			statement.mask = mask;
			statement.comment = std::move(comment);
			statement.unexaminedCode = std::move(unexaminedCode);
			if (callExpr != nullptr && !callExpr->items.empty()) {
				statement.condition = std::move(callExpr->items.front());
			}
			if (!ParseStatementBlock(reader, statement.block, outError) ||
				!ReadExpectedMarker(reader, 0x52, "if_true_end_missing", outError) ||
				!ReadExpectedMarker(reader, 0x73, "if_true_tail_marker_missing", outError)) {
				return false;
			}
			outBlock->items.push_back(std::move(statement));
			continue;
		}

		if (type == 0x6B) {
			Statement statement;
			statement.kind = StatementKind::IfElse;
			statement.mask = mask;
			statement.comment = std::move(comment);
			statement.unexaminedCode = std::move(unexaminedCode);
			if (callExpr != nullptr && !callExpr->items.empty()) {
				statement.condition = std::move(callExpr->items.front());
			}
			if (!ParseStatementBlock(reader, statement.block, outError) ||
				!ReadExpectedMarker(reader, 0x50, "if_else_marker_missing", outError) ||
				!ParseStatementBlock(reader, statement.elseBlock, outError) ||
				!ReadExpectedMarker(reader, 0x51, "if_end_missing", outError) ||
				!ReadExpectedMarker(reader, 0x72, "if_tail_marker_missing", outError)) {
				return false;
			}
			outBlock->items.push_back(std::move(statement));
			continue;
		}

		if (type == 0x6A) {
			Statement statement;
			statement.mask = mask;
			statement.comment = std::move(comment);
			if (!unexaminedCode.empty()) {
				statement.kind = StatementKind::Unexamined;
				statement.unexaminedCode = std::move(unexaminedCode);
			}
			else {
				statement.kind = StatementKind::Expression;
				if (callExpr != nullptr && callExpr->shortValue1 != -1) {
					statement.expression = std::move(callExpr);
				}
			}
			outBlock->items.push_back(std::move(statement));
			continue;
		}

		if (outError != nullptr) {
			*outError = "statement_type_not_supported";
		}
		return false;
	}

	return true;
}

void AppendStatementLines(const StatementBlock& block, SymbolResolver& resolver, int indent, std::vector<std::string>& outLines)
{
	for (const auto& statement : block.items) {
		switch (statement.kind) {
		case StatementKind::Expression: {
			std::string line = BuildIndent(indent);
			if (statement.mask) {
				line += "' ";
			}
			if (statement.expression != nullptr) {
				line += RenderExpr(*statement.expression, resolver);
			}
			if (!statement.comment.empty()) {
				line += statement.expression == nullptr ? "' " : "  ' ";
				line += statement.comment;
			}
			outLines.push_back(line);
			break;
		}
		case StatementKind::Unexamined: {
			std::string line = BuildIndent(indent);
			if (statement.mask) {
				line += "' ";
			}
			line += statement.unexaminedCode;
			outLines.push_back(line);
			break;
		}
		case StatementKind::IfTrue: {
			std::string line = BuildIndent(indent);
			if (statement.mask) {
				line += "' ";
			}
			line += statement.unexaminedCode.empty()
				? ".如果真 (" + (statement.condition == nullptr ? std::string() : RenderExpr(*statement.condition, resolver)) + ")"
				: "." + statement.unexaminedCode;
			if (!statement.comment.empty()) {
				line += "  ' " + statement.comment;
			}
			outLines.push_back(line);
			if (statement.block != nullptr) {
				AppendStatementLines(*statement.block, resolver, indent + 1, outLines);
			}
			outLines.push_back(BuildIndent(indent) + (statement.mask ? "' " : "") + ".如果真结束");
			break;
		}
		case StatementKind::IfElse: {
			std::string line = BuildIndent(indent);
			if (statement.mask) {
				line += "' ";
			}
			line += statement.unexaminedCode.empty()
				? ".如果 (" + (statement.condition == nullptr ? std::string() : RenderExpr(*statement.condition, resolver)) + ")"
				: "." + statement.unexaminedCode;
			if (!statement.comment.empty()) {
				line += "  ' " + statement.comment;
			}
			outLines.push_back(line);
			if (statement.block != nullptr) {
				AppendStatementLines(*statement.block, resolver, indent + 1, outLines);
			}
			outLines.push_back(BuildIndent(indent) + (statement.mask ? "' " : "") + ".否则");
			if (statement.elseBlock != nullptr) {
				AppendStatementLines(*statement.elseBlock, resolver, indent + 1, outLines);
			}
			outLines.push_back(BuildIndent(indent) + (statement.mask ? "' " : "") + ".如果结束");
			break;
		}
		case StatementKind::WhileLoop: {
			std::string line = BuildIndent(indent);
			if (statement.mask) {
				line += "' ";
			}
			line += statement.unexaminedCode.empty()
				? ".判断循环首 (" + (statement.condition == nullptr ? std::string() : RenderExpr(*statement.condition, resolver)) + ")"
				: "." + statement.unexaminedCode;
			if (!statement.comment.empty()) {
				line += "  ' " + statement.comment;
			}
			outLines.push_back(line);
			if (statement.block != nullptr) {
				AppendStatementLines(*statement.block, resolver, indent + 1, outLines);
			}
			std::string endLine = BuildIndent(indent);
			if (statement.maskOnEnd) {
				endLine += "' ";
			}
			endLine += ".判断循环尾 ()";
			if (!statement.commentOnEnd.empty()) {
				endLine += "  ' " + statement.commentOnEnd;
			}
			outLines.push_back(endLine);
			break;
		}
		case StatementKind::DoWhileLoop: {
			std::string line = BuildIndent(indent);
			if (statement.mask) {
				line += "' ";
			}
			line += ".循环判断首 ()";
			if (!statement.comment.empty()) {
				line += "  ' " + statement.comment;
			}
			outLines.push_back(line);
			if (statement.block != nullptr) {
				AppendStatementLines(*statement.block, resolver, indent + 1, outLines);
			}
			std::string endLine = BuildIndent(indent);
			if (statement.mask) {
				endLine += "' ";
			}
			endLine += statement.unexaminedCode.empty()
				? ".循环判断尾 (" + (statement.condition == nullptr ? std::string() : RenderExpr(*statement.condition, resolver)) + ")"
				: "." + statement.unexaminedCode;
			if (!statement.comment.empty()) {
				endLine += "  ' " + statement.comment;
			}
			outLines.push_back(endLine);
			break;
		}
		case StatementKind::CounterLoop: {
			std::string line = BuildIndent(indent);
			if (statement.mask) {
				line += "' ";
			}
			line += statement.unexaminedCode.empty()
				? ".计次循环首 (" +
					(statement.condition == nullptr ? std::string() : RenderExpr(*statement.condition, resolver)) +
					", " +
					(statement.expr2 == nullptr ? std::string() : RenderExpr(*statement.expr2, resolver)) +
					")"
				: "." + statement.unexaminedCode;
			if (!statement.comment.empty()) {
				line += "  ' " + statement.comment;
			}
			outLines.push_back(line);
			if (statement.block != nullptr) {
				AppendStatementLines(*statement.block, resolver, indent + 1, outLines);
			}
			std::string endLine = BuildIndent(indent);
			if (statement.maskOnEnd) {
				endLine += "' ";
			}
			endLine += ".计次循环尾 ()";
			if (!statement.commentOnEnd.empty()) {
				endLine += "  ' " + statement.commentOnEnd;
			}
			outLines.push_back(endLine);
			break;
		}
		case StatementKind::ForLoop: {
			std::string line = BuildIndent(indent);
			if (statement.mask) {
				line += "' ";
			}
			line += statement.unexaminedCode.empty()
				? ".变量循环首 (" +
					(statement.condition == nullptr ? std::string() : RenderExpr(*statement.condition, resolver)) +
					", " +
					(statement.expr2 == nullptr ? std::string() : RenderExpr(*statement.expr2, resolver)) +
					", " +
					(statement.expr3 == nullptr ? std::string() : RenderExpr(*statement.expr3, resolver)) +
					", " +
					(statement.expr4 == nullptr ? std::string() : RenderExpr(*statement.expr4, resolver)) +
					")"
				: "." + statement.unexaminedCode;
			if (!statement.comment.empty()) {
				line += "  ' " + statement.comment;
			}
			outLines.push_back(line);
			if (statement.block != nullptr) {
				AppendStatementLines(*statement.block, resolver, indent + 1, outLines);
			}
			std::string endLine = BuildIndent(indent);
			if (statement.maskOnEnd) {
				endLine += "' ";
			}
			endLine += ".变量循环尾 ()";
			if (!statement.commentOnEnd.empty()) {
				endLine += "  ' " + statement.commentOnEnd;
			}
			outLines.push_back(endLine);
			break;
		}
		case StatementKind::SwitchBlock: {
			if (statement.switchCases.empty()) {
				break;
			}
			for (size_t i = 0; i < statement.switchCases.size(); ++i) {
				const auto& caseItem = statement.switchCases[i];
				std::string line = BuildIndent(indent);
				if (caseItem.mask) {
					line += "' ";
				}
				if (caseItem.unexaminedCode.empty()) {
					line += i == 0 ? ".判断开始 (" : ".判断 (";
					line += caseItem.condition == nullptr ? std::string() : RenderExpr(*caseItem.condition, resolver);
					line += ")";
				}
				else if (i == 0) {
					line += ".判断开始";
					std::string suffix = caseItem.unexaminedCode;
					if (suffix.rfind("判断", 0) == 0) {
						suffix.erase(0, std::string("判断").size());
					}
					line += suffix;
				}
				else {
					line += "." + caseItem.unexaminedCode;
				}
				if (!caseItem.comment.empty()) {
					line += "  ' " + caseItem.comment;
				}
				outLines.push_back(line);
				if (caseItem.block != nullptr) {
					AppendStatementLines(*caseItem.block, resolver, indent + 1, outLines);
				}
			}
			outLines.push_back(BuildIndent(indent) + (statement.switchCases[0].mask ? "' " : "") + ".默认");
			if (statement.defaultBlock != nullptr) {
				AppendStatementLines(*statement.defaultBlock, resolver, indent + 1, outLines);
			}
			outLines.push_back(BuildIndent(indent) + (statement.switchCases[0].mask ? "' " : "") + ".判断结束");
			break;
		}
		}
	}
}

bool TryRenderFunctionBody(const FunctionInfo& functionInfo, SymbolResolver& resolver, std::vector<std::string>& outLines, std::string* outError)
{
	outLines.clear();
	if (outError != nullptr) {
		outError->clear();
	}
	if (functionInfo.expressionData.empty()) {
		return true;
	}

	ByteReader reader(functionInfo.expressionData);
	std::unique_ptr<StatementBlock> block;
	std::string parseError;
	if (!ParseStatementBlock(reader, block, &parseError) || block == nullptr) {
		if (outError != nullptr) {
			*outError = parseError.empty() ? "function_body_parse_failed" : parseError;
		}
		return false;
	}

	AppendStatementLines(*block, resolver, 0, outLines);
	return true;
}

bool IsProgramPagePublic(const ModuleSections& sections, std::int32_t pageId)
{
	for (const auto& item : sections.classPublicities) {
		if (item.classId == pageId) {
			return (item.flags & 0x1) != 0;
		}
	}
	return false;
}

void BuildProgramPages(const ModuleSections& sections, const GenerateOptions& options, Document& outDocument)
{
	SymbolResolver resolver(sections.program, sections.resources, outDocument.sourcePath);
	bool firstProgramPage = true;
	TraceLine("BuildProgramPages begin");
	for (const auto& pageInfo : sections.program.codePages) {
		const auto functions = CollectPageFunctions(sections.program, pageInfo);
		if (!ShouldKeepPage(pageInfo, functions, options.includeImportedPages)) {
			continue;
		}
		TraceLine("BuildProgramPages page=" + TrimAsciiCopy(pageInfo.name) + " funcs=" + std::to_string(functions.size()));

		Page page;
		page.typeName = "程序集";
		page.name = TrimAsciiCopy(pageInfo.name);
		AppendLine(page, ".版本 2");
		AppendLine(page, "");
		if (firstProgramPage) {
			for (const auto& item : outDocument.dependencies) {
				if (item.kind == DependencyKind::ELib && !item.name.empty()) {
					AppendLine(page, ".支持库 " + item.name);
				}
			}
			if (!page.lines.empty() && page.lines.back() != "") {
				AppendLine(page, "");
			}
			firstProgramPage = false;
		}

		std::string baseClassName;
		if (pageInfo.baseClass != 0 && pageInfo.baseClass != -1) {
			baseClassName = TrimAsciiCopy(resolver.ResolveType(pageInfo.baseClass));
		}
		if (baseClassName == "对象") {
			baseClassName.clear();
		}
		AppendLine(page, BuildDefinitionLine(
			"程序集",
			{
				page.name,
				baseClassName,
				IsProgramPagePublic(sections, pageInfo.header.dwId) ? "公开" : std::string(),
				TrimAsciiCopy(pageInfo.comment),
			}));
		AppendLine(page, "");

		for (const auto& pageVar : pageInfo.pageVars) {
			AppendLine(page, BuildClassVariableLine(pageVar, resolver));
		}
		if (!pageInfo.pageVars.empty()) {
			AppendLine(page, "");
		}

		for (const auto* functionInfo : functions) {
			if (functionInfo == nullptr) {
				continue;
			}
			AppendLine(page, BuildDefinitionLine(
				"子程序",
				{
					TrimAsciiCopy(functionInfo->name),
					TrimAsciiCopy(resolver.ResolveType(functionInfo->returnType)),
					(functionInfo->attr & 0x8) != 0 ? "公开" : std::string(),
					TrimAsciiCopy(functionInfo->comment),
				}));

			for (const auto& param : functionInfo->params) {
				AppendLine(page, BuildMethodParameterLine(param, resolver));
			}

			for (const auto& local : functionInfo->locals) {
				AppendLine(page, BuildLocalVariableLine(local, resolver));
			}

			std::vector<std::string> bodyLines;
			std::string bodyError;
			if (TryRenderFunctionBody(*functionInfo, resolver, bodyLines, &bodyError)) {
				if (!bodyLines.empty()) {
					AppendLine(page, "");
					for (const auto& bodyLine : bodyLines) {
						AppendLine(page, bodyLine);
					}
				}
			}
			else {
				TraceLine("BuildProgramPages body_parse_failed func=" + TrimAsciiCopy(functionInfo->name) + " error=" + bodyError);
				AppendLine(page, "");
				AppendLine(page, "' native_parse_failed: " + bodyError);
			}

			AppendLine(page, "");
			AppendLine(page, "");
		}

		outDocument.pages.push_back(std::move(page));
	}
	TraceLine("BuildProgramPages end count=" + std::to_string(outDocument.pages.size()));
}

void BuildGlobalPage(const ModuleSections& sections, Document& outDocument)
{
	if (sections.program.globals.empty()) {
		return;
	}

	SymbolResolver resolver(sections.program, sections.resources, outDocument.sourcePath);
	Page page;
	page.typeName = "全局变量";
	page.name = "全局变量";
	AppendLine(page, ".版本 2");
	AppendLine(page, "");
	for (const auto& item : sections.program.globals) {
		AppendLine(page, BuildGlobalVariableLine(item, resolver));
	}
	outDocument.pages.push_back(std::move(page));
}

void BuildStructPage(const ModuleSections& sections, const AnonymousTypeHints& hints, Document& outDocument)
{
	if (sections.program.dataTypes.empty()) {
		return;
	}

	SymbolResolver resolver(sections.program, sections.resources, outDocument.sourcePath);
	Page page;
	page.typeName = "自定义数据类型";
	page.name = "自定义数据类型";
	AppendLine(page, ".版本 2");
	AppendLine(page, "");
	for (const auto& item : sections.program.dataTypes) {
		std::string typeName = TrimAsciiCopy(item.name);
		if (typeName.empty()) {
			if (const auto it = hints.localTypeAliases.find(item.header.dwId); it != hints.localTypeAliases.end()) {
				typeName = it->second;
			}
		}
		AppendLine(page, BuildDefinitionLine(
			"数据类型",
			{
				typeName,
				(item.attr & 0x2) != 0 ? "公开" : std::string(),
				TrimAsciiCopy(item.comment),
			}));
		for (const auto& member : item.members) {
			const std::string resolvedType = ResolveMemberTypeWithHints(member, typeName, resolver, hints);
			AppendLine(page, BuildStructMemberLine(member, resolvedType));
		}
		AppendLine(page, "");
	}
	outDocument.pages.push_back(std::move(page));
}

void BuildDllPage(const ModuleSections& sections, const AnonymousTypeHints& hints, Document& outDocument)
{
	if (sections.program.dlls.empty()) {
		return;
	}

	SymbolResolver resolver(sections.program, sections.resources, outDocument.sourcePath);
	Page page;
	page.typeName = "DLL命令";
	page.name = "Dll命令";
	AppendLine(page, ".版本 2");
	AppendLine(page, "");
	for (const auto& item : sections.program.dlls) {
		const std::string dllName = TrimAsciiCopy(item.name);
		AppendLine(page, BuildDefinitionLine(
			"DLL命令",
			{
				dllName,
				TrimAsciiCopy(resolver.ResolveType(item.returnType)),
				QuoteIfNotEmpty(item.fileName),
				QuoteIfNotEmpty(item.commandName),
				(item.attr & 0x2) != 0 ? "公开" : std::string(),
				TrimAsciiCopy(item.comment),
			}));
		for (const auto& param : item.params) {
			std::string typeText = ResolveDllParamTypeWithHints(param, dllName, resolver, hints);
			AppendLine(page, BuildDllParameterLine(param, typeText));
		}
		AppendLine(page, "");
	}
	outDocument.pages.push_back(std::move(page));
}

void BuildConstantPage(const ModuleSections& sections, Document& outDocument)
{
	const bool hasValueConstants = std::any_of(
		sections.resources.constants.begin(),
		sections.resources.constants.end(),
		[](const ConstantInfo& item) { return item.pageType == kConstPageValue; });
	if (!hasValueConstants) {
		return;
	}

	Page page;
	page.typeName = "常量资源";
	page.name = "常量表...";
	AppendLine(page, ".版本 2");
	AppendLine(page, "");
	for (const auto& item : sections.resources.constants) {
		if (item.pageType != kConstPageValue || item.valueText == "<图片>" || item.valueText == "<声音>") {
			continue;
		}
		std::string valueText = item.valueText;
		if (item.longText && item.textByteLength > 64) {
			valueText = "<文本长度: " + std::to_string(item.textByteLength) + ">";
		}
		else if (!valueText.empty() &&
			valueText.find_first_of("eE") != std::string::npos &&
			valueText.find_first_not_of("0123456789+-.eE") == std::string::npos) {
			try {
				valueText = FormatNumberLiteral(std::stod(valueText));
			}
			catch (...) {
			}
		}
		AppendLine(page, BuildDefinitionLine(
			"常量",
			{
				TrimAsciiCopy(item.name),
				Quote(valueText),
				(item.attr & 0x2) != 0 ? "公开" : std::string(),
				TrimAsciiCopy(item.comment),
			}));
	}
	outDocument.pages.push_back(std::move(page));
}

void BuildFormPage(const ModuleSections& sections, Document& outDocument)
{
	if (sections.resources.forms.empty()) {
		return;
	}

	Page page;
	page.typeName = "窗口/表单";
	page.name = "窗口";
	AppendLine(page, ".版本 2");
	AppendLine(page, "");
	for (const auto& item : sections.resources.forms) {
		AppendLine(page, BuildDefinitionLine("窗口", { TrimAsciiCopy(item.name), TrimAsciiCopy(item.comment) }));
	}
	outDocument.pages.push_back(std::move(page));
}

}  // namespace

bool Generator::GenerateDocument(const std::string& inputPath, Document& outDocument, std::string* outError) const
{
	return GenerateDocumentInternal(inputPath, {}, outDocument, outError);
}

bool Generator::GenerateDocumentInternal(
	const std::string& inputPath,
	const GenerateOptions& options,
	Document& outDocument,
	std::string* outError) const
{
	TraceLine("GenerateDocumentInternal begin input=" + inputPath);
	if (outError != nullptr) {
		outError->clear();
	}

	ModuleSections sections;
	if (!ParseModuleSections(inputPath, sections, outError)) {
		TraceLine("GenerateDocumentInternal parse_failed");
		return false;
	}
	TraceLine(
		"GenerateDocumentInternal parsed code_pages=" + std::to_string(sections.program.codePages.size()) +
		" functions=" + std::to_string(sections.program.functions.size()) +
		" globals=" + std::to_string(sections.program.globals.size()) +
		" data_types=" + std::to_string(sections.program.dataTypes.size()) +
		" dlls=" + std::to_string(sections.program.dlls.size()) +
		" forms=" + std::to_string(sections.resources.forms.size()) +
		" constants=" + std::to_string(sections.resources.constants.size()));

	Document document;
	document.sourcePath = inputPath;
	document.projectName = sections.hasUserInfo && !TrimAsciiCopy(sections.userInfo.programName).empty()
		? TrimAsciiCopy(sections.userInfo.programName)
		: std::filesystem::path(inputPath).stem().string();
	document.versionText = BuildVersionText(sections);

	BuildDependencies(sections, document);
	const AnonymousTypeHints anonymousTypeHints = BuildAnonymousTypeHints(sections, inputPath);
	BuildProgramPages(sections, options, document);
	BuildGlobalPage(sections, document);
	BuildStructPage(sections, anonymousTypeHints, document);
	BuildDllPage(sections, anonymousTypeHints, document);
	BuildFormPage(sections, document);
	BuildConstantPage(sections, document);
	TraceLine("GenerateDocumentInternal pages_done count=" + std::to_string(document.pages.size()));

	outDocument = std::move(document);
	TraceLine("GenerateDocumentInternal end");
	return true;
}

bool Generator::GenerateText(const Document& document, std::string& outText, std::string* outError) const
{
	if (outError != nullptr) {
		outError->clear();
	}

	std::ostringstream stream;
	stream << "e2txt Generated Dump\r\n";
	stream << "source=" << document.sourcePath << "\r\n";
	if (!document.outputPath.empty()) {
		stream << "output=" << document.outputPath << "\r\n";
	}
	stream << "project=" << document.projectName << "\r\n";
	stream << "version=" << document.versionText << "\r\n";
	stream << "total_pages=" << document.pages.size() << "\r\n";

	if (!document.dependencies.empty()) {
		stream << "\r\n";
		stream << "================================================================================\r\n";
		stream << "[Dependencies]\r\n";
		stream << "--------------------------------------------------------------------------------\r\n";
		for (const auto& item : document.dependencies) {
			if (item.kind == DependencyKind::ELib) {
				stream << "ELib name=" << item.name
					<< " file=" << item.fileName
					<< " guid=" << item.guid
					<< " version=" << item.versionText << "\r\n";
			}
			else {
				stream << "ECom name=" << item.name
					<< " path=" << item.path
					<< " re_export=" << (item.reExport ? "true" : "false") << "\r\n";
			}
		}
	}

	for (size_t i = 0; i < document.pages.size(); ++i) {
		const auto& page = document.pages[i];
		stream << "\r\n";
		stream << "================================================================================\r\n";
		stream << "[" << (i + 1) << "/" << document.pages.size() << "] type=" << page.typeName << " name=" << page.name << "\r\n";
		stream << "--------------------------------------------------------------------------------\r\n";
		for (const auto& line : page.lines) {
			stream << line << "\r\n";
		}
	}

	outText = stream.str();
	return true;
}

bool Generator::GenerateToFile(
	const std::string& inputPath,
	const std::string& outputPath,
	std::string* outSummary,
	std::string* outError,
	const GenerateOptions& options) const
{
	TraceLine("GenerateToFile begin");
	if (outError != nullptr) {
		outError->clear();
	}
	if (outSummary != nullptr) {
		outSummary->clear();
	}

	Document document;
	if (!GenerateDocumentInternal(inputPath, options, document, outError)) {
		TraceLine("GenerateToFile generate_document_failed");
		return false;
	}

	document.outputPath = outputPath;
	std::string text;
	if (!GenerateText(document, text, outError)) {
		TraceLine("GenerateToFile generate_text_failed");
		return false;
	}

	std::ofstream out(outputPath, std::ios::binary);
	if (!out.is_open()) {
		if (outError != nullptr) {
			*outError = "open_output_file_failed";
		}
		TraceLine("GenerateToFile open_output_failed");
		return false;
	}
	out.write(text.data(), static_cast<std::streamsize>(text.size()));
	out.close();
	TraceLine("GenerateToFile write_done");

	if (outSummary != nullptr) {
		*outSummary = "pages=" + std::to_string(document.pages.size()) + ", output=" + outputPath;
	}
	return true;
}

}  // namespace e2txt
