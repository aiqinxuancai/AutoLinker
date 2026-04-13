#include "e2txt.h"

#include <Windows.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <charconv>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <limits>
#include <optional>
#include <sstream>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>

#include <lib2.h>

#include "PathHelper.h"

namespace e2txt {

namespace {

constexpr std::uint32_t kMagicFileHeader1 = 1415007811u;
constexpr std::uint32_t kMagicFileHeader2 = 1196576837u;
constexpr std::uint32_t kMagicSection = 353465113u;
constexpr std::uint32_t kSectionEndOfFile = 0x07007319u;

constexpr std::uint32_t kSectionSystemInfo = 0x02007319u;
constexpr std::uint32_t kSectionProjectConfig = 0x01007319u;
constexpr std::uint32_t kSectionResource = 0x04007319u;
constexpr std::uint32_t kSectionCode = 0x03007319u;
constexpr std::uint32_t kSectionLosable = 0x05007319u;
constexpr std::uint32_t kSectionClassPublicity = 0x0B007319u;
constexpr std::uint32_t kSectionEcDependencies = 0x0C007319u;

constexpr std::array<std::uint8_t, 4> kSectionNameNoKey = { 25, 115, 0, 7 };

constexpr std::int32_t kProgramHeaderVersionFlag1 = 66279;
constexpr std::int32_t kProgramHeaderUnk1 = 51113791;

constexpr std::uint8_t kConstTypeEmpty = 22;
constexpr std::uint8_t kConstTypeNumber = 23;
constexpr std::uint8_t kConstTypeBool = 24;
constexpr std::uint8_t kConstTypeDate = 25;
constexpr std::uint8_t kConstTypeText = 26;

constexpr std::int16_t kVarAttrStatic = 0x0001;
constexpr std::int16_t kVarAttrByRef = 0x0002;
constexpr std::int16_t kVarAttrNullable = 0x0004;
constexpr std::int16_t kVarAttrArray = 0x0008;
constexpr std::int16_t kConstAttrPublic = 0x0002;
constexpr std::int16_t kConstAttrLongText = 0x0010;
constexpr std::int16_t kGlobalAttrPublic = 0x0100;

constexpr std::int32_t kConstPageValue = 1;
constexpr std::int32_t kConstPageImage = 2;
constexpr std::int32_t kConstPageSound = 3;

struct RestoreDependencyInfo {
	std::string name;
	std::string fileName;
	std::string guid;
	std::string versionText;
	std::string path;
	bool reExport = false;
	bool isSupportLibrary = false;
};

struct RestoreVariable {
	std::int32_t id = 0;
	std::int32_t dataType = 0;
	std::int16_t attr = 0;
	std::vector<std::int32_t> arrayBounds;
	std::string name;
	std::string comment;
};

struct RestoreMethod {
	std::int32_t id = 0;
	std::int32_t memoryAddress = 0;
	std::int32_t ownerClass = 0;
	std::int32_t attr = 0;
	std::int32_t returnType = 0;
	std::string name;
	std::string comment;
	std::vector<RestoreVariable> params;
	std::vector<RestoreVariable> locals;
	std::vector<std::uint8_t> lineOffset;
	std::vector<std::uint8_t> blockOffset;
	std::vector<std::uint8_t> methodReference;
	std::vector<std::uint8_t> variableReference;
	std::vector<std::uint8_t> constantReference;
	std::vector<std::uint8_t> expressionData;
};

struct RestoreClass {
	std::int32_t id = 0;
	std::int32_t memoryAddress = 0;
	std::int32_t formId = 0;
	std::int32_t baseClass = -1;
	std::string name;
	std::string comment;
	std::vector<std::int32_t> functionIds;
	std::vector<RestoreVariable> vars;
	bool isFormClass = false;
	bool isPublic = false;
};

struct RestoreStruct {
	std::int32_t id = 0;
	std::int32_t memoryAddress = 0;
	std::int32_t attr = 0;
	std::string name;
	std::string comment;
	std::vector<RestoreVariable> members;
	bool isPlaceholder = false;
};

struct RestoreDll {
	std::int32_t id = 0;
	std::int32_t memoryAddress = 0;
	std::int32_t attr = 0;
	std::int32_t returnType = 0;
	std::string name;
	std::string comment;
	std::string fileName;
	std::string commandName;
	std::vector<RestoreVariable> params;
};

struct RestoreConstant {
	std::int32_t id = 0;
	std::int16_t attr = 0;
	std::int32_t pageType = kConstPageValue;
	std::string name;
	std::string comment;
	std::string valueText;
};

struct RestoreFormElement {
	std::int32_t id = 0;
	std::int32_t dataType = 0;
	bool isMenu = false;
	std::string name;
	bool visible = true;
	bool disable = false;

	std::string comment;
	std::int32_t cWndAddress = 0;
	std::int32_t left = 0;
	std::int32_t top = 0;
	std::int32_t width = 0;
	std::int32_t height = 0;
	std::int32_t unknownBeforeParent = 0;
	std::int32_t parent = 0;
	std::vector<std::int32_t> children;
	std::vector<std::uint8_t> cursor;
	std::string tag;
	std::int32_t unknownBeforeVisible = 0;
	bool tabStop = true;
	bool locked = false;
	std::int32_t tabIndex = 0;
	std::vector<std::pair<std::int32_t, std::int32_t>> events;
	std::vector<std::uint8_t> extensionData;

	std::int32_t hotKey = 0;
	std::int32_t level = 0;
	bool selected = false;
	std::string text;
	std::int32_t clickEvent = 0;
};

struct RestoreForm {
	std::int32_t id = 0;
	std::int32_t memoryAddress = 0;
	std::int32_t unknown1 = 0;
	std::int32_t classId = 0;
	std::string name;
	std::string comment;
	std::vector<RestoreFormElement> elements;
};

struct RestoreDocumentModel {
	std::string sourcePath;
	std::string projectName;
	std::string versionText;
	std::vector<RestoreDependencyInfo> dependencies;
	std::vector<RestoreClass> classes;
	std::vector<RestoreMethod> methods;
	std::vector<RestoreVariable> globals;
	std::vector<RestoreStruct> structs;
	std::vector<RestoreDll> dlls;
	std::vector<RestoreConstant> constants;
	std::vector<RestoreForm> forms;
};

namespace epl_system_id {

constexpr std::int32_t kTypeMethod = 0x04000000;
constexpr std::int32_t kTypeGlobal = 0x05000000;
constexpr std::int32_t kTypeFormSelf = 0x06000000;
constexpr std::int32_t kTypeStaticClass = 0x09000000;
constexpr std::int32_t kTypeDll = 0x0A000000;
constexpr std::int32_t kTypeClassMember = 0x15000000;
constexpr std::int32_t kTypeFormControl = 0x16000000;
constexpr std::int32_t kTypeConstant = 0x18000000;
constexpr std::int32_t kTypeFormClass = 0x19000000;
constexpr std::int32_t kTypeLocal = 0x25000000;
constexpr std::int32_t kTypeFormMenu = 0x26000000;
constexpr std::int32_t kTypeStructMember = 0x35000000;
constexpr std::int32_t kTypeStruct = 0x41000000;
constexpr std::int32_t kTypeDllParameter = 0x45000000;
constexpr std::int32_t kTypeForm = 0x52000000;

constexpr std::int32_t kMaskType = static_cast<std::int32_t>(0xFF000000u);
constexpr std::int32_t kMaskNum = 0x00FFFFFF;

inline std::int32_t GetType(const std::int32_t id)
{
	return id & kMaskType;
}

inline bool IsLibDataType(const std::int32_t id)
{
	return (id & kMaskType) == 0 && id != 0;
}

}  // namespace epl_system_id

class ByteWriter {
public:
	void WriteU8(const std::uint8_t value)
	{
		m_bytes.push_back(value);
	}

	void WriteI16(const std::int16_t value)
	{
		WritePod(value);
	}

	void WriteU16(const std::uint16_t value)
	{
		WritePod(value);
	}

	void WriteI32(const std::int32_t value)
	{
		WritePod(value);
	}

	void WriteU32(const std::uint32_t value)
	{
		WritePod(value);
	}

	void WriteI64(const std::int64_t value)
	{
		WritePod(value);
	}

	void WriteDouble(const double value)
	{
		WritePod(value);
	}

	void WriteBool32(const bool value)
	{
		WriteI32(value ? 1 : 0);
	}

	void WriteRaw(const void* data, const size_t size)
	{
		if (size == 0) {
			return;
		}
		const auto* begin = static_cast<const std::uint8_t*>(data);
		m_bytes.insert(m_bytes.end(), begin, begin + static_cast<std::ptrdiff_t>(size));
	}

	void WriteBytes(const std::vector<std::uint8_t>& data)
	{
		if (!data.empty()) {
			m_bytes.insert(m_bytes.end(), data.begin(), data.end());
		}
	}

	void WriteDynamicBytes(const std::vector<std::uint8_t>& data)
	{
		WriteI32(static_cast<std::int32_t>(data.size()));
		WriteBytes(data);
	}

	void WriteDynamicText(const std::string& text)
	{
		WriteI32(static_cast<std::int32_t>(text.size()));
		WriteRaw(text.data(), text.size());
	}

	void WriteStandardText(const std::string& text)
	{
		WriteRaw(text.data(), text.size());
		WriteU8(0);
	}

	void WriteBStr(const std::optional<std::string>& text)
	{
		if (!text.has_value()) {
			WriteI32(0);
			return;
		}
		WriteI32(static_cast<std::int32_t>(text->size() + 1));
		WriteRaw(text->data(), text->size());
		WriteU8(0);
	}

	void WriteTextArray(const std::vector<std::string>& values)
	{
		WriteI16(static_cast<std::int16_t>(values.size()));
		for (const auto& value : values) {
			WriteDynamicText(value);
		}
	}

	size_t position() const
	{
		return m_bytes.size();
	}

	void PatchI32(const size_t offset, const std::int32_t value)
	{
		if (offset + sizeof(value) > m_bytes.size()) {
			return;
		}
		std::memcpy(m_bytes.data() + offset, &value, sizeof(value));
	}

	std::vector<std::uint8_t> TakeBytes()
	{
		return std::move(m_bytes);
	}

	const std::vector<std::uint8_t>& bytes() const
	{
		return m_bytes;
	}

private:
	template <typename T>
	void WritePod(const T value)
	{
		const auto* begin = reinterpret_cast<const std::uint8_t*>(&value);
		m_bytes.insert(m_bytes.end(), begin, begin + sizeof(T));
	}

	std::vector<std::uint8_t> m_bytes;
};

std::string TrimAsciiCopy(std::string text)
{
	size_t begin = 0;
	while (begin < text.size() && static_cast<unsigned char>(text[begin]) <= 0x20) {
		++begin;
	}

	size_t end = text.size();
	while (end > begin && static_cast<unsigned char>(text[end - 1]) <= 0x20) {
		--end;
	}
	return text.substr(begin, end - begin);
}

bool StartsWith(const std::string_view text, const std::string_view prefix)
{
	return text.size() >= prefix.size() && text.substr(0, prefix.size()) == prefix;
}

bool EndsWith(const std::string_view text, const std::string_view suffix)
{
	return text.size() >= suffix.size() &&
		text.substr(text.size() - suffix.size(), suffix.size()) == suffix;
}

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

std::vector<std::string> SplitLines(const std::string& text)
{
	std::vector<std::string> lines;
	size_t start = 0;
	size_t index = 0;
	while (index < text.size()) {
		if (text[index] == '\r' || text[index] == '\n') {
			lines.push_back(text.substr(start, index - start));
			if (text[index] == '\r' && index + 1 < text.size() && text[index + 1] == '\n') {
				++index;
			}
			start = index + 1;
		}
		++index;
	}
	lines.push_back(text.substr(start));
	return lines;
}

std::string RemoveUtf8Bom(const std::string& text)
{
	if (text.size() >= 3 &&
		static_cast<unsigned char>(text[0]) == 0xEF &&
		static_cast<unsigned char>(text[1]) == 0xBB &&
		static_cast<unsigned char>(text[2]) == 0xBF) {
		return text.substr(3);
	}
	return text;
}

std::vector<std::string> SplitTopLevelCommaFields(const std::string& text)
{
	std::vector<std::string> fields;
	std::string current;
	bool inQuote = false;
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '"') {
			inQuote = !inQuote;
			current.push_back(ch);
			continue;
		}
		if (ch == ',' && !inQuote) {
			fields.push_back(TrimAsciiCopy(current));
			current.clear();
			continue;
		}
		current.push_back(ch);
	}
	fields.push_back(TrimAsciiCopy(current));
	return fields;
}

std::string Unquote(const std::string& text)
{
	if (text.size() >= 2 && text.front() == '"' && text.back() == '"') {
		return text.substr(1, text.size() - 2);
	}
	return text;
}

constexpr const char* kTextLiteralLeftQuote = "“";
constexpr const char* kTextLiteralRightQuote = "”";
constexpr const char* kEscapedTextLiteralPrefix = "#e2txt_text#";
constexpr const char* kEscapedLongTextLiteralPrefix = "#e2txt_long_text#";

int ParseHexNibble(const char ch)
{
	if (ch >= '0' && ch <= '9') {
		return ch - '0';
	}
	if (ch >= 'a' && ch <= 'f') {
		return ch - 'a' + 10;
	}
	if (ch >= 'A' && ch <= 'F') {
		return ch - 'A' + 10;
	}
	return -1;
}

std::string UnescapeTextLiteralPayload(const std::string& text)
{
	std::string out;
	out.reserve(text.size());
	for (size_t i = 0; i < text.size(); ++i) {
		if (text[i] != '\\') {
			out.push_back(text[i]);
			continue;
		}
		if (i + 1 >= text.size()) {
			out.push_back('\\');
			break;
		}

		const char next = text[++i];
		switch (next) {
		case '\\':
			out.push_back('\\');
			break;
		case 'r':
			out.push_back('\r');
			break;
		case 'n':
			out.push_back('\n');
			break;
		case 't':
			out.push_back('\t');
			break;
		case 'x':
			if (i + 2 < text.size()) {
				const int high = ParseHexNibble(text[i + 1]);
				const int low = ParseHexNibble(text[i + 2]);
				if (high >= 0 && low >= 0) {
					out.push_back(static_cast<char>((high << 4) | low));
					i += 2;
					break;
				}
			}
			out += "\\x";
			break;
		default:
			out.push_back(next);
			break;
		}
	}
	return out;
}

bool TryParseInt32(const std::string& text, std::int32_t& outValue)
{
	const std::string trimmed = TrimAsciiCopy(text);
	if (trimmed.empty()) {
		return false;
	}
	const char* begin = trimmed.data();
	const char* end = trimmed.data() + trimmed.size();
	const auto [ptr, ec] = std::from_chars(begin, end, outValue);
	return ec == std::errc() && ptr == end;
}

bool TryParseDouble(const std::string& text, double& outValue)
{
	const std::string trimmed = TrimAsciiCopy(text);
	if (trimmed.empty()) {
		return false;
	}

	char* end = nullptr;
	outValue = std::strtod(trimmed.c_str(), &end);
	return end != nullptr && *end == '\0';
}

std::optional<bool> ParseBoolLiteral(const std::string& text)
{
	const std::string trimmed = TrimAsciiCopy(text);
	if (trimmed == "真" || trimmed == "true" || trimmed == "TRUE") {
		return true;
	}
	if (trimmed == "假" || trimmed == "false" || trimmed == "FALSE") {
		return false;
	}
	return std::nullopt;
}

std::string StripWrappedText(const std::string& text, const std::string& left, const std::string& right)
{
	if (text.size() < left.size() + right.size()) {
		return text;
	}
	if (!StartsWith(text, left) || !EndsWith(text, right)) {
		return text;
	}
	return text.substr(left.size(), text.size() - left.size() - right.size());
}

bool TryDecodeDumpTextLiteral(const std::string& valueText, std::string& outRawText, bool& outIsLongText)
{
	outRawText.clear();
	outIsLongText = false;
	if (!StartsWith(valueText, kTextLiteralLeftQuote) || !EndsWith(valueText, kTextLiteralRightQuote)) {
		return false;
	}

	const std::string payload = StripWrappedText(valueText, kTextLiteralLeftQuote, kTextLiteralRightQuote);
	if (StartsWith(payload, kEscapedLongTextLiteralPrefix)) {
		outIsLongText = true;
		outRawText = UnescapeTextLiteralPayload(payload.substr(std::strlen(kEscapedLongTextLiteralPrefix)));
		return true;
	}
	if (StartsWith(payload, kEscapedTextLiteralPrefix)) {
		outRawText = UnescapeTextLiteralPayload(payload.substr(std::strlen(kEscapedTextLiteralPrefix)));
		return true;
	}

	outRawText = payload;
	return true;
}

std::vector<std::uint8_t> DecodeBase64(const std::string& text)
{
	static constexpr std::array<std::uint8_t, 256> kMap = [] {
		std::array<std::uint8_t, 256> map = {};
		map.fill(0xFFu);
		for (std::uint8_t i = 0; i < 26; ++i) {
			map[static_cast<unsigned char>('A' + i)] = i;
			map[static_cast<unsigned char>('a' + i)] = static_cast<std::uint8_t>(26 + i);
		}
		for (std::uint8_t i = 0; i < 10; ++i) {
			map[static_cast<unsigned char>('0' + i)] = static_cast<std::uint8_t>(52 + i);
		}
		map[static_cast<unsigned char>('+')] = 62;
		map[static_cast<unsigned char>('/')] = 63;
		return map;
	}();

	std::vector<std::uint8_t> out;
	std::array<std::uint8_t, 4> chunk = {};
	int chunkSize = 0;
	int padding = 0;
	for (const unsigned char ch : text) {
		if (std::isspace(ch) != 0) {
			continue;
		}
		if (ch == '=') {
			chunk[chunkSize++] = 0;
			++padding;
		}
		else {
			const std::uint8_t value = kMap[ch];
			if (value == 0xFFu) {
				return {};
			}
			chunk[chunkSize++] = value;
		}

		if (chunkSize == 4) {
			const std::uint32_t bits =
				(static_cast<std::uint32_t>(chunk[0]) << 18) |
				(static_cast<std::uint32_t>(chunk[1]) << 12) |
				(static_cast<std::uint32_t>(chunk[2]) << 6) |
				static_cast<std::uint32_t>(chunk[3]);
			out.push_back(static_cast<std::uint8_t>((bits >> 16) & 0xFFu));
			if (padding < 2) {
				out.push_back(static_cast<std::uint8_t>((bits >> 8) & 0xFFu));
			}
			if (padding == 0) {
				out.push_back(static_cast<std::uint8_t>(bits & 0xFFu));
			}
			chunkSize = 0;
			padding = 0;
		}
	}
	return out;
}

std::string DecodeXmlEntities(const std::string& text)
{
	std::string out;
	out.reserve(text.size());
	for (size_t i = 0; i < text.size(); ++i) {
		if (text[i] != '&') {
			out.push_back(text[i]);
			continue;
		}
		if (StartsWith(std::string_view(text).substr(i), "&amp;")) {
			out.push_back('&');
			i += 4;
		}
		else if (StartsWith(std::string_view(text).substr(i), "&lt;")) {
			out.push_back('<');
			i += 3;
		}
		else if (StartsWith(std::string_view(text).substr(i), "&gt;")) {
			out.push_back('>');
			i += 3;
		}
		else if (StartsWith(std::string_view(text).substr(i), "&quot;")) {
			out.push_back('"');
			i += 5;
		}
		else if (StartsWith(std::string_view(text).substr(i), "&apos;")) {
			out.push_back('\'');
			i += 5;
		}
		else {
			out.push_back(text[i]);
		}
	}
	return out;
}

struct DumpBlock {
	enum class Kind {
		Dependencies,
		Page,
		FormXml,
	};

	Kind kind = Kind::Page;
	std::string pageType;
	std::string name;
	std::vector<std::string> lines;
};

bool ParseHeaderValueLine(
	const std::string& line,
	const std::string& key,
	std::string& outValue)
{
	const std::string prefix = key + "=";
	if (!StartsWith(line, prefix)) {
		return false;
	}
	outValue = line.substr(prefix.size());
	return true;
}

std::string ExtractNamedSegment(
	const std::string& text,
	const std::string& key,
	const std::optional<std::string>& nextKey)
{
	const std::string marker = key + "=";
	const size_t begin = text.find(marker);
	if (begin == std::string::npos) {
		return std::string();
	}

	const size_t valueBegin = begin + marker.size();
	size_t valueEnd = text.size();
	if (nextKey.has_value()) {
		const size_t nextPos = text.find(" " + *nextKey + "=", valueBegin);
		if (nextPos != std::string::npos) {
			valueEnd = nextPos;
		}
	}
	return text.substr(valueBegin, valueEnd - valueBegin);
}

bool ParseDumpBlocks(const std::vector<std::string>& lines, std::vector<DumpBlock>& outBlocks, std::string* outError)
{
	outBlocks.clear();
	size_t index = 0;
	while (index < lines.size()) {
		if (lines[index] != "================================================================================") {
			++index;
			continue;
		}
		++index;
		if (index >= lines.size()) {
			break;
		}

		DumpBlock block;
		const std::string& titleLine = lines[index++];
		if (titleLine == "[Dependencies]") {
			block.kind = DumpBlock::Kind::Dependencies;
		}
		else if (StartsWith(titleLine, "[Form XML] ")) {
			block.kind = DumpBlock::Kind::FormXml;
			block.name = titleLine.substr(std::string("[Form XML] ").size());
		}
		else if (!titleLine.empty() && titleLine.front() == '[') {
			const size_t typePos = titleLine.find("] type=");
			const size_t namePos = titleLine.find(" name=");
			if (typePos == std::string::npos || namePos == std::string::npos || namePos <= typePos + 7) {
				if (outError != nullptr) {
					*outError = "dump_page_header_invalid";
				}
				return false;
			}
			block.kind = DumpBlock::Kind::Page;
			block.pageType = titleLine.substr(typePos + 7, namePos - (typePos + 7));
			block.name = titleLine.substr(namePos + 6);
		}
		else {
			if (outError != nullptr) {
				*outError = "dump_block_header_invalid";
			}
			return false;
		}

		if (index < lines.size() && lines[index] == "--------------------------------------------------------------------------------") {
			++index;
		}
		while (index < lines.size() && lines[index] != "================================================================================") {
			block.lines.push_back(lines[index++]);
		}
		outBlocks.push_back(std::move(block));
	}
	return true;
}

std::string ReadSupportLibraryName(const char* text)
{
	return text == nullptr ? std::string() : std::string(text);
}

std::vector<std::string> BuildSupportTypeMemberNames(const LIB_DATA_TYPE_INFO& dataType)
{
	std::vector<std::string> memberNames;
	if (dataType.m_nPropertyCount > 0 && dataType.m_pPropertyBegin != nullptr) {
		memberNames.reserve(static_cast<size_t>(dataType.m_nPropertyCount));
		for (int propertyIndex = 0; propertyIndex < dataType.m_nPropertyCount; ++propertyIndex) {
			memberNames.emplace_back(ReadSupportLibraryName(dataType.m_pPropertyBegin[propertyIndex].m_szName));
		}
		return memberNames;
	}

	if (dataType.m_nElementCount > 0 && dataType.m_pElementBegin != nullptr) {
		memberNames.reserve(static_cast<size_t>(dataType.m_nElementCount));
		for (int memberIndex = 0; memberIndex < dataType.m_nElementCount; ++memberIndex) {
			memberNames.emplace_back(ReadSupportLibraryName(dataType.m_pElementBegin[memberIndex].m_szName));
		}
	}
	return memberNames;
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

const std::vector<std::pair<std::string, std::int32_t>>& GetBuiltinTypes()
{
	static const std::vector<std::pair<std::string, std::int32_t>> kTypes = {
		{ "通用型", static_cast<std::int32_t>(0x80000000u) },
		{ "字节型", static_cast<std::int32_t>(0x80000101u) },
		{ "短整数型", static_cast<std::int32_t>(0x80000201u) },
		{ "整数型", static_cast<std::int32_t>(0x80000301u) },
		{ "长整数型", static_cast<std::int32_t>(0x80000401u) },
		{ "小数型", static_cast<std::int32_t>(0x80000501u) },
		{ "双精度小数型", static_cast<std::int32_t>(0x80000601u) },
		{ "逻辑型", static_cast<std::int32_t>(0x80000002u) },
		{ "日期时间型", static_cast<std::int32_t>(0x80000003u) },
		{ "文本型", static_cast<std::int32_t>(0x80000004u) },
		{ "字节集", static_cast<std::int32_t>(0x80000005u) },
		{ "子程序指针", static_cast<std::int32_t>(0x80000006u) },
		{ "条件语句型", static_cast<std::int32_t>(0x80000008u) },
		{ "窗口", 65537 },
		{ "菜单", 65539 },
		{ "字体", 65540 },
		{ "编辑框", 65541 },
		{ "图片框", 65542 },
		{ "外形框", 65543 },
		{ "画板", 65544 },
		{ "分组框", 65545 },
		{ "标签", 65546 },
		{ "按钮", 65547 },
		{ "选择框", 65548 },
		{ "单选框", 65549 },
		{ "组合框", 65550 },
		{ "列表框", 65551 },
		{ "选择列表框", 65552 },
		{ "横向滚动条", 65553 },
		{ "纵向滚动条", 65554 },
		{ "进度条", 65555 },
		{ "滑块条", 65556 },
		{ "选择夹", 65557 },
		{ "影像框", 65558 },
		{ "日期框", 65559 },
		{ "月历", 65560 },
		{ "驱动器框", 65561 },
		{ "目录框", 65562 },
		{ "文件框", 65563 },
		{ "颜色选择器", 65564 },
		{ "超级链接框", 65565 },
		{ "调节器", 65566 },
		{ "通用对话框", 65567 },
		{ "时钟", 65568 },
		{ "打印机", 65569 },
		{ "字段信息", 65570 },
		{ "数据报", 65572 },
		{ "客户", 65573 },
		{ "服务器", 65574 },
		{ "端口", 65575 },
		{ "打印设置信息", 65576 },
		{ "表格", 65577 },
		{ "数据源", 65578 },
		{ "通用提供者", 65579 },
		{ "数据库提供者", 65580 },
		{ "图形按钮", 65581 },
		{ "外部数据库", 65582 },
		{ "外部数据提供者", 65583 },
		{ "对象", 65584 },
		{ "变体型", 65585 },
		{ "变体类型", 65586 },
		{ "工具条", 196611 },
		{ "超级列表框", 196612 },
		{ "高级表格", 262145 },
	};
	return kTypes;
}

class IdAllocator {
public:
	std::int32_t Alloc(const std::int32_t typeMask)
	{
		return typeMask | (++m_next);
	}

private:
	std::int32_t m_next = 0xFFFF;
};

struct SupportLibraryTypeInfo {
	std::int32_t typeId = 0;
	bool isTabControl = false;
};

class TypeResolver {
public:
	TypeResolver(const std::string& sourcePath, const std::vector<RestoreDependencyInfo>& dependencies)
		: m_sourcePath(sourcePath)
	{
		for (const auto& [name, value] : GetBuiltinTypes()) {
			m_builtinTypes.emplace(name, value);
		}

		int supportIndex = 1;
		for (const auto& dependency : dependencies) {
			if (!dependency.isSupportLibrary) {
				continue;
			}
			m_supportLibraryOrder.push_back(&dependency);
			LoadSupportLibrary(dependency, supportIndex++);
		}
	}

	void RegisterUserType(const std::string& name, const std::int32_t typeId)
	{
		const std::string key = NormalizeTypeName(name);
		if (!key.empty()) {
			m_userTypes[key] = typeId;
		}
	}

	void RegisterPlaceholderType(const std::string& name, const std::int32_t typeId)
	{
		RegisterUserType(name, typeId);
		m_placeholderTypeNames.insert(NormalizeTypeName(name));
	}

	bool IsPlaceholderType(const std::string& name) const
	{
		return m_placeholderTypeNames.contains(NormalizeTypeName(name));
	}

	std::int32_t ResolveTypeId(const std::string& rawTypeName) const
	{
		const std::string typeName = NormalizeTypeName(rawTypeName);
		if (typeName.empty()) {
			return 0;
		}
		if (const auto it = m_builtinTypes.find(typeName); it != m_builtinTypes.end()) {
			return it->second;
		}
		if (const auto it = m_userTypes.find(typeName); it != m_userTypes.end()) {
			return it->second;
		}
		if (const auto it = m_supportTypes.find(typeName); it != m_supportTypes.end()) {
			return it->second.typeId;
		}
		return 0;
	}

	bool IsTabControlType(const std::int32_t typeId) const
	{
		if (typeId == 65557) {
			return true;
		}
		for (const auto& [_, info] : m_supportTypes) {
			if (info.typeId == typeId) {
				return info.isTabControl;
			}
		}
		return false;
	}

	static std::string NormalizeTypeName(std::string value)
	{
		value = TrimAsciiCopy(std::move(value));
		if (value.size() >= 2 && value.front() == '<' && value.back() == '>') {
			value = TrimAsciiCopy(value.substr(1, value.size() - 2));
		}
		return value;
	}

private:
	void LoadSupportLibrary(const RestoreDependencyInfo& dependency, const int supportIndex)
	{
		const auto candidates = BuildSupportLibraryCandidatePaths(m_sourcePath, dependency.fileName);
		HMODULE module = nullptr;
		for (const auto& path : candidates) {
			std::error_code ec;
			if (!std::filesystem::exists(path, ec)) {
				continue;
			}
			module = LoadLibraryExA(path.string().c_str(), nullptr, 0);
			if (module != nullptr) {
				break;
			}
		}
		if (module == nullptr) {
			return;
		}

		const auto* getInfoProc = reinterpret_cast<PFN_GET_LIB_INFO>(GetProcAddress(module, FUNCNAME_GET_LIB_INFO));
		if (getInfoProc == nullptr) {
			FreeLibrary(module);
			return;
		}

		const auto* libInfo = getInfoProc();
		if (libInfo == nullptr || libInfo->m_nDataTypeCount <= 0 || libInfo->m_pDataType == nullptr) {
			FreeLibrary(module);
			return;
		}

		for (int i = 0; i < libInfo->m_nDataTypeCount; ++i) {
			const LIB_DATA_TYPE_INFO& dataType = libInfo->m_pDataType[i];
			SupportLibraryTypeInfo info;
			info.typeId = (supportIndex << 16) | (i + 1);
			info.isTabControl = (dataType.m_dwState & LDT_IS_TAB_UNIT) != 0;
			const std::string name = NormalizeTypeName(ReadSupportLibraryName(dataType.m_szName));
			if (!name.empty()) {
				m_supportTypes.insert_or_assign(name, info);
			}
		}
		FreeLibrary(module);
	}

	std::string m_sourcePath;
	std::unordered_map<std::string, std::int32_t> m_builtinTypes;
	std::unordered_map<std::string, std::int32_t> m_userTypes;
	std::unordered_map<std::string, SupportLibraryTypeInfo> m_supportTypes;
	std::unordered_set<std::string> m_placeholderTypeNames;
	std::vector<const RestoreDependencyInfo*> m_supportLibraryOrder;
};

std::optional<std::pair<std::string, std::string>> SplitFixedCodeComment(const std::string& text)
{
	const size_t pos = text.find("  ' ");
	if (pos == std::string::npos) {
		return std::make_pair(text, std::string());
	}
	return std::make_pair(text.substr(0, pos), text.substr(pos + 4));
}

struct BodyStatement;

struct BodySwitchCase {
	bool mask = false;
	std::string code;
	std::vector<BodyStatement> block;
};

enum class BodyStatementKind {
	Raw,
	IfTrue,
	IfElse,
	WhileLoop,
	DoWhileLoop,
	CounterLoop,
	ForLoop,
	SwitchBlock,
};

struct BodyStatement {
	BodyStatementKind kind = BodyStatementKind::Raw;
	bool mask = false;
	bool maskOnEnd = false;
	std::string code;
	std::string fixedComment;
	std::string fixedEndComment;
	std::string endCode;
	std::vector<BodyStatement> block;
	std::vector<BodyStatement> elseBlock;
	std::vector<BodySwitchCase> cases;
	std::vector<BodyStatement> defaultBlock;
};

int CountIndentLevel(const std::string& line)
{
	size_t count = 0;
	while (count < line.size() && line[count] == ' ') {
		++count;
	}
	return static_cast<int>(count / 4);
}

std::string StripIndent(const std::string& line)
{
	size_t index = 0;
	while (index < line.size() && line[index] == ' ') {
		++index;
	}
	return line.substr(index);
}

bool ExtractMaskPrefix(const std::string& line, bool& outMask, std::string& outCode)
{
	outMask = false;
	outCode = StripIndent(line);
	if (StartsWith(outCode, "' ")) {
		outMask = true;
		outCode.erase(0, 2);
	}
	return true;
}

bool IsBlankLine(const std::string& line)
{
	return TrimAsciiCopy(line).empty();
}

bool StartsWithControl(const std::string& line, const std::string& token)
{
	return StartsWith(TrimAsciiCopy(line), token);
}

bool MatchesEndToken(const std::string& code, const std::unordered_set<std::string>& endTokens)
{
	for (const auto& token : endTokens) {
		if (!token.empty() && token.back() == '*') {
			if (StartsWith(code, token.substr(0, token.size() - 1))) {
				return true;
			}
			continue;
		}
		if (code == token) {
			return true;
		}
	}
	return false;
}

class MethodCodeWriter {
public:
	void BeginBlock(const std::uint8_t type)
	{
		m_blockOffset.WriteU8(type);
		m_blockOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
		m_blockStack.push_back(m_blockOffset.position());
		m_blockOffset.WriteI32(0);
	}

	void EndBlock()
	{
		if (m_blockStack.empty()) {
			return;
		}
		const size_t patchPos = m_blockStack.back();
		m_blockStack.pop_back();
		m_blockOffset.PatchI32(patchPos, static_cast<std::int32_t>(m_expressionData.position()));
	}

	void WriteRawStatement(const bool mask, const std::string& code)
	{
		m_lineOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
		WriteUnexaminedCall(0x6A, -1, 0, mask, code);
	}

	void WriteIfTrue(const BodyStatement& statement)
	{
		BeginBlock(2);
		m_lineOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
		WriteUnexaminedCall(0x6C, 0, 1, statement.mask, statement.code);
		WriteBlock(statement.block);
		m_expressionData.WriteU8(0x52);
		EndBlock();
		m_expressionData.WriteU8(0x73);
	}

	void WriteIfElse(const BodyStatement& statement)
	{
		BeginBlock(1);
		m_lineOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
		WriteUnexaminedCall(0x6B, 0, 0, statement.mask, statement.code);
		WriteBlock(statement.block);
		m_expressionData.WriteU8(0x50);
		WriteBlock(statement.elseBlock);
		m_expressionData.WriteU8(0x51);
		EndBlock();
		m_expressionData.WriteU8(0x72);
	}

	void WriteWhile(const BodyStatement& statement)
	{
		BeginBlock(3);
		m_lineOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
		WriteUnexaminedCall(0x70, 0, 3, statement.mask, statement.code);
		WriteBlock(statement.block);
		m_expressionData.WriteU8(0x55);
		EndBlock();
		WriteFixedCall(0x71, 0, 4, statement.maskOnEnd, statement.fixedEndComment);
	}

	void WriteDoWhile(const BodyStatement& statement)
	{
		BeginBlock(3);
		WriteFixedCall(0x70, 0, 5, statement.mask, statement.fixedComment);
		WriteBlock(statement.block);
		m_expressionData.WriteU8(0x55);
		EndBlock();
		m_lineOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
		WriteUnexaminedCall(0x71, 0, 6, statement.maskOnEnd, statement.endCode);
	}

	void WriteCounter(const BodyStatement& statement)
	{
		BeginBlock(3);
		m_lineOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
		WriteUnexaminedCall(0x70, 0, 7, statement.mask, statement.code);
		WriteBlock(statement.block);
		m_expressionData.WriteU8(0x55);
		EndBlock();
		WriteFixedCall(0x71, 0, 8, statement.maskOnEnd, statement.fixedEndComment);
	}

	void WriteFor(const BodyStatement& statement)
	{
		BeginBlock(3);
		m_lineOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
		WriteUnexaminedCall(0x70, 0, 9, statement.mask, statement.code);
		WriteBlock(statement.block);
		m_expressionData.WriteU8(0x55);
		EndBlock();
		WriteFixedCall(0x71, 0, 10, statement.maskOnEnd, statement.fixedEndComment);
	}

	void WriteSwitch(const BodyStatement& statement)
	{
		BeginBlock(4);
		m_expressionData.WriteU8(0x6D);
		for (const auto& caseItem : statement.cases) {
			m_lineOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
			WriteUnexaminedCall(0x6E, 0, 2, caseItem.mask, caseItem.code);
			WriteBlock(caseItem.block);
			m_expressionData.WriteU8(0x53);
		}
		m_expressionData.WriteU8(0x6F);
		WriteBlock(statement.defaultBlock);
		m_expressionData.WriteU8(0x54);
		EndBlock();
		m_expressionData.WriteU8(0x74);
	}

	void WriteBlock(const std::vector<BodyStatement>& statements)
	{
		for (const auto& statement : statements) {
			switch (statement.kind) {
			case BodyStatementKind::Raw:
				WriteRawStatement(statement.mask, statement.code);
				break;
			case BodyStatementKind::IfTrue:
				WriteIfTrue(statement);
				break;
			case BodyStatementKind::IfElse:
				WriteIfElse(statement);
				break;
			case BodyStatementKind::WhileLoop:
				WriteWhile(statement);
				break;
			case BodyStatementKind::DoWhileLoop:
				WriteDoWhile(statement);
				break;
			case BodyStatementKind::CounterLoop:
				WriteCounter(statement);
				break;
			case BodyStatementKind::ForLoop:
				WriteFor(statement);
				break;
			case BodyStatementKind::SwitchBlock:
				WriteSwitch(statement);
				break;
			}
		}
	}

	std::vector<std::uint8_t> TakeLineOffset() { return m_lineOffset.TakeBytes(); }
	std::vector<std::uint8_t> TakeBlockOffset() { return m_blockOffset.TakeBytes(); }
	std::vector<std::uint8_t> TakeMethodReference() { return m_methodReference.TakeBytes(); }
	std::vector<std::uint8_t> TakeVariableReference() { return m_variableReference.TakeBytes(); }
	std::vector<std::uint8_t> TakeConstantReference() { return m_constantReference.TakeBytes(); }
	std::vector<std::uint8_t> TakeExpressionData() { return m_expressionData.TakeBytes(); }

private:
	void WriteFixedCall(
		const std::uint8_t type,
		const std::int16_t libraryId,
		const std::int32_t methodId,
		const bool mask,
		const std::string& comment)
	{
		m_lineOffset.WriteI32(static_cast<std::int32_t>(m_expressionData.position()));
		m_expressionData.WriteU8(type);
		m_expressionData.WriteI32(methodId);
		m_expressionData.WriteI16(libraryId);
		m_expressionData.WriteI16(static_cast<std::int16_t>(mask ? 0x20 : 0));
		m_expressionData.WriteBStr(std::nullopt);
		m_expressionData.WriteBStr(comment.empty() ? std::nullopt : std::make_optional(comment));
		m_expressionData.WriteU8(0x36);
		m_expressionData.WriteU8(0x01);
	}

	void WriteUnexaminedCall(
		const std::uint8_t type,
		const std::int16_t libraryId,
		const std::int32_t methodId,
		const bool mask,
		const std::string& code)
	{
		m_expressionData.WriteU8(type);
		m_expressionData.WriteI32(methodId);
		m_expressionData.WriteI16(libraryId);
		m_expressionData.WriteI16(static_cast<std::int16_t>(mask ? 0x20 : 0x40));
		m_expressionData.WriteBStr(std::make_optional(code));
		m_expressionData.WriteBStr(std::nullopt);
		m_expressionData.WriteU8(0x36);
		m_expressionData.WriteU8(0x01);
	}

	ByteWriter m_lineOffset;
	ByteWriter m_blockOffset;
	ByteWriter m_methodReference;
	ByteWriter m_variableReference;
	ByteWriter m_constantReference;
	ByteWriter m_expressionData;
	std::vector<size_t> m_blockStack;
};

bool ParseBodyBlock(
	const std::vector<std::string>& lines,
	size_t& index,
	const int expectedIndent,
	const std::unordered_set<std::string>& endTokens,
	std::vector<BodyStatement>& outStatements,
	std::string* outError);

bool ParseSwitchBlock(
	const std::vector<std::string>& lines,
	size_t& index,
	const int indent,
	const bool firstMask,
	const std::string& firstCaseCode,
	std::vector<BodyStatement>& outStatements,
	std::string* outError)
{
	BodyStatement statement;
	statement.kind = BodyStatementKind::SwitchBlock;
	statement.cases.push_back(BodySwitchCase{ firstMask, firstCaseCode, {} });

	if (!ParseBodyBlock(lines, index, indent + 1, { ".判断*", ".默认", ".判断结束" }, statement.cases.back().block, outError)) {
		return false;
	}

	while (index < lines.size()) {
		if (IsBlankLine(lines[index])) {
			++index;
			continue;
		}

		bool mask = false;
		std::string code;
		ExtractMaskPrefix(lines[index], mask, code);
		code = TrimAsciiCopy(code);
		if (!StartsWith(code, ".")) {
			if (outError != nullptr) {
				*outError = "switch_marker_invalid";
			}
			return false;
		}
		if (code == ".默认") {
			++index;
			break;
		}
		if (!StartsWith(code, ".判断")) {
			if (outError != nullptr) {
				*outError = "switch_case_missing";
			}
			return false;
		}

		BodySwitchCase nextCase;
		nextCase.mask = mask;
		nextCase.code = code.substr(1);
		++index;
		if (!ParseBodyBlock(lines, index, indent + 1, { ".判断*", ".默认", ".判断结束" }, nextCase.block, outError)) {
			return false;
		}
		statement.cases.push_back(std::move(nextCase));
	}

	if (!ParseBodyBlock(lines, index, indent + 1, { ".判断结束" }, statement.defaultBlock, outError)) {
		return false;
	}
	if (index >= lines.size()) {
		if (outError != nullptr) {
			*outError = "switch_end_missing";
		}
		return false;
	}

	bool endMask = false;
	std::string endCode;
	ExtractMaskPrefix(lines[index], endMask, endCode);
	endCode = TrimAsciiCopy(endCode);
	if (endCode != ".判断结束") {
		if (outError != nullptr) {
			*outError = "switch_end_invalid";
		}
		return false;
	}
	++index;
	outStatements.push_back(std::move(statement));
	return true;
}

bool ParseBodyBlock(
	const std::vector<std::string>& lines,
	size_t& index,
	const int expectedIndent,
	const std::unordered_set<std::string>& endTokens,
	std::vector<BodyStatement>& outStatements,
	std::string* outError)
{
	outStatements.clear();
	while (index < lines.size()) {
		const std::string& rawLine = lines[index];
		if (IsBlankLine(rawLine)) {
			++index;
			continue;
		}
		if (CountIndentLevel(rawLine) < expectedIndent) {
			break;
		}

		bool mask = false;
		std::string code;
		ExtractMaskPrefix(rawLine, mask, code);
		code = TrimAsciiCopy(code);
		if (MatchesEndToken(code, endTokens)) {
			break;
		}
		if (!StartsWith(code, ".")) {
			outStatements.push_back(BodyStatement{ BodyStatementKind::Raw, mask, false, code });
			++index;
			continue;
		}

		if (StartsWith(code, ".如果真 ")) {
			BodyStatement statement;
			statement.kind = BodyStatementKind::IfTrue;
			statement.mask = mask;
			statement.code = code.substr(1);
			++index;
			if (!ParseBodyBlock(lines, index, expectedIndent + 1, { ".如果真结束" }, statement.block, outError)) {
				return false;
			}
			if (index >= lines.size()) {
				if (outError != nullptr) {
					*outError = "if_true_end_missing";
				}
				return false;
			}
			bool endMask = false;
			std::string endCode;
			ExtractMaskPrefix(lines[index], endMask, endCode);
			if (TrimAsciiCopy(endCode) != ".如果真结束") {
				if (outError != nullptr) {
					*outError = "if_true_end_invalid";
				}
				return false;
			}
			++index;
			outStatements.push_back(std::move(statement));
			continue;
		}

		if (StartsWith(code, ".如果 ")) {
			BodyStatement statement;
			statement.kind = BodyStatementKind::IfElse;
			statement.mask = mask;
			statement.code = code.substr(1);
			++index;
			if (!ParseBodyBlock(lines, index, expectedIndent + 1, { ".否则" }, statement.block, outError)) {
				return false;
			}
			if (index >= lines.size()) {
				if (outError != nullptr) {
					*outError = "if_else_marker_missing";
				}
				return false;
			}
			bool elseMask = false;
			std::string elseCode;
			ExtractMaskPrefix(lines[index], elseMask, elseCode);
			if (TrimAsciiCopy(elseCode) != ".否则") {
				if (outError != nullptr) {
					*outError = "if_else_invalid";
				}
				return false;
			}
			++index;
			if (!ParseBodyBlock(lines, index, expectedIndent + 1, { ".如果结束" }, statement.elseBlock, outError)) {
				return false;
			}
			if (index >= lines.size()) {
				if (outError != nullptr) {
					*outError = "if_end_missing";
				}
				return false;
			}
			bool endMask = false;
			std::string endCode;
			ExtractMaskPrefix(lines[index], endMask, endCode);
			if (TrimAsciiCopy(endCode) != ".如果结束") {
				if (outError != nullptr) {
					*outError = "if_end_invalid";
				}
				return false;
			}
			++index;
			outStatements.push_back(std::move(statement));
			continue;
		}

		if (StartsWith(code, ".判断循环首 ")) {
			BodyStatement statement;
			statement.kind = BodyStatementKind::WhileLoop;
			statement.mask = mask;
			statement.code = code.substr(1);
			++index;
			if (!ParseBodyBlock(lines, index, expectedIndent + 1, { ".判断循环尾 ()" }, statement.block, outError)) {
				return false;
			}
			if (index >= lines.size()) {
				if (outError != nullptr) {
					*outError = "while_end_missing";
				}
				return false;
			}
			bool endMask = false;
			std::string endCode;
			ExtractMaskPrefix(lines[index], endMask, endCode);
			const auto split = SplitFixedCodeComment(TrimAsciiCopy(endCode));
			if (!split.has_value() || split->first != ".判断循环尾 ()") {
				if (outError != nullptr) {
					*outError = "while_end_invalid";
				}
				return false;
			}
			statement.maskOnEnd = endMask;
			statement.fixedEndComment = split->second;
			++index;
			outStatements.push_back(std::move(statement));
			continue;
		}

		if (StartsWith(code, ".循环判断首 ()")) {
			BodyStatement statement;
			statement.kind = BodyStatementKind::DoWhileLoop;
			statement.mask = mask;
			const auto split = SplitFixedCodeComment(code);
			statement.fixedComment = split.has_value() ? split->second : std::string();
			++index;
			if (!ParseBodyBlock(lines, index, expectedIndent + 1, { ".循环判断尾*" }, statement.block, outError)) {
				return false;
			}
			if (index >= lines.size()) {
				if (outError != nullptr) {
					*outError = "do_while_end_missing";
				}
				return false;
			}
			bool endMask = false;
			std::string endCode;
			ExtractMaskPrefix(lines[index], endMask, endCode);
			statement.maskOnEnd = endMask;
			statement.endCode = TrimAsciiCopy(endCode).substr(1);
			++index;
			outStatements.push_back(std::move(statement));
			continue;
		}

		if (StartsWith(code, ".计次循环首 ")) {
			BodyStatement statement;
			statement.kind = BodyStatementKind::CounterLoop;
			statement.mask = mask;
			statement.code = code.substr(1);
			++index;
			if (!ParseBodyBlock(lines, index, expectedIndent + 1, { ".计次循环尾 ()" }, statement.block, outError)) {
				return false;
			}
			if (index >= lines.size()) {
				if (outError != nullptr) {
					*outError = "counter_end_missing";
				}
				return false;
			}
			bool endMask = false;
			std::string endCode;
			ExtractMaskPrefix(lines[index], endMask, endCode);
			const auto split = SplitFixedCodeComment(TrimAsciiCopy(endCode));
			if (!split.has_value() || split->first != ".计次循环尾 ()") {
				if (outError != nullptr) {
					*outError = "counter_end_invalid";
				}
				return false;
			}
			statement.maskOnEnd = endMask;
			statement.fixedEndComment = split->second;
			++index;
			outStatements.push_back(std::move(statement));
			continue;
		}

		if (StartsWith(code, ".变量循环首 ")) {
			BodyStatement statement;
			statement.kind = BodyStatementKind::ForLoop;
			statement.mask = mask;
			statement.code = code.substr(1);
			++index;
			if (!ParseBodyBlock(lines, index, expectedIndent + 1, { ".变量循环尾 ()" }, statement.block, outError)) {
				return false;
			}
			if (index >= lines.size()) {
				if (outError != nullptr) {
					*outError = "for_end_missing";
				}
				return false;
			}
			bool endMask = false;
			std::string endCode;
			ExtractMaskPrefix(lines[index], endMask, endCode);
			const auto split = SplitFixedCodeComment(TrimAsciiCopy(endCode));
			if (!split.has_value() || split->first != ".变量循环尾 ()") {
				if (outError != nullptr) {
					*outError = "for_end_invalid";
				}
				return false;
			}
			statement.maskOnEnd = endMask;
			statement.fixedEndComment = split->second;
			++index;
			outStatements.push_back(std::move(statement));
			continue;
		}

		if (StartsWith(code, ".判断开始")) {
			++index;
			std::string firstCaseCode = "判断" + code.substr(std::string(".判断开始").size());
			if (!ParseSwitchBlock(lines, index, expectedIndent, mask, firstCaseCode, outStatements, outError)) {
				return false;
			}
			continue;
		}

		outStatements.push_back(BodyStatement{ BodyStatementKind::Raw, mask, false, code });
		++index;
	}
	return true;
}

bool BuildMethodCodeData(const std::vector<std::string>& lines, RestoreMethod& outMethod, std::string* outError)
{
	std::vector<BodyStatement> statements;
	size_t index = 0;
	if (!ParseBodyBlock(lines, index, 0, {}, statements, outError)) {
		return false;
	}

	MethodCodeWriter writer;
	writer.WriteBlock(statements);
	outMethod.lineOffset = writer.TakeLineOffset();
	outMethod.blockOffset = writer.TakeBlockOffset();
	outMethod.methodReference = writer.TakeMethodReference();
	outMethod.variableReference = writer.TakeVariableReference();
	outMethod.constantReference = writer.TakeConstantReference();
	outMethod.expressionData = writer.TakeExpressionData();
	return true;
}

struct XmlNode {
	std::string name;
	std::unordered_map<std::string, std::string> attributes;
	std::vector<XmlNode> children;
};

class SimpleXmlParser {
public:
	explicit SimpleXmlParser(const std::string& text)
		: m_text(text)
	{
	}

	bool Parse(XmlNode& outRoot, std::string* outError)
	{
		SkipWhitespace();
		if (StartsWith(std::string_view(m_text).substr(m_pos), "<?xml")) {
			const size_t end = m_text.find("?>", m_pos);
			if (end == std::string::npos) {
				if (outError != nullptr) {
					*outError = "xml_declaration_invalid";
				}
				return false;
			}
			m_pos = end + 2;
		}
		SkipWhitespace();
		if (!ParseNode(outRoot, outError)) {
			return false;
		}
		SkipWhitespace();
		return true;
	}

private:
	void SkipWhitespace()
	{
		while (m_pos < m_text.size() && std::isspace(static_cast<unsigned char>(m_text[m_pos])) != 0) {
			++m_pos;
		}
	}

	bool ParseName(std::string& outName)
	{
		const size_t start = m_pos;
		while (m_pos < m_text.size()) {
			const unsigned char ch = static_cast<unsigned char>(m_text[m_pos]);
			if (std::isspace(ch) != 0 || ch == '/' || ch == '>' || ch == '=' || ch == '?') {
				break;
			}
			++m_pos;
		}
		if (m_pos == start) {
			return false;
		}
		outName = m_text.substr(start, m_pos - start);
		return true;
	}

	bool ParseQuotedValue(std::string& outValue)
	{
		if (m_pos >= m_text.size() || m_text[m_pos] != '"') {
			return false;
		}
		++m_pos;
		const size_t start = m_pos;
		while (m_pos < m_text.size() && m_text[m_pos] != '"') {
			++m_pos;
		}
		if (m_pos >= m_text.size()) {
			return false;
		}
		outValue = DecodeXmlEntities(m_text.substr(start, m_pos - start));
		++m_pos;
		return true;
	}

	bool ParseAttributes(
		std::unordered_map<std::string, std::string>& outAttributes,
		bool& outSelfClosing,
		std::string* outError)
	{
		outSelfClosing = false;
		while (m_pos < m_text.size()) {
			SkipWhitespace();
			if (m_pos >= m_text.size()) {
				break;
			}
			if (m_text[m_pos] == '/') {
				++m_pos;
				if (m_pos >= m_text.size() || m_text[m_pos] != '>') {
					if (outError != nullptr) {
						*outError = "xml_self_closing_invalid";
					}
					return false;
				}
				++m_pos;
				outSelfClosing = true;
				return true;
			}
			if (m_text[m_pos] == '>') {
				++m_pos;
				return true;
			}

			std::string key;
			if (!ParseName(key)) {
				if (outError != nullptr) {
					*outError = "xml_attr_name_invalid";
				}
				return false;
			}
			SkipWhitespace();
			if (m_pos >= m_text.size() || m_text[m_pos] != '=') {
				if (outError != nullptr) {
					*outError = "xml_attr_assign_missing";
				}
				return false;
			}
			++m_pos;
			SkipWhitespace();
			std::string value;
			if (!ParseQuotedValue(value)) {
				if (outError != nullptr) {
					*outError = "xml_attr_value_invalid";
				}
				return false;
			}
			outAttributes.insert_or_assign(key, value);
		}

		if (outError != nullptr) {
			*outError = "xml_attr_eof";
		}
		return false;
	}

	bool ParseNode(XmlNode& outNode, std::string* outError)
	{
		if (m_pos >= m_text.size() || m_text[m_pos] != '<') {
			if (outError != nullptr) {
				*outError = "xml_tag_missing";
			}
			return false;
		}
		++m_pos;
		if (!ParseName(outNode.name)) {
			if (outError != nullptr) {
				*outError = "xml_tag_name_invalid";
			}
			return false;
		}

		bool selfClosing = false;
		if (!ParseAttributes(outNode.attributes, selfClosing, outError)) {
			return false;
		}
		if (selfClosing) {
			return true;
		}

		while (m_pos < m_text.size()) {
			SkipWhitespace();
			if (StartsWith(std::string_view(m_text).substr(m_pos), "</")) {
				m_pos += 2;
				std::string closeName;
				if (!ParseName(closeName) || closeName != outNode.name) {
					if (outError != nullptr) {
						*outError = "xml_close_tag_invalid";
					}
					return false;
				}
				SkipWhitespace();
				if (m_pos >= m_text.size() || m_text[m_pos] != '>') {
					if (outError != nullptr) {
						*outError = "xml_close_tag_end_missing";
					}
					return false;
				}
				++m_pos;
				return true;
			}
			if (m_pos < m_text.size() && m_text[m_pos] == '<') {
				XmlNode child;
				if (!ParseNode(child, outError)) {
					return false;
				}
				outNode.children.push_back(std::move(child));
				continue;
			}
			while (m_pos < m_text.size() && m_text[m_pos] != '<') {
				++m_pos;
			}
		}

		if (outError != nullptr) {
			*outError = "xml_close_tag_missing";
		}
		return false;
	}

	const std::string& m_text;
	size_t m_pos = 0;
};

struct ParsedVariableDef {
	std::string name;
	std::string typeName;
	std::string flagsText;
	std::string arrayText;
	std::string comment;
};

struct ParsedMethodDef {
	std::string name;
	std::string returnTypeName;
	bool isPublic = false;
	std::string comment;
	std::vector<ParsedVariableDef> params;
	std::vector<ParsedVariableDef> locals;
	std::vector<std::string> bodyLines;
};

struct ParsedClassDef {
	std::string name;
	std::string baseClassName;
	bool isPublic = false;
	bool isFormClass = false;
	std::string comment;
	std::vector<ParsedVariableDef> vars;
	std::vector<ParsedMethodDef> methods;
};

struct ParsedStructDef {
	std::string name;
	bool isPublic = false;
	std::string comment;
	std::vector<ParsedVariableDef> members;
};

struct ParsedDllDef {
	std::string name;
	std::string returnTypeName;
	std::string fileName;
	std::string commandName;
	bool isPublic = false;
	std::string comment;
	std::vector<ParsedVariableDef> params;
};

struct ParsedConstantDef {
	std::string name;
	std::string valueText;
	bool isLongText = false;
	bool isPublic = false;
	std::string comment;
};

struct ParsedFormDef {
	std::string name;
	std::string comment;
	const FormXml* formXml = nullptr;
};

bool IsLikelyFormClassName(const std::string& rawName)
{
	return StartsWith(TypeResolver::NormalizeTypeName(rawName), "窗口程序集");
}

bool ParseDefinitionFields(
	const std::string& line,
	const std::string& keyword,
	std::vector<std::string>& outFields)
{
	const std::string prefix = "." + keyword;
	if (!StartsWith(line, prefix)) {
		return false;
	}
	std::string rest = TrimAsciiCopy(line.substr(prefix.size()));
	if (rest.empty()) {
		outFields.clear();
		return true;
	}
	outFields = SplitTopLevelCommaFields(rest);
	return true;
}

std::vector<std::int32_t> ParseArrayBounds(const std::string& text)
{
	std::vector<std::int32_t> bounds;
	const std::string raw = Unquote(TrimAsciiCopy(text));
	if (raw.empty()) {
		return bounds;
	}

	size_t start = 0;
	while (start <= raw.size()) {
		const size_t commaPos = raw.find(',', start);
		const std::string part = raw.substr(start, commaPos == std::string::npos ? std::string::npos : commaPos - start);
		if (TrimAsciiCopy(part).empty()) {
			bounds.push_back(0);
		}
		else {
			std::int32_t value = 0;
			bounds.push_back(TryParseInt32(part, value) ? value : 0);
		}
		if (commaPos == std::string::npos) {
			break;
		}
		start = commaPos + 1;
	}
	return bounds;
}

bool HasWordFlag(const std::string& flagsText, const std::string& word)
{
	if (flagsText.empty()) {
		return false;
	}
	std::istringstream stream(flagsText);
	std::string token;
	while (stream >> token) {
		if (token == word) {
			return true;
		}
	}
	return false;
}

bool ParseProgramPage(const Page& page, const std::unordered_set<std::string>& formNames, ParsedClassDef& outClass, std::string* outError)
{
	outClass = {};
	size_t index = 0;
	while (index < page.lines.size() && TrimAsciiCopy(page.lines[index]) != ".版本 2") {
		++index;
	}
	if (index < page.lines.size()) {
		++index;
	}
	while (index < page.lines.size()) {
		const std::string trimmed = TrimAsciiCopy(page.lines[index]);
		if (trimmed.empty() || StartsWith(trimmed, ".支持库 ")) {
			++index;
			continue;
		}
		break;
	}

	std::vector<std::string> fields;
	if (index >= page.lines.size() || !ParseDefinitionFields(TrimAsciiCopy(page.lines[index]), "程序集", fields)) {
		if (outError != nullptr) {
			*outError = "program_page_header_missing";
		}
		return false;
	}

	outClass.name = fields.size() > 0 ? fields[0] : page.name;
	outClass.baseClassName = fields.size() > 1 ? fields[1] : std::string();
	outClass.isPublic = fields.size() > 2 && fields[2] == "公开";
	outClass.comment = fields.size() > 3 ? fields[3] : std::string();
	outClass.isFormClass =
		formNames.contains(outClass.name) ||
		TypeResolver::NormalizeTypeName(outClass.baseClassName) == "窗口" ||
		IsLikelyFormClassName(outClass.name);
	++index;

	while (index < page.lines.size()) {
		const std::string trimmed = TrimAsciiCopy(page.lines[index]);
		if (trimmed.empty()) {
			++index;
			continue;
		}
		if (StartsWith(trimmed, ".程序集变量")) {
			if (!ParseDefinitionFields(trimmed, "程序集变量", fields)) {
				++index;
				continue;
			}
			ParsedVariableDef variable;
			variable.name = fields.size() > 0 ? fields[0] : std::string();
			variable.typeName = fields.size() > 1 ? fields[1] : std::string();
			variable.arrayText = fields.size() > 3 ? fields[3] : std::string();
			variable.comment = fields.size() > 4 ? fields[4] : std::string();
			outClass.vars.push_back(std::move(variable));
			++index;
			continue;
		}
		if (StartsWith(trimmed, ".子程序")) {
			ParsedMethodDef method;
			ParseDefinitionFields(trimmed, "子程序", fields);
			method.name = fields.size() > 0 ? fields[0] : std::string();
			method.returnTypeName = fields.size() > 1 ? fields[1] : std::string();
			method.isPublic = fields.size() > 2 && fields[2] == "公开";
			method.comment = fields.size() > 3 ? fields[3] : std::string();
			++index;

			while (index < page.lines.size()) {
				const std::string line = TrimAsciiCopy(page.lines[index]);
				if (StartsWith(line, ".参数")) {
					ParseDefinitionFields(line, "参数", fields);
					ParsedVariableDef variable;
					variable.name = fields.size() > 0 ? fields[0] : std::string();
					variable.typeName = fields.size() > 1 ? fields[1] : std::string();
					variable.flagsText = fields.size() > 2 ? fields[2] : std::string();
					variable.comment = fields.size() > 3 ? fields[3] : std::string();
					method.params.push_back(std::move(variable));
					++index;
					continue;
				}
				if (StartsWith(line, ".局部变量")) {
					ParseDefinitionFields(line, "局部变量", fields);
					ParsedVariableDef variable;
					variable.name = fields.size() > 0 ? fields[0] : std::string();
					variable.typeName = fields.size() > 1 ? fields[1] : std::string();
					variable.flagsText = fields.size() > 2 ? fields[2] : std::string();
					variable.arrayText = fields.size() > 3 ? fields[3] : std::string();
					variable.comment = fields.size() > 4 ? fields[4] : std::string();
					method.locals.push_back(std::move(variable));
					++index;
					continue;
				}
				break;
			}

			while (index < page.lines.size()) {
				const std::string line = page.lines[index];
				const std::string trimmedLine = TrimAsciiCopy(line);
				if (StartsWith(trimmedLine, ".子程序")) {
					break;
				}
				method.bodyLines.push_back(line);
				++index;
			}
			outClass.methods.push_back(std::move(method));
			continue;
		}
		++index;
	}
	return true;
}

void ParseGlobalPage(const Page& page, std::vector<ParsedVariableDef>& outGlobals)
{
	std::vector<std::string> fields;
	for (const auto& line : page.lines) {
		const std::string trimmed = TrimAsciiCopy(line);
		if (!StartsWith(trimmed, ".全局变量")) {
			continue;
		}
		ParseDefinitionFields(trimmed, "全局变量", fields);
		ParsedVariableDef variable;
		variable.name = fields.size() > 0 ? fields[0] : std::string();
		variable.typeName = fields.size() > 1 ? fields[1] : std::string();
		variable.flagsText = fields.size() > 2 ? fields[2] : std::string();
		variable.arrayText = fields.size() > 3 ? fields[3] : std::string();
		variable.comment = fields.size() > 4 ? fields[4] : std::string();
		outGlobals.push_back(std::move(variable));
	}
}

void ParseStructPage(const Page& page, std::vector<ParsedStructDef>& outStructs)
{
	std::vector<std::string> fields;
	ParsedStructDef* current = nullptr;
	for (const auto& line : page.lines) {
		const std::string trimmed = TrimAsciiCopy(line);
		if (StartsWith(trimmed, ".数据类型")) {
			ParseDefinitionFields(trimmed, "数据类型", fields);
			ParsedStructDef item;
			item.name = fields.size() > 0 ? fields[0] : std::string();
			item.isPublic = fields.size() > 1 && fields[1] == "公开";
			item.comment = fields.size() > 2 ? fields[2] : std::string();
			outStructs.push_back(std::move(item));
			current = &outStructs.back();
			continue;
		}
		if (current != nullptr && StartsWith(trimmed, ".成员")) {
			ParseDefinitionFields(trimmed, "成员", fields);
			ParsedVariableDef member;
			member.name = fields.size() > 0 ? fields[0] : std::string();
			member.typeName = fields.size() > 1 ? fields[1] : std::string();
			member.flagsText = fields.size() > 2 ? fields[2] : std::string();
			member.arrayText = fields.size() > 3 ? fields[3] : std::string();
			member.comment = fields.size() > 4 ? fields[4] : std::string();
			current->members.push_back(std::move(member));
		}
	}
}

void ParseDllPage(const Page& page, std::vector<ParsedDllDef>& outDlls)
{
	std::vector<std::string> fields;
	ParsedDllDef* current = nullptr;
	for (const auto& line : page.lines) {
		const std::string trimmed = TrimAsciiCopy(line);
		if (StartsWith(trimmed, ".DLL命令")) {
			ParseDefinitionFields(trimmed, "DLL命令", fields);
			ParsedDllDef dll;
			dll.name = fields.size() > 0 ? fields[0] : std::string();
			dll.returnTypeName = fields.size() > 1 ? fields[1] : std::string();
			dll.fileName = fields.size() > 2 ? Unquote(fields[2]) : std::string();
			dll.commandName = fields.size() > 3 ? Unquote(fields[3]) : std::string();
			dll.isPublic = fields.size() > 4 && fields[4] == "公开";
			dll.comment = fields.size() > 5 ? fields[5] : std::string();
			outDlls.push_back(std::move(dll));
			current = &outDlls.back();
			continue;
		}
		if (current != nullptr && StartsWith(trimmed, ".参数")) {
			ParseDefinitionFields(trimmed, "参数", fields);
			ParsedVariableDef param;
			param.name = fields.size() > 0 ? fields[0] : std::string();
			param.typeName = fields.size() > 1 ? fields[1] : std::string();
			param.flagsText = fields.size() > 2 ? fields[2] : std::string();
			param.comment = fields.size() > 3 ? fields[3] : std::string();
			current->params.push_back(std::move(param));
		}
	}
}

void ParseConstantPage(const Page& page, std::vector<ParsedConstantDef>& outConstants)
{
	std::vector<std::string> fields;
	for (const auto& line : page.lines) {
		const std::string trimmed = TrimAsciiCopy(line);
		if (!StartsWith(trimmed, ".常量")) {
			continue;
		}
		ParseDefinitionFields(trimmed, "常量", fields);
		ParsedConstantDef item;
		item.name = fields.size() > 0 ? fields[0] : std::string();
		item.valueText = fields.size() > 1 ? Unquote(fields[1]) : std::string();
		std::string decodedText;
		bool decodedLongText = false;
		if (TryDecodeDumpTextLiteral(item.valueText, decodedText, decodedLongText)) {
			item.isLongText = decodedLongText;
		}
		else if (StartsWith(item.valueText, "<文本长度:") && EndsWith(item.valueText, ">")) {
			item.isLongText = true;
		}
		(void)decodedText;
		item.isPublic = fields.size() > 2 && fields[2] == "公开";
		item.comment = fields.size() > 3 ? fields[3] : std::string();
		outConstants.push_back(std::move(item));
	}
}

void ParseWindowPage(const Page& page, std::vector<ParsedFormDef>& outForms)
{
	std::vector<std::string> fields;
	for (const auto& line : page.lines) {
		const std::string trimmed = TrimAsciiCopy(line);
		if (!StartsWith(trimmed, ".窗口")) {
			continue;
		}
		ParseDefinitionFields(trimmed, "窗口", fields);
		ParsedFormDef item;
		item.name = fields.size() > 0 ? fields[0] : std::string();
		item.comment = fields.size() > 1 ? fields[1] : std::string();
		outForms.push_back(std::move(item));
	}
}

std::string GetXmlAttribute(const XmlNode& node, const std::string& key)
{
	if (const auto it = node.attributes.find(key); it != node.attributes.end()) {
		return it->second;
	}
	return std::string();
}

std::int32_t GetXmlIntAttribute(const XmlNode& node, const std::string& key, const std::int32_t defaultValue)
{
	std::int32_t value = 0;
	return TryParseInt32(GetXmlAttribute(node, key), value) ? value : defaultValue;
}

bool GetXmlBoolAttribute(const XmlNode& node, const std::string& key, const bool defaultValue)
{
	const auto value = ParseBoolLiteral(GetXmlAttribute(node, key));
	return value.has_value() ? *value : defaultValue;
}

std::int16_t BuildVariableAttr(const ParsedVariableDef& definition, const bool allowStatic, const bool allowPublic)
{
	std::int16_t attr = 0;
	if (allowStatic && HasWordFlag(definition.flagsText, "静态")) {
		attr |= kVarAttrStatic;
	}
	if (HasWordFlag(definition.flagsText, "参考") || HasWordFlag(definition.flagsText, "传址")) {
		attr |= kVarAttrByRef;
	}
	if (HasWordFlag(definition.flagsText, "可空")) {
		attr |= kVarAttrNullable;
	}
	if (HasWordFlag(definition.flagsText, "数组")) {
		attr |= kVarAttrArray;
	}
	if (allowPublic && HasWordFlag(definition.flagsText, "公开")) {
		attr |= kGlobalAttrPublic;
	}
	if (!definition.arrayText.empty()) {
		attr |= kVarAttrArray;
	}
	return attr;
}

std::int32_t ResolveFormElementTypeId(const std::string& tagName, TypeResolver& resolver)
{
	const std::string normalized = TypeResolver::NormalizeTypeName(tagName);
	if (StartsWith(normalized, "未知类型.Lib")) {
		const size_t dotPos = normalized.find('.', std::string("未知类型.Lib").size());
		if (dotPos != std::string::npos) {
			std::int32_t libraryId = 0;
			std::int32_t typeId = 0;
			if (TryParseInt32(normalized.substr(std::string("未知类型.Lib").size(), dotPos - std::string("未知类型.Lib").size()), libraryId) &&
				TryParseInt32(normalized.substr(dotPos + 1), typeId)) {
				return ((libraryId + 1) << 16) | (typeId + 1);
			}
		}
	}
	if (StartsWith(normalized, "未知类型.")) {
		std::int32_t rawType = 0;
		if (TryParseInt32(normalized.substr(std::string("未知类型.").size()), rawType)) {
			return rawType;
		}
	}
	return resolver.ResolveTypeId(normalized);
}

void BuildFormControlTree(
	const XmlNode& node,
	const std::int32_t parentId,
	TypeResolver& resolver,
	IdAllocator& allocator,
	std::vector<RestoreFormElement>& outElements,
	std::vector<std::int32_t>& outChildren)
{
	RestoreFormElement element;
	element.id = allocator.Alloc(epl_system_id::kTypeFormControl);
	element.dataType = ResolveFormElementTypeId(node.name, resolver);
	element.name = GetXmlAttribute(node, "名称");
	element.comment = GetXmlAttribute(node, "备注");
	element.parent = parentId;
	element.left = GetXmlIntAttribute(node, "左边", 0);
	element.top = GetXmlIntAttribute(node, "顶边", 0);
	element.width = GetXmlIntAttribute(node, "宽度", 0);
	element.height = GetXmlIntAttribute(node, "高度", 0);
	element.tag = GetXmlAttribute(node, "标记");
	element.disable = GetXmlBoolAttribute(node, "禁止", false);
	element.visible = GetXmlBoolAttribute(node, "可视", true);
	element.cursor = DecodeBase64(GetXmlAttribute(node, "鼠标指针"));
	element.tabStop = GetXmlBoolAttribute(node, "可停留焦点", true);
	element.tabIndex = GetXmlIntAttribute(node, "停留顺序", 0);
	element.extensionData = DecodeBase64(GetXmlAttribute(node, "扩展属性数据"));

	std::vector<std::int32_t> childIds;
	const bool isTabControl = resolver.IsTabControlType(element.dataType);
	if (isTabControl) {
		bool firstTab = true;
		for (const auto& child : node.children) {
			if (child.name != node.name + ".子夹") {
				continue;
			}
			if (!firstTab) {
				childIds.push_back(0);
			}
			firstTab = false;
			for (const auto& tabChild : child.children) {
				if (StartsWith(tabChild.name, node.name + ".")) {
					continue;
				}
				BuildFormControlTree(tabChild, element.id, resolver, allocator, outElements, childIds);
			}
		}
	}
	else {
		for (const auto& child : node.children) {
			if (StartsWith(child.name, node.name + ".")) {
				continue;
			}
			BuildFormControlTree(child, element.id, resolver, allocator, outElements, childIds);
		}
	}
	element.children = std::move(childIds);
	outChildren.push_back(element.id);
	outElements.push_back(std::move(element));
}

void BuildFormMenus(
	const XmlNode& node,
	const int level,
	IdAllocator& allocator,
	std::vector<RestoreFormElement>& outElements)
{
	for (const auto& child : node.children) {
		if (child.name != "菜单") {
			continue;
		}
		RestoreFormElement element;
		element.id = allocator.Alloc(epl_system_id::kTypeFormMenu);
		element.dataType = 65539;
		element.isMenu = true;
		element.name = GetXmlAttribute(child, "名称");
		element.text = GetXmlAttribute(child, "标题");
		element.visible = GetXmlBoolAttribute(child, "可视", true);
		element.disable = GetXmlBoolAttribute(child, "禁止", false);
		element.selected = GetXmlBoolAttribute(child, "选中", false);
		element.hotKey = GetXmlIntAttribute(child, "快捷键", 0);
		element.level = level;
		outElements.push_back(std::move(element));
		BuildFormMenus(child, level + 1, allocator, outElements);
	}
}

bool BuildFormsFromXml(
	const std::vector<ParsedFormDef>& parsedForms,
	const std::unordered_map<std::string, std::int32_t>& formClassIds,
	TypeResolver& resolver,
	IdAllocator& allocator,
	std::vector<RestoreForm>& outForms,
	std::string* outError)
{
	outForms.clear();
	for (const auto& formDef : parsedForms) {
		RestoreForm form;
		form.id = allocator.Alloc(epl_system_id::kTypeForm);
		form.classId = 0;
		if (const auto it = formClassIds.find(TypeResolver::NormalizeTypeName(formDef.name)); it != formClassIds.end()) {
			form.classId = it->second;
		}
		form.name = formDef.name;
		form.comment = formDef.comment;

		RestoreFormElement selfElement;
		selfElement.id = allocator.Alloc(epl_system_id::kTypeFormSelf);
		selfElement.dataType = 65537;

		if (formDef.formXml != nullptr) {
			const std::string xmlText = [&]() {
				std::ostringstream stream;
				for (size_t i = 0; i < formDef.formXml->lines.size(); ++i) {
					if (i != 0) {
						stream << "\n";
					}
					stream << formDef.formXml->lines[i];
				}
				return stream.str();
			}();

			XmlNode root;
			SimpleXmlParser parser(xmlText);
			if (!parser.Parse(root, outError)) {
				return false;
			}
			form.name = GetXmlAttribute(root, "名称").empty() ? form.name : GetXmlAttribute(root, "名称");
			form.comment = GetXmlAttribute(root, "备注").empty() ? form.comment : GetXmlAttribute(root, "备注");
			selfElement.left = GetXmlIntAttribute(root, "左边", 0);
			selfElement.top = GetXmlIntAttribute(root, "顶边", 0);
			selfElement.width = GetXmlIntAttribute(root, "宽度", 0);
			selfElement.height = GetXmlIntAttribute(root, "高度", 0);
			selfElement.tag = GetXmlAttribute(root, "标记");
			selfElement.disable = GetXmlBoolAttribute(root, "禁止", false);
			selfElement.visible = GetXmlBoolAttribute(root, "可视", true);
			selfElement.cursor = DecodeBase64(GetXmlAttribute(root, "鼠标指针"));
			selfElement.tabStop = GetXmlBoolAttribute(root, "可停留焦点", true);
			selfElement.tabIndex = GetXmlIntAttribute(root, "停留顺序", 0);
			selfElement.extensionData = DecodeBase64(GetXmlAttribute(root, "扩展属性数据"));

			form.elements.push_back(selfElement);
			for (const auto& child : root.children) {
				if (child.name == "窗口.菜单") {
					BuildFormMenus(child, 0, allocator, form.elements);
				}
			}
			std::vector<std::int32_t> rootChildren;
			for (const auto& child : root.children) {
				if (child.name == "窗口.菜单" || StartsWith(child.name, root.name + ".")) {
					continue;
				}
				BuildFormControlTree(child, 0, resolver, allocator, form.elements, rootChildren);
			}
		}
		else {
			form.elements.push_back(selfElement);
		}
		outForms.push_back(std::move(form));
	}
	return true;
}

std::unordered_map<std::string, size_t> BuildFormClassMatchTable(
	const std::vector<ParsedFormDef>& parsedForms,
	const std::vector<ParsedClassDef>& parsedClasses)
{
	std::unordered_map<std::string, size_t> matches;
	std::vector<bool> classAssigned(parsedClasses.size(), false);

	const auto tryMatch = [&](const size_t formIndex, const std::string& candidateClassName) -> bool {
		const std::string normalizedFormName = TypeResolver::NormalizeTypeName(parsedForms[formIndex].name);
		if (normalizedFormName.empty()) {
			return false;
		}

		const std::string normalizedCandidate = TypeResolver::NormalizeTypeName(candidateClassName);
		if (normalizedCandidate.empty()) {
			return false;
		}

		for (size_t classIndex = 0; classIndex < parsedClasses.size(); ++classIndex) {
			if (classAssigned[classIndex] || !parsedClasses[classIndex].isFormClass) {
				continue;
			}
			if (TypeResolver::NormalizeTypeName(parsedClasses[classIndex].name) != normalizedCandidate) {
				continue;
			}
			classAssigned[classIndex] = true;
			matches.insert_or_assign(normalizedFormName, classIndex);
			return true;
		}
		return false;
	};

	for (size_t formIndex = 0; formIndex < parsedForms.size(); ++formIndex) {
		tryMatch(formIndex, parsedForms[formIndex].name);
	}

	for (size_t formIndex = 0; formIndex < parsedForms.size(); ++formIndex) {
		const std::string normalizedFormName = TypeResolver::NormalizeTypeName(parsedForms[formIndex].name);
		if (matches.contains(normalizedFormName)) {
			continue;
		}
		tryMatch(formIndex, "窗口程序集_" + parsedForms[formIndex].name);
	}

	for (size_t formIndex = 0; formIndex < parsedForms.size(); ++formIndex) {
		const std::string normalizedFormName = TypeResolver::NormalizeTypeName(parsedForms[formIndex].name);
		if (matches.contains(normalizedFormName)) {
			continue;
		}

		if (StartsWith(normalizedFormName, "窗口")) {
			tryMatch(formIndex, "窗口程序集" + normalizedFormName.substr(std::string("窗口").size()));
		}
	}

	std::vector<size_t> remainingForms;
	std::vector<size_t> remainingClasses;
	for (size_t formIndex = 0; formIndex < parsedForms.size(); ++formIndex) {
		const std::string normalizedFormName = TypeResolver::NormalizeTypeName(parsedForms[formIndex].name);
		if (!matches.contains(normalizedFormName)) {
			remainingForms.push_back(formIndex);
		}
	}
	for (size_t classIndex = 0; classIndex < parsedClasses.size(); ++classIndex) {
		if (!classAssigned[classIndex] && parsedClasses[classIndex].isFormClass) {
			remainingClasses.push_back(classIndex);
		}
	}

	if (remainingForms.size() == remainingClasses.size()) {
		for (size_t i = 0; i < remainingForms.size(); ++i) {
			const size_t formIndex = remainingForms[i];
			const size_t classIndex = remainingClasses[i];
			matches.insert_or_assign(TypeResolver::NormalizeTypeName(parsedForms[formIndex].name), classIndex);
			classAssigned[classIndex] = true;
		}
	}

	return matches;
}

bool BuildRestoreModel(const Document& document, RestoreDocumentModel& outModel, std::string* outError)
{
	RestoreDocumentModel model;
	model.sourcePath = document.sourcePath;
	model.projectName = document.projectName.empty() ? "txt2e_project" : document.projectName;
	model.versionText = document.versionText.empty() ? "1.0" : document.versionText;
	for (const auto& dependency : document.dependencies) {
		RestoreDependencyInfo item;
		item.name = dependency.name;
		item.fileName = dependency.fileName;
		item.guid = dependency.guid;
		item.versionText = dependency.versionText;
		item.path = dependency.path;
		item.reExport = dependency.reExport;
		item.isSupportLibrary = dependency.kind == DependencyKind::ELib;
		model.dependencies.push_back(std::move(item));
	}

	std::vector<ParsedClassDef> parsedClasses;
	std::vector<ParsedVariableDef> parsedGlobals;
	std::vector<ParsedStructDef> parsedStructs;
	std::vector<ParsedDllDef> parsedDlls;
	std::vector<ParsedConstantDef> parsedConstants;
	std::vector<ParsedFormDef> parsedForms;
	for (const auto& page : document.pages) {
		if (page.typeName == "窗口/表单") {
			ParseWindowPage(page, parsedForms);
		}
	}
	for (const auto& formXml : document.formXmls) {
		const std::string normalized = TypeResolver::NormalizeTypeName(formXml.name);
		auto it = std::find_if(
			parsedForms.begin(),
			parsedForms.end(),
			[&](const ParsedFormDef& item) { return TypeResolver::NormalizeTypeName(item.name) == normalized; });
		if (it == parsedForms.end()) {
			ParsedFormDef item;
			item.name = formXml.name;
			item.formXml = &formXml;
			parsedForms.push_back(item);
		}
		else {
			it->formXml = &formXml;
		}
	}

	std::unordered_set<std::string> formNames;
	for (const auto& form : parsedForms) {
		formNames.insert(TypeResolver::NormalizeTypeName(form.name));
	}

	for (const auto& page : document.pages) {
		if (page.typeName == "程序集") {
			ParsedClassDef parsedClass;
			if (!ParseProgramPage(page, formNames, parsedClass, outError)) {
				return false;
			}
			parsedClasses.push_back(std::move(parsedClass));
		}
		else if (page.typeName == "全局变量") {
			ParseGlobalPage(page, parsedGlobals);
		}
		else if (page.typeName == "自定义数据类型") {
			ParseStructPage(page, parsedStructs);
		}
		else if (page.typeName == "DLL命令") {
			ParseDllPage(page, parsedDlls);
		}
		else if (page.typeName == "常量资源") {
			ParseConstantPage(page, parsedConstants);
		}
	}

	const auto formClassMatches = BuildFormClassMatchTable(parsedForms, parsedClasses);

	IdAllocator allocator;
	TypeResolver resolver(document.sourcePath, model.dependencies);
	std::unordered_map<std::string, std::int32_t> classIds;
	for (const auto& parsedClass : parsedClasses) {
		RestoreClass item;
		item.id = allocator.Alloc(parsedClass.isFormClass ? epl_system_id::kTypeFormClass : epl_system_id::kTypeStaticClass);
		item.name = parsedClass.name;
		item.comment = parsedClass.comment;
		item.isPublic = parsedClass.isPublic;
		item.isFormClass = parsedClass.isFormClass;
		model.classes.push_back(std::move(item));
		classIds.insert_or_assign(TypeResolver::NormalizeTypeName(parsedClass.name), model.classes.back().id);
		resolver.RegisterUserType(parsedClass.name, model.classes.back().id);
	}

	for (const auto& parsedStruct : parsedStructs) {
		RestoreStruct item;
		item.id = allocator.Alloc(epl_system_id::kTypeStruct);
		item.name = parsedStruct.name;
		item.comment = parsedStruct.comment;
		item.attr = parsedStruct.isPublic ? 0x2 : 0;
		model.structs.push_back(std::move(item));
		resolver.RegisterUserType(parsedStruct.name, model.structs.back().id);
	}

	auto ensureTypeId = [&](const std::string& rawTypeName) -> std::int32_t {
		const std::string typeName = TypeResolver::NormalizeTypeName(rawTypeName);
		if (typeName.empty()) {
			return 0;
		}
		if (const std::int32_t typeId = resolver.ResolveTypeId(typeName); typeId != 0) {
			return typeId;
		}

		RestoreStruct placeholder;
		placeholder.id = allocator.Alloc(epl_system_id::kTypeStruct);
		placeholder.name = typeName;
		placeholder.comment = "txt2e placeholder";
		placeholder.attr = 0;
		placeholder.isPlaceholder = true;
		model.structs.push_back(std::move(placeholder));
		resolver.RegisterPlaceholderType(typeName, model.structs.back().id);
		return model.structs.back().id;
	};

	auto convertVariable = [&](const ParsedVariableDef& definition, const std::int32_t idType, const bool allowStatic, const bool allowPublic) {
		RestoreVariable variable;
		variable.id = allocator.Alloc(idType);
		variable.name = definition.name;
		variable.comment = definition.comment;
		variable.dataType = ensureTypeId(definition.typeName);
		variable.attr = BuildVariableAttr(definition, allowStatic, allowPublic);
		variable.arrayBounds = ParseArrayBounds(definition.arrayText);
		return variable;
	};

	for (size_t classIndex = 0; classIndex < parsedClasses.size(); ++classIndex) {
		const auto& parsedClass = parsedClasses[classIndex];
		auto& targetClass = model.classes[classIndex];
		targetClass.baseClass = parsedClass.isFormClass && TypeResolver::NormalizeTypeName(parsedClass.baseClassName).empty()
			? 65537
			: (TypeResolver::NormalizeTypeName(parsedClass.baseClassName) == "对象" ? -1 : ensureTypeId(parsedClass.baseClassName));
		for (const auto& variable : parsedClass.vars) {
			targetClass.vars.push_back(convertVariable(variable, epl_system_id::kTypeClassMember, false, false));
		}

		for (const auto& parsedMethod : parsedClass.methods) {
			RestoreMethod method;
			method.id = allocator.Alloc(epl_system_id::kTypeMethod);
			method.ownerClass = targetClass.id;
			method.attr = parsedMethod.isPublic ? 0x8 : 0;
			method.returnType = ensureTypeId(parsedMethod.returnTypeName);
			method.name = parsedMethod.name;
			method.comment = parsedMethod.comment;
			for (const auto& param : parsedMethod.params) {
				method.params.push_back(convertVariable(param, epl_system_id::kTypeLocal, false, false));
			}
			for (const auto& local : parsedMethod.locals) {
				method.locals.push_back(convertVariable(local, epl_system_id::kTypeLocal, true, false));
			}
			if (!BuildMethodCodeData(parsedMethod.bodyLines, method, outError)) {
				return false;
			}
			targetClass.functionIds.push_back(method.id);
			model.methods.push_back(std::move(method));
		}
	}

	for (const auto& variable : parsedGlobals) {
		model.globals.push_back(convertVariable(variable, epl_system_id::kTypeGlobal, false, true));
	}

	for (size_t structIndex = 0; structIndex < parsedStructs.size(); ++structIndex) {
		const auto& parsedStruct = parsedStructs[structIndex];
		auto& targetStruct = model.structs[structIndex];
		for (const auto& member : parsedStruct.members) {
			targetStruct.members.push_back(convertVariable(member, epl_system_id::kTypeStructMember, false, false));
		}
	}

	for (const auto& parsedDll : parsedDlls) {
		RestoreDll dll;
		dll.id = allocator.Alloc(epl_system_id::kTypeDll);
		dll.attr = parsedDll.isPublic ? 0x2 : 0;
		dll.returnType = ensureTypeId(parsedDll.returnTypeName);
		dll.name = parsedDll.name;
		dll.comment = parsedDll.comment;
		dll.fileName = parsedDll.fileName;
		dll.commandName = parsedDll.commandName;
		for (const auto& param : parsedDll.params) {
			dll.params.push_back(convertVariable(param, epl_system_id::kTypeDllParameter, false, false));
		}
		model.dlls.push_back(std::move(dll));
	}

	for (const auto& parsedConstant : parsedConstants) {
		RestoreConstant constant;
		constant.id = allocator.Alloc(epl_system_id::kTypeConstant);
		constant.attr = parsedConstant.isPublic ? kConstAttrPublic : 0;
		if (parsedConstant.isLongText) {
			constant.attr |= kConstAttrLongText;
		}
		constant.name = parsedConstant.name;
		constant.comment = parsedConstant.comment;
		constant.valueText = parsedConstant.valueText;
		model.constants.push_back(std::move(constant));
	}

	std::unordered_map<std::string, std::int32_t> formClassIds;
	for (const auto& [formName, classIndex] : formClassMatches) {
		if (classIndex < model.classes.size()) {
			formClassIds.insert_or_assign(formName, model.classes[classIndex].id);
		}
	}
	if (!BuildFormsFromXml(parsedForms, formClassIds, resolver, allocator, model.forms, outError)) {
		return false;
	}

	for (auto& item : model.classes) {
		if (!item.isFormClass) {
			continue;
		}
		for (const auto& form : model.forms) {
			if (form.classId == item.id) {
				item.formId = form.id;
				break;
			}
		}
	}

	outModel = std::move(model);
	return true;
}

void WriteInt32ArrayPayload(ByteWriter& writer, const std::vector<std::int32_t>& values)
{
	for (const auto value : values) {
		writer.WriteI32(value);
	}
}

void WriteInt16ArrayPayload(ByteWriter& writer, const std::vector<std::int16_t>& values)
{
	for (const auto value : values) {
		writer.WriteI16(value);
	}
}

void WriteInt32ArrayWithByteSizePrefix(ByteWriter& writer, const std::vector<std::int32_t>& values)
{
	ByteWriter payload;
	WriteInt32ArrayPayload(payload, values);
	writer.WriteDynamicBytes(payload.bytes());
}

void WriteInt16ArrayWithByteSizePrefix(ByteWriter& writer, const std::vector<std::int16_t>& values)
{
	ByteWriter payload;
	WriteInt16ArrayPayload(payload, values);
	writer.WriteDynamicBytes(payload.bytes());
}

void WriteVariableBlockPayload(ByteWriter& writer, const std::vector<RestoreVariable>& variables)
{
	if (variables.empty()) {
		writer.WriteI32(0);
		writer.WriteDynamicBytes({});
		return;
	}

	ByteWriter payload;
	for (const auto& variable : variables) {
		payload.WriteI32(variable.id);
	}

	std::vector<std::int32_t> offsets;
	offsets.reserve(variables.size());
	ByteWriter body;
	for (const auto& variable : variables) {
		offsets.push_back(static_cast<std::int32_t>(body.position()));
		ByteWriter item;
		item.WriteI32(0);
		item.WriteI32(variable.dataType);
		item.WriteI16(variable.attr);
		item.WriteU8(static_cast<std::uint8_t>(variable.arrayBounds.size()));
		for (const auto bound : variable.arrayBounds) {
			item.WriteI32(bound);
		}
		item.WriteStandardText(variable.name);
		item.WriteStandardText(variable.comment);
		const std::int32_t itemLength = static_cast<std::int32_t>(item.position());
		item.PatchI32(0, itemLength);
		body.WriteBytes(item.bytes());
	}
	for (const auto offset : offsets) {
		payload.WriteI32(offset);
	}
	payload.WriteBytes(body.bytes());

	writer.WriteI32(static_cast<std::int32_t>(variables.size()));
	writer.WriteDynamicBytes(payload.bytes());
}

template <typename TItem, typename TWriter>
void WriteBlocksWithIdAndMemoryAddress(
	ByteWriter& writer,
	const std::vector<TItem>& items,
	TWriter&& itemWriter)
{
	writer.WriteI32(static_cast<std::int32_t>(items.size() * 8));
	for (const auto& item : items) {
		writer.WriteI32(item.id);
	}
	for (const auto& item : items) {
		writer.WriteI32(item.memoryAddress);
	}
	for (const auto& item : items) {
		itemWriter(writer, item);
	}
}

template <typename TItem, typename TWriter>
void WriteBlocksWithIdAndOffset(
	ByteWriter& writer,
	const std::vector<TItem>& items,
	TWriter&& itemWriter)
{
	writer.WriteI32(static_cast<std::int32_t>(items.size()));
	if (items.empty()) {
		writer.WriteI32(0);
		return;
	}

	std::vector<std::vector<std::uint8_t>> encodedItems;
	encodedItems.reserve(items.size());
	for (const auto& item : items) {
		ByteWriter itemData;
		itemWriter(itemData, item);
		ByteWriter withLength;
		withLength.WriteI32(static_cast<std::int32_t>(itemData.position()));
		withLength.WriteBytes(itemData.bytes());
		encodedItems.push_back(withLength.TakeBytes());
	}

	std::int32_t totalSize = static_cast<std::int32_t>(items.size() * 8);
	std::vector<std::int32_t> offsets;
	offsets.reserve(items.size());
	std::int32_t offset = 0;
	for (const auto& encoded : encodedItems) {
		offsets.push_back(offset);
		offset += static_cast<std::int32_t>(encoded.size());
		totalSize += static_cast<std::int32_t>(encoded.size());
	}

	writer.WriteI32(totalSize);
	for (const auto& item : items) {
		writer.WriteI32(item.id);
	}
	for (const auto itemOffset : offsets) {
		writer.WriteI32(itemOffset);
	}
	for (const auto& encoded : encodedItems) {
		writer.WriteBytes(encoded);
	}
}

void WriteFormElements(ByteWriter& writer, const std::vector<RestoreFormElement>& elements)
{
	WriteBlocksWithIdAndOffset(writer, elements, [](ByteWriter& itemWriter, const RestoreFormElement& element) {
		itemWriter.WriteI32(element.dataType);
		itemWriter.WriteBytes(std::vector<std::uint8_t>(20, 0));
		itemWriter.WriteStandardText(element.name);
		if (element.isMenu) {
			itemWriter.WriteStandardText(std::string());
			itemWriter.WriteI32(element.hotKey);
			itemWriter.WriteI32(element.level);
			std::int32_t showStatus = (element.visible ? 0 : 0x1) | (element.disable ? 0x2 : 0) | (element.selected ? 0x4 : 0);
			itemWriter.WriteI32(showStatus);
			itemWriter.WriteStandardText(element.text);
			itemWriter.WriteI32(element.clickEvent);
			itemWriter.WriteBytes(std::vector<std::uint8_t>(16, 0));
			return;
		}

		itemWriter.WriteStandardText(element.comment);
		itemWriter.WriteI32(element.cWndAddress);
		itemWriter.WriteI32(element.left);
		itemWriter.WriteI32(element.top);
		itemWriter.WriteI32(element.width);
		itemWriter.WriteI32(element.height);
		itemWriter.WriteI32(element.unknownBeforeParent);
		itemWriter.WriteI32(element.parent);
		itemWriter.WriteI32(static_cast<std::int32_t>(element.children.size()));
		for (const auto child : element.children) {
			itemWriter.WriteI32(child);
		}
		itemWriter.WriteDynamicBytes(element.cursor);
		itemWriter.WriteStandardText(element.tag);
		itemWriter.WriteI32(element.unknownBeforeVisible);
		std::int32_t showStatus = (element.visible ? 0x1 : 0) | (element.disable ? 0x2 : 0) |
			(element.tabStop ? 0x4 : 0) | (element.locked ? 0x10 : 0);
		itemWriter.WriteI32(showStatus);
		itemWriter.WriteI32(element.tabIndex);
		itemWriter.WriteI32(static_cast<std::int32_t>(element.events.size()));
		for (const auto& [key, value] : element.events) {
			itemWriter.WriteI32(key);
			itemWriter.WriteI32(value);
		}
		itemWriter.WriteBytes(std::vector<std::uint8_t>(20, 0));
		itemWriter.WriteBytes(element.extensionData);
	});
}

void WriteForms(ByteWriter& writer, const std::vector<RestoreForm>& forms)
{
	WriteBlocksWithIdAndMemoryAddress(writer, forms, [](ByteWriter& out, const RestoreForm& form) {
		out.WriteI32(form.unknown1);
		out.WriteI32(form.classId);
		out.WriteDynamicText(form.name);
		out.WriteDynamicText(form.comment);
		WriteFormElements(out, form.elements);
	});
}

void WriteConstants(ByteWriter& writer, const std::vector<RestoreConstant>& constants, std::string* outError)
{
	WriteBlocksWithIdAndOffset(writer, constants, [&](ByteWriter& out, const RestoreConstant& constant) {
		out.WriteI16(constant.attr);
		out.WriteStandardText(constant.name);
		out.WriteStandardText(constant.comment);
		if (constant.pageType == kConstPageImage || constant.pageType == kConstPageSound) {
			out.WriteDynamicBytes({});
			return;
		}

		const std::string valueText = constant.valueText;
		std::string decodedText;
		bool decodedLongText = false;
		if (TryDecodeDumpTextLiteral(valueText, decodedText, decodedLongText)) {
			out.WriteU8(kConstTypeText);
			out.WriteBStr(std::make_optional(decodedText));
			return;
		}
		(void)decodedLongText;
		if (valueText.empty()) {
			out.WriteU8(kConstTypeEmpty);
			return;
		}
		if (StartsWith(valueText, "[") && EndsWith(valueText, "]")) {
			// v1 仅保留格式，可继续扩展成完整日期解析。
			out.WriteU8(kConstTypeText);
			out.WriteBStr(std::make_optional(valueText));
			return;
		}
		if (StartsWith(valueText, "“") && EndsWith(valueText, "”")) {
			out.WriteU8(kConstTypeText);
			out.WriteBStr(std::make_optional(StripWrappedText(valueText, "“", "”")));
			return;
		}

		if (StartsWith(valueText, "<文本长度:") && EndsWith(valueText, ">")) {
			const size_t colonPos = valueText.find(':');
			const size_t endPos = valueText.rfind('>');
			std::int32_t length = 0;
			if (colonPos != std::string::npos && endPos != std::string::npos && colonPos + 1 < endPos) {
				TryParseInt32(valueText.substr(colonPos + 1, endPos - colonPos - 1), length);
			}
			out.WriteU8(kConstTypeText);
			out.WriteBStr(std::make_optional(std::string((std::max)(length, 0), ' ')));
			return;
		}

		double numberValue = 0.0;
		if (TryParseDouble(valueText, numberValue)) {
			out.WriteU8(kConstTypeNumber);
			out.WriteDouble(numberValue);
			return;
		}
		if (const auto boolValue = ParseBoolLiteral(valueText); boolValue.has_value()) {
			out.WriteU8(kConstTypeBool);
			out.WriteBool32(*boolValue);
			return;
		}

		out.WriteU8(kConstTypeText);
		out.WriteBStr(std::make_optional(valueText));
	});
	(void)outError;
}

std::uint32_t ComputeChecksum(const std::vector<std::uint8_t>& data)
{
	std::array<std::uint8_t, 4> checksum = {};
	for (size_t i = 0; i < data.size(); ++i) {
		checksum[i & 0x3] ^= data[i];
	}
	return
		(static_cast<std::uint32_t>(checksum[3]) << 24) |
		(static_cast<std::uint32_t>(checksum[2]) << 16) |
		(static_cast<std::uint32_t>(checksum[1]) << 8) |
		static_cast<std::uint32_t>(checksum[0]);
}

std::array<std::uint8_t, 30> EncodeSectionName(const std::uint32_t key, const std::string& name)
{
	std::array<std::uint8_t, 30> encoded = {};
	if (!name.empty()) {
		std::memcpy(encoded.data(), name.data(), (std::min)(encoded.size(), name.size()));
	}
	if (key != kSectionEndOfFile) {
		const auto* keyBytes = reinterpret_cast<const std::uint8_t*>(&key);
		for (size_t i = 0; i < encoded.size(); ++i) {
			encoded[i] ^= keyBytes[(i + 1) % 4];
		}
	}
	return encoded;
}

void WriteSection(
	ByteWriter& writer,
	const std::uint32_t key,
	const std::string& name,
	const bool isOptional,
	const std::int32_t index,
	const std::vector<std::uint8_t>& data)
{
	ByteWriter headerInfo;
	headerInfo.WriteU32(key);
	const auto encodedName = EncodeSectionName(key, name);
	headerInfo.WriteRaw(encodedName.data(), encodedName.size());
	headerInfo.WriteI16(0);
	headerInfo.WriteI32(index);
	headerInfo.WriteI32(isOptional ? 1 : 0);
	headerInfo.WriteI32(static_cast<std::int32_t>(ComputeChecksum(data)));
	headerInfo.WriteI32(static_cast<std::int32_t>(data.size()));
	for (int i = 0; i < 10; ++i) {
		headerInfo.WriteI32(0);
	}

	writer.WriteU32(kMagicSection);
	writer.WriteU32(ComputeChecksum(headerInfo.bytes()));
	writer.WriteBytes(headerInfo.bytes());
	writer.WriteBytes(data);
}

std::pair<std::int32_t, std::int32_t> ParseVersionPair(const std::string& versionText)
{
	std::int32_t major = 1;
	std::int32_t minor = 0;
	const size_t dotPos = versionText.find('.');
	if (dotPos == std::string::npos) {
		TryParseInt32(versionText, major);
		return { major, minor };
	}
	TryParseInt32(versionText.substr(0, dotPos), major);
	TryParseInt32(versionText.substr(dotPos + 1), minor);
	return { major, minor };
}

std::vector<std::string> BuildSupportLibraryInfoText(const std::vector<RestoreDependencyInfo>& dependencies)
{
	std::vector<std::string> values;
	for (const auto& dependency : dependencies) {
		if (!dependency.isSupportLibrary) {
			continue;
		}
		auto [major, minor] = ParseVersionPair(dependency.versionText);
		values.push_back(
			dependency.fileName + "\r" +
			dependency.guid + "\r" +
			std::to_string(major) + "\r" +
			std::to_string(minor) + "\r" +
			dependency.name);
	}
	return values;
}

std::int32_t ComputeAllocatedIdNum(const RestoreDocumentModel& model)
{
	std::int32_t maxId = 0xFFFF;
	const auto update = [&](const std::int32_t id) {
		if ((id & epl_system_id::kMaskType) != 0) {
			maxId = (std::max)(maxId, id & epl_system_id::kMaskNum);
		}
	};

	for (const auto& item : model.classes) {
		update(item.id);
		for (const auto& variable : item.vars) {
			update(variable.id);
		}
	}
	for (const auto& item : model.methods) {
		update(item.id);
		for (const auto& variable : item.params) {
			update(variable.id);
		}
		for (const auto& variable : item.locals) {
			update(variable.id);
		}
	}
	for (const auto& item : model.globals) {
		update(item.id);
	}
	for (const auto& item : model.structs) {
		update(item.id);
		for (const auto& member : item.members) {
			update(member.id);
		}
	}
	for (const auto& item : model.dlls) {
		update(item.id);
		for (const auto& param : item.params) {
			update(param.id);
		}
	}
	for (const auto& item : model.constants) {
		update(item.id);
	}
	for (const auto& item : model.forms) {
		update(item.id);
		for (const auto& element : item.elements) {
			update(element.id);
		}
	}
	return maxId;
}

std::vector<std::uint8_t> BuildSystemInfoSection(const RestoreDocumentModel&)
{
	ByteWriter writer;
	writer.WriteI16(5);
	writer.WriteI16(6);
	writer.WriteI32(1);
	writer.WriteI32(1);
	writer.WriteI16(1);
	writer.WriteI16(7);
	writer.WriteI32(1);
	writer.WriteI32(0);
	writer.WriteI32(0);
	for (int i = 0; i < 8; ++i) {
		writer.WriteI32(0);
	}
	return writer.TakeBytes();
}

std::vector<std::uint8_t> BuildProjectConfigSection(const RestoreDocumentModel& model)
{
	const auto [major, minor] = ParseVersionPair(model.versionText);
	ByteWriter writer;
	writer.WriteDynamicText(model.projectName);
	writer.WriteDynamicText(std::string());
	writer.WriteDynamicText(std::string());
	writer.WriteDynamicText(std::string());
	writer.WriteDynamicText(std::string());
	writer.WriteDynamicText(std::string());
	writer.WriteDynamicText(std::string());
	writer.WriteDynamicText(std::string());
	writer.WriteDynamicText(std::string());
	writer.WriteDynamicText(std::string());
	writer.WriteI32(major);
	writer.WriteI32(minor);
	writer.WriteI32(0);
	writer.WriteI32(0);
	writer.WriteI32(0);
	writer.WriteRaw(std::vector<std::uint8_t>(20, 0).data(), 20);
	writer.WriteI32(0);
	writer.WriteI32(0);
	return writer.TakeBytes();
}

std::vector<std::uint8_t> BuildCodeSection(const RestoreDocumentModel& model)
{
	ByteWriter writer;
	writer.WriteI32(ComputeAllocatedIdNum(model));
	writer.WriteI32(kProgramHeaderUnk1);

	const size_t supportCount = std::count_if(
		model.dependencies.begin(),
		model.dependencies.end(),
		[](const RestoreDependencyInfo& item) { return item.isSupportLibrary; });
	std::vector<std::int32_t> minCmd(static_cast<size_t>(supportCount), 0);
	std::vector<std::int16_t> minType(static_cast<size_t>(supportCount), 0);
	std::vector<std::int16_t> minConst(static_cast<size_t>(supportCount), 0);
	WriteInt32ArrayWithByteSizePrefix(writer, minCmd);
	WriteInt16ArrayWithByteSizePrefix(writer, minType);
	WriteInt16ArrayWithByteSizePrefix(writer, minConst);
	writer.WriteTextArray(BuildSupportLibraryInfoText(model.dependencies));
	writer.WriteI32(0);
	writer.WriteI32(0);
	writer.WriteDynamicBytes({});
	writer.WriteDynamicText(std::string());

	WriteBlocksWithIdAndMemoryAddress(writer, model.classes, [](ByteWriter& out, const RestoreClass& item) {
		out.WriteI32(item.formId);
		out.WriteI32(item.baseClass);
		out.WriteDynamicText(item.name);
		out.WriteDynamicText(item.comment);
		out.WriteI32(static_cast<std::int32_t>(item.functionIds.size() * 4));
		for (const auto functionId : item.functionIds) {
			out.WriteI32(functionId);
		}
		WriteVariableBlockPayload(out, item.vars);
	});

	WriteBlocksWithIdAndMemoryAddress(writer, model.methods, [](ByteWriter& out, const RestoreMethod& item) {
		out.WriteI32(item.ownerClass);
		out.WriteI32(item.attr);
		out.WriteI32(item.returnType);
		out.WriteDynamicText(item.name);
		out.WriteDynamicText(item.comment);
		WriteVariableBlockPayload(out, item.locals);
		WriteVariableBlockPayload(out, item.params);
		out.WriteDynamicBytes(item.lineOffset);
		out.WriteDynamicBytes(item.blockOffset);
		out.WriteDynamicBytes(item.methodReference);
		out.WriteDynamicBytes(item.variableReference);
		out.WriteDynamicBytes(item.constantReference);
		out.WriteDynamicBytes(item.expressionData);
	});

	WriteVariableBlockPayload(writer, model.globals);

	WriteBlocksWithIdAndMemoryAddress(writer, model.structs, [](ByteWriter& out, const RestoreStruct& item) {
		out.WriteI32(item.attr);
		out.WriteDynamicText(item.name);
		out.WriteDynamicText(item.comment);
		WriteVariableBlockPayload(out, item.members);
	});

	WriteBlocksWithIdAndMemoryAddress(writer, model.dlls, [](ByteWriter& out, const RestoreDll& item) {
		out.WriteI32(item.attr);
		out.WriteI32(item.returnType);
		out.WriteDynamicText(item.name);
		out.WriteDynamicText(item.comment);
		out.WriteDynamicText(item.fileName);
		out.WriteDynamicText(item.commandName);
		WriteVariableBlockPayload(out, item.params);
	});

	writer.WriteBytes(std::vector<std::uint8_t>(40, 0));
	return writer.TakeBytes();
}

std::vector<std::uint8_t> BuildResourceSection(const RestoreDocumentModel& model, std::string* outError)
{
	ByteWriter writer;
	WriteForms(writer, model.forms);
	WriteConstants(writer, model.constants, outError);
	writer.WriteI32(0);
	return writer.TakeBytes();
}

std::vector<std::uint8_t> BuildClassPublicitySection(const RestoreDocumentModel& model)
{
	ByteWriter writer;
	for (const auto& item : model.classes) {
		if (!item.isPublic) {
			continue;
		}
		writer.WriteI32(item.id);
		writer.WriteI32(1);
	}
	return writer.TakeBytes();
}

std::vector<std::uint8_t> BuildLosableSection()
{
	ByteWriter writer;
	writer.WriteDynamicText(std::string());
	writer.WriteI32(0);
	writer.WriteI32(0);
	writer.WriteBytes(std::vector<std::uint8_t>(16, 0));
	writer.WriteI32(-1);
	return writer.TakeBytes();
}

std::vector<std::uint8_t> BuildEcDependenciesSection(const RestoreDocumentModel& model)
{
	ByteWriter writer;
	std::vector<const RestoreDependencyInfo*> ecomDependencies;
	for (const auto& dependency : model.dependencies) {
		if (!dependency.isSupportLibrary) {
			ecomDependencies.push_back(&dependency);
		}
	}

	writer.WriteI32(static_cast<std::int32_t>(ecomDependencies.size()));
	for (const auto* dependency : ecomDependencies) {
		writer.WriteI32(2);

		std::int64_t fileTime = 0;
		std::int32_t fileSize = 0;
		if (!dependency->path.empty()) {
			std::error_code ec;
			const std::filesystem::path path(dependency->path);
			if (std::filesystem::exists(path, ec)) {
				fileSize = static_cast<std::int32_t>(std::filesystem::file_size(path, ec));
				if (!ec) {
					const auto writeTime = std::filesystem::last_write_time(path, ec);
					if (!ec) {
						const auto systemNow = std::chrono::system_clock::now();
						const auto fileNow = decltype(writeTime)::clock::now();
						const auto systemTime = systemNow + (writeTime - fileNow);
						fileTime = std::chrono::duration_cast<std::chrono::nanoseconds>(systemTime.time_since_epoch()).count() / 100 + 116444736000000000LL;
					}
				}
			}
		}
		writer.WriteI32(fileSize);
		writer.WriteI64(fileTime);
		writer.WriteI32(dependency->reExport ? 1 : 0);
		writer.WriteDynamicText(dependency->name);
		writer.WriteDynamicText(dependency->path);
		WriteInt32ArrayWithByteSizePrefix(writer, {});
		WriteInt32ArrayWithByteSizePrefix(writer, {});
	}
	return writer.TakeBytes();
}

bool SerializeToModuleBytes(const RestoreDocumentModel& model, std::vector<std::uint8_t>& outBytes, std::string* outError)
{
	if (outError != nullptr) {
		outError->clear();
	}

	ByteWriter file;
	file.WriteU32(kMagicFileHeader1);
	file.WriteU32(kMagicFileHeader2);

	int sectionIndex = 1;
	WriteSection(file, kSectionSystemInfo, "系统信息段", false, sectionIndex++, BuildSystemInfoSection(model));
	WriteSection(file, kSectionProjectConfig, "用户信息段", false, sectionIndex++, BuildProjectConfigSection(model));

	std::string resourceError;
	const std::vector<std::uint8_t> resourceBytes = BuildResourceSection(model, &resourceError);
	if (!resourceError.empty()) {
		if (outError != nullptr) {
			*outError = resourceError;
		}
		return false;
	}
	WriteSection(file, kSectionResource, "程序资源段", false, sectionIndex++, resourceBytes);
	WriteSection(file, kSectionCode, "程序段", false, sectionIndex++, BuildCodeSection(model));
	WriteSection(file, kSectionEcDependencies, "易模块记录段", false, sectionIndex++, BuildEcDependenciesSection(model));

	const std::vector<std::uint8_t> publicityBytes = BuildClassPublicitySection(model);
	if (!publicityBytes.empty()) {
		WriteSection(file, kSectionClassPublicity, "辅助信息段2", true, sectionIndex++, publicityBytes);
	}

	WriteSection(file, kSectionLosable, "可丢失程序段", true, sectionIndex++, BuildLosableSection());
	WriteSection(file, kSectionEndOfFile, "", false, sectionIndex++, {});
	outBytes = file.TakeBytes();
	return true;
}

}  // namespace

bool Restorer::ParseText(const std::string& inputPath, Document& outDocument, std::string* outError) const
{
	if (outError != nullptr) {
		outError->clear();
	}

	std::vector<std::uint8_t> bytes;
	if (!ReadFileBytes(inputPath, bytes)) {
		if (outError != nullptr) {
			*outError = "read_input_failed";
		}
		return false;
	}

	const std::string text = RemoveUtf8Bom(std::string(reinterpret_cast<const char*>(bytes.data()), bytes.size()));
	const auto lines = SplitLines(text);
	if (lines.empty() || lines.front() != "e2txt Generated Dump") {
		if (outError != nullptr) {
			*outError = "dump_header_invalid";
		}
		return false;
	}

	Document document;
	document.outputPath = inputPath;
	size_t index = 1;
	for (; index < lines.size(); ++index) {
		const std::string& line = lines[index];
		if (line == "================================================================================") {
			break;
		}
		if (line.empty()) {
			continue;
		}

		std::string value;
		if (ParseHeaderValueLine(line, "source", value)) {
			document.sourcePath = value;
		}
		else if (ParseHeaderValueLine(line, "output", value)) {
			document.outputPath = value;
		}
		else if (ParseHeaderValueLine(line, "project", value)) {
			document.projectName = value;
		}
		else if (ParseHeaderValueLine(line, "version", value)) {
			document.versionText = value;
		}
	}

	std::vector<DumpBlock> blocks;
	if (!ParseDumpBlocks(lines, blocks, outError)) {
		return false;
	}

	for (const auto& block : blocks) {
		if (block.kind == DumpBlock::Kind::Dependencies) {
			for (const auto& line : block.lines) {
				const std::string trimmed = TrimAsciiCopy(line);
				if (StartsWith(trimmed, "ELib ")) {
					Dependency dependency;
					dependency.kind = DependencyKind::ELib;
					dependency.name = ExtractNamedSegment(trimmed, "name", std::make_optional(std::string("file")));
					dependency.fileName = ExtractNamedSegment(trimmed, "file", std::make_optional(std::string("guid")));
					dependency.guid = ExtractNamedSegment(trimmed, "guid", std::make_optional(std::string("version")));
					dependency.versionText = ExtractNamedSegment(trimmed, "version", std::nullopt);
					document.dependencies.push_back(std::move(dependency));
				}
				else if (StartsWith(trimmed, "ECom ")) {
					Dependency dependency;
					dependency.kind = DependencyKind::ECom;
					dependency.name = ExtractNamedSegment(trimmed, "name", std::make_optional(std::string("path")));
					dependency.path = ExtractNamedSegment(trimmed, "path", std::make_optional(std::string("re_export")));
					const std::string reExportText = ExtractNamedSegment(trimmed, "re_export", std::nullopt);
					dependency.reExport = reExportText == "true" || reExportText == "1";
					document.dependencies.push_back(std::move(dependency));
				}
			}
			continue;
		}

		if (block.kind == DumpBlock::Kind::Page) {
			Page page;
			page.typeName = block.pageType;
			page.name = block.name;
			page.lines = block.lines;
			document.pages.push_back(std::move(page));
			continue;
		}

		FormXml formXml;
		formXml.name = block.name;
		formXml.lines = block.lines;
		document.formXmls.push_back(std::move(formXml));
	}

	outDocument = std::move(document);
	return true;
}

bool Restorer::RestoreToBytes(const Document& document, std::vector<std::uint8_t>& outBytes, std::string* outError) const
{
	if (outError != nullptr) {
		outError->clear();
	}

	RestoreDocumentModel model;
	if (!BuildRestoreModel(document, model, outError)) {
		return false;
	}
	return SerializeToModuleBytes(model, outBytes, outError);
}

bool Restorer::RestoreToFile(
	const std::string& inputPath,
	const std::string& outputPath,
	std::string* outSummary,
	std::string* outError) const
{
	if (outError != nullptr) {
		outError->clear();
	}
	if (outSummary != nullptr) {
		outSummary->clear();
	}

	Document document;
	if (!ParseText(inputPath, document, outError)) {
		return false;
	}

	std::vector<std::uint8_t> bytes;
	if (!RestoreToBytes(document, bytes, outError)) {
		return false;
	}

	std::ofstream out(outputPath, std::ios::binary);
	if (!out.is_open()) {
		if (outError != nullptr) {
			*outError = "open_output_failed";
		}
		return false;
	}
	out.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
	if (!out.good()) {
		if (outError != nullptr) {
			*outError = "write_output_failed";
		}
		return false;
	}

	if (outSummary != nullptr) {
		*outSummary = "bytes=" + std::to_string(bytes.size()) + ", output=" + outputPath;
	}
	return true;
}

}  // namespace e2txt
