#include "RealPageCodeToolSupport.h"

#include <Windows.h>
#include <bcrypt.h>

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <format>
#include <limits>
#include <mutex>
#include <string_view>
#include <vector>

#pragma comment(lib, "bcrypt.lib")

namespace {

constexpr std::string_view kDirectiveSubroutine = ".子程序";
constexpr std::string_view kDirectiveParameter = ".参数";
constexpr std::string_view kDirectiveLocalVariable = ".局部变量";
constexpr std::string_view kDirectiveAssemblyVariable = ".程序集变量";
// 类声明指令（注意必须区别于 .程序集变量：后者是变量声明）。
constexpr std::string_view kDirectiveAssembly = ".程序集";
// 易语言默认基类标注。声明 `.程序集 X, <对象>` 与 `.程序集 X` 完全等价：
// IDE 会把显式的默认基类省略掉再存，读回时即变成无基类形式。比较时需归一化两者。
constexpr std::string_view kDefaultBaseClassToken = "<对象>";
// 结构指纹里的条件语句 token（已去空白后的形态）。
constexpr std::string_view kFingerprintTokenVersionPrefix = ".版本";
constexpr std::string_view kFingerprintTokenOtherwise = ".否则";
constexpr std::string_view kFingerprintTokenEndIf = ".如果结束";
constexpr std::string_view kFingerprintTokenDefault = ".默认";
constexpr std::string_view kFingerprintTokenEndSwitch = ".判断结束";
// 去空白后「无子程序名」的裸 .子程序 token（真实子程序声明必带名字，去空白后形如 ".子程序xxx"）。
constexpr std::string_view kFingerprintTokenBareSubroutine = ".子程序";
constexpr std::string_view kIdeLeftDoubleQuote = "“";
constexpr std::string_view kIdeRightDoubleQuote = "”";

struct Sha256Provider {
	BCRYPT_ALG_HANDLE algorithm = nullptr;
	DWORD objectBytes = 0;
	DWORD hashBytes = 0;
	bool ready = false;

	~Sha256Provider()
	{
		if (algorithm != nullptr) {
			BCryptCloseAlgorithmProvider(algorithm, 0);
		}
	}
};

Sha256Provider& GetSha256Provider()
{
	static Sha256Provider provider;
	static std::once_flag once;
	Sha256Provider* providerPtr = &provider;
	std::call_once(once, [providerPtr]() {
		Sha256Provider& provider = *providerPtr;
		DWORD resultBytes = 0;
		if (BCryptOpenAlgorithmProvider(&provider.algorithm, BCRYPT_SHA256_ALGORITHM, nullptr, 0) < 0 ||
			BCryptGetProperty(
				provider.algorithm,
				BCRYPT_OBJECT_LENGTH,
				reinterpret_cast<PUCHAR>(&provider.objectBytes),
				sizeof(provider.objectBytes),
				&resultBytes,
				0) < 0 ||
			BCryptGetProperty(
				provider.algorithm,
				BCRYPT_HASH_LENGTH,
				reinterpret_cast<PUCHAR>(&provider.hashBytes),
				sizeof(provider.hashBytes),
				&resultBytes,
				0) < 0 ||
			provider.objectBytes == 0 || provider.hashBytes == 0) {
			if (provider.algorithm != nullptr) {
				BCryptCloseAlgorithmProvider(provider.algorithm, 0);
				provider.algorithm = nullptr;
			}
			return;
		}
		provider.ready = true;
	});
	return provider;
}

bool TryBuildSha256Hex(const std::string& text, std::string& outHex)
{
	outHex.clear();
	Sha256Provider& provider = GetSha256Provider();
	if (!provider.ready) {
		return false;
	}
	std::vector<unsigned char> hashObject(provider.objectBytes);
	std::vector<unsigned char> hash(provider.hashBytes);
	BCRYPT_HASH_HANDLE handle = nullptr;
	if (BCryptCreateHash(
			provider.algorithm,
			&handle,
			hashObject.data(),
			static_cast<ULONG>(hashObject.size()),
			nullptr,
			0,
			0) < 0) {
		return false;
	}
	bool ok = true;
	size_t offset = 0;
	while (offset < text.size()) {
		const ULONG chunk = static_cast<ULONG>((std::min)(
			text.size() - offset,
			static_cast<size_t>((std::numeric_limits<ULONG>::max)())));
		if (BCryptHashData(
				handle,
				reinterpret_cast<PUCHAR>(const_cast<char*>(text.data() + offset)),
				chunk,
				0) < 0) {
			ok = false;
			break;
		}
		offset += chunk;
	}
	if (ok && BCryptFinishHash(handle, hash.data(), static_cast<ULONG>(hash.size()), 0) < 0) {
		ok = false;
	}
	BCryptDestroyHash(handle);
	if (!ok) {
		return false;
	}
	static constexpr char kHex[] = "0123456789abcdef";
	outHex.resize(hash.size() * 2);
	for (size_t i = 0; i < hash.size(); ++i) {
		outHex[i * 2] = kHex[(hash[i] >> 4) & 0x0F];
		outHex[i * 2 + 1] = kHex[hash[i] & 0x0F];
	}
	return true;
}

struct DiffOp {
	char type = ' ';
	std::string line;
	int beforeOldLine = 0;
	int beforeNewLine = 0;
};

struct OperatorForm {
	std::string_view text;
	char canonical = '\0';
};

std::string TrimAsciiCopyLocal(const std::string& text)
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

std::string ToLowerAsciiCopyLocal(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

bool StartsWithLocal(const std::string& text, std::string_view prefix)
{
	return text.size() >= prefix.size() &&
		std::equal(prefix.begin(), prefix.end(), text.begin());
}

bool HasIdeControlTokenBoundary(const std::string& token, size_t keywordSize)
{
	if (token.size() == keywordSize) {
		return true;
	}
	const unsigned char next = static_cast<unsigned char>(token[keywordSize]);
	return next == '(';
}

bool IsIdeOptionalLeadingDotControlToken(const std::string& token)
{
	if (token.empty() || token.front() == '.') {
		return false;
	}

	static constexpr std::string_view kIdeControlKeywords[] = {
		"如果真",
		"如果",
		"否则",
		"如果结束",
		"计次循环首",
		"计次循环尾",
		"判断开始",
		"判断",
		"默认",
		"判断结束",
		"循环判断首",
		"循环判断尾",
		"变量循环首",
		"变量循环尾",
		"到循环尾",
		"跳出循环",
		"返回",
	};

	for (const std::string_view keyword : kIdeControlKeywords) {
		if (StartsWithLocal(token, keyword) &&
			HasIdeControlTokenBoundary(token, keyword.size())) {
			return true;
		}
	}
	return false;
}

std::string NormalizeIdeOptionalLeadingDotToken(const std::string& token)
{
	if (!IsIdeOptionalLeadingDotControlToken(token)) {
		return token;
	}
	return "." + token;
}

std::string NormalizeIdeTrailingEmptyDirectiveFields(std::string token)
{
	// 易语言会省略声明行尾部的空字段，例如：
	//   .子程序 Foo,        -> .子程序 Foo
	//   .参数 value, 整数型, -> .参数 value, 整数型
	// 结构指纹已经去除了空白；仅对指令 token 去掉尾随 ASCII 逗号，
	// 不会影响说明文字或字符串内部的逗号（它们不会位于 token 末尾）。
	if (token.empty() || token.front() != '.') {
		return token;
	}
	while (!token.empty() && token.back() == ',') {
		token.pop_back();
	}
	return token;
}

bool IsVersionFingerprintToken(const std::string& token)
{
	return StartsWithLocal(token, kFingerprintTokenVersionPrefix);
}

struct DoubleQuoteNormalizedText {
	std::string text;
	std::vector<size_t> sourceBoundaries;
};

struct EquivalentTextMatchSpan {
	size_t offset = 0;
	size_t length = 0;
};

// 将三种双引号折叠为同一字符，并保留规范化文本到原文本的字节边界映射。
bool StartsWithAtLocal(const std::string& text, size_t offset, std::string_view value)
{
	return offset <= text.size() &&
		value.size() <= text.size() - offset &&
		text.compare(offset, value.size(), value.data(), value.size()) == 0;
}

size_t GetEquivalentDoubleQuoteBytes(const std::string& text, size_t offset)
{
	if (offset < text.size() && text[offset] == '"') {
		return 1;
	}
	if (StartsWithAtLocal(text, offset, kIdeLeftDoubleQuote)) {
		return kIdeLeftDoubleQuote.size();
	}
	if (StartsWithAtLocal(text, offset, kIdeRightDoubleQuote)) {
		return kIdeRightDoubleQuote.size();
	}
	return 0;
}

size_t GetLocalCharacterBytes(const std::string& text, size_t offset)
{
	if (offset + 1 < text.size() &&
		IsDBCSLeadByteEx(CP_ACP, static_cast<BYTE>(text[offset])) != FALSE) {
		return 2;
	}
	return 1;
}

DoubleQuoteNormalizedText BuildDoubleQuoteNormalizedText(const std::string& source)
{
	DoubleQuoteNormalizedText result;
	result.text.reserve(source.size());
	result.sourceBoundaries.reserve(source.size() + 1);
	result.sourceBoundaries.push_back(0);

	size_t offset = 0;
	while (offset < source.size()) {
		const size_t quoteBytes = GetEquivalentDoubleQuoteBytes(source, offset);
		if (quoteBytes != 0) {
			result.text.push_back('"');
			offset += quoteBytes;
			result.sourceBoundaries.push_back(offset);
		}
		else {
			const size_t characterBytes = GetLocalCharacterBytes(source, offset);
			result.text.append(source, offset, characterBytes);
			for (size_t byteIndex = 1; byteIndex <= characterBytes; ++byteIndex) {
				result.sourceBoundaries.push_back(offset + byteIndex);
			}
			offset += characterBytes;
		}
	}
	return result;
}

std::vector<EquivalentTextMatchSpan> FindEquivalentTextMatches(
	const std::string& source,
	const std::string& needle)
{
	const DoubleQuoteNormalizedText normalizedSource = BuildDoubleQuoteNormalizedText(source);
	const DoubleQuoteNormalizedText normalizedNeedle = BuildDoubleQuoteNormalizedText(needle);
	std::vector<EquivalentTextMatchSpan> matches;
	if (normalizedNeedle.text.empty()) {
		return matches;
	}

	size_t searchOffset = 0;
	while ((searchOffset = normalizedSource.text.find(normalizedNeedle.text, searchOffset)) != std::string::npos) {
		const size_t sourceBegin = normalizedSource.sourceBoundaries[searchOffset];
		const size_t sourceEnd = normalizedSource.sourceBoundaries[searchOffset + normalizedNeedle.text.size()];
		matches.push_back({sourceBegin, sourceEnd - sourceBegin});
		searchOffset += normalizedNeedle.text.size();
	}
	return matches;
}

bool IsDirectiveLine(const std::string& trimmed, std::string_view directive)
{
	if (!StartsWithLocal(trimmed, directive)) {
		return false;
	}
	if (trimmed.size() == directive.size()) {
		return true;
	}

	const unsigned char next = static_cast<unsigned char>(trimmed[directive.size()]);
	return std::isspace(next) != 0 || next == '(' || next == ',' || next == '\t';
}

bool TryMatchDirectiveAtLineStart(
	const std::string& line,
	std::string_view directive,
	size_t* outDirectivePos = nullptr)
{
	size_t directivePos = 0;
	while (directivePos < line.size() && (line[directivePos] == ' ' || line[directivePos] == '\t')) {
		++directivePos;
	}
	if (line.size() < directivePos + directive.size()) {
		return false;
	}
	if (!std::equal(directive.begin(), directive.end(), line.begin() + static_cast<std::ptrdiff_t>(directivePos))) {
		return false;
	}
	if (line.size() != directivePos + directive.size()) {
		const unsigned char next = static_cast<unsigned char>(line[directivePos + directive.size()]);
		if (std::isspace(next) == 0 && next != '(' && next != ',' && next != '\t') {
			return false;
		}
	}
	if (outDirectivePos != nullptr) {
		*outDirectivePos = directivePos;
	}
	return true;
}

bool TrySplitSubroutineHeaderFields(
	const std::string& line,
	std::vector<std::string>& outFields,
	size_t* outDirectivePos = nullptr)
{
	outFields.clear();
	size_t directivePos = 0;
	if (!TryMatchDirectiveAtLineStart(line, kDirectiveSubroutine, &directivePos)) {
		return false;
	}

	const std::string rest = line.substr(directivePos + kDirectiveSubroutine.size());
	size_t fieldStart = 0;
	// 只解析前三个结构逗号；第三个逗号后的全部内容都是第四字段说明，
	// 说明中的逗号属于正文，不能再产生新的结构字段。
	for (size_t structuralComma = 0; structuralComma < 3; ++structuralComma) {
		const size_t comma = rest.find(',', fieldStart);
		if (comma == std::string::npos) {
			outFields.push_back(TrimAsciiCopyLocal(rest.substr(fieldStart)));
			if (outDirectivePos != nullptr) {
				*outDirectivePos = directivePos;
			}
			return true;
		}
		outFields.push_back(TrimAsciiCopyLocal(rest.substr(fieldStart, comma - fieldStart)));
		fieldStart = comma + 1;
	}
	outFields.push_back(TrimAsciiCopyLocal(rest.substr(fieldStart)));
	if (outDirectivePos != nullptr) {
		*outDirectivePos = directivePos;
	}
	return true;
}

std::vector<RealPageStructuredPatchHunk> BuildFallbackStructuredPatch(
	const std::vector<std::string>& oldLines,
	const std::vector<std::string>& newLines)
{
	size_t prefix = 0;
	while (prefix < oldLines.size() &&
		prefix < newLines.size() &&
		oldLines[prefix] == newLines[prefix]) {
		++prefix;
	}

	size_t oldSuffix = oldLines.size();
	size_t newSuffix = newLines.size();
	while (oldSuffix > prefix &&
		newSuffix > prefix &&
		oldLines[oldSuffix - 1] == newLines[newSuffix - 1]) {
		--oldSuffix;
		--newSuffix;
	}

	if (prefix == oldLines.size() && prefix == newLines.size()) {
		return {};
	}

	RealPageStructuredPatchHunk hunk;
	hunk.oldStart = static_cast<int>(prefix) + 1;
	hunk.newStart = static_cast<int>(prefix) + 1;
	hunk.oldLines = static_cast<int>(oldSuffix - prefix);
	hunk.newLines = static_cast<int>(newSuffix - prefix);

	for (size_t index = prefix; index < oldSuffix; ++index) {
		hunk.lines.push_back("-" + oldLines[index]);
	}
	for (size_t index = prefix; index < newSuffix; ++index) {
		hunk.lines.push_back("+" + newLines[index]);
	}
	return {hunk};
}

}  // namespace

std::string NormalizeRealCodeLineBreaksToCrLf(const std::string& text)
{
	std::string normalized;
	normalized.reserve(text.size() + 8);
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
				++i;
			}
			normalized += "\r\n";
			continue;
		}
		if (ch == '\n') {
			normalized += "\r\n";
			continue;
		}
		normalized.push_back(ch);
	}
	return normalized;
}

std::string NormalizeRealCodeLineBreaksToLf(const std::string& text)
{
	std::string normalized;
	normalized.reserve(text.size());
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
				++i;
			}
			normalized.push_back('\n');
			continue;
		}
		normalized.push_back(ch);
	}
	return normalized;
}

std::vector<std::string> SplitRealCodeLines(const std::string& text)
{
	const std::string normalized = NormalizeRealCodeLineBreaksToLf(text);
	std::vector<std::string> lines;

	size_t start = 0;
	while (start <= normalized.size()) {
		size_t end = normalized.find('\n', start);
		if (end == std::string::npos) {
			end = normalized.size();
		}

		lines.push_back(normalized.substr(start, end - start));
		if (end == normalized.size()) {
			break;
		}
		start = end + 1;
	}

	if (lines.size() == 1 && lines.front().empty() && text.empty()) {
		lines.clear();
	}
	return lines;
}

std::string JoinRealCodeLines(const std::vector<std::string>& lines)
{
	std::string text;
	for (size_t i = 0; i < lines.size(); ++i) {
		if (i != 0) {
			text += "\r\n";
		}
		text += lines[i];
	}
	return text;
}

std::string PrepareNewClassPageLifecycleFunctions(
	const std::string& requestedCode,
	const std::string& ideDefaultCode,
	bool* outChanged,
	bool* outComplete)
{
	if (outChanged != nullptr) {
		*outChanged = false;
	}
	if (outComplete != nullptr) {
		*outComplete = false;
	}

	struct FunctionBlock {
		std::string name;
		std::vector<std::string> lines;
	};
	const auto parseFunctionName = [](const std::string& line, std::string& outName) {
		outName.clear();
		size_t directivePos = 0;
		if (!TryMatchDirectiveAtLineStart(line, kDirectiveSubroutine, &directivePos)) {
			return false;
		}
		size_t nameStart = directivePos + kDirectiveSubroutine.size();
		while (nameStart < line.size() &&
			(line[nameStart] == ' ' || line[nameStart] == '\t')) {
			++nameStart;
		}
		size_t nameEnd = nameStart;
		while (nameEnd < line.size() && line[nameEnd] != ',' && line[nameEnd] != '\t') {
			++nameEnd;
		}
		outName = TrimAsciiCopyLocal(line.substr(nameStart, nameEnd - nameStart));
		return !outName.empty();
	};
	const auto splitFunctions = [&parseFunctionName](
			const std::string& code,
			std::vector<std::string>& outPrefix,
			std::vector<FunctionBlock>& outBlocks) {
		outPrefix.clear();
		outBlocks.clear();
		const std::vector<std::string> lines = SplitRealCodeLines(code);
		std::vector<size_t> starts;
		std::vector<std::string> names;
		for (size_t i = 0; i < lines.size(); ++i) {
			std::string name;
			if (parseFunctionName(lines[i], name)) {
				starts.push_back(i);
				names.push_back(std::move(name));
			}
		}
		const size_t prefixEnd = starts.empty() ? lines.size() : starts.front();
		outPrefix.assign(lines.begin(), lines.begin() + prefixEnd);
		for (size_t i = 0; i < starts.size(); ++i) {
			const size_t end = i + 1 < starts.size() ? starts[i + 1] : lines.size();
			FunctionBlock block;
			block.name = names[i];
			block.lines.assign(lines.begin() + starts[i], lines.begin() + end);
			outBlocks.push_back(std::move(block));
		}
	};

	const std::string normalizedRequested = NormalizeRealCodeLineBreaksToCrLf(requestedCode);
	const std::string normalizedDefault = NormalizeRealCodeLineBreaksToCrLf(ideDefaultCode);
	std::vector<std::string> requestedPrefix;
	std::vector<FunctionBlock> requestedBlocks;
	std::vector<std::string> defaultPrefix;
	std::vector<FunctionBlock> defaultBlocks;
	splitFunctions(normalizedRequested, requestedPrefix, requestedBlocks);
	splitFunctions(normalizedDefault, defaultPrefix, defaultBlocks);

	FunctionBlock initializer;
	FunctionBlock destructor;
	std::vector<FunctionBlock> otherBlocks;
	for (const auto& block : requestedBlocks) {
		if (block.name == "_初始化") {
			if (initializer.lines.empty()) {
				initializer = block;
			}
			continue;
		}
		if (block.name == "_销毁") {
			if (destructor.lines.empty()) {
				destructor = block;
			}
			continue;
		}
		otherBlocks.push_back(block);
	}
	for (const auto& block : defaultBlocks) {
		if (initializer.lines.empty() && block.name == "_初始化") {
			initializer = block;
		}
		else if (destructor.lines.empty() && block.name == "_销毁") {
			destructor = block;
		}
	}

	const bool complete = !initializer.lines.empty() && !destructor.lines.empty();
	if (outComplete != nullptr) {
		*outComplete = complete;
	}
	if (!complete) {
		return normalizedRequested;
	}

	std::vector<std::string> rebuiltLines = requestedPrefix;
	const auto appendBlock = [&rebuiltLines](const FunctionBlock& block) {
		rebuiltLines.insert(rebuiltLines.end(), block.lines.begin(), block.lines.end());
	};
	appendBlock(initializer);
	appendBlock(destructor);
	for (const auto& block : otherBlocks) {
		appendBlock(block);
	}
	const std::string rebuilt = JoinRealCodeLines(rebuiltLines);
	if (outChanged != nullptr) {
		*outChanged = rebuilt != normalizedRequested;
	}
	return rebuilt;
}

std::string PrepareAssemblyVariablesForRealPageWrite(const std::string& text)
{
	std::vector<std::string> lines = SplitRealCodeLines(text);
	bool changed = false;
	for (auto& line : lines) {
		size_t directivePos = 0;
		if (!TryMatchDirectiveAtLineStart(line, kDirectiveAssemblyVariable, &directivePos)) {
			continue;
		}

		line.replace(
			directivePos,
			kDirectiveAssemblyVariable.size(),
			kDirectiveLocalVariable.data(),
			kDirectiveLocalVariable.size());
		changed = true;
	}

	return changed ? JoinRealCodeLines(lines) : text;
}

bool ValidateRealPageSubroutineHeaders(const std::string& text, std::string& outError)
{
	outError.clear();
	const std::vector<std::string> lines = SplitRealCodeLines(text);
	bool insideSubroutine = false;
	bool sawLocalVariable = false;
	bool sawExecutableStatement = false;
	for (size_t lineIndex = 0; lineIndex < lines.size(); ++lineIndex) {
		std::vector<std::string> fields;
		if (TrySplitSubroutineHeaderFields(lines[lineIndex], fields, nullptr)) {
			insideSubroutine = true;
			sawLocalVariable = false;
			sawExecutableStatement = false;
			// 固定槽位：fields[0] 名称，fields[1] 返回值，fields[2] 公开属性，
			// fields[3] 及其后为说明。第三项只能为空或“公开”；不能把说明挤到这里。
			if (fields.size() >= 3 && !fields[2].empty() && fields[2] != "公开") {
				outError = "invalid subroutine header at line " + std::to_string(lineIndex + 1) +
					": field 3 (public attribute) must be empty or public; description must start after the third comma in field 4";
				return false;
			}
			continue;
		}
		if (!insideSubroutine) {
			continue;
		}

		const std::string trimmed = TrimAsciiCopyLocal(lines[lineIndex]);
		if (trimmed.empty() || (!trimmed.empty() && trimmed.front() == '\'')) {
			continue;
		}

		size_t directivePos = 0;
		if (TryMatchDirectiveAtLineStart(lines[lineIndex], kDirectiveParameter, &directivePos)) {
			if (sawLocalVariable || sawExecutableStatement) {
				outError = "invalid subroutine declaration order at line " + std::to_string(lineIndex + 1) +
					": all parameters must precede local variables and executable statements";
				return false;
			}
			continue;
		}
		if (TryMatchDirectiveAtLineStart(lines[lineIndex], kDirectiveLocalVariable, &directivePos)) {
			if (sawExecutableStatement) {
				outError = "invalid subroutine declaration order at line " + std::to_string(lineIndex + 1) +
					": all local variables must precede executable statements";
				return false;
			}
			sawLocalVariable = true;
			continue;
		}
		// 任何非空、非注释且非声明行都视为执行语句/流程指令。
		sawExecutableStatement = true;
	}
	return true;
}

bool ValidateRealPageControlFlow(const std::string& text, std::string& outError)
{
	enum class BlockKind {
		If,
		IfTrue,
		Switch,
		CountLoop,
		WhileLoop,
		VariableLoop
	};
	struct BlockState {
		BlockKind kind;
		size_t lineIndex;
		bool sawBranch = false;
	};

	const auto blockName = [](BlockKind kind) -> const char* {
		switch (kind) {
		case BlockKind::If:
			return ".如果";
		case BlockKind::IfTrue:
			return ".如果真";
		case BlockKind::Switch:
			return ".判断开始";
		case BlockKind::CountLoop:
			return ".计次循环首";
		case BlockKind::WhileLoop:
			return ".循环判断首/.判断循环首";
		default:
			return ".变量循环首";
		}
	};
	const auto matchesDirective = [](const std::string& sourceLine, std::string_view keyword) {
		std::string line = TrimAsciiCopyLocal(sourceLine);
		if (line.empty() || line.front() == '\'') {
			return false;
		}
		if (line.front() == '.') {
			line.erase(line.begin());
		}
		if (!StartsWithLocal(line, keyword)) {
			return false;
		}
		if (line.size() == keyword.size()) {
			return true;
		}
		const unsigned char next = static_cast<unsigned char>(line[keyword.size()]);
		return next == '(' || std::isspace(next) != 0;
	};

	outError.clear();
	std::vector<BlockState> stack;
	const std::vector<std::string> lines = SplitRealCodeLines(text);
	const auto failUnexpected = [&outError, &stack, &blockName](
		size_t lineIndex,
		const char* directive) {
		outError = std::format(
			"invalid control flow at line {}: {} does not match the active {} block",
			lineIndex + 1,
			directive,
			stack.empty() ? "top-level" : blockName(stack.back().kind));
		return false;
	};
	const auto closeBlock = [&stack, &failUnexpected](
		size_t lineIndex,
		BlockKind expected,
		const char* directive) {
		if (stack.empty() || stack.back().kind != expected) {
			return failUnexpected(lineIndex, directive);
		}
		stack.pop_back();
		return true;
	};

	for (size_t lineIndex = 0; lineIndex < lines.size(); ++lineIndex) {
		const std::string& line = lines[lineIndex];
		if (matchesDirective(line, "如果真结束")) {
			if (!closeBlock(lineIndex, BlockKind::IfTrue, ".如果真结束")) {
				return false;
			}
		}
		else if (matchesDirective(line, "如果结束")) {
			if (!closeBlock(lineIndex, BlockKind::If, ".如果结束")) {
				return false;
			}
		}
		else if (matchesDirective(line, "否则")) {
			if (!stack.empty() && stack.back().kind == BlockKind::IfTrue) {
				outError = std::format(
					"invalid control flow at line {}: .否则 cannot be used with .如果真; use .如果/.如果结束 when an else branch is required",
					lineIndex + 1);
				return false;
			}
			if (stack.empty() || stack.back().kind != BlockKind::If || stack.back().sawBranch) {
				return failUnexpected(lineIndex, ".否则");
			}
			stack.back().sawBranch = true;
		}
		else if (matchesDirective(line, "如果真")) {
			stack.push_back({BlockKind::IfTrue, lineIndex});
		}
		else if (matchesDirective(line, "如果")) {
			stack.push_back({BlockKind::If, lineIndex});
		}
		else if (matchesDirective(line, "判断结束")) {
			if (!closeBlock(lineIndex, BlockKind::Switch, ".判断结束")) {
				return false;
			}
		}
		else if (matchesDirective(line, "判断开始")) {
			stack.push_back({BlockKind::Switch, lineIndex});
		}
		else if (matchesDirective(line, "默认")) {
			if (stack.empty() || stack.back().kind != BlockKind::Switch || stack.back().sawBranch) {
				return failUnexpected(lineIndex, ".默认");
			}
			stack.back().sawBranch = true;
		}
		else if (matchesDirective(line, "判断") &&
			(stack.empty() || stack.back().kind != BlockKind::Switch)) {
			return failUnexpected(lineIndex, ".判断");
		}
		else if (matchesDirective(line, "计次循环尾")) {
			if (!closeBlock(lineIndex, BlockKind::CountLoop, ".计次循环尾")) {
				return false;
			}
		}
		else if (matchesDirective(line, "计次循环首")) {
			stack.push_back({BlockKind::CountLoop, lineIndex});
		}
		else if (matchesDirective(line, "循环判断尾") || matchesDirective(line, "判断循环尾")) {
			if (!closeBlock(lineIndex, BlockKind::WhileLoop, ".循环判断尾/.判断循环尾")) {
				return false;
			}
		}
		else if (matchesDirective(line, "循环判断首") || matchesDirective(line, "判断循环首")) {
			stack.push_back({BlockKind::WhileLoop, lineIndex});
		}
		else if (matchesDirective(line, "变量循环尾")) {
			if (!closeBlock(lineIndex, BlockKind::VariableLoop, ".变量循环尾")) {
				return false;
			}
		}
		else if (matchesDirective(line, "变量循环首")) {
			stack.push_back({BlockKind::VariableLoop, lineIndex});
		}
	}

	if (!stack.empty()) {
		outError = std::format(
			"invalid control flow: {} block opened at line {} is not closed",
			blockName(stack.back().kind),
			stack.back().lineIndex + 1);
		return false;
	}
	return true;
}

// 把 IDE 读回时会省略的程序集头部字段归一化掉：
// 保留类名和真实的非默认基类，忽略公开标记、说明等不会出现在真实页格式化结果中的字段。
// 例如 ".程序集 X, , 公开, 说明" → ".程序集 X"，
//      ".程序集 X, Base, 公开, 说明" → ".程序集 X, Base"。
// 这样可以避免成功写入被头部元数据差异误判为 verify_mismatch，进而触发回滚重写。
// 仅处理 .程序集（类声明）行，不会匹配 .程序集变量（指令边界由 TryMatchDirectiveAtLineStart 校验）。
// 返回是否修改了 line。
bool TryNormalizeProgramUnitHeaderForCompare(std::string& line)
{
	size_t directivePos = 0;
	if (!TryMatchDirectiveAtLineStart(line, kDirectiveAssembly, &directivePos)) {
		return false;
	}

	const std::string indent = line.substr(0, directivePos);
	const std::string rest = line.substr(directivePos + kDirectiveAssembly.size());

	std::vector<std::string> fields;
	size_t fieldStart = 0;
	while (true) {
		const size_t comma = rest.find(',', fieldStart);
		const std::string field = comma == std::string::npos
			? rest.substr(fieldStart)
			: rest.substr(fieldStart, comma - fieldStart);
		fields.push_back(TrimAsciiCopyLocal(field));
		if (comma == std::string::npos) {
			break;
		}
		fieldStart = comma + 1;
	}

	// fields[0]=程序集/类名；fields[1]=基类（若存在）。IDE 的真实页格式化
	// 只可靠保留名称和非默认基类，后续字段（公开、说明等）统一视为不可观测元数据。
	bool changed = false;
	if (fields.size() >= 2 &&
		(fields[1].empty() || fields[1] == kDefaultBaseClassToken)) {
		fields.resize(1);
		changed = true;
	}
	else if (fields.size() > 2) {
		fields.resize(2);
		changed = true;
	}
	// 移除基类后可能遗留的尾部空字段（主要覆盖没有进入上面分支的异常/简写输入）。
	while (fields.size() > 1 && fields.back().empty()) {
		fields.pop_back();
		changed = true;
	}
	if (!changed) {
		return false;
	}

	std::string rebuilt = indent + std::string(kDirectiveAssembly);
	if (!fields.empty()) {
		rebuilt += " " + fields[0];
		for (size_t i = 1; i < fields.size(); ++i) {
			rebuilt += ", " + fields[i];
		}
	}
	line = rebuilt;
	return true;
}

bool TryNormalizeSubroutineHeaderForCompare(std::string& line)
{
	std::vector<std::string> fields;
	size_t directivePos = 0;
	if (!TrySplitSubroutineHeaderFields(line, fields, &directivePos)) {
		return false;
	}

	const std::string indent = line.substr(0, directivePos);
	if (fields[0].empty()) {
		return false;
	}
	if (fields.size() >= 3 && !fields[2].empty() && fields[2] != "公开") {
		// 字段错位的源码不是 IDE 省略注释后的等价形式，不能静默归一化。
		return false;
	}

	// 真实页格式化可靠保留：子程序名、返回类型、公开标记。
	// 说明文字及其后续字段不会出现在 direct formatter 结果中。
	size_t keepCount = 1;
	if (fields.size() >= 2 && !fields[1].empty()) {
		keepCount = 2;
	}
	if (fields.size() >= 3 && fields[2] == "公开") {
		keepCount = 3;
	}
	if (fields.size() <= keepCount) {
		return false;
	}

	std::string rebuilt = indent + std::string(kDirectiveSubroutine) + " " + fields[0];
	for (size_t i = 1; i < keepCount; ++i) {
		rebuilt += ", " + fields[i];
	}
	line = std::move(rebuilt);
	return true;
}

std::string NormalizeRealPageAssemblyVariableAliasesForCompare(const std::string& text)
{
	std::vector<std::string> lines = SplitRealCodeLines(text);
	bool changed = false;
	bool reachedSubroutineBlock = false;
	for (auto& line : lines) {
		const std::string trimmed = TrimAsciiCopyLocal(line);
		if (TryNormalizeSubroutineHeaderForCompare(line)) {
			changed = true;
		}
		if (!reachedSubroutineBlock) {
			size_t directivePos = 0;
			if (TryMatchDirectiveAtLineStart(line, kDirectiveLocalVariable, &directivePos)) {
				line.replace(
					directivePos,
					kDirectiveLocalVariable.size(),
					kDirectiveAssemblyVariable.data(),
					kDirectiveAssemblyVariable.size());
				changed = true;
			}
			else if (TryNormalizeProgramUnitHeaderForCompare(line)) {
				changed = true;
			}
		}

		if (IsDirectiveLine(trimmed, kDirectiveSubroutine)) {
			reachedSubroutineBlock = true;
		}
	}

	return changed ? JoinRealCodeLines(lines) : text;
}

std::string NormalizeRealPageOperatorFormsForCompare(const std::string& text)
{
	static constexpr OperatorForm kOperatorForms[] = {
		{ "＋", '+' },
		{ "－", '-' },
		{ "＊", '*' },
		{ "×", '*' },
		{ "／", '/' },
		{ "÷", '/' },
		{ "＝", '=' },
	};

	std::string normalized;
	normalized.reserve(text.size());
	bool inString = false;

	for (size_t i = 0; i < text.size();) {
		const char ch = text[i];
		if (ch == '"') {
			normalized.push_back(ch);
			if (inString && i + 1 < text.size() && text[i + 1] == '"') {
				normalized.push_back(text[i + 1]);
				i += 2;
				continue;
			}
			inString = !inString;
			++i;
			continue;
		}

		if (!inString) {
			bool matched = false;
			for (const auto& form : kOperatorForms) {
				if (form.text.empty() || i + form.text.size() > text.size()) {
					continue;
				}
				if (text.compare(i, form.text.size(), form.text.data(), form.text.size()) == 0) {
					normalized.push_back(form.canonical);
					i += form.text.size();
					matched = true;
					break;
				}
			}
			if (matched) {
				continue;
			}
		}

		normalized.push_back(ch);
		++i;
	}

	return normalized;
}

// 消除 IDE 存盘对结构指纹的等价性改写，使写入与读回的指纹可比。
// 传入的每个元素是「已去空白」的单行 token（见 NormalizePageCodeLineForStructuralCompare）。
// 处理五类 IDE 自动改写（均为语义等价、非内容丢失）：
//  1) 有些读回路径会带 .版本 页头，有些不会：结构比较时忽略该页头 token。
//  2) e5.95 会把部分控制语句前导点省略：把 如果真(...) 等归一为 .如果真(...)。
//  3) 给无 else 的 .如果 块补空 .否则：去掉紧邻 .如果结束 之前的 .否则 token。
//     该规则对写入与读回对称——真实的空 .否则 会在两侧同样被去掉，不会凭空造出或消除差异。
//  4) 给无 default 的 .判断开始块补空 .默认：去掉紧邻 .判断结束 之前的 .默认 token。
//  5) 页尾追加裸 .子程序（无名字）+ 把程序集级注释搬到页尾孤儿化：去掉去空白后恰为 ".子程序" 的 token。
//     真实子程序声明必带名字，去空白后形如 ".子程序名字"，绝不会精确等于 ".子程序"。
//  6) 声明尾部的空字段会被省略：去掉指令 token 末尾的 ASCII 逗号。
std::vector<std::string> NormalizeStructuralFingerprintForIdeRewrite(
	const std::vector<std::string>& fingerprint)
{
	std::vector<std::string> normalized;
	normalized.reserve(fingerprint.size());
	for (size_t i = 0; i < fingerprint.size(); ++i) {
		std::string token = NormalizeIdeTrailingEmptyDirectiveFields(
			NormalizeIdeOptionalLeadingDotToken(fingerprint[i]));
		if (IsVersionFingerprintToken(token)) {
			continue;
		}
		if (token == kFingerprintTokenOtherwise &&
			i + 1 < fingerprint.size() &&
			NormalizeIdeTrailingEmptyDirectiveFields(
				NormalizeIdeOptionalLeadingDotToken(fingerprint[i + 1])) == kFingerprintTokenEndIf) {
			continue;
		}
		if (token == kFingerprintTokenDefault &&
			i + 1 < fingerprint.size() &&
			NormalizeIdeTrailingEmptyDirectiveFields(
				NormalizeIdeOptionalLeadingDotToken(fingerprint[i + 1])) == kFingerprintTokenEndSwitch) {
			continue;
		}
		if (token == kFingerprintTokenBareSubroutine) {
			continue;
		}
		normalized.push_back(std::move(token));
	}
	return normalized;
}

std::string BuildStableTextHashForRealCode(const std::string& text)
{
	const std::string normalized = NormalizeRealCodeLineBreaksToLf(text);
	std::string sha256;
	if (TryBuildSha256Hex(normalized, sha256)) {
		return sha256;
	}

	// BCrypt 极端失败时保留旧的稳定哈希作为降级路径，保证工具仍可工作。
	std::uint64_t hash = 1469598103934665603ull;
	for (unsigned char ch : normalized) {
		hash ^= static_cast<std::uint64_t>(ch);
		hash *= 1099511628211ull;
	}

	static constexpr char kHex[] = "0123456789abcdef";
	std::string out(16, '0');
	for (int i = 15; i >= 0; --i) {
		out[static_cast<size_t>(i)] = kHex[hash & 0x0F];
		hash >>= 4;
	}
	return out;
}

bool ApplyRealPageTextEdits(
	const std::string& sourceCode,
	const std::vector<RealPageTextEditRequest>& edits,
	bool failOnUnmatched,
	std::string& outCode,
	std::vector<RealPageTextEditApplyResult>& outResults,
	std::string& outError)
{
	outCode = sourceCode;
	outResults.clear();
	outError.clear();

	for (const auto& edit : edits) {
		RealPageTextEditApplyResult result;
		const std::string normalizedOldText = NormalizeRealCodeLineBreaksToCrLf(edit.oldText);
		const std::string normalizedNewText = NormalizeRealCodeLineBreaksToCrLf(edit.newText);
		if (normalizedOldText.empty()) {
			result.error = "old_text is required";
			outResults.push_back(result);
			if (failOnUnmatched) {
				outError = result.error;
				return false;
			}
			continue;
		}

		const std::vector<EquivalentTextMatchSpan> matches =
			FindEquivalentTextMatches(outCode, normalizedOldText);
		result.matchCount = matches.size();
		if (result.matchCount == 0) {
			result.error = "old_text not found";
			outResults.push_back(result);
			if (failOnUnmatched) {
				outError = result.error;
				return false;
			}
			continue;
		}

		if (!edit.replaceAll && result.matchCount != 1) {
			result.error = "old_text matched multiple times";
			outResults.push_back(result);
			if (failOnUnmatched) {
				outError = result.error;
				return false;
			}
			continue;
		}

		if (edit.replaceAll) {
			for (auto it = matches.rbegin(); it != matches.rend(); ++it) {
				outCode.replace(it->offset, it->length, normalizedNewText);
			}
		}
		else {
			outCode.replace(matches.front().offset, matches.front().length, normalizedNewText);
		}
		result.applied = true;
		outResults.push_back(result);
	}

	outCode = NormalizeRealCodeLineBreaksToCrLf(outCode);
	return true;
}

std::vector<RealPageStructuredPatchHunk> BuildRealPageStructuredPatch(
	const std::string& oldCode,
	const std::string& newCode)
{
	const std::vector<std::string> oldLines = SplitRealCodeLines(oldCode);
	const std::vector<std::string> newLines = SplitRealCodeLines(newCode);
	if (oldLines == newLines) {
		return {};
	}

	if (oldLines.size() * newLines.size() > 4000000ull) {
		return BuildFallbackStructuredPatch(oldLines, newLines);
	}

	const size_t n = oldLines.size();
	const size_t m = newLines.size();

	std::vector<int> dp((n + 1) * (m + 1), 0);
	const auto cell = [&](size_t i, size_t j) -> int& {
		return dp[i * (m + 1) + j];
	};

	for (size_t i = n; i-- > 0;) {
		for (size_t j = m; j-- > 0;) {
			if (oldLines[i] == newLines[j]) {
				cell(i, j) = cell(i + 1, j + 1) + 1;
			}
			else {
				cell(i, j) = (std::max)(cell(i + 1, j), cell(i, j + 1));
			}
		}
	}

	std::vector<DiffOp> ops;
	ops.reserve(n + m);
	size_t i = 0;
	size_t j = 0;
	int oldLineNumber = 1;
	int newLineNumber = 1;
	while (i < n && j < m) {
		if (oldLines[i] == newLines[j]) {
			ops.push_back({' ', oldLines[i], oldLineNumber, newLineNumber});
			++i;
			++j;
			++oldLineNumber;
			++newLineNumber;
			continue;
		}

		if (cell(i + 1, j) >= cell(i, j + 1)) {
			ops.push_back({'-', oldLines[i], oldLineNumber, newLineNumber});
			++i;
			++oldLineNumber;
		}
		else {
			ops.push_back({'+', newLines[j], oldLineNumber, newLineNumber});
			++j;
			++newLineNumber;
		}
	}
	while (i < n) {
		ops.push_back({'-', oldLines[i], oldLineNumber, newLineNumber});
		++i;
		++oldLineNumber;
	}
	while (j < m) {
		ops.push_back({'+', newLines[j], oldLineNumber, newLineNumber});
		++j;
		++newLineNumber;
	}

	std::vector<RealPageStructuredPatchHunk> hunks;
	size_t index = 0;
	while (index < ops.size()) {
		while (index < ops.size() && ops[index].type == ' ') {
			++index;
		}
		if (index >= ops.size()) {
			break;
		}

		RealPageStructuredPatchHunk hunk;
		hunk.oldStart = ops[index].beforeOldLine;
		hunk.newStart = ops[index].beforeNewLine;
		while (index < ops.size() && ops[index].type != ' ') {
			const DiffOp& op = ops[index];
			if (op.type == '-') {
				++hunk.oldLines;
			}
			else if (op.type == '+') {
				++hunk.newLines;
			}
			hunk.lines.push_back(std::string(1, op.type) + op.line);
			++index;
		}
		hunks.push_back(std::move(hunk));
	}
	return hunks;
}

int CountStructuredPatchChangedLines(const std::vector<RealPageStructuredPatchHunk>& hunks)
{
	int changed = 0;
	for (const auto& hunk : hunks) {
		changed += hunk.oldLines;
		changed += hunk.newLines;
	}
	return changed;
}
