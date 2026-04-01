#include "RealPageCodeToolSupport.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <regex>
#include <sstream>
#include <string_view>
#include <vector>

namespace {

constexpr std::string_view kDirectiveSubroutine = ".子程序";
constexpr std::string_view kDirectiveLocalVariable = ".局部变量";
constexpr std::string_view kDirectiveParameter = ".参数";
constexpr std::string_view kDirectiveAssemblyVariable = ".程序集变量";
constexpr std::string_view kDirectiveMemberVariable = ".成员变量";

struct DiffOp {
	char type = ' ';
	std::string line;
	int beforeOldLine = 0;
	int beforeNewLine = 0;
};

struct TopLevelRealPageSymbolSeed {
	int lineNumber = 0;
	std::string kind;
	std::string name;
	bool isBlock = false;
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

std::string ExtractDirectiveName(const std::string& rest)
{
	const std::string trimmed = TrimAsciiCopyLocal(rest);
	const size_t splitPos = trimmed.find_first_of(" \t,(");
	if (splitPos == std::string::npos) {
		return trimmed;
	}
	return TrimAsciiCopyLocal(trimmed.substr(0, splitPos));
}

size_t CountOccurrencesLocal(const std::string& text, const std::string& needle)
{
	if (needle.empty()) {
		return 0;
	}

	size_t count = 0;
	size_t pos = 0;
	while ((pos = text.find(needle, pos)) != std::string::npos) {
		++count;
		pos += needle.size();
	}
	return count;
}

std::string ReplaceAllLocal(std::string text, const std::string& oldText, const std::string& newText)
{
	if (oldText.empty()) {
		return text;
	}

	size_t pos = 0;
	while ((pos = text.find(oldText, pos)) != std::string::npos) {
		text.replace(pos, oldText.size(), newText);
		pos += newText.size();
	}
	return text;
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

bool TryParseTopLevelSymbolSeed(const std::string& trimmed, int lineNumber, TopLevelRealPageSymbolSeed& outSeed)
{
	outSeed = {};
	outSeed.lineNumber = lineNumber;

	if (IsDirectiveLine(trimmed, kDirectiveSubroutine)) {
		outSeed.kind = "subroutine";
		outSeed.name = ExtractDirectiveName(trimmed.substr(kDirectiveSubroutine.size()));
		outSeed.isBlock = true;
		return true;
	}
	if (IsDirectiveLine(trimmed, kDirectiveAssemblyVariable)) {
		outSeed.kind = "assembly_variable";
		outSeed.name = ExtractDirectiveName(trimmed.substr(kDirectiveAssemblyVariable.size()));
		outSeed.isBlock = false;
		return true;
	}
	if (IsDirectiveLine(trimmed, kDirectiveMemberVariable)) {
		outSeed.kind = "member_variable";
		outSeed.name = ExtractDirectiveName(trimmed.substr(kDirectiveMemberVariable.size()));
		outSeed.isBlock = false;
		return true;
	}
	return false;
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

std::string BuildStableTextHashForRealCode(const std::string& text)
{
	const std::string normalized = NormalizeRealCodeLineBreaksToLf(text);
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

std::string BuildRealCodeView(
	const std::string& code,
	int offsetLines,
	int limitLines,
	bool withLineNumbers,
	int* outTotalLines,
	int* outReturnedLines,
	int* outStartLine)
{
	if (outTotalLines != nullptr) {
		*outTotalLines = 0;
	}
	if (outReturnedLines != nullptr) {
		*outReturnedLines = 0;
	}
	if (outStartLine != nullptr) {
		*outStartLine = 1;
	}

	const std::vector<std::string> lines = SplitRealCodeLines(code);
	const int totalLines = static_cast<int>(lines.size());
	if (outTotalLines != nullptr) {
		*outTotalLines = totalLines;
	}
	if (totalLines <= 0) {
		return std::string();
	}

	const int startIndex = (std::max)(0, offsetLines);
	const int endIndex = limitLines > 0
		? (std::min)(totalLines, startIndex + limitLines)
		: totalLines;
	if (outReturnedLines != nullptr) {
		*outReturnedLines = (std::max)(0, endIndex - startIndex);
	}
	if (outStartLine != nullptr) {
		*outStartLine = startIndex + 1;
	}

	std::ostringstream out;
	for (int i = startIndex; i < endIndex; ++i) {
		if (i != startIndex) {
			out << "\r\n";
		}
		if (withLineNumbers) {
			out << (i + 1) << '\t';
		}
		out << lines[static_cast<size_t>(i)];
	}
	return out.str();
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
		if (edit.oldText.empty()) {
			result.error = "old_text is required";
			outResults.push_back(result);
			if (failOnUnmatched) {
				outError = result.error;
				return false;
			}
			continue;
		}

		result.matchCount = CountOccurrencesLocal(outCode, edit.oldText);
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
			outCode = ReplaceAllLocal(outCode, edit.oldText, edit.newText);
		}
		else {
			const size_t pos = outCode.find(edit.oldText);
			outCode.replace(pos, edit.oldText.size(), edit.newText);
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

std::vector<RealPageSymbolInfo> ParseRealPageSymbols(const std::string& pageCode)
{
	const std::vector<std::string> lines = SplitRealCodeLines(pageCode);
	std::vector<TopLevelRealPageSymbolSeed> topLevelSeeds;
	for (size_t i = 0; i < lines.size(); ++i) {
		TopLevelRealPageSymbolSeed seed;
		if (TryParseTopLevelSymbolSeed(TrimAsciiCopyLocal(lines[i]), static_cast<int>(i) + 1, seed)) {
			topLevelSeeds.push_back(std::move(seed));
		}
	}

	std::vector<RealPageSymbolInfo> symbols;
	const int totalLines = static_cast<int>(lines.size());
	const int headerEndLine = topLevelSeeds.empty() ? totalLines : (topLevelSeeds.front().lineNumber - 1);
	if (headerEndLine > 0) {
		RealPageSymbolInfo header;
		header.name = "__page_header__";
		header.kind = "page_header";
		header.displayName = "页头";
		header.startLine = 1;
		header.endLine = headerEndLine;
		header.declarationLine = 1;
		header.code = JoinRealCodeLines(std::vector<std::string>(
			lines.begin(),
			lines.begin() + headerEndLine));
		symbols.push_back(std::move(header));
	}

	for (size_t index = 0; index < topLevelSeeds.size(); ++index) {
		const TopLevelRealPageSymbolSeed& seed = topLevelSeeds[index];
		const int startLine = seed.lineNumber;
		const int endLine = seed.isBlock
			? ((index + 1 < topLevelSeeds.size()) ? (topLevelSeeds[index + 1].lineNumber - 1) : totalLines)
			: startLine;

		RealPageSymbolInfo symbol;
		symbol.name = seed.name;
		symbol.kind = seed.kind;
		symbol.displayName = seed.name;
		symbol.declarationLine = startLine;
		symbol.startLine = startLine;
		symbol.endLine = endLine;
		symbol.isEventHandler = seed.kind == "subroutine" && StartsWithLocal(seed.name, "__");
		symbol.code = JoinRealCodeLines(std::vector<std::string>(
			lines.begin() + (startLine - 1),
			lines.begin() + endLine));
		symbols.push_back(symbol);

		if (!seed.isBlock) {
			continue;
		}

		for (int lineIndex = startLine; lineIndex <= endLine; ++lineIndex) {
			const std::string trimmed = TrimAsciiCopyLocal(lines[static_cast<size_t>(lineIndex - 1)]);
			if (IsDirectiveLine(trimmed, kDirectiveParameter)) {
				RealPageSymbolInfo parameter;
				parameter.name = ExtractDirectiveName(trimmed.substr(kDirectiveParameter.size()));
				parameter.kind = "parameter";
				parameter.displayName = parameter.name;
				parameter.parentName = seed.name;
				parameter.declarationLine = lineIndex;
				parameter.startLine = lineIndex;
				parameter.endLine = lineIndex;
				parameter.code = lines[static_cast<size_t>(lineIndex - 1)];
				symbols.push_back(std::move(parameter));
				continue;
			}
			if (IsDirectiveLine(trimmed, kDirectiveLocalVariable)) {
				RealPageSymbolInfo localVariable;
				localVariable.name = ExtractDirectiveName(trimmed.substr(kDirectiveLocalVariable.size()));
				localVariable.kind = "local_variable";
				localVariable.displayName = localVariable.name;
				localVariable.parentName = seed.name;
				localVariable.declarationLine = lineIndex;
				localVariable.startLine = lineIndex;
				localVariable.endLine = lineIndex;
				localVariable.code = lines[static_cast<size_t>(lineIndex - 1)];
				symbols.push_back(std::move(localVariable));
			}
		}
	}

	return symbols;
}

bool FindRealPageSymbol(
	const std::vector<RealPageSymbolInfo>& symbols,
	const std::string& symbolName,
	const std::string& symbolKind,
	const std::string& parentSymbolName,
	int occurrence,
	RealPageSymbolInfo& outSymbol,
	std::string& outError)
{
	outSymbol = {};
	outError.clear();

	const std::string expectedKind = ToLowerAsciiCopyLocal(TrimAsciiCopyLocal(symbolKind));
	const int expectedOccurrence = occurrence <= 0 ? 1 : occurrence;
	int currentOccurrence = 0;

	for (const auto& symbol : symbols) {
		if (!expectedKind.empty() && symbol.kind != expectedKind) {
			continue;
		}
		if (!symbolName.empty() && symbol.name != symbolName) {
			continue;
		}
		if (!parentSymbolName.empty() && symbol.parentName != parentSymbolName) {
			continue;
		}

		++currentOccurrence;
		if (currentOccurrence == expectedOccurrence) {
			outSymbol = symbol;
			return true;
		}
	}

	outError = "symbol not found";
	return false;
}

bool ReplaceRealPageSymbolCode(
	const std::string& sourceCode,
	const RealPageSymbolInfo& symbol,
	const std::string& newSymbolCode,
	std::string& outCode,
	std::string& outError)
{
	outCode.clear();
	outError.clear();
	if (symbol.startLine <= 0 || symbol.endLine < symbol.startLine) {
		outError = "symbol range invalid";
		return false;
	}

	std::vector<std::string> lines = SplitRealCodeLines(sourceCode);
	std::vector<std::string> replacementLines = SplitRealCodeLines(newSymbolCode);
	if (symbol.kind == "subroutine" && !replacementLines.empty()) {
		const bool hasBodyLine = replacementLines.size() >= 2 &&
			!TrimAsciiCopyLocal(replacementLines[1]).empty();
		if (hasBodyLine) {
			replacementLines.insert(replacementLines.begin() + 1, std::string());
		}
		if (!replacementLines.empty() &&
			!TrimAsciiCopyLocal(replacementLines.back()).empty()) {
			replacementLines.push_back(std::string());
		}
	}
	if (symbol.endLine > static_cast<int>(lines.size())) {
		outError = "symbol range out of bounds";
		return false;
	}

	lines.erase(
		lines.begin() + (symbol.startLine - 1),
		lines.begin() + symbol.endLine);
	lines.insert(
		lines.begin() + (symbol.startLine - 1),
		replacementLines.begin(),
		replacementLines.end());
	outCode = JoinRealCodeLines(lines);
	return true;
}

bool InsertRealPageCodeBlock(
	const std::string& sourceCode,
	const std::vector<RealPageSymbolInfo>& symbols,
	const RealPageInsertSpec& spec,
	const std::string& codeBlock,
	std::string& outCode,
	std::string& outError)
{
	outCode.clear();
	outError.clear();
	if (TrimAsciiCopyLocal(codeBlock).empty()) {
		outError = "code_block is required";
		return false;
	}

	const std::string mode = ToLowerAsciiCopyLocal(TrimAsciiCopyLocal(spec.mode));
	if (mode == "top" || mode == "bottom" || mode == "before_symbol" || mode == "after_symbol") {
		std::vector<std::string> lines = SplitRealCodeLines(sourceCode);
		std::vector<std::string> blockLines = SplitRealCodeLines(codeBlock);
		if (!blockLines.empty() && IsDirectiveLine(TrimAsciiCopyLocal(blockLines.front()), kDirectiveSubroutine)) {
			const bool hasBodyLine = blockLines.size() >= 2 &&
				!TrimAsciiCopyLocal(blockLines[1]).empty();
			if (hasBodyLine) {
				blockLines.insert(blockLines.begin() + 1, std::string());
			}
			if (!blockLines.empty() &&
				!TrimAsciiCopyLocal(blockLines.back()).empty()) {
				blockLines.push_back(std::string());
			}
		}

		size_t insertIndex = 0;
		if (mode == "top") {
			insertIndex = 0;
		}
		else if (mode == "bottom") {
			insertIndex = lines.size();
		}
		else {
			RealPageSymbolInfo symbol;
			if (!FindRealPageSymbol(
					symbols,
					spec.symbolName,
					spec.symbolKind,
					spec.parentSymbolName,
					spec.occurrence,
					symbol,
					outError)) {
				return false;
			}
			insertIndex = mode == "before_symbol"
				? static_cast<size_t>(symbol.startLine - 1)
				: static_cast<size_t>(symbol.endLine);
		}

		lines.insert(lines.begin() + static_cast<std::ptrdiff_t>(insertIndex), blockLines.begin(), blockLines.end());
		outCode = JoinRealCodeLines(lines);
		return true;
	}

	if (mode == "before_text" || mode == "after_text") {
		if (spec.anchorText.empty()) {
			outError = "anchor_text is required";
			return false;
		}

		size_t found = std::string::npos;
		size_t scan = 0;
		const int expectedOccurrence = spec.occurrence <= 0 ? 1 : spec.occurrence;
		for (int current = 0; current < expectedOccurrence; ++current) {
			found = sourceCode.find(spec.anchorText, scan);
			if (found == std::string::npos) {
				outError = "anchor_text not found";
				return false;
			}
			scan = found + spec.anchorText.size();
		}

		outCode = sourceCode;
		const size_t insertPos = mode == "before_text" ? found : (found + spec.anchorText.size());
		outCode.insert(insertPos, codeBlock);
		outCode = NormalizeRealCodeLineBreaksToCrLf(outCode);
		return true;
	}

	outError = "insert mode unsupported";
	return false;
}

std::vector<RealPageSearchMatch> SearchRealPageCode(
	const std::string& sourceCode,
	const std::string& keyword,
	bool caseSensitive,
	bool useRegex,
	int contextLines,
	size_t limit,
	std::string* outError)
{
	if (outError != nullptr) {
		*outError = "";
	}

	std::vector<RealPageSearchMatch> matches;
	if (TrimAsciiCopyLocal(keyword).empty()) {
		if (outError != nullptr) {
			*outError = "keyword is required";
		}
		return matches;
	}

	const std::vector<std::string> lines = SplitRealCodeLines(sourceCode);
	std::regex regex;
	if (useRegex) {
		try {
			regex = std::regex(keyword, caseSensitive ? std::regex::ECMAScript : (std::regex::ECMAScript | std::regex::icase));
		}
		catch (const std::exception& ex) {
			if (outError != nullptr) {
				*outError = std::string("regex invalid: ") + ex.what();
			}
			return matches;
		}
	}

	const std::string loweredKeyword = caseSensitive ? keyword : ToLowerAsciiCopyLocal(keyword);
	for (size_t i = 0; i < lines.size(); ++i) {
		bool matched = false;
		size_t matchColumn = 0;
		if (useRegex) {
			std::smatch sm;
			if (std::regex_search(lines[i], sm, regex)) {
				matched = true;
				matchColumn = sm.position() + 1;
			}
		}
		else {
			const std::string haystack = caseSensitive ? lines[i] : ToLowerAsciiCopyLocal(lines[i]);
			const size_t pos = haystack.find(loweredKeyword);
			if (pos != std::string::npos) {
				matched = true;
				matchColumn = pos + 1;
			}
		}

		if (!matched) {
			continue;
		}

		RealPageSearchMatch match;
		match.lineNumber = static_cast<int>(i) + 1;
		match.matchColumn = matchColumn;
		match.lineText = lines[i];

		const int context = (std::max)(0, contextLines);
		for (int before = (std::max)(0, static_cast<int>(i) - context); before < static_cast<int>(i); ++before) {
			match.beforeContextLines.push_back(lines[static_cast<size_t>(before)]);
		}
		for (int after = static_cast<int>(i) + 1; after <= (std::min)(static_cast<int>(lines.size()) - 1, static_cast<int>(i) + context); ++after) {
			match.afterContextLines.push_back(lines[static_cast<size_t>(after)]);
		}

		matches.push_back(std::move(match));
		if (limit > 0 && matches.size() >= limit) {
			break;
		}
	}

	return matches;
}
