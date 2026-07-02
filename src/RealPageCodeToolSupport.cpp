#include "RealPageCodeToolSupport.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <string_view>
#include <vector>

namespace {

constexpr std::string_view kDirectiveSubroutine = ".子程序";
constexpr std::string_view kDirectiveLocalVariable = ".局部变量";
constexpr std::string_view kDirectiveAssemblyVariable = ".程序集变量";
// 类声明指令（注意必须区别于 .程序集变量：后者是变量声明）。
constexpr std::string_view kDirectiveAssembly = ".程序集";
// 易语言默认基类标注。声明 `.程序集 X, <对象>` 与 `.程序集 X` 完全等价：
// IDE 会把显式的默认基类省略掉再存，读回时即变成无基类形式。比较时需归一化两者。
constexpr std::string_view kDefaultBaseClassToken = "<对象>";

struct DiffOp {
	char type = ' ';
	std::string line;
	int beforeOldLine = 0;
	int beforeNewLine = 0;
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

// 把类声明里等价的默认基类标注归一化掉：".程序集 X, <对象>" → ".程序集 X"。
// IDE 存盘时会省略显式写出的默认基类 <对象>，读回后两种写法语义完全相同，
// 因此比较前需消除该差异，避免把成功写入误判为 verify_mismatch 而触发回滚重写。
// 仅处理 .程序集（类声明）行，不会匹配 .程序集变量（分隔符校验已排除）。
// 返回是否修改了 line。
bool TryNormalizeAssemblyClassDefaultBase(std::string& line)
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

	// fields[0]=类名；fields[1]=基类（若存在）。只在基类恰为默认 <对象> 时移除。
	bool changed = false;
	if (fields.size() >= 2 && fields[1] == kDefaultBaseClassToken) {
		fields.erase(fields.begin() + 1);
		changed = true;
	}
	// 移除基类后可能遗留的尾部空字段。
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

std::string NormalizeRealPageAssemblyVariableAliasesForCompare(const std::string& text)
{
	std::vector<std::string> lines = SplitRealCodeLines(text);
	bool changed = false;
	bool reachedSubroutineBlock = false;
	for (auto& line : lines) {
		const std::string trimmed = TrimAsciiCopyLocal(line);
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
			else if (TryNormalizeAssemblyClassDefaultBase(line)) {
				changed = true;
			}
		}

		if (IsDirectiveLine(trimmed, kDirectiveSubroutine)) {
			reachedSubroutineBlock = true;
		}
	}

	return changed ? JoinRealCodeLines(lines) : text;
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

		result.matchCount = CountOccurrencesLocal(outCode, normalizedOldText);
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
			outCode = ReplaceAllLocal(outCode, normalizedOldText, normalizedNewText);
		}
		else {
			const size_t pos = outCode.find(normalizedOldText);
			outCode.replace(pos, normalizedOldText.size(), normalizedNewText);
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
