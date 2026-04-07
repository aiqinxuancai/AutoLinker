#pragma once

#include <cstddef>
#include <string>
#include <vector>

// 真实页文本替换请求。
struct RealPageTextEditRequest {
	std::string oldText;
	std::string newText;
	bool replaceAll = false;
};

// 单条文本替换结果。
struct RealPageTextEditApplyResult {
	size_t matchCount = 0;
	bool applied = false;
	std::string error;
};

// 真实页结构化差异块。
struct RealPageStructuredPatchHunk {
	int oldStart = 0;
	int oldLines = 0;
	int newStart = 0;
	int newLines = 0;
	std::vector<std::string> lines;
};

// 真实页符号信息。
struct RealPageSymbolInfo {
	std::string name;
	std::string kind;
	std::string displayName;
	std::string parentName;
	int declarationLine = 0;
	int startLine = 0;
	int endLine = 0;
	bool isEventHandler = false;
	std::string code;
};

// 真实页搜索命中。
struct RealPageSearchMatch {
	int lineNumber = 0;
	size_t matchColumn = 0;
	std::string lineText;
	std::vector<std::string> beforeContextLines;
	std::vector<std::string> afterContextLines;
};

// 真实页代码块插入描述。
struct RealPageInsertSpec {
	std::string mode;
	std::string symbolName;
	std::string symbolKind;
	std::string parentSymbolName;
	std::string anchorText;
	int occurrence = 1;
};

std::string BuildStableTextHashForRealCode(const std::string& text);
std::string NormalizeRealCodeLineBreaksToCrLf(const std::string& text);
std::string NormalizeRealCodeLineBreaksToLf(const std::string& text);
std::string PrepareAssemblyVariablesForRealPageWrite(const std::string& text);
std::string NormalizeRealPageAssemblyVariableAliasesForCompare(const std::string& text);
std::vector<std::string> SplitRealCodeLines(const std::string& text);
std::string JoinRealCodeLines(const std::vector<std::string>& lines);
std::string BuildRealCodeView(
	const std::string& code,
	int offsetLines,
	int limitLines,
	bool withLineNumbers,
	int* outTotalLines = nullptr,
	int* outReturnedLines = nullptr,
	int* outStartLine = nullptr);

bool ApplyRealPageTextEdits(
	const std::string& sourceCode,
	const std::vector<RealPageTextEditRequest>& edits,
	bool failOnUnmatched,
	std::string& outCode,
	std::vector<RealPageTextEditApplyResult>& outResults,
	std::string& outError);

std::vector<RealPageStructuredPatchHunk> BuildRealPageStructuredPatch(
	const std::string& oldCode,
	const std::string& newCode);
int CountStructuredPatchChangedLines(const std::vector<RealPageStructuredPatchHunk>& hunks);

std::vector<RealPageSymbolInfo> ParseRealPageSymbols(const std::string& pageCode);
bool FindRealPageSymbol(
	const std::vector<RealPageSymbolInfo>& symbols,
	const std::string& symbolName,
	const std::string& symbolKind,
	const std::string& parentSymbolName,
	int occurrence,
	RealPageSymbolInfo& outSymbol,
	std::string& outError);

bool ReplaceRealPageSymbolCode(
	const std::string& sourceCode,
	const RealPageSymbolInfo& symbol,
	const std::string& newSymbolCode,
	std::string& outCode,
	std::string& outError);

bool InsertRealPageCodeBlock(
	const std::string& sourceCode,
	const std::vector<RealPageSymbolInfo>& symbols,
	const RealPageInsertSpec& spec,
	const std::string& codeBlock,
	std::string& outCode,
	std::string& outError);

std::vector<RealPageSearchMatch> SearchRealPageCode(
	const std::string& sourceCode,
	const std::string& keyword,
	bool caseSensitive,
	bool useRegex,
	int contextLines,
	size_t limit,
	std::string* outError = nullptr);
