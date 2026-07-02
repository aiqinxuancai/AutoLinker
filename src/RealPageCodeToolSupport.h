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

std::string BuildStableTextHashForRealCode(const std::string& text);
std::string NormalizeRealCodeLineBreaksToCrLf(const std::string& text);
std::string NormalizeRealCodeLineBreaksToLf(const std::string& text);
std::string PrepareAssemblyVariablesForRealPageWrite(const std::string& text);
std::string NormalizeRealPageAssemblyVariableAliasesForCompare(const std::string& text);
std::vector<std::string> SplitRealCodeLines(const std::string& text);
std::string JoinRealCodeLines(const std::vector<std::string>& lines);

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
