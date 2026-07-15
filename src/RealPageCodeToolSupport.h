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
// 合并 IDE 新建类时生成的生命周期函数，并保证 _初始化、_销毁依次为前两个函数。
std::string PrepareNewClassPageLifecycleFunctions(
	const std::string& requestedCode,
	const std::string& ideDefaultCode,
	bool* outChanged = nullptr,
	bool* outComplete = nullptr);
std::string NormalizeRealPageAssemblyVariableAliasesForCompare(const std::string& text);
// 比较真实页源码时消除 IDE 对运算符符号的自动换形差异。
std::string NormalizeRealPageOperatorFormsForCompare(const std::string& text);
// 消除 IDE 存盘对结构指纹的等价改写（.版本页头 / 控制语句前导点 / 补空 .否则 / 页尾裸 .子程序），使写入与读回指纹可比。
// 入参为「已去空白」的逐行 token 序列。
std::vector<std::string> NormalizeStructuralFingerprintForIdeRewrite(
	const std::vector<std::string>& fingerprint);
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
