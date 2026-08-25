#pragma once

#include <cstddef>
#include <string>
#include <vector>

namespace eide_const {

// 易语言常量声明的结构化信息。
struct ConstDeclaration {
	size_t lineIndex = 0;
	size_t valueBegin = 0;
	size_t valueEnd = 0;
	std::string rawLine;
	std::string name;
	std::string value;
	std::vector<std::string> fields;
	bool isLongTextPlaceholder = false;
	size_t longTextLength = 0;
};

// 常量页解析结果。
struct ConstPage {
	std::vector<ConstDeclaration> declarations;
};

// 单个长文本常量在保护写入中的替换信息。
struct LongTextPreservationEntry {
	std::string name;
	size_t longTextLength = 0;
	std::string markerValue;
	std::string placeholderValue;
	size_t baseLineIndex = 0;
	size_t targetLineIndex = 0;
	size_t baseRowIndex = 0;
	size_t targetRowIndex = 0;
};

// 长文本常量保护写入计划。
struct LongTextPreservationPlan {
	bool protectionRequired = false;
	std::string stagedCode;
	size_t baseRowCount = 0;
	size_t targetRowCount = 0;
	std::vector<LongTextPreservationEntry> entries;
};

// 解析完整常量页；非常量行会被忽略，格式损坏的常量声明会返回失败。
bool TryParseConstPage(
	const std::string& pageCode,
	ConstPage& outPage,
	std::string& outError);

// 校验长文本集合和元数据，并生成以普通文本标记暂代长文本值的写入代码。
bool TryBuildLongTextPreservationPlan(
	const std::string& baseCode,
	const std::string& targetCode,
	LongTextPreservationPlan& outPlan,
	std::string& outError);

}  // namespace eide_const
