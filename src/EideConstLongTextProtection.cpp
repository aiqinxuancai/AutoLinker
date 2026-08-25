#include "EideConstLongTextProtection.h"

#include <algorithm>
#include <limits>
#include <string_view>
#include <unordered_map>
#include <utility>

namespace eide_const {
namespace {

constexpr std::string_view kConstDirectiveUtf8 = ".\xE5\xB8\xB8\xE9\x87\x8F";
constexpr std::string_view kConstDirectiveGbk = ".\xB3\xA3\xC1\xBF";
constexpr std::string_view kLongTextPrefixUtf8 =
	"<\xE6\x96\x87\xE6\x9C\xAC\xE9\x95\xBF\xE5\xBA\xA6: ";
constexpr std::string_view kLongTextPrefixGbk = "<\xCE\xC4\xB1\xBE\xB3\xA4\xB6\xC8: ";
constexpr std::string_view kLongTextTokenUtf8 =
	"<\xE6\x96\x87\xE6\x9C\xAC\xE9\x95\xBF\xE5\xBA\xA6:";
constexpr std::string_view kLongTextTokenGbk = "<\xCE\xC4\xB1\xBE\xB3\xA4\xB6\xC8:";

struct FieldRange {
	size_t begin = 0;
	size_t end = 0;
};

bool IsAsciiSpace(char ch)
{
	return ch == ' ' || ch == '\t' || ch == '\v' || ch == '\f';
}

FieldRange TrimRange(std::string_view text, size_t begin, size_t end)
{
	while (begin < end && IsAsciiSpace(text[begin])) {
		++begin;
	}
	while (end > begin && IsAsciiSpace(text[end - 1])) {
		--end;
	}
	return FieldRange{begin, end};
}

bool TryMatchConstDirective(std::string_view line, size_t& outAfterDirective)
{
	outAfterDirective = 0;
	size_t begin = 0;
	while (begin < line.size() && IsAsciiSpace(line[begin])) {
		++begin;
	}

	for (const std::string_view directive : {kConstDirectiveUtf8, kConstDirectiveGbk}) {
		if (line.substr(begin).starts_with(directive)) {
			const size_t after = begin + directive.size();
			if (after == line.size() || IsAsciiSpace(line[after])) {
				outAfterDirective = after;
				return true;
			}
		}
	}
	return false;
}

bool TryParseLongTextPlaceholder(std::string_view value, size_t& outLength)
{
	outLength = 0;
	if (value.size() < 4 || value.front() != '"' || value.back() != '"') {
		return false;
	}

	const std::string_view inner = value.substr(1, value.size() - 2);
	std::string_view digits;
	for (const std::string_view prefix : {kLongTextPrefixUtf8, kLongTextPrefixGbk}) {
		if (inner.starts_with(prefix) && inner.size() > prefix.size() && inner.back() == '>') {
			digits = inner.substr(prefix.size(), inner.size() - prefix.size() - 1);
			break;
		}
	}
	if (digits.empty()) {
		return false;
	}

	size_t valueLength = 0;
	for (const char ch : digits) {
		if (ch < '0' || ch > '9') {
			return false;
		}
		const size_t digit = static_cast<size_t>(ch - '0');
		if (valueLength > ((std::numeric_limits<size_t>::max)() - digit) / 10) {
			return false;
		}
		valueLength = valueLength * 10 + digit;
	}
	outLength = valueLength;
	return true;
}

bool ContainsLongTextToken(std::string_view text)
{
	return text.find(kLongTextTokenUtf8) != std::string_view::npos ||
		text.find(kLongTextTokenGbk) != std::string_view::npos;
}

bool TryParseConstDeclaration(
	std::string_view line,
	size_t lineIndex,
	size_t lineBegin,
	ConstDeclaration& outDeclaration,
	bool& outIsConstDirective,
	std::string& outError)
{
	outDeclaration = {};
	outIsConstDirective = false;
	outError.clear();

	size_t afterDirective = 0;
	if (!TryMatchConstDirective(line, afterDirective)) {
		return true;
	}
	outIsConstDirective = true;

	size_t declarationBegin = afterDirective;
	while (declarationBegin < line.size() && IsAsciiSpace(line[declarationBegin])) {
		++declarationBegin;
	}
	if (declarationBegin == line.size()) {
		outError = "const declaration is empty at line " + std::to_string(lineIndex + 1);
		return false;
	}

	std::vector<FieldRange> ranges;
	bool inQuotes = false;
	size_t fieldBegin = declarationBegin;
	for (size_t i = declarationBegin; i < line.size(); ++i) {
		if (line[i] == '"') {
			if (inQuotes && i + 1 < line.size() && line[i + 1] == '"') {
				++i;
				continue;
			}
			inQuotes = !inQuotes;
			continue;
		}
		if (line[i] == ',' && !inQuotes) {
			ranges.push_back(TrimRange(line, fieldBegin, i));
			fieldBegin = i + 1;
		}
	}
	if (inQuotes) {
		outError = "unterminated quoted value at line " + std::to_string(lineIndex + 1);
		return false;
	}
	ranges.push_back(TrimRange(line, fieldBegin, line.size()));
	if (ranges.size() < 2 || ranges[0].begin == ranges[0].end) {
		outError = "const declaration requires name and value at line " + std::to_string(lineIndex + 1);
		return false;
	}

	outDeclaration.lineIndex = lineIndex;
	outDeclaration.rawLine.assign(line);
	outDeclaration.valueBegin = lineBegin + ranges[1].begin;
	outDeclaration.valueEnd = lineBegin + ranges[1].end;
	outDeclaration.fields.reserve(ranges.size());
	for (const FieldRange& range : ranges) {
		outDeclaration.fields.emplace_back(line.substr(range.begin, range.end - range.begin));
	}
	outDeclaration.name = outDeclaration.fields[0];
	outDeclaration.value = outDeclaration.fields[1];
	outDeclaration.isLongTextPlaceholder =
		TryParseLongTextPlaceholder(outDeclaration.value, outDeclaration.longTextLength);
	return true;
}

bool MetadataMatches(const ConstDeclaration& base, const ConstDeclaration& target)
{
	if (base.fields.size() != target.fields.size()) {
		return false;
	}
	for (size_t i = 0; i < base.fields.size(); ++i) {
		if (i != 1 && base.fields[i] != target.fields[i]) {
			return false;
		}
	}
	return true;
}

std::string MakeUniqueMarkerValue(const std::string& targetCode, size_t index)
{
	std::string markerBody =
		"__AUTOLINKER_PRESERVE_LONG_TEXT_" + std::to_string(index) + "__";
	while (targetCode.find(markerBody) != std::string::npos) {
		markerBody.push_back('_');
	}
	return "\"" + markerBody + "\"";
}

}  // namespace

bool TryParseConstPage(
	const std::string& pageCode,
	ConstPage& outPage,
	std::string& outError)
{
	outPage = {};
	outError.clear();

	size_t lineIndex = 0;
	size_t lineBegin = 0;
	while (lineBegin <= pageCode.size()) {
		size_t lineEnd = pageCode.find_first_of("\r\n", lineBegin);
		if (lineEnd == std::string::npos) {
			lineEnd = pageCode.size();
		}

		ConstDeclaration declaration;
		bool isConstDirective = false;
		if (!TryParseConstDeclaration(
				std::string_view(pageCode).substr(lineBegin, lineEnd - lineBegin),
				lineIndex,
				lineBegin,
				declaration,
				isConstDirective,
				outError)) {
			return false;
		}
		if (isConstDirective) {
			outPage.declarations.push_back(std::move(declaration));
		}

		if (lineEnd == pageCode.size()) {
			break;
		}
		lineBegin = lineEnd + 1;
		if (pageCode[lineEnd] == '\r' && lineBegin < pageCode.size() && pageCode[lineBegin] == '\n') {
			++lineBegin;
		}
		++lineIndex;
	}
	return true;
}

bool TryBuildLongTextPreservationPlan(
	const std::string& baseCode,
	const std::string& targetCode,
	LongTextPreservationPlan& outPlan,
	std::string& outError)
{
	outPlan = {};
	outPlan.stagedCode = targetCode;
	outError.clear();

	ConstPage basePage;
	ConstPage targetPage;
	std::string baseError;
	std::string targetError;
	const bool baseParsed = TryParseConstPage(baseCode, basePage, baseError);
	const bool targetParsed = TryParseConstPage(targetCode, targetPage, targetError);
	if (!baseParsed || !targetParsed) {
		outPlan.protectionRequired = ContainsLongTextToken(baseCode) || ContainsLongTextToken(targetCode);
		outError = !baseParsed ? baseError : targetError;
		return false;
	}
	outPlan.baseRowCount = basePage.declarations.size();
	outPlan.targetRowCount = targetPage.declarations.size();

	std::unordered_map<std::string, std::vector<const ConstDeclaration*>> baseByName;
	std::unordered_map<std::string, std::vector<const ConstDeclaration*>> targetByName;
	std::vector<const ConstDeclaration*> baseLongText;
	std::vector<const ConstDeclaration*> targetLongText;
	for (const ConstDeclaration& declaration : basePage.declarations) {
		baseByName[declaration.name].push_back(&declaration);
		if (declaration.isLongTextPlaceholder) {
			baseLongText.push_back(&declaration);
		}
	}
	for (const ConstDeclaration& declaration : targetPage.declarations) {
		targetByName[declaration.name].push_back(&declaration);
		if (declaration.isLongTextPlaceholder) {
			targetLongText.push_back(&declaration);
		}
	}

	outPlan.protectionRequired = !baseLongText.empty() || !targetLongText.empty();
	if (!outPlan.protectionRequired) {
		return true;
	}

	for (const ConstDeclaration* target : targetLongText) {
		const auto baseIt = baseByName.find(target->name);
		if (baseIt == baseByName.end()) {
			outError = "new long-text constant is not supported by text write: " + target->name;
			return false;
		}
		if (baseIt->second.size() != 1 || !baseIt->second.front()->isLongTextPlaceholder) {
			outError = "ambiguous long-text constant in base page: " + target->name;
			return false;
		}
	}

	struct Replacement {
		size_t begin = 0;
		size_t end = 0;
		std::string value;
	};
	std::vector<Replacement> replacements;
	for (size_t index = 0; index < baseLongText.size(); ++index) {
		const ConstDeclaration& base = *baseLongText[index];
		const auto baseIt = baseByName.find(base.name);
		const auto targetIt = targetByName.find(base.name);
		if (baseIt == baseByName.end() || baseIt->second.size() != 1) {
			outError = "duplicate long-text constant in base page: " + base.name;
			return false;
		}
		if (targetIt == targetByName.end()) {
			outError = "existing long-text constant cannot be removed by text write: " + base.name;
			return false;
		}
		if (targetIt->second.size() != 1) {
			outError = "duplicate long-text constant in target page: " + base.name;
			return false;
		}

		const ConstDeclaration& target = *targetIt->second.front();
		if (!target.isLongTextPlaceholder) {
			outError = "long-text value cannot be changed by text write: " + base.name;
			return false;
		}
		if (base.longTextLength != target.longTextLength) {
			outError = "long-text placeholder length changed: " + base.name;
			return false;
		}
		if (!MetadataMatches(base, target)) {
			outError = "long-text metadata cannot be changed by text write: " + base.name;
			return false;
		}

		LongTextPreservationEntry entry;
		entry.name = base.name;
		entry.longTextLength = base.longTextLength;
		entry.placeholderValue = base.value;
		entry.baseLineIndex = base.lineIndex;
		entry.targetLineIndex = target.lineIndex;
		entry.baseRowIndex = static_cast<size_t>(&base - basePage.declarations.data());
		entry.targetRowIndex = static_cast<size_t>(&target - targetPage.declarations.data());
		entry.markerValue = MakeUniqueMarkerValue(targetCode, index);
		replacements.push_back(Replacement{target.valueBegin, target.valueEnd, entry.markerValue});
		outPlan.entries.push_back(std::move(entry));
	}

	std::sort(
		replacements.begin(),
		replacements.end(),
		[](const Replacement& left, const Replacement& right) { return left.begin > right.begin; });
	for (const Replacement& replacement : replacements) {
		outPlan.stagedCode.replace(
			replacement.begin,
			replacement.end - replacement.begin,
			replacement.value);
	}
	return true;
}

}  // namespace eide_const
