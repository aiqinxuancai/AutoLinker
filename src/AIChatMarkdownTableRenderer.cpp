#include "AIChatMarkdownTableRenderer.h"

#include <cctype>

namespace AIChatMarkdownTableRenderer {
namespace {

enum class Alignment {
	Default,
	Left,
	Center,
	Right,
};

struct TableRow {
	std::vector<std::string> cells;
	bool hasSeparator = false;
};

std::string TrimAsciiCopy(const std::string& text)
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

TableRow SplitTableRow(const std::string& line)
{
	const std::string trimmed = TrimAsciiCopy(line);
	TableRow row;
	std::string cell;
	size_t codeDelimiterLength = 0;
	bool endedWithSeparator = false;

	for (size_t i = 0; i < trimmed.size();) {
		if (trimmed[i] == '`') {
			size_t runLength = 1;
			while (i + runLength < trimmed.size() && trimmed[i + runLength] == '`') {
				++runLength;
			}
			if (codeDelimiterLength == 0) {
				codeDelimiterLength = runLength;
			}
			else if (codeDelimiterLength == runLength) {
				codeDelimiterLength = 0;
			}
			cell.append(runLength, '`');
			i += runLength;
			endedWithSeparator = false;
			continue;
		}

		if (trimmed[i] == '|' && codeDelimiterLength == 0) {
			size_t precedingBackslashes = 0;
			while (precedingBackslashes < cell.size() &&
				cell[cell.size() - precedingBackslashes - 1] == '\\') {
				++precedingBackslashes;
			}
			if ((precedingBackslashes % 2) != 0) {
				cell.pop_back();
				cell.push_back('|');
				++i;
				endedWithSeparator = false;
				continue;
			}

			row.cells.push_back(TrimAsciiCopy(cell));
			cell.clear();
			row.hasSeparator = true;
			endedWithSeparator = true;
			++i;
			continue;
		}

		cell.push_back(trimmed[i]);
		endedWithSeparator = false;
		++i;
	}

	row.cells.push_back(TrimAsciiCopy(cell));
	if (!row.cells.empty() && !trimmed.empty() && trimmed.front() == '|') {
		row.cells.erase(row.cells.begin());
	}
	if (!row.cells.empty() && endedWithSeparator) {
		row.cells.pop_back();
	}
	return row;
}

bool TryParseDelimiter(const TableRow& row, std::vector<Alignment>& alignments)
{
	if (!row.hasSeparator || row.cells.empty()) {
		return false;
	}

	alignments.clear();
	alignments.reserve(row.cells.size());
	for (const std::string& rawCell : row.cells) {
		const std::string cell = TrimAsciiCopy(rawCell);
		const bool alignLeft = !cell.empty() && cell.front() == ':';
		const bool alignRight = !cell.empty() && cell.back() == ':';
		const size_t dashBegin = alignLeft ? 1 : 0;
		const size_t dashEnd = cell.size() - (alignRight ? 1 : 0);
		if (dashEnd < dashBegin + 3) {
			return false;
		}
		for (size_t i = dashBegin; i < dashEnd; ++i) {
			if (cell[i] != '-') {
				return false;
			}
		}

		if (alignLeft && alignRight) {
			alignments.push_back(Alignment::Center);
		}
		else if (alignRight) {
			alignments.push_back(Alignment::Right);
		}
		else if (alignLeft) {
			alignments.push_back(Alignment::Left);
		}
		else {
			alignments.push_back(Alignment::Default);
		}
	}
	return true;
}

const char* AlignmentClass(Alignment alignment)
{
	switch (alignment) {
	case Alignment::Left: return " class=\"markdown-align-left\"";
	case Alignment::Center: return " class=\"markdown-align-center\"";
	case Alignment::Right: return " class=\"markdown-align-right\"";
	default: return "";
	}
}

} // namespace

bool TryRenderTable(
	const std::vector<std::string>& lines,
	size_t headerLineIndex,
	InlineRenderer inlineRenderer,
	std::string& html,
	size_t& lastRenderedLineIndex)
{
	if (inlineRenderer == nullptr || headerLineIndex + 1 >= lines.size()) {
		return false;
	}

	const TableRow header = SplitTableRow(lines[headerLineIndex]);
	if (!header.hasSeparator || header.cells.empty()) {
		return false;
	}

	std::vector<Alignment> alignments;
	const TableRow delimiter = SplitTableRow(lines[headerLineIndex + 1]);
	if (!TryParseDelimiter(delimiter, alignments) || alignments.size() != header.cells.size()) {
		return false;
	}

	html += "<div class=\"markdown-table-wrap\"><table><thead><tr>";
	for (size_t column = 0; column < header.cells.size(); ++column) {
		html += "<th scope=\"col\"";
		html += AlignmentClass(alignments[column]);
		html += ">";
		html += inlineRenderer(header.cells[column]);
		html += "</th>";
	}
	html += "</tr></thead><tbody>";

	size_t bodyLineIndex = headerLineIndex + 2;
	for (; bodyLineIndex < lines.size(); ++bodyLineIndex) {
		if (TrimAsciiCopy(lines[bodyLineIndex]).empty()) {
			break;
		}
		const TableRow bodyRow = SplitTableRow(lines[bodyLineIndex]);
		if (!bodyRow.hasSeparator) {
			break;
		}

		html += "<tr>";
		for (size_t column = 0; column < header.cells.size(); ++column) {
			html += "<td";
			html += AlignmentClass(alignments[column]);
			html += ">";
			if (column < bodyRow.cells.size()) {
				html += inlineRenderer(bodyRow.cells[column]);
			}
			html += "</td>";
		}
		html += "</tr>";
	}

	html += "</tbody></table></div>";
	lastRenderedLineIndex = bodyLineIndex - 1;
	return true;
}

} // namespace AIChatMarkdownTableRenderer
