#include "UnicodeTextCodec.h"

#include <Windows.h>

#include <charconv>
#include <cstdint>
#include <format>

namespace UnicodeTextCodec {
namespace {

bool IsValidUtf8(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0) > 0;
}

std::wstring Utf8ToWide(const std::string& text)
{
	const int length = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (length <= 0) {
		return {};
	}

	std::wstring wide(static_cast<size_t>(length), L'\0');
	if (MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		length) <= 0) {
		return {};
	}
	return wide;
}

std::string LocalToUtf8(const std::string& text)
{
	const int wideLength = MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLength <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLength), L'\0');
	if (MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLength) <= 0) {
		return text;
	}

	const int utf8Length = WideCharToMultiByte(
		CP_UTF8,
		0,
		wide.data(),
		wideLength,
		nullptr,
		0,
		nullptr,
		nullptr);
	if (utf8Length <= 0) {
		return text;
	}

	std::string utf8(static_cast<size_t>(utf8Length), '\0');
	if (WideCharToMultiByte(
		CP_UTF8,
		0,
		wide.data(),
		wideLength,
		utf8.data(),
		utf8Length,
		nullptr,
		nullptr) <= 0) {
		return text;
	}
	return utf8;
}

bool IsHighSurrogate(wchar_t value)
{
	return value >= 0xD800 && value <= 0xDBFF;
}

bool IsLowSurrogate(wchar_t value)
{
	return value >= 0xDC00 && value <= 0xDFFF;
}

uint32_t DecodeCodePoint(const wchar_t* chars, int charCount)
{
	if (charCount == 2) {
		return 0x10000u +
			((static_cast<uint32_t>(chars[0]) - 0xD800u) << 10) +
			(static_cast<uint32_t>(chars[1]) - 0xDC00u);
	}
	return static_cast<uint32_t>(chars[0]);
}

void AppendUtf8CodePoint(uint32_t codePoint, std::string& out)
{
	if (codePoint <= 0x7Fu) {
		out.push_back(static_cast<char>(codePoint));
	}
	else if (codePoint <= 0x7FFu) {
		out.push_back(static_cast<char>(0xC0u | (codePoint >> 6)));
		out.push_back(static_cast<char>(0x80u | (codePoint & 0x3Fu)));
	}
	else if (codePoint <= 0xFFFFu) {
		out.push_back(static_cast<char>(0xE0u | (codePoint >> 12)));
		out.push_back(static_cast<char>(0x80u | ((codePoint >> 6) & 0x3Fu)));
		out.push_back(static_cast<char>(0x80u | (codePoint & 0x3Fu)));
	}
	else {
		out.push_back(static_cast<char>(0xF0u | (codePoint >> 18)));
		out.push_back(static_cast<char>(0x80u | ((codePoint >> 12) & 0x3Fu)));
		out.push_back(static_cast<char>(0x80u | ((codePoint >> 6) & 0x3Fu)));
		out.push_back(static_cast<char>(0x80u | (codePoint & 0x3Fu)));
	}
}

bool TryParseNumericCharacterReference(
	std::string_view text,
	size_t offset,
	uint32_t& codePoint,
	size_t& length)
{
	codePoint = 0;
	length = 0;
	if (offset + 4 > text.size() || text[offset] != '&' || text[offset + 1] != '#') {
		return false;
	}

	size_t digitsBegin = offset + 2;
	int base = 10;
	if (digitsBegin < text.size() && (text[digitsBegin] == 'x' || text[digitsBegin] == 'X')) {
		base = 16;
		++digitsBegin;
	}
	const size_t semicolon = text.find(';', digitsBegin);
	if (semicolon == std::string_view::npos || semicolon == digitsBegin || semicolon - digitsBegin > 8) {
		return false;
	}

	uint32_t parsed = 0;
	const char* begin = text.data() + digitsBegin;
	const char* end = text.data() + semicolon;
	const auto parseResult = std::from_chars(begin, end, parsed, base);
	if (parseResult.ec != std::errc() || parseResult.ptr != end ||
		parsed == 0 || parsed > 0x10FFFFu ||
		(parsed >= 0xD800u && parsed <= 0xDFFFu)) {
		return false;
	}

	codePoint = parsed;
	length = semicolon - offset + 1;
	return true;
}

std::string DecodeNumericCharacterReferences(const std::string& text)
{
	std::string out;
	out.reserve(text.size());
	for (size_t i = 0; i < text.size();) {
		uint32_t codePoint = 0;
		size_t length = 0;
		if (TryParseNumericCharacterReference(text, i, codePoint, length)) {
			AppendUtf8CodePoint(codePoint, out);
			i += length;
			continue;
		}
		out.push_back(text[i]);
		++i;
	}
	return out;
}

} // namespace

std::string Utf8ToLocalPreservingUnicode(const std::string& text)
{
	if (text.empty() || !IsValidUtf8(text)) {
		return text;
	}
	if (GetACP() == CP_UTF8) {
		return text;
	}

	const std::wstring wide = Utf8ToWide(text);
	if (wide.empty()) {
		return text;
	}

	std::string out;
	out.reserve(text.size());
	for (size_t i = 0; i < wide.size();) {
		int charCount = 1;
		if (IsHighSurrogate(wide[i]) && i + 1 < wide.size() && IsLowSurrogate(wide[i + 1])) {
			charCount = 2;
		}

		BOOL usedDefaultChar = FALSE;
		const char defaultChar = '?';
		const int encodedLength = WideCharToMultiByte(
			CP_ACP,
			WC_NO_BEST_FIT_CHARS,
			wide.data() + i,
			charCount,
			nullptr,
			0,
			&defaultChar,
			&usedDefaultChar);
		if (encodedLength > 0 && !usedDefaultChar) {
			std::string encoded(static_cast<size_t>(encodedLength), '\0');
			usedDefaultChar = FALSE;
			if (WideCharToMultiByte(
					CP_ACP,
					WC_NO_BEST_FIT_CHARS,
					wide.data() + i,
					charCount,
					encoded.data(),
					encodedLength,
					&defaultChar,
					&usedDefaultChar) > 0 &&
				!usedDefaultChar) {
				out += encoded;
				i += static_cast<size_t>(charCount);
				continue;
			}
		}

		out += std::format("&#x{:X};", DecodeCodePoint(wide.data() + i, charCount));
		i += static_cast<size_t>(charCount);
	}
	return out;
}

std::string LocalToUtf8RestoringUnicode(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const std::string utf8 = IsValidUtf8(text) ? text : LocalToUtf8(text);
	return DecodeNumericCharacterReferences(utf8);
}

size_t NumericCharacterReferenceLength(std::string_view text, size_t offset)
{
	uint32_t codePoint = 0;
	size_t length = 0;
	return TryParseNumericCharacterReference(text, offset, codePoint, length) ? length : 0;
}

std::string EscapeHtmlTextPreservingUnicode(const std::string& text)
{
	std::string out;
	out.reserve(text.size() + 32);
	for (size_t i = 0; i < text.size();) {
		const size_t referenceLength = NumericCharacterReferenceLength(text, i);
		if (referenceLength > 0) {
			out.append(text, i, referenceLength);
			i += referenceLength;
			continue;
		}

		switch (text[i]) {
		case '&': out += "&amp;"; break;
		case '<': out += "&lt;"; break;
		case '>': out += "&gt;"; break;
		case '"': out += "&quot;"; break;
		default: out.push_back(text[i]); break;
		}
		++i;
	}
	return out;
}

} // namespace UnicodeTextCodec
