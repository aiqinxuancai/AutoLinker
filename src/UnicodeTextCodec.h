#pragma once

#include <cstddef>
#include <string>
#include <string_view>

// Unicode 文本编码桥接：在 UTF-8 与系统本地代码页之间无损保存不可表示字符。
namespace UnicodeTextCodec {

// 将 UTF-8 转为本地编码，不可表示字符使用安全的数字字符引用暂存。
std::string Utf8ToLocalPreservingUnicode(const std::string& text);

// 将本地编码转为 UTF-8，并还原暂存的数字字符引用。
std::string LocalToUtf8RestoringUnicode(const std::string& text);

// 返回指定位置合法数字字符引用的字节数，不匹配时返回 0。
size_t NumericCharacterReferenceLength(std::string_view text, size_t offset);

// 转义 HTML 文本，同时保留合法数字字符引用供 WebView 恢复 Unicode。
std::string EscapeHtmlTextPreservingUnicode(const std::string& text);

} // namespace UnicodeTextCodec
