// ResourceTextLoader.h - 资源文本加载工具，统一读取内嵌到 DLL 的文本资源
#pragma once

#include <string>

// 从当前模块加载 UTF-8 HTML 资源文本，失败时返回空字符串。
std::string LoadUtf8HtmlResourceText(int resourceId);
