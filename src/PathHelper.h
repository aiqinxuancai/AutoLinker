#pragma once

#include <string>
#include <vector>
#include <optional>
#include <Windows.h>

std::string GetBasePath();

std::string ExtractBetweenDashes(const std::string& text);

/// <summary>
/// 读取
/// </summary>
/// <param name="filename"></param>
/// <param name="search_bytes"></param>
/// <returns></returns>
std::optional<size_t> FindByteInFile(const std::string& filename, const std::vector<char>& search_bytes);


/// <summary>
/// 获取out的文件名
/// </summary>
/// <param name="s"></param>
/// <returns></returns>
std::string GetLinkerCommandOutFileName(const std::string& s);


std::string GetLinkerCommandKrnlnFileName(const std::string& s);