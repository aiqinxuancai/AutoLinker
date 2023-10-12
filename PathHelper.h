#pragma once

#include <string>
#include <vector>
#include <optional>
#include <Windows.h>

std::string GetBasePath();

std::string ExtractBetweenDashes(const std::string& text);

/// <summary>
/// ∂¡»°
/// </summary>
/// <param name="filename"></param>
/// <param name="search_bytes"></param>
/// <returns></returns>
std::optional<size_t> FindByteInFile(const std::string& filename, const std::vector<char>& search_bytes);