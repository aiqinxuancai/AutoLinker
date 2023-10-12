#pragma once

#include <string>
#include <Windows.h>

std::string GetBasePath();

std::string ExtractBetweenDashes(const std::string& text);