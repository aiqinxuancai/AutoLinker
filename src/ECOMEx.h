#pragma once
#include <format>

#include <string>
#include <Windows.h>

///Ä£¿é²Ù×÷À©Õ¹

int GetECOMCount();

std::string GetECOMPath(int index);

BOOL AddECOM(std::string filePath);

BOOL AddECOM2(std::string filePath);

BOOL RemoveECOM(int index);

BOOL RemoveECOM(std::string filePah);

int FindECOMIndex(std::string filePah);

int FindECOMNameIndex(std::string ecomName);


void RunChangeECOM(bool isDebug);