#pragma once
#include <format>

#include <string>
#include <Windows.h>

///Ä£¿é²Ù×÷À©Õ¹

int GetEModelCount();

std::string GetEModelPath(int index);

BOOL AddEModel(std::string filePath);

BOOL AddEModel2(std::string filePath);

BOOL RemoveEModel(int index);

BOOL RemoveEModel(std::string filePah);

int FindEModelIndex(std::string filePah);

int FindEModelNameIndex(std::string ecomName);


void RunChangeECOM(bool isDebug);