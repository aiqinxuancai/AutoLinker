#pragma once
#include <format>

#include <string>
#include <Windows.h>

// 易模块导入及动静态切换扩展。

int GetECOMCount();

std::string GetECOMPath(int index);

BOOL AddECOM(std::string filePath);

BOOL AddECOM2(std::string filePath);

BOOL RemoveECOM(int index);

BOOL RemoveECOM(std::string filePah);

int FindECOMIndex(std::string filePah);

int FindECOMNameIndex(std::string ecomName);


// 根据阶段切换 EC：true 切到静态编译版，false 切回动态调试版；返回成功切换数量。
int RunChangeECOM(bool useCompileModules);
