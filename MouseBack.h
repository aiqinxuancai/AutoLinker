#pragma once
#include <Windows.h>

void StartEditViewSubclassTask();

//处理输入框代码编辑框
LRESULT CALLBACK EditViewSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);