#pragma once
#include <Windows.h>

void StartEditViewSubclassTask();

//������������༭��
LRESULT CALLBACK EditViewSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);