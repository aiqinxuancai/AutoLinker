
#include "MouseBack.h"
#include <vector>
#include <format>
#include "Global.h"
#include "StringHelper.h"
#include <thread>
#include <CommCtrl.h>

//本代码用于在使用鼠标后退键时，按下Ctrl+J快捷键，实现其他IDE的鼠标键快速后退到上次编辑的位置

std::jthread g_mouseBackThread;
std::vector<HWND> g_subclassedWindows;

//处理代码编辑框的消息
LRESULT CALLBACK EditViewSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	switch (uMsg) {
	case WM_XBUTTONDOWN:
		break;
	case WM_XBUTTONUP:
		WORD xButton = GET_XBUTTON_WPARAM(wParam);
		if (xButton == XBUTTON1) { //后退
			INPUT inputs[4] = {};

			inputs[0].type = INPUT_KEYBOARD;
			inputs[0].ki.wVk = VK_CONTROL;

			inputs[1].type = INPUT_KEYBOARD;
			inputs[1].ki.wVk = 'J';

			inputs[2].type = INPUT_KEYBOARD;
			inputs[2].ki.wVk = 'J';
			inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;

			inputs[3].type = INPUT_KEYBOARD;
			inputs[3].ki.wVk = VK_CONTROL;
			inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;

			// Send the input
			UINT uSent = SendInput(ARRAYSIZE(inputs), inputs, sizeof(INPUT));
			if (uSent != ARRAYSIZE(inputs)) {
				// Not all inputs were sent successfully
			}
		}
		else if (xButton == XBUTTON2) {
			OutputStringToELog("2");
		}
		break;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}


BOOL CALLBACK EnumChildProcEditView(HWND hwndChild, LPARAM lParam) {
	char className[100];
	GetClassName(hwndChild, className, sizeof(className));
	if (std::string(className) == "Afx:400000:b:10003:900015:0") {
		if (std::find(g_subclassedWindows.begin(), g_subclassedWindows.end(), hwndChild) == g_subclassedWindows.end()) {
			int result = SendMessage(g_hwnd, 20708, (WPARAM)hwndChild, 0);
			if (result != 1) {
				// Handle error
				std::string s = std::format("子类化失败 {}", (int)hwndChild);
				OutputStringToELog(s);
			}
			else {
				std::string s = std::format("子类化 {}", (int)hwndChild);
				OutputStringToELog(s);
				g_subclassedWindows.push_back(hwndChild);
			}
		}
	}
	EnumChildWindows(hwndChild, EnumChildProcEditView, 0);

	return TRUE;
}


void StartEditViewSubclassTask() {
	g_mouseBackThread = std::jthread([&](std::stop_token stoken) {
		while (!stoken.stop_requested()) {
			//OutputStringToELog("检查代码框");
			//g_allFindWindow = 0;
			EnumChildWindows(g_hwnd, EnumChildProcEditView, 0);
			//std::string s = std::format("检查代码框完毕{} {}", (int)g_hwnd, g_allFindWindow);
			//OutputStringToELog(s);
			std::this_thread::sleep_for(std::chrono::seconds(3));
		}
		});
}