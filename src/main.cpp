#include <windows.h>
#include "PathHelper.h"
#include <filesystem>
#include <psapi.h>
#include <vector>
#include <algorithm>
#include "Global.h"
#include "MemFind.h"


int g_debugStartAddress;
int g_compileStartAddress;


BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH: {

		//寻找编译及调试开始的位置
		g_debugStartAddress = FindSelfModelMemory("55 8B EC 6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 81 EC ?? ?? 00 00 56 89 8D ?? FB FF FF 8B 8D ?? FB FF FF 81 C1 C0 00 00 00 E8 ?? ?? ?? ?? 85 C0 74 05 E9 ?? 08 00 00");
		g_compileStartAddress = FindSelfModelMemory("55 8B EC 6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 81 EC C4 00");

		//std::string s = std::format("找到内存地址 {:X} {:X}", g_debugStartAddress, g_compileStartAddress);
		//OutputDebugString(s.c_str());

		break;
	}
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}