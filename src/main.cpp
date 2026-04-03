#include <windows.h>
// 编译/调试入口地址在 IDE 就绪后再解析。
int g_debugStartAddress = 0;
int g_compileStartAddress = 0;


extern "C" BOOL WINAPI DllMain(
	HMODULE hModule,
	DWORD ul_reason_for_call,
	LPVOID lpReserved)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		if (hModule != nullptr) {
			DisableThreadLibraryCalls(hModule);
		}
		break;
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}
