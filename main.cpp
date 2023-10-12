#include <windows.h>
#include "PathHelper.h"
#include <filesystem>
#include <psapi.h>
#include <vector>
#include <algorithm>
#include "Global.h"

int g_compilerAddress;


std::vector<byte> get_module_bytes(HMODULE h) {
	MODULEINFO module_info;
	GetModuleInformation(GetCurrentProcess(), h, &module_info, sizeof(module_info));

	std::vector<byte> bytes(module_info.SizeOfImage);
	CopyMemory(bytes.data(), h, bytes.size());

	return bytes;
}

ptrdiff_t find_bytes(const std::vector<byte>& bytes, const std::vector<byte>& pattern) {
	auto it = std::search(bytes.begin(), bytes.end(), pattern.begin(), pattern.end());
	if (it != bytes.end()) {
		return std::distance(bytes.begin(), it);
	}
	else {
		return -1;
	}
}


int get_build_base() {
	HMODULE h = GetModuleHandle(NULL);
	if (h == NULL) {
		//MessageBox(NULL, "h=0", NULL, MB_OK);
		return 0;
	}
	std::vector<byte> module_bytes = get_module_bytes(h);
	std::vector<byte> pattern = { 139, 69, 240, 139, 78, 8, 87, 80, 131, 193, 72 };  // ÌØÕ÷Âë

	ptrdiff_t address = find_bytes(module_bytes, pattern);
	if (address != -1) {
		address -= 1;
		return reinterpret_cast<int>(h) + address;
	}
	return 0;
}



BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH: {
		g_compilerAddress = get_build_base();
		break;
	}
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}