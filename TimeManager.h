#ifndef TIME_MANAGER_H
#define TIME_MANAGER_H

#include <windows.h>
#include <map>
#include <string>

typedef ULONGLONG(WINAPI *GetTickCount64Proc)();

typedef std::map<std::string, DWORD> TimeMap;

class TimeManager
{
private:
	CRITICAL_SECTION cs; 
    TimeMap timeMap;
    GetTickCount64Proc pGetTickCount64;
	
public:
    TimeManager()
    {
        // Try to get GetTickCount64 from kernel32.dll
		InitializeCriticalSection(&cs);
        HMODULE hKernel32 = GetModuleHandle("kernel32.dll");
        pGetTickCount64 = (GetTickCount64Proc)GetProcAddress(hKernel32, "GetTickCount64");
    }
	
	~TimeManager()
    {
        // Delete the critical section
        DeleteCriticalSection(&cs);
    }
	
    DWORD getTickCount()
    {
        if (pGetTickCount64 != NULL)
        {
            // If GetTickCount64 is available, use it
            return (DWORD)pGetTickCount64();
        }
        else
        {
            // Otherwise, use GetTickCount
            return GetTickCount();
        }
    }
	
	bool TimeManager::isCanCall(std::string key, int intervalTime) {
    EnterCriticalSection(&cs); // Enter critical section

    DWORD currentTime = getTickCount();
    TimeMap::iterator iter = timeMap.find(key);

    if (iter != timeMap.end()) {
        if (currentTime - iter->second < intervalTime) { // If the time since the last call is less than intervalTime
            LeaveCriticalSection(&cs); // Leave critical section
            return false; // Function call is too soon, so return false
        }
    }

    // Update the last call time for this key
    timeMap[key] = currentTime;

    LeaveCriticalSection(&cs); // Leave critical section

    return true; // Function can be called, so return true
	}
};


#endif // TIME_MANAGER_H