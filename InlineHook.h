#pragma once
#include <Windows.h>
#include <cstdint>

class InlineHook {
public:
    InlineHook(void* original, void* hook)
        : original_(original), hook_(hook) {}

    bool Hook() {
        DWORD oldProtect;
        if (!VirtualProtect(original_, 5, PAGE_EXECUTE_READWRITE, &oldProtect)) {
            return false;
        }

        *reinterpret_cast<uint8_t*>(original_) = 0xE9;  // jmp÷∏¡Ó
        *reinterpret_cast<int*>(reinterpret_cast<uintptr_t>(original_) + 1) = reinterpret_cast<int>(hook_) - reinterpret_cast<int>(original_) - 5;

        return VirtualProtect(original_, 5, oldProtect, &oldProtect);
    }

private:
    void* original_;
    void* hook_;
};