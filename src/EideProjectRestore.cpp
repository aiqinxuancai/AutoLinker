#include "EideProjectRestore.h"
#include <Windows.h>
#include <cstdint>
#include <cstring>

namespace EideProjectRestore {
namespace {
struct Signature { uintptr_t address; const char* hex; };
constexpr Signature signatures[] = {
    // 同步验证加载分发、收尾及 IDE 自身容器 ABI。
    {0x46B310, "558bec6aff68f0755e0064a10000000050648925"},
    {0x46B4D0, "6aff68077b5e0064a1000000005064892500"},
    {0x5BC5AA, "56578b7c240c8bf185ff750433c0eb0757ff"},
    {0x46B010, "568bf1e828daffff8b864c09000085c0"},
    {0x457490, "558bec6aff680d5b5e0064a100000000"},
    {0x46D0C0, "64a1000000006aff68e87b5e00508b44"},
    {0x46D3C0, "568bf1578dbec00000008bcfe8ff38ff"},
    {0x5BB2C6, "8bc133c9c700a0006000894804894810"},
    {0x5BB2F9, "b8b44d5f00e87553f7ff51894df0c701"},
    {0x5BC48F, "a1bc076700538b5c2408568bf15785db"},
    {0x5BC421, "568bf18b068d48f43b0dc00767007418"},
    {0x5BC55A, "56578b7c240c8bf18b0e8b073bc87439"},
    {0x460C80, "8b9138030000b801000000f7d223d03a"}
};
int Hex(char c) { return c <= '9' ? c - '0' : c - 'a' + 10; }
#if defined(_M_IX86)
constexpr uintptr_t kDocument = 0x6756E0;
constexpr size_t kPathOffsets[] = {0x3FC, 0x400, 0x404};

bool ReadIdentityNative(char (*paths)[32768])
{
    __try {
        const auto doc = reinterpret_cast<const unsigned char*>(kDocument);
        if ((*reinterpret_cast<const unsigned*>(doc + 0x3F8) & 1) == 0) return false;
        for (size_t i = 0; i < 3; ++i) {
            const auto value = *reinterpret_cast<const char* const*>(doc + kPathOffsets[i]);
            if (!value) return false;
            const auto length = strnlen_s(value, 32768);
            if (length >= 32768) return false;
            memcpy(paths[i], value, length + 1);
        }
        return paths[0][0] != '\0';
    } __except(EXCEPTION_EXECUTE_HANDLER) { return false; }
}

// 复现 0x46B310 的二进制分支：关闭旧视图 -> 加载 -> 恢复身份 -> post-load。
// 不调用完整“打开工程”包装器，避免其先将暂存路径发布到标题、MRU 和插件。
// CPtrArray 必须由 IDE 构造和析构，不能传入插件版本 MFC/STL 容器。
int ReloadNative(const char* path, const char* const* originalPaths)
{
    using Unary = void(__thiscall*)(void*);
    using Load = int(__thiscall*)(void*, const char*, void*, int);
    using Finish = int(__thiscall*)(void*, int, int, void*);
    using Assign = void*(__thiscall*)(void*, const char*);
    auto doc = reinterpret_cast<unsigned char*>(kDocument);
    unsigned errors[5] = {};
    bool arrayConstructed = false;
    int result = -1;
    __try {
        __try {
            reinterpret_cast<Unary>(0x5BB2C6)(errors);
            arrayConstructed = true;
            reinterpret_cast<Unary>(0x46B010)(doc);
            const int loaded = reinterpret_cast<Load>(0x46B4D0)(doc, path, errors, 1);
            // 成功与失败均先恢复事务开始时的身份，回滚不能继承失败现场路径。
            for (size_t i = 0; i < 3; ++i)
                reinterpret_cast<Assign>(0x5BC5AA)(doc + kPathOffsets[i], originalPaths[i]);
            result = reinterpret_cast<Finish>(0x46D0C0)(doc, loaded, 0, errors);
            // post-load 的失败分支会清空工程，再次确保后续回滚使用的身份一致。
            if (result != 1) {
                for (size_t i = 0; i < 3; ++i)
                    reinterpret_cast<Assign>(0x5BC5AA)(doc + kPathOffsets[i], originalPaths[i]);
            } else {
                *reinterpret_cast<unsigned*>(doc + 0x418) = 1;
                reinterpret_cast<Unary>(0x46D3C0)(doc);
                for (size_t i = 0; i < 3; ++i)
                    if (strcmp(*reinterpret_cast<const char**>(doc + kPathOffsets[i]), originalPaths[i]) != 0)
                        result = -2;
            }
        } __finally {
            if (arrayConstructed) reinterpret_cast<Unary>(0x5BB2F9)(errors);
        }
    } __except(EXCEPTION_EXECUTE_HANDLER) { return -1; }
    return result;
}
#endif
}
bool IsSupported()
{
#if defined(_M_IX86)
    __try {
        const auto base = reinterpret_cast<uintptr_t>(GetModuleHandleW(nullptr));
        // 本版本签名包含绝对地址；其他装载布局同样安全拒绝。
        if (base != 0x400000) return false;
        const auto dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(base);
        if (dos->e_magic != IMAGE_DOS_SIGNATURE) return false;
        const auto nt = reinterpret_cast<const IMAGE_NT_HEADERS*>(base + dos->e_lfanew);
        if (nt->Signature != IMAGE_NT_SIGNATURE || nt->FileHeader.Machine != IMAGE_FILE_MACHINE_I386 ||
            nt->OptionalHeader.SizeOfImage < 0x276300) return false;
        for (const auto& signature : signatures) {
            const auto bytes = reinterpret_cast<const unsigned char*>(signature.address);
            for (size_t i = 0; i < strlen(signature.hex) / 2; ++i)
                if (bytes[i] != (Hex(signature.hex[i * 2]) * 16 + Hex(signature.hex[i * 2 + 1]))) return false;
        }
        return true;
    } __except(EXCEPTION_EXECUTE_HANDLER) { return false; }
#else
    return false;
#endif
}
bool CaptureIdentity(Identity& identity, std::string& error)
{
    identity = {};
    error.clear();
    if (!IsSupported()) { error = "unsupported_ide_requires_e5.95"; return false; }
#if defined(_M_IX86)
    char paths[3][32768] = {};
    if (!ReadIdentityNative(paths)) { error = "current_project_identity_unavailable"; return false; }
    for (size_t i = 0; i < 3; ++i) identity.paths[i] = paths[i];
    return true;
#else
    return false;
#endif
}
void* CurrentSerializerObject()
{
#if defined(_M_IX86)
    if (!IsSupported()) return nullptr;
    char paths[3][32768] = {};
    return ReadIdentityNative(paths) ? reinterpret_cast<void*>(kDocument + 0xC0) : nullptr;
#else
    return nullptr;
#endif
}
bool Reload(const std::string& path, const Identity& identity, std::string& error)
{
    error.clear();
    if (!IsSupported()) { error = "unsupported_ide_requires_e5.95"; return false; }
    if (identity.paths[0].empty()) { error = "restore_identity_missing"; return false; }
#if defined(_M_IX86)
    const char* paths[] = {identity.paths[0].c_str(), identity.paths[1].c_str(), identity.paths[2].c_str()};
    const int result = ReloadNative(path.c_str(), paths);
    if (result != 1) {
        error = result == -2 ? "restore_identity_mismatch" : result < 0 ? "native_reload_exception" : "native_reload_failed";
        return false;
    }
    return true;
#else
    return false;
#endif
}
}
