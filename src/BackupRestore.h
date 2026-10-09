#pragma once
#include <filesystem>
#include <string>
// 当前工程恢复事务：先备份当前内容，失败回滚，原源码文件不写入。
namespace BackupRestore {
std::filesystem::path CurrentSource();
bool Restore(const std::filesystem::path& source, const std::filesystem::path& backup,
    std::filesystem::path& safetyBackup, std::wstring& message);
}

