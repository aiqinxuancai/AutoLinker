#pragma once

#include <string>

// e-packager 集成：负责工具更新、解包当前易语言工程和打开输出目录。
namespace EPackagerIntegration {

// 弹出目录选择窗口，并把当前 .e 文件解包到所选目录下的 {文件名}.unpack。
void RunCurrentSourceUnpackToDirectory();

// 构建“将[xxx.e]反编译到目录”菜单标题。
std::wstring BuildUnpackMenuTitle();

// 当前是否有可解包的 .e 文件。
bool CanUnpackCurrentSource();

} // namespace EPackagerIntegration
