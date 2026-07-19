#pragma once

// 易语言 IDE 静默编译弹窗守卫：仅处理“是否写出相关依赖文件”确认框。

#include <string>

namespace IdeCompileDialogGuard {

void BeginCompileSession();
void EndCompileSession();

// 从非 IDE 主线程调用；精确匹配依赖写出提示后异步点击“不写出”。
bool TryDismissDependencyWriteDialog();
bool WasDependencyWriteDialogDismissed();

// 无需 IDE 的提示语义匹配自检。
std::string BuildSelfTestJson();

} // namespace IdeCompileDialogGuard
