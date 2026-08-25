#pragma once

// 易语言 IDE 静默编译弹窗守卫：处理编译期间不会改变编译语义的确认框。

#include <string>

namespace IdeCompileDialogGuard {

void BeginCompileSession();
void EndCompileSession();
bool IsCompileSessionActive();

// 从非 IDE 主线程调用；精确匹配依赖写出提示后异步点击“不写出”。
bool TryDismissDependencyWriteDialog();
bool WasDependencyWriteDialogDismissed();

// 精确匹配“多个名称与指定拼音输入字相对应”的选择框后异步点击“确定”。
// 一个编译会话可能连续出现多个选择框，因此每次发现都尝试处理。
bool TryDismissNameConflictDialog();
bool WasNameConflictDialogDismissed();

// 无需 IDE 的提示语义匹配自检。
std::string BuildSelfTestJson();

} // namespace IdeCompileDialogGuard
