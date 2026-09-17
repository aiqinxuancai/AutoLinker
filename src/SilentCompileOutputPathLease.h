#pragma once
#include "Global.h"

// 作用域拥有静默输出路径请求，保证异常和提前返回时也会取消请求。
class SilentCompileOutputPathLease final {
public:
	SilentCompileOutputPathLease() = default;
	~SilentCompileOutputPathLease() { CancelSilentCompileOutputPathRequest(); }
	SilentCompileOutputPathLease(const SilentCompileOutputPathLease&) = delete;
	SilentCompileOutputPathLease& operator=(const SilentCompileOutputPathLease&) = delete;
};
