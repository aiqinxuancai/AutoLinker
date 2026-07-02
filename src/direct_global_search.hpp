#pragma once

#include <cstdint>

namespace e571 {

static_assert(sizeof(void*) == 4, "DirectGlobalSearch requires a 32-bit build.");

// 易语言IDE全局搜索内部结构的最小类型定义。
class DirectGlobalSearch {
public:
	struct SearchContext {
		int type;
		int data;
		int flag;
		int owner;
	};

	struct GlobalSearchHit {
		int type;
		int extra;
		int outerIndex;
		int innerIndex;
		int matchOffset;
	};
};

} // namespace e571
