#pragma once

#include <Windows.h>
#include <afx.h>

#include <cstddef>
#include <cstdint>
#include <vector>

namespace e571 {

static_assert(sizeof(void*) == 4, "DirectGlobalSearch requires a 32-bit build.");

// 这是一个“整体搜寻”桥接层。
// 约束：
// 1. 必须在目标进程内调用。
// 2. 必须使用 32 位编译。
// 3. 必须与目标程序的 MFC / CStringA ABI 保持一致。
class DirectGlobalSearch {
public:
    // 搜索上下文，对应目标程序里被 sub_4C74E0 初始化的 16 字节对象。
    struct SearchContext {
        int type;   // 搜索类别，已知会遍历 1/2/3/4/6/7/8
        int data;   // 当前类别对应的数据源指针；type 1 时为 bucket 数据
        int flag;   // 原始逻辑里保持为 0
        int owner;  // 固定指向 0x5CAF30 对象
    };

    // 原生搜索结果记录，布局与程序内部追加到结果数组的 0x14 字节一致。
    struct GlobalSearchHit {
        int type;         // 搜索类别
        int extra;        // type 1 时为 bucketId，其他类别通常为 0
        int outerIndex;   // 外层索引
        int innerIndex;   // 内层索引
        int matchOffset;  // 命中串在 text 中的偏移
    };

    // 便于外部直接消费的增强结果。
    struct SearchResult {
        GlobalSearchHit hit;  // 可直接用于跳转
        CStringA prefix;      // FetchSearchText 输出的前缀
        CStringA text;        // FetchSearchText 输出的正文
        CStringA displayText; // 方便展示的拼接文本
    };

    struct SearchOptions {
        bool matchCase = false;  // true=区分大小写，false=不区分大小写
    };

    explicit DirectGlobalSearch(std::uintptr_t moduleBase = 0x400000);

    // 执行“整体搜寻”，返回所有命中结果。
    // moduleBase 默认按 e571.exe 的 0x400000 处理；若有重定位，传实际模块基址。
    std::vector<SearchResult> Search(const char* keyword, const SearchOptions& options = {}) const;

    // 按命中记录直接执行跳转。
    bool JumpToResult(const GlobalSearchHit& hit) const;

    // SearchResult 只是对 GlobalSearchHit 的包装，跳转时直接复用底层记录。
    bool JumpToResult(const SearchResult& result) const;

private:
    struct DwordContainer;

    std::uintptr_t moduleBase_;

    template <typename T>
    T Bind(std::uintptr_t absoluteAddress) const;

    template <typename T>
    T* Ptr(std::uintptr_t absoluteAddress) const;

    void InitContext(SearchContext* ctx) const;
    int GetOuterCount(SearchContext* ctx) const;
    int GetInnerCount(SearchContext* ctx, int outerIndex) const;
    int FetchSearchText(
        SearchContext* ctx,
        int outerIndex,
        int innerIndex,
        CStringA* outText,
        CStringA* outPrefix,
        int filter,
        int* optionalOut) const;
    int ContainerGetAt(void* container, int index, int* outPtr) const;
    int ContainerGetId(void* container, int index) const;
    int ResolveBucketIndex(void* container, int bucketId, int* outValue, int* outPos) const;
    HWND OpenCodeTarget(
        void* mainEditorHost,
        int type,
        int arg2,
        int arg3,
        int outerIndex,
        int innerIndex,
        int activate,
        int arg7) const;
    void MoveToLine(HWND hwnd, int outerIndex, int innerIndex, int arg3, int arg4, int arg5) const;
    void EnsureVisible(HWND hwnd, int arg1, int arg2, int arg3) const;
    void MoveCaretToOffset(HWND hwnd, int matchOffset, int force, void* maybeNull, int redraw) const;
    void ActivateWindow(HWND hwnd) const;
    void NotifyOpenFailure(int messageId) const;

    // 将搜索类别映射到对应的数据源。
    // type 1 不走这里，而是单独枚举 0x5CB184 容器中的 bucket。
    int ResolveTypeData(int type) const;

    // type 1 不是单一类别，而是需要枚举 bucket 容器中的每一项，
    // 每个 bucket 都要单独构造 SearchContext 并执行完整搜索。
    void EnumerateType1(
        const char* keyword,
        std::size_t keywordLen,
        const SearchOptions& options,
        std::vector<SearchResult>& results) const;

    // 对一个 SearchContext 执行完整搜索遍历：
    // outer -> inner -> FetchSearchText -> 字符串匹配 -> 生成命中记录。
    void SearchOneContext(
        SearchContext& ctx,
        const char* keyword,
        std::size_t keywordLen,
        const SearchOptions& options,
        int extra,
        std::vector<SearchResult>& results) const;

    // 在单条文本上执行滑动匹配，支持一条文本命中多次。
    void CollectMatches(
        int type,
        int extra,
        int outerIndex,
        int innerIndex,
        const char* keyword,
        std::size_t keywordLen,
        const SearchOptions& options,
        const CStringA& prefix,
        const CStringA& text,
        std::vector<SearchResult>& results) const;

    // 目标程序使用 MBCS 比较函数，因此这里保持与原逻辑一致。
    bool MatchAt(const char* haystack, const char* needle, std::size_t count, bool matchCase) const;

    // 这里没有强行复刻 UI 文本格式，只做最简单的“prefix + text”展示。
    CStringA BuildDisplayText(const CStringA& prefix, const CStringA& text) const;

    // 复刻双击搜索结果后的落点校验逻辑：
    // 1. 当前打开项是否与命中记录一致
    // 2. 光标是否成功移动到 matchOffset
    bool FocusResult(HWND hwnd, const GlobalSearchHit& hit) const;
};

}  // namespace e571
