#pragma once

#include <Windows.h>
#include <afx.h>

#include <cstddef>
#include <cstdint>
#include <vector>

namespace e571 {

static_assert(sizeof(void*) == 4, "DirectGlobalSearch requires a 32-bit build.");

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

	struct SearchResult {
		GlobalSearchHit hit;
		CStringA prefix;
		CStringA text;
		CStringA displayText;
	};

	struct SearchOptions {
		bool matchCase = false;
	};

	explicit DirectGlobalSearch(std::uintptr_t moduleBase = 0x400000);

	std::vector<SearchResult> Search(const char* keyword, const SearchOptions& options = {}) const;
	bool JumpToResult(const GlobalSearchHit& hit) const;
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

	int ResolveTypeData(int type) const;
	void EnumerateType1(
		const char* keyword,
		std::size_t keywordLen,
		const SearchOptions& options,
		std::vector<SearchResult>& results) const;
	void SearchOneContext(
		SearchContext& ctx,
		const char* keyword,
		std::size_t keywordLen,
		const SearchOptions& options,
		int extra,
		std::vector<SearchResult>& results) const;
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
	bool MatchAt(const char* haystack, const char* needle, std::size_t count, bool matchCase) const;
	CStringA BuildDisplayText(const CStringA& prefix, const CStringA& text) const;
	bool FocusResult(HWND hwnd, const GlobalSearchHit& hit) const;
};

} // namespace e571
