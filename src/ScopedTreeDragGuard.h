#pragma once
#include <Windows.h>
#include <CommCtrl.h>

// 防止同步模拟按下时进入树控件拖拽等待；退出时只恢复本类更改的样式位。
class ScopedTreeDragGuard {
public:
    explicit ScopedTreeDragGuard(HWND window) : window_(window),
        originallyDisabled_((GetWindowLongPtr(window, GWL_STYLE) & TVS_DISABLEDRAGDROP) != 0) {
        SetWindowLongPtr(window_, GWL_STYLE, GetWindowLongPtr(window_, GWL_STYLE) | TVS_DISABLEDRAGDROP);
    }
    ~ScopedTreeDragGuard() {
        if (IsWindow(window_) && !originallyDisabled_)
            SetWindowLongPtr(window_, GWL_STYLE, GetWindowLongPtr(window_, GWL_STYLE) & ~static_cast<LONG_PTR>(TVS_DISABLEDRAGDROP));
        if (GetCapture() == window_) ReleaseCapture();
    }
    ScopedTreeDragGuard(const ScopedTreeDragGuard&) = delete;
    ScopedTreeDragGuard& operator=(const ScopedTreeDragGuard&) = delete;
private:
    HWND window_;
    bool originallyDisabled_;
};
