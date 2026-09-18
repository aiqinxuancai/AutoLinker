#pragma once
#include "AIChatToolingInternal.h"
#include "WorkspaceMirror.h"
#include <atomic>
#include <deque>
#include <functional>
#include <thread>

// 仅主线程操作队列；后台线程只构建候选结果，通过消息通知继续事务。
class IdeToolScheduler {
public:
    UINT Message() const { static const UINT id = RegisterWindowMessageW(L"AutoLinker.ToolScheduler.Continue"); return id; }
    static constexpr UINT_PTR TimerId = 0x414C5451;
    std::deque<std::shared_ptr<ToolExecutionRequest>> queued;
    std::shared_ptr<ToolExecutionRequest> active;
    std::shared_ptr<WorkspaceMirror::RefreshTask> refresh;
    bool stopped = false;
    bool executing = false;
    bool HasWorker() const { return worker_.joinable(); }
    bool WorkerDone() const { return done_.load(std::memory_order_acquire); }
    void Start(HWND window, std::function<void()> work) {
        done_ = false;
        worker_ = std::thread([this, window, work = std::move(work)] {
            try { work(); } catch (...) {}
            done_.store(true, std::memory_order_release);
            PostMessageW(window, Message(), 0, 0);
        });
        // 消息投递失败也可由计时器回收；主线程从不等待尚未完成的工作线程。
        SetTimer(window, TimerId, 100, nullptr);
    }
    void JoinCompleted(HWND window) {
        if (worker_.joinable() && WorkerDone()) worker_.join();
        KillTimer(window, TimerId);
    }
    void Stop(HWND window) {
        stopped = true;
        WorkspaceMirror::CancelRefresh(refresh);
        if (worker_.joinable()) worker_.join();
        KillTimer(window, TimerId);
        refresh.reset();
    }
    ~IdeToolScheduler() { if (worker_.joinable()) { WorkspaceMirror::CancelRefresh(refresh); worker_.join(); } }
private:
    std::atomic_bool done_ = false;
    std::thread worker_;
};
