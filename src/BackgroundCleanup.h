#pragma once
#include <condition_variable>
#include <deque>
#include <functional>
#include <mutex>
#include <thread>
#include <memory>

// 独立文件清理队列；任务只允许访问文件系统，不得访问 IDE。
class BackgroundCleanup {
public:
    static std::shared_ptr<BackgroundCleanup> Shared() {
        static auto instance = std::shared_ptr<BackgroundCleanup>(new BackgroundCleanup()); return instance;
    }
    void Enqueue(std::function<void()> work) {
        { std::lock_guard<std::mutex> lock(mutex_);
          if (!stopping_) { queue_.push_back(std::move(work)); cv_.notify_one(); return; } }
        // 退出后最后一个租约仍可安全清理，不再创建或唤醒线程。
        try { work(); } catch (...) {}
    }
    void Stop() {
        { std::lock_guard<std::mutex> lock(mutex_); stopping_ = true; }
        cv_.notify_one();
        if (worker_.joinable()) worker_.join();
    }
    ~BackgroundCleanup() { Stop(); }
private:
    BackgroundCleanup() : worker_([this] {
        for (;;) {
            std::function<void()> work;
            { std::unique_lock<std::mutex> lock(mutex_);
              cv_.wait(lock, [this] { return stopping_ || !queue_.empty(); });
              if (queue_.empty()) return;
              work = std::move(queue_.front()); queue_.pop_front(); }
            try { work(); } catch (...) {}
        }
    }) {}
    std::mutex mutex_;
    std::condition_variable cv_;
    std::deque<std::function<void()>> queue_;
    bool stopping_ = false;
    std::thread worker_;
};
