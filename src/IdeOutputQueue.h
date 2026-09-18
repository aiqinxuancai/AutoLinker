#pragma once
#include <Windows.h>
#include <deque>
#include <mutex>
#include <string>
#include <vector>

// IDE 输出只异步投递；调用者持有业务锁时也不会同步等待窗口线程。
class IdeOutputQueue {
public:
    static IdeOutputQueue& Instance() { static IdeOutputQueue queue; return queue; }
    UINT Message() const { static const UINT message = RegisterWindowMessageW(L"AutoLinker.Output.Drain"); return message; }
    void Push(HWND window, std::string line) {
        std::lock_guard<std::mutex> lock(mutex_);
        if (stopped_ || !window) return;
        // 超大单条日志仅保留文件副本，避免突破 UI 队列的总字节上限。
        if (line.size() > 4 * 1024 * 1024) {
            ++dropped_;
            if (!posted_) posted_ = PostMessageW(window, Message(), 0, 0) != FALSE;
            return;
        }
        while (!lines_.empty() && (lines_.size() >= 1024 || bytes_ + line.size() > 4 * 1024 * 1024)) {
            bytes_ -= lines_.front().size(); lines_.pop_front(); ++dropped_;
        }
        bytes_ += line.size(); lines_.push_back(std::move(line));
        if (!posted_) posted_ = PostMessageW(window, Message(), 0, 0) != FALSE;
        if (!posted_) { lines_.clear(); bytes_ = 0; }
    }
    std::vector<std::string> Take(HWND window) {
        std::vector<std::string> batch;
        std::lock_guard<std::mutex> lock(mutex_);
        posted_ = false;
        if (dropped_) { batch.push_back("[AutoLinker] IDE output queue overflow: " + std::to_string(dropped_) + " lines omitted; see file log."); dropped_ = 0; }
        while (!lines_.empty() && batch.size() < 64) {
            bytes_ -= lines_.front().size(); batch.push_back(std::move(lines_.front())); lines_.pop_front();
        }
        if (!lines_.empty()) posted_ = PostMessageW(window, Message(), 0, 0) != FALSE;
        if (!posted_ && !lines_.empty()) { lines_.clear(); bytes_ = 0; }
        return batch;
    }
    void Stop() { std::lock_guard<std::mutex> lock(mutex_); stopped_ = true; lines_.clear(); bytes_ = 0; }
private:
    std::mutex mutex_;
    std::deque<std::string> lines_;
    size_t bytes_ = 0, dropped_ = 0;
    bool posted_ = false, stopped_ = false;
};
