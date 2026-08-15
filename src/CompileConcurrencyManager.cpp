#include "CompileConcurrencyManager.h"

#include <atomic>
#include <chrono>
#include <cstdint>
#include <format>
#include <thread>

#include "..\\thirdparty\\json.hpp"

namespace {

std::atomic_bool g_processCompileActive = false;

std::wstring NormalizeOutputPathForLock(const std::string& path)
{
	if (path.empty()) {
		return std::wstring();
	}
	const int wideLength = MultiByteToWideChar(
		CP_ACP,
		0,
		path.data(),
		static_cast<int>(path.size()),
		nullptr,
		0);
	if (wideLength <= 0) {
		return std::wstring(path.begin(), path.end());
	}
	std::wstring normalized(static_cast<size_t>(wideLength), L'\0');
	MultiByteToWideChar(
		CP_ACP,
		0,
		path.data(),
		static_cast<int>(path.size()),
		normalized.data(),
		wideLength);
	for (wchar_t& ch : normalized) {
		if (ch == L'/') {
			ch = L'\\';
		}
	}
	CharLowerBuffW(normalized.data(), static_cast<DWORD>(normalized.size()));
	return normalized;
}

std::uint64_t HashNormalizedOutputPath(const std::string& path)
{
	constexpr std::uint64_t kOffsetBasis = 14695981039346656037ull;
	constexpr std::uint64_t kPrime = 1099511628211ull;
	std::uint64_t hash = kOffsetBasis;
	for (const wchar_t ch : NormalizeOutputPathForLock(path)) {
		hash ^= static_cast<std::uint16_t>(ch);
		hash *= kPrime;
	}
	return hash;
}

std::wstring BuildMutexName(const std::string& normalizedOutputPath)
{
	return std::format(
		L"Local\\AutoLinker.CompileOutput.{:016X}",
		HashNormalizedOutputPath(normalizedOutputPath));
}

} // namespace

namespace CompileConcurrencyManager {

ProcessExecutionGuard::ProcessExecutionGuard()
{
	bool expected = false;
	acquired_ = g_processCompileActive.compare_exchange_strong(expected, true);
}

ProcessExecutionGuard::~ProcessExecutionGuard()
{
	if (acquired_) {
		g_processCompileActive.store(false);
	}
}

bool ProcessExecutionGuard::Acquired() const
{
	return acquired_;
}

OutputPathLock::~OutputPathLock()
{
	Release();
}

bool OutputPathLock::Acquire(
	const std::string& normalizedOutputPath,
	DWORD timeoutMilliseconds,
	std::string& outError)
{
	Release();
	outError.clear();
	if (normalizedOutputPath.empty()) {
		outError = "compile output path is empty";
		return false;
	}

	const std::wstring mutexName = BuildMutexName(normalizedOutputPath);
	mutex_ = CreateMutexW(nullptr, FALSE, mutexName.c_str());
	if (mutex_ == nullptr) {
		outError = std::format("CreateMutexW failed: {}", GetLastError());
		return false;
	}

	const auto waitStartedAt = std::chrono::steady_clock::now();
	const DWORD waitResult = WaitForSingleObject(mutex_, timeoutMilliseconds);
	waitElapsedMilliseconds_ = static_cast<DWORD>(
		std::chrono::duration_cast<std::chrono::milliseconds>(
			std::chrono::steady_clock::now() - waitStartedAt).count());
	if (waitResult == WAIT_OBJECT_0 || waitResult == WAIT_ABANDONED) {
		acquired_ = true;
		return true;
	}

	if (waitResult == WAIT_TIMEOUT) {
		outError = "wait compile output lock timed out";
	}
	else {
		outError = std::format("WaitForSingleObject failed: {}", GetLastError());
	}
	CloseHandle(mutex_);
	mutex_ = nullptr;
	return false;
}

void OutputPathLock::Release()
{
	if (mutex_ == nullptr) {
		return;
	}
	if (acquired_) {
		ReleaseMutex(mutex_);
	}
	CloseHandle(mutex_);
	mutex_ = nullptr;
	acquired_ = false;
	waitElapsedMilliseconds_ = 0;
}

DWORD OutputPathLock::WaitElapsedMilliseconds() const
{
	return waitElapsedMilliseconds_;
}

std::string BuildSelfTestJson()
{
	const std::string sharedPath = "C:\\AutoLinkerSelfTest\\shared-output.exe";
	const std::string otherPath = "C:\\AutoLinkerSelfTest\\other-output.exe";

	ProcessExecutionGuard firstProcessGuard;
	ProcessExecutionGuard nestedProcessGuard;
	const bool processReentryBlocked = firstProcessGuard.Acquired() && !nestedProcessGuard.Acquired();

	OutputPathLock firstOutputLock;
	std::string firstError;
	const bool firstOutputAcquired = firstOutputLock.Acquire(sharedPath, 1000, firstError);

	bool samePathTimedOut = false;
	std::string samePathError;
	std::thread contender([&]() {
		OutputPathLock secondOutputLock;
		samePathTimedOut = !secondOutputLock.Acquire(sharedPath, 50, samePathError) &&
			samePathError == "wait compile output lock timed out";
	});
	contender.join();

	OutputPathLock otherOutputLock;
	std::string otherError;
	const bool differentPathAcquired = otherOutputLock.Acquire(otherPath, 1000, otherError);
	firstOutputLock.Release();

	OutputPathLock retryOutputLock;
	std::string retryError;
	const bool samePathAcquiredAfterRelease = retryOutputLock.Acquire(sharedPath, 1000, retryError);

	const bool ok = processReentryBlocked &&
		firstOutputAcquired &&
		samePathTimedOut &&
		differentPathAcquired &&
		samePathAcquiredAfterRelease;
	return nlohmann::json({
		{"name", "compile-concurrency-manager"},
		{"ok", ok},
		{"process_reentry_blocked", processReentryBlocked},
		{"same_path_serialized", samePathTimedOut && samePathAcquiredAfterRelease},
		{"different_path_independent", differentPathAcquired},
		{"first_error", firstError},
		{"same_path_error", samePathError},
		{"other_error", otherError},
		{"retry_error", retryError}
	}).dump();
}

} // namespace CompileConcurrencyManager
