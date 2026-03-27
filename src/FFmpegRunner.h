#pragma once

#include <windows.h>

#include <functional>
#include <atomic>
#include <string>
#include <thread>

class FFmpegRunner {
public:
    using LogCallback = std::function<void(const std::wstring&)>;
    using FinishCallback = std::function<void(int)>;

    FFmpegRunner();
    ~FFmpegRunner();

    bool IsRunning() const;
    void RequestStop();
    bool Start(const std::wstring& ffmpegPath,
               const std::wstring& arguments,
               LogCallback logCallback,
               FinishCallback finishCallback);

private:
    void RunProcess(std::wstring ffmpegPath,
                    std::wstring arguments,
                    LogCallback logCallback,
                    FinishCallback finishCallback);

    std::thread worker_;
    std::atomic<bool> running_;
    std::atomic<HANDLE> processHandle_;
};
