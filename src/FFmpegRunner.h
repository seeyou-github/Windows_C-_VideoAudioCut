#pragma once

#include <functional>
#include <string>
#include <thread>

class FFmpegRunner {
public:
    using LogCallback = std::function<void(const std::wstring&)>;
    using FinishCallback = std::function<void(int)>;

    FFmpegRunner();
    ~FFmpegRunner();

    bool IsRunning() const;
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
    bool running_;
};
