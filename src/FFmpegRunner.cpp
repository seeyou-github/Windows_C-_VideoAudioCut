#include "FFmpegRunner.h"

#include "ProcessUtils.h"

FFmpegRunner::FFmpegRunner() : running_(false), processHandle_(nullptr) {
}

FFmpegRunner::~FFmpegRunner() {
    RequestStop();
    if (worker_.joinable()) {
        worker_.join();
    }
}

bool FFmpegRunner::IsRunning() const {
    return running_.load();
}

void FFmpegRunner::RequestStop() {
    HANDLE processHandle = processHandle_.load();
    if (processHandle != nullptr) {
        ::TerminateProcess(processHandle, 1);
    }
}

bool FFmpegRunner::Start(const std::wstring& ffmpegPath,
                         const std::wstring& arguments,
                         LogCallback logCallback,
                         FinishCallback finishCallback) {
    if (running_.load()) {
        return false;
    }

    if (worker_.joinable()) {
        worker_.join();
    }

    running_.store(true);
    worker_ = std::thread(&FFmpegRunner::RunProcess, this, ffmpegPath, arguments, logCallback, finishCallback);
    return true;
}

void FFmpegRunner::RunProcess(std::wstring ffmpegPath,
                              std::wstring arguments,
                              LogCallback logCallback,
                              FinishCallback finishCallback) {
    HANDLE readPipe = nullptr;
    PROCESS_INFORMATION processInfo{};
    if (!ProcessUtils::CreateRedirectedProcess(ffmpegPath, arguments, processInfo, readPipe)) {
        running_.store(false);
        finishCallback(-1);
        return;
    }
    processHandle_.store(processInfo.hProcess);

    ProcessUtils::DrainProcessOutput(readPipe, logCallback);

    ::WaitForSingleObject(processInfo.hProcess, INFINITE);

    DWORD exitCode = 0;
    ::GetExitCodeProcess(processInfo.hProcess, &exitCode);

    processHandle_.store(nullptr);
    ::CloseHandle(processInfo.hThread);
    ::CloseHandle(processInfo.hProcess);
    ::CloseHandle(readPipe);

    running_.store(false);
    finishCallback(static_cast<int>(exitCode));
}
