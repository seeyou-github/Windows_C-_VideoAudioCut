#include "FFmpegRunner.h"

#include <windows.h>

#include <vector>

namespace {

std::wstring BuildCommandLine(const std::wstring& executable, const std::wstring& arguments) {
    return L"\"" + executable + L"\" " + arguments;
}

}  // namespace

FFmpegRunner::FFmpegRunner() : running_(false) {
}

FFmpegRunner::~FFmpegRunner() {
    if (worker_.joinable()) {
        worker_.join();
    }
}

bool FFmpegRunner::IsRunning() const {
    return running_.load();
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
    SECURITY_ATTRIBUTES securityAttributes{};
    securityAttributes.nLength = sizeof(securityAttributes);
    securityAttributes.bInheritHandle = TRUE;

    HANDLE readPipe = nullptr;
    HANDLE writePipe = nullptr;

    if (!::CreatePipe(&readPipe, &writePipe, &securityAttributes, 0)) {
        running_.store(false);
        finishCallback(-1);
        return;
    }

    ::SetHandleInformation(readPipe, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW startupInfo{};
    startupInfo.cb = sizeof(startupInfo);
    startupInfo.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    startupInfo.wShowWindow = SW_HIDE;
    startupInfo.hStdOutput = writePipe;
    startupInfo.hStdError = writePipe;

    PROCESS_INFORMATION processInfo{};
    std::wstring commandLine = BuildCommandLine(ffmpegPath, arguments);
    std::vector<wchar_t> mutableCommand(commandLine.begin(), commandLine.end());
    mutableCommand.push_back(L'\0');

    const BOOL created = ::CreateProcessW(
        nullptr,
        mutableCommand.data(),
        nullptr,
        nullptr,
        TRUE,
        CREATE_NO_WINDOW,
        nullptr,
        nullptr,
        &startupInfo,
        &processInfo);

    ::CloseHandle(writePipe);

    if (!created) {
        ::CloseHandle(readPipe);
        running_.store(false);
        finishCallback(-1);
        return;
    }

    char buffer[512];
    DWORD bytesRead = 0;
    std::string textBuffer;

    while (::ReadFile(readPipe, buffer, sizeof(buffer) - 1, &bytesRead, nullptr) && bytesRead > 0) {
        buffer[bytesRead] = '\0';
        textBuffer.append(buffer, bytesRead);

        size_t newlinePos = 0;
        while ((newlinePos = textBuffer.find('\n')) != std::string::npos) {
            std::string line = textBuffer.substr(0, newlinePos + 1);
            textBuffer.erase(0, newlinePos + 1);

            const int utf8Length = ::MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, nullptr, 0);
            if (utf8Length > 0) {
                std::wstring wideLine(utf8Length - 1, L'\0');
                ::MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, wideLine.data(), utf8Length);
                logCallback(wideLine);
            } else {
                const int ansiLength = ::MultiByteToWideChar(CP_ACP, 0, line.c_str(), -1, nullptr, 0);
                std::wstring ansiLine(ansiLength - 1, L'\0');
                ::MultiByteToWideChar(CP_ACP, 0, line.c_str(), -1, ansiLine.data(), ansiLength);
                logCallback(ansiLine);
            }
        }
    }

    if (!textBuffer.empty()) {
        const int ansiLength = ::MultiByteToWideChar(CP_ACP, 0, textBuffer.c_str(), -1, nullptr, 0);
        if (ansiLength > 0) {
            std::wstring ansiLine(ansiLength - 1, L'\0');
            ::MultiByteToWideChar(CP_ACP, 0, textBuffer.c_str(), -1, ansiLine.data(), ansiLength);
            logCallback(ansiLine);
        }
    }

    ::WaitForSingleObject(processInfo.hProcess, INFINITE);

    DWORD exitCode = 0;
    ::GetExitCodeProcess(processInfo.hProcess, &exitCode);

    ::CloseHandle(processInfo.hThread);
    ::CloseHandle(processInfo.hProcess);
    ::CloseHandle(readPipe);

    running_.store(false);
    finishCallback(static_cast<int>(exitCode));
}
