#include "ProcessUtils.h"

#include <vector>

namespace ProcessUtils {

std::wstring QuoteCommandLineArgument(const std::wstring& value) {
    if (value.empty()) {
        return L"\"\"";
    }

    const bool needsQuotes = value.find_first_of(L" \t\n\v\"") != std::wstring::npos;
    if (!needsQuotes) {
        return value;
    }

    std::wstring quoted;
    quoted.push_back(L'"');
    size_t backslashCount = 0;
    for (wchar_t ch : value) {
        if (ch == L'\\') {
            ++backslashCount;
            continue;
        }
        if (ch == L'"') {
            quoted.append(backslashCount * 2 + 1, L'\\');
            quoted.push_back(L'"');
            backslashCount = 0;
            continue;
        }
        if (backslashCount > 0) {
            quoted.append(backslashCount, L'\\');
            backslashCount = 0;
        }
        quoted.push_back(ch);
    }
    if (backslashCount > 0) {
        quoted.append(backslashCount * 2, L'\\');
    }
    quoted.push_back(L'"');
    return quoted;
}

std::wstring BuildCommandLine(const std::wstring& executable, const std::wstring& arguments) {
    return QuoteCommandLineArgument(executable) + L" " + arguments;
}

std::wstring DecodeProcessText(const std::string& text) {
    auto decode = [&](UINT codePage) -> std::wstring {
        const int length = ::MultiByteToWideChar(codePage, 0, text.c_str(), -1, nullptr, 0);
        if (length <= 0) {
            return L"";
        }
        std::wstring wide(length - 1, L'\0');
        ::MultiByteToWideChar(codePage, 0, text.c_str(), -1, wide.data(), length);
        return wide;
    };

    std::wstring utf8 = decode(CP_UTF8);
    if (!utf8.empty()) {
        return utf8;
    }
    return decode(CP_ACP);
}

bool CreateRedirectedProcess(const std::wstring& executable,
                             const std::wstring& arguments,
                             PROCESS_INFORMATION& processInfo,
                             HANDLE& readPipe) {
    readPipe = nullptr;

    SECURITY_ATTRIBUTES securityAttributes{};
    securityAttributes.nLength = sizeof(securityAttributes);
    securityAttributes.bInheritHandle = TRUE;

    HANDLE writePipe = nullptr;
    if (!::CreatePipe(&readPipe, &writePipe, &securityAttributes, 0)) {
        return false;
    }
    ::SetHandleInformation(readPipe, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW startupInfo{};
    startupInfo.cb = sizeof(startupInfo);
    startupInfo.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    startupInfo.wShowWindow = SW_HIDE;
    startupInfo.hStdOutput = writePipe;
    startupInfo.hStdError = writePipe;

    std::wstring commandLine = BuildCommandLine(executable, arguments);
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
        readPipe = nullptr;
        return false;
    }
    return true;
}

void DrainProcessOutput(HANDLE readPipe, const std::function<void(const std::wstring&)>& onLine) {
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

            std::wstring wideLine = DecodeProcessText(line);
            if (!wideLine.empty()) {
                onLine(wideLine);
            }
        }
    }

    if (!textBuffer.empty()) {
        std::wstring wideLine = DecodeProcessText(textBuffer);
        if (!wideLine.empty()) {
            onLine(wideLine);
        }
    }
}

}  // namespace ProcessUtils
