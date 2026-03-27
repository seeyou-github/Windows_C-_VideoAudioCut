#pragma once

#include <windows.h>

#include <functional>
#include <string>

namespace ProcessUtils {

std::wstring QuoteCommandLineArgument(const std::wstring& value);
std::wstring BuildCommandLine(const std::wstring& executable, const std::wstring& arguments);
std::wstring DecodeProcessText(const std::string& text);
bool CreateRedirectedProcess(const std::wstring& executable,
                             const std::wstring& arguments,
                             PROCESS_INFORMATION& processInfo,
                             HANDLE& readPipe);
void DrainProcessOutput(HANDLE readPipe, const std::function<void(const std::wstring&)>& onLine);

}  // namespace ProcessUtils
