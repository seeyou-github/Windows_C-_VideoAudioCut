#include "StringResources.h"

#include <windows.h>

std::wstring LoadStringResource(unsigned int id) {
    wchar_t buffer[512] = {};
    const int length = ::LoadStringW(::GetModuleHandleW(nullptr), id, buffer, static_cast<int>(std::size(buffer)));
    return std::wstring(buffer, buffer + length);
}
