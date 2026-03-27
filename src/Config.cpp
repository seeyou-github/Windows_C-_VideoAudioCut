#include "Config.h"

#include <windows.h>

#include <filesystem>
#include <fstream>

namespace {

std::wstring Trim(const std::wstring& value) {
    const auto first = value.find_first_not_of(L" \t\r\n");
    if (first == std::wstring::npos) {
        return L"";
    }

    const auto last = value.find_last_not_of(L" \t\r\n");
    return value.substr(first, last - first + 1);
}

std::wstring GetModuleDirectory() {
    wchar_t path[MAX_PATH] = {};
    ::GetModuleFileNameW(nullptr, path, MAX_PATH);

    std::wstring fullPath(path);
    const auto pos = fullPath.find_last_of(L"\\/");
    return pos == std::wstring::npos ? L"." : fullPath.substr(0, pos);
}

}  // namespace

ConfigManager::ConfigManager() : configPath_(GetModuleDirectory() + L"\\VideoAudioCut.ini") {
}

bool ConfigManager::Load(AppConfig& config) const {
    config = {};

    std::wifstream input(configPath_.c_str());
    if (!input.is_open()) {
        return true;
    }

    std::wstring line;
    while (std::getline(input, line)) {
        const auto pos = line.find(L'=');
        if (pos == std::wstring::npos) {
            continue;
        }

        const std::wstring key = Trim(line.substr(0, pos));
        const std::wstring value = Trim(line.substr(pos + 1));
        if (key == L"ffmpeg_path") {
            config.ffmpegPath = value;
        } else if (key == L"ffprobe_path") {
            config.ffprobePath = value;
        } else if (key == L"cut_suffix") {
            config.cutSuffix = value;
        } else if (key == L"save_to_source_folder") {
            config.saveToSourceFolder = value != L"0" && value != L"false" && value != L"False";
        } else if (key == L"window_width") {
            config.windowWidth = _wtoi(value.c_str());
        } else if (key == L"window_height") {
            config.windowHeight = _wtoi(value.c_str());
        }
    }

    if (config.windowWidth < 760) {
        config.windowWidth = 980;
    }
    if (config.windowHeight < 560) {
        config.windowHeight = 720;
    }
    if (config.ffprobePath.empty() && !config.ffmpegPath.empty()) {
        const std::filesystem::path ffmpegPath(config.ffmpegPath);
        config.ffprobePath = (ffmpegPath.parent_path() / L"ffprobe.exe").wstring();
    }
    if (config.cutSuffix.empty()) {
        config.cutSuffix = L"_cut";
    }

    return true;
}

bool ConfigManager::SaveIfChanged(const AppConfig& config) const {
    AppConfig current;
    Load(current);
    if (current.ffmpegPath == config.ffmpegPath &&
        current.ffprobePath == config.ffprobePath &&
        current.cutSuffix == config.cutSuffix &&
        current.saveToSourceFolder == config.saveToSourceFolder &&
        current.windowWidth == config.windowWidth &&
        current.windowHeight == config.windowHeight) {
        return true;
    }

    std::wofstream output(configPath_.c_str(), std::ios::trunc);
    if (!output.is_open()) {
        return false;
    }

    output << L"ffmpeg_path=" << config.ffmpegPath << L"\n";
    output << L"ffprobe_path=" << config.ffprobePath << L"\n";
    output << L"cut_suffix=" << config.cutSuffix << L"\n";
    output << L"save_to_source_folder=" << (config.saveToSourceFolder ? 1 : 0) << L"\n";
    output << L"window_width=" << config.windowWidth << L"\n";
    output << L"window_height=" << config.windowHeight << L"\n";
    return output.good();
}

std::wstring ConfigManager::GetConfigPath() const {
    return configPath_;
}

std::wstring ConfigManager::GetExecutableDirectory() const {
    return GetModuleDirectory();
}
