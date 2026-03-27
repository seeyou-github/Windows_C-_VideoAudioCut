#pragma once

#include <string>

struct AppConfig {
    std::wstring ffmpegPath;
    std::wstring ffprobePath;
    std::wstring cutSuffix = L"_cut";
    bool saveToSourceFolder = true;
    int windowWidth = 980;
    int windowHeight = 720;
};

class ConfigManager {
public:
    ConfigManager();

    bool Load(AppConfig& config) const;
    bool SaveIfChanged(const AppConfig& config) const;
    std::wstring GetConfigPath() const;
    std::wstring GetExecutableDirectory() const;

private:
    std::wstring configPath_;
};
