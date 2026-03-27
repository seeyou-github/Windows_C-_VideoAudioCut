#pragma once

#include "Config.h"
#include "FFmpegRunner.h"

#include <windows.h>

#include <atomic>
#include <functional>
#include <string>
#include <thread>
#include <vector>

class MainWindow {
public:
    MainWindow();
    ~MainWindow();

    bool Create(HINSTANCE instance, int showCommand);
    HWND Handle() const;

private:
    struct ControlSet {
        HWND sidebar = nullptr;
        HWND navCutButton = nullptr;
        HWND navConcatButton = nullptr;
        HWND navFadeButton = nullptr;
        HWND navConvertButton = nullptr;
        HWND settingsButton = nullptr;
        HWND contentTitle = nullptr;
        HWND cutGroup = nullptr;
        HWND inputLabel = nullptr;
        HWND inputEdit = nullptr;
        HWND inputBrowse = nullptr;
        HWND startLabel = nullptr;
        HWND startDecrease = nullptr;
        HWND startTrack = nullptr;
        HWND fileDurationValue = nullptr;
        HWND startValue = nullptr;
        HWND startIncrease = nullptr;
        HWND cutDurationValue = nullptr;
        HWND endLabel = nullptr;
        HWND endTrack = nullptr;
        HWND endDecrease = nullptr;
        HWND endValue = nullptr;
        HWND endIncrease = nullptr;
        HWND runButton = nullptr;
        HWND clearListButton = nullptr;
        HWND openFolderButton = nullptr;
        HWND concatList = nullptr;
        HWND convertModeVideoButton = nullptr;
        HWND convertModeAudioButton = nullptr;
        HWND convertHint = nullptr;
        HWND convertInfoLabelFile = nullptr;
        HWND convertInfoValueFile = nullptr;
        HWND convertInfoLabelFormat = nullptr;
        HWND convertInfoValueFormat = nullptr;
        HWND convertInfoLabelVideoCodec = nullptr;
        HWND convertInfoValueVideoCodec = nullptr;
        HWND convertInfoLabelVideoBitrate = nullptr;
        HWND convertInfoValueVideoBitrate = nullptr;
        HWND convertInfoLabelResolution = nullptr;
        HWND convertInfoValueResolution = nullptr;
        HWND convertInfoLabelAudioCodec = nullptr;
        HWND convertInfoValueAudioCodec = nullptr;
        HWND convertInfoLabelAudioBitrate = nullptr;
        HWND convertInfoValueAudioBitrate = nullptr;
        HWND convertToMp3Button = nullptr;
        HWND convertToMp4Button = nullptr;
        HWND fadeInTrack = nullptr;
        HWND fadeOutTrack = nullptr;
        HWND logEdit = nullptr;
        HWND placeholderStatic = nullptr;
    };

    struct RunLogMessage {
        std::wstring text;
    };

    struct DurationProbeResult {
        int requestId = 0;
        bool success = false;
        int durationMilliseconds = 0;
    };

    struct FadeItemStatusMessage {
        int index = -1;
        bool success = false;
        bool applyLastOutputPath = false;
        std::wstring outputPath;
        bool applyConvertItemResult = false;
        std::wstring convertResultText;
        bool convertHasError = false;
        std::wstring errorText;
    };

    struct ConcatListItem {
        std::wstring filePath;
        std::wstring fileName;
        std::wstring durationText;
        std::wstring resolutionText;
        std::wstring videoCodec;
        std::wstring videoBitrate;
        std::wstring audioCodec;
        std::wstring audioBitrate;
        std::wstring resultText;
        bool hasError = false;
    };

    struct MediaProbeItemMessage {
        int requestId = 0;
        int module = 0;
        int index = -1;
        bool singleItem = false;
        ConcatListItem item;
    };

    struct MediaProbeFinishedMessage {
        int requestId = 0;
        int module = 0;
    };

    bool RegisterWindowClass(HINSTANCE instance);
    bool RegisterRangeSliderClass(HINSTANCE instance) const;
    void CreateControls();
    void ApplyFonts();
    void LayoutControls(int width, int height);
    void RefreshText();
    void RefreshModule();
    void AppendLog(const std::wstring& line);
    void ClearLog();
    void OpenInputFileDialog();
    void OpenConcatFilesDialog();
    void OpenFadeFilesDialog();
    void OpenConvertFilesDialog();
    void StartMediaProbe(int module, const std::vector<std::wstring>& inputPaths, bool singleItem);
    bool PromptOutputFolder(std::wstring& folderPath);
    void OpenOutputFolder();
    void OpenSettingsDialog();
    void StartCutTask();
    void StartConcatTask();
    void StartFadeTask();
    void StartConvertTask();
    void StartConvertToMp3Task();
    void StartConvertToMp4Task();
    void AutoDetectInputDuration();
    void ResetDurationState();
    void UpdateFileDurationDisplay();
    void AdjustRangeValueByMilliseconds(bool adjustStart, int deltaMilliseconds);
    void ResetCutResultState();
    void UpdateTrackbarPositions();
    void UpdateTimeDisplays();
    void UpdateLogPanelState();
    void UpdatePrimaryActionState();
    void ConfigureMediaListColumns();
    void InitializeConcatList();
    void RefreshConcatList();
    void RefreshConvertInfo();
    void SetLogExpanded(bool expanded);
    void SpawnBackgroundTask(std::function<void()> task);
    std::wstring FormatMilliseconds(int totalMilliseconds) const;
    std::wstring FormatClockTextNoMilliseconds(int totalMilliseconds) const;
    std::wstring FormatBitrateText(long long bitrate) const;
    std::wstring BuildFadeOutputPath(const std::wstring& inputPath) const;
    std::wstring BuildConvertOutputPath(const std::wstring& inputPath) const;
    std::wstring BuildConvertMp3OutputPath(const std::wstring& inputPath) const;
    std::wstring BuildConvertMp4OutputPath(const std::wstring& inputPath) const;
    void SaveWindowPlacement();
    std::wstring BuildDefaultOutputPath(const std::wstring& inputPath) const;
    std::wstring BuildOutputPathInFolder(const std::wstring& inputPath, const std::wstring& folderPath) const;
    std::wstring BuildConcatOutputPath(const std::vector<std::wstring>& inputPaths) const;
    std::wstring BuildFfprobePath() const;
    bool TryProbeDurationMilliseconds(const std::wstring& inputPath, int& durationMilliseconds) const;
    bool BuildConcatListFile(const std::vector<std::wstring>& inputPaths,
                             std::wstring& listFilePath,
                             int& fileCount) const;
    bool TryProbeConcatItem(const std::wstring& inputPath, ConcatListItem& item) const;
    int RunProcessCapture(const std::wstring& executablePath,
                          const std::wstring& arguments,
                          const std::function<void(const std::wstring&)>& logCallback) const;
    bool ValidateFfmpegPath() const;
    void SetDarkTitleBar() const;
    void DrawOwnerDrawButton(const DRAWITEMSTRUCT* drawItem) const;
    void DrawSidebarItem(const DRAWITEMSTRUCT* drawItem) const;
    void PaintFrame(HDC hdc) const;
    void PaintRangeSlider(HDC hdc, const RECT& clientRect) const;
    LRESULT HandleRangeSliderMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK RangeSliderProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
    static std::wstring EscapeArgument(const std::wstring& value);
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK SettingsProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
    LRESULT HandleMessage(UINT message, WPARAM wParam, LPARAM lParam);

    HINSTANCE instance_;
    HWND hwnd_;
    HFONT uiFont_;
    HFONT navFont_;
    HFONT totalTimeFont_;
    HBRUSH backgroundBrush_;
    HBRUSH panelBrush_;
    HBRUSH editBrush_;
    HBRUSH buttonBrush_;
    HBRUSH navBrush_;
    HPEN borderPen_;
    HPEN accentPen_;
    COLORREF textColor_;
    COLORREF subtleTextColor_;
    COLORREF accentColor_;
    COLORREF accentHotColor_;
    COLORREF borderColor_;
    COLORREF buttonTextColor_;
    ConfigManager configManager_;
    AppConfig config_;
    FFmpegRunner runner_;
    ControlSet controls_;
    int selectedModule_;
    bool durationAvailable_;
    int mediaDurationMilliseconds_;
    int rangeStartMilliseconds_;
    int rangeEndMilliseconds_;
    int activeRangeThumb_;
    bool logExpanded_;
    bool cutSucceeded_;
    int durationProbeRequestId_;
    bool updatingTrackbars_;
    std::wstring inputFilePath_;
    std::vector<std::wstring> concatInputPaths_;
    std::vector<ConcatListItem> concatItems_;
    std::vector<std::wstring> fadeInputPaths_;
    std::vector<ConcatListItem> fadeItems_;
    std::vector<std::wstring> convertInputPaths_;
    std::vector<ConcatListItem> convertItems_;
    ConcatListItem convertItem_;
    std::wstring concatListFilePath_;
    std::wstring lastOutputPath_;
    int activeTaskModule_;
    int convertMode_;
    int fadeInMilliseconds_;
    int fadeOutMilliseconds_;
    bool taskRunning_;
    bool convertCanToMp3_;
    bool convertCanToMp4_;
    int mediaProbeRequestId_;
    std::atomic<bool> shuttingDown_;
    std::vector<std::thread> backgroundThreads_;
};
