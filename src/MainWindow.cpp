#include "MainWindow.h"

#include "AppIds.h"
#include "ProcessUtils.h"
#include "ResourceIds.h"
#include "StringResources.h"

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <commdlg.h>
#include <dwmapi.h>
#include <shellapi.h>
#include <shlobj.h>

#include <algorithm>
#include <array>
#include <cstdlib>
#include <cstdio>
#include <cwchar>
#include <cwctype>
#include <filesystem>
#include <memory>
#include <fstream>
#include <sstream>
#include <thread>
#include <vector>

#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif

namespace {

struct RunLogBatchMessage {
    std::vector<std::wstring> lines;
};

constexpr wchar_t kWindowClassName[] = L"VideoAudioCutMainWindow";
constexpr wchar_t kRangeSliderClassName[] = L"VideoAudioCutRangeSlider";
constexpr int kMaxLogChars = 64 * 1024;
constexpr int kTrimLogChars = 16 * 1024;
constexpr int kSidebarWidth = 228;
constexpr int kMargin = 24;
constexpr int kNavPadding = 14;
constexpr int kNavItemHeight = 62;
constexpr int kMinimumWindowWidth = 1440;
constexpr int kMinimumWindowHeight = 960;
constexpr int kLabelHeight = 26;
constexpr int kEditHeight = 60;
constexpr int kButtonHeight = 40;
constexpr int kTrackMin = 0;
constexpr int kTrackMax = 24 * 60 * 60 * 1000;
constexpr int kRangeSliderHeight = 96;
constexpr int kRangeTrackHeight = 6;
constexpr int kRangeThumbRadius = 9;
constexpr int kRangeTrackPadding = 18;
constexpr int kRangeThumbNone = 0;
constexpr int kRangeThumbStart = 1;
constexpr int kRangeThumbEnd = 2;
constexpr int kLogCollapsedHeight = 0;
constexpr int kLogExpandedHeight = 110;

std::wstring GetWindowTextString(HWND hwnd) {
    const int length = ::GetWindowTextLengthW(hwnd);
    std::wstring value(length + 1, L'\0');
    const int copied = ::GetWindowTextW(hwnd, value.data(), length + 1);
    value.resize(std::max(0, copied));
    return value;
}

template <typename MessageType>
bool PostOwnedMessage(HWND hwnd, UINT message, WPARAM wParam, std::unique_ptr<MessageType> payload) {
    if (hwnd == nullptr || payload == nullptr || !::IsWindow(hwnd)) {
        return false;
    }
    MessageType* raw = payload.release();
    if (!::PostMessageW(hwnd, message, wParam, reinterpret_cast<LPARAM>(raw))) {
        delete raw;
        return false;
    }
    return true;
}

constexpr size_t kLogBatchSize = 16;

void FlushPendingLogLines(HWND hwnd, std::vector<std::wstring>& pendingLines) {
    if (pendingLines.empty()) {
        return;
    }
    auto batch = std::make_unique<RunLogBatchMessage>();
    batch->lines.swap(pendingLines);
    PostOwnedMessage(hwnd, WM_APP_RUN_LOG_BATCH, 0, std::move(batch));
}

void QueueLogLine(HWND hwnd, std::vector<std::wstring>& pendingLines, const std::wstring& line) {
    pendingLines.push_back(line);
    if (pendingLines.size() >= kLogBatchSize) {
        FlushPendingLogLines(hwnd, pendingLines);
    }
}

bool FileExists(const std::wstring& path) {
    const DWORD attributes = ::GetFileAttributesW(path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

void CenterWindowToScreen(HWND hwnd, int width, int height) {
    const int screenWidth = ::GetSystemMetrics(SM_CXSCREEN);
    const int screenHeight = ::GetSystemMetrics(SM_CYSCREEN);
    const int x = (screenWidth - width) / 2;
    const int y = (screenHeight - height) / 2;
    ::MoveWindow(hwnd, x, y, width, height, FALSE);
}

}  // namespace

MainWindow::MainWindow()
    : instance_(nullptr),
      hwnd_(nullptr),
      uiFont_(nullptr),
      navFont_(nullptr),
      totalTimeFont_(nullptr),
      backgroundBrush_(::CreateSolidBrush(RGB(34, 36, 40))),
      panelBrush_(::CreateSolidBrush(RGB(42, 45, 50))),
      editBrush_(::CreateSolidBrush(RGB(50, 53, 58))),
      buttonBrush_(::CreateSolidBrush(RGB(56, 60, 66))),
      navBrush_(::CreateSolidBrush(RGB(46, 49, 54))),
      borderPen_(::CreatePen(PS_SOLID, 1, RGB(72, 76, 82))),
      accentPen_(::CreatePen(PS_SOLID, 1, RGB(110, 136, 166))),
      textColor_(RGB(226, 229, 233)),
      subtleTextColor_(RGB(188, 193, 199)),
      accentColor_(RGB(110, 136, 166)),
      accentHotColor_(RGB(142, 164, 190)),
      borderColor_(RGB(72, 76, 82)),
      buttonTextColor_(RGB(226, 229, 233)),
      selectedModule_(0),
      durationAvailable_(false),
      mediaDurationMilliseconds_(0),
      rangeStartMilliseconds_(0),
      rangeEndMilliseconds_(0),
      activeRangeThumb_(kRangeThumbNone),
      logExpanded_(true),
      cutSucceeded_(false),
      durationProbeRequestId_(0),
      updatingTrackbars_(false),
      inputFilePath_(L""),
      concatInputPaths_(),
      concatListFilePath_(L""),
      lastOutputPath_(L""),
      activeTaskModule_(-1),
      convertMode_(0),
      fadeInMilliseconds_(6000),
      fadeOutMilliseconds_(6000),
      taskRunning_(false),
      convertCanToMp3_(false),
      convertCanToMp4_(false),
      mediaProbeRequestId_(0),
      shuttingDown_(false),
      cancelBackgroundWork_(false),
      backgroundThreads_(),
      backgroundProcessMutex_(),
      backgroundProcesses_() {
}

MainWindow::~MainWindow() {
    shuttingDown_.store(true);
    cancelBackgroundWork_.store(true);
    runner_.RequestStop();
    RequestBackgroundStop();
    for (std::thread& worker : backgroundThreads_) {
        if (worker.joinable()) {
            worker.join();
        }
    }
    if (uiFont_ != nullptr && uiFont_ != ::GetStockObject(DEFAULT_GUI_FONT)) {
        ::DeleteObject(uiFont_);
    }
    if (navFont_ != nullptr && navFont_ != ::GetStockObject(DEFAULT_GUI_FONT)) {
        ::DeleteObject(navFont_);
    }
    if (totalTimeFont_ != nullptr && totalTimeFont_ != ::GetStockObject(DEFAULT_GUI_FONT)) {
        ::DeleteObject(totalTimeFont_);
    }

    ::DeleteObject(backgroundBrush_);
    ::DeleteObject(panelBrush_);
    ::DeleteObject(editBrush_);
    ::DeleteObject(buttonBrush_);
    ::DeleteObject(navBrush_);
    ::DeleteObject(borderPen_);
    ::DeleteObject(accentPen_);
}

bool MainWindow::Create(HINSTANCE instance, int showCommand) {
    instance_ = instance;
    configManager_.Load(config_);

    INITCOMMONCONTROLSEX controlsEx{};
    controlsEx.dwSize = sizeof(controlsEx);
    controlsEx.dwICC = ICC_STANDARD_CLASSES | ICC_BAR_CLASSES;
    ::InitCommonControlsEx(&controlsEx);

    if (!RegisterWindowClass(instance)) {
        return false;
    }
    if (!RegisterRangeSliderClass(instance)) {
        return false;
    }

    hwnd_ = ::CreateWindowExW(
        0,
        kWindowClassName,
        LoadStringResource(IDS_APP_TITLE).c_str(),
        (WS_OVERLAPPEDWINDOW & ~WS_MAXIMIZEBOX),
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        config_.windowWidth,
        config_.windowHeight,
        nullptr,
        nullptr,
        instance,
        this);

    if (hwnd_ == nullptr) {
        return false;
    }

    if (HICON largeIcon = static_cast<HICON>(::LoadImageW(instance, MAKEINTRESOURCEW(ID_APP_ICON), IMAGE_ICON,
                                                          ::GetSystemMetrics(SM_CXICON), ::GetSystemMetrics(SM_CYICON), LR_SHARED))) {
        ::SendMessageW(hwnd_, WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(largeIcon));
    }
    if (HICON smallIcon = static_cast<HICON>(::LoadImageW(instance, MAKEINTRESOURCEW(ID_APP_ICON), IMAGE_ICON,
                                                          ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), LR_SHARED))) {
        ::SendMessageW(hwnd_, WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(smallIcon));
    }

    CenterWindowToScreen(hwnd_, config_.windowWidth, config_.windowHeight);
    SetDarkTitleBar();
    ::ShowWindow(hwnd_, showCommand);
    ::UpdateWindow(hwnd_);
    return true;
}

HWND MainWindow::Handle() const {
    return hwnd_;
}

bool MainWindow::RegisterWindowClass(HINSTANCE instance) {
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = MainWindow::WindowProc;
    wc.hInstance = instance;
    wc.lpszClassName = kWindowClassName;
    wc.hCursor = ::LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = backgroundBrush_;
    wc.hIcon = static_cast<HICON>(::LoadImageW(instance, MAKEINTRESOURCEW(ID_APP_ICON), IMAGE_ICON,
                                               ::GetSystemMetrics(SM_CXICON), ::GetSystemMetrics(SM_CYICON), LR_SHARED));
    wc.hIconSm = static_cast<HICON>(::LoadImageW(instance, MAKEINTRESOURCEW(ID_APP_ICON), IMAGE_ICON,
                                                 ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), LR_SHARED));
    return ::RegisterClassExW(&wc) != 0;
}

bool MainWindow::RegisterRangeSliderClass(HINSTANCE instance) const {
    WNDCLASSW wc{};
    wc.lpfnWndProc = MainWindow::RangeSliderProc;
    wc.hInstance = instance;
    wc.lpszClassName = kRangeSliderClassName;
    wc.hCursor = ::LoadCursorW(nullptr, IDC_HAND);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
    const ATOM result = ::RegisterClassW(&wc);
    return result != 0 || ::GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

void MainWindow::CreateControls() {
    controls_.sidebar = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_SIDEBAR), instance_, nullptr);

    controls_.navCutButton = ::CreateWindowExW(
        0, WC_BUTTONW, nullptr,
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_NAV_CUT), instance_, nullptr);

    controls_.navConcatButton = ::CreateWindowExW(
        0, WC_BUTTONW, nullptr,
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_NAV_CONCAT), instance_, nullptr);

    controls_.navFadeButton = ::CreateWindowExW(
        0, WC_BUTTONW, nullptr,
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_NAV_FADE), instance_, nullptr);

    controls_.navConvertButton = ::CreateWindowExW(
        0, WC_BUTTONW, nullptr,
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_NAV_CONVERT), instance_, nullptr);

    controls_.settingsButton = ::CreateWindowExW(
        0, WC_BUTTONW, nullptr,
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_SETTINGS_BUTTON), instance_, nullptr);

    controls_.contentTitle = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONTENT_TITLE), instance_, nullptr);

    controls_.cutGroup = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_GROUP), instance_, nullptr);

    controls_.inputLabel = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_INPUT_LABEL), instance_, nullptr);

    controls_.inputEdit = ::CreateWindowExW(
        WS_EX_CLIENTEDGE, WC_EDITW, nullptr,
        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_READONLY,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_INPUT_EDIT), instance_, nullptr);

    controls_.inputBrowse = ::CreateWindowExW(
        0, WC_BUTTONW, nullptr,
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_INPUT_BROWSE), instance_, nullptr);

    controls_.startLabel = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_START_LABEL), instance_, nullptr);

    controls_.startDecrease = ::CreateWindowExW(
        0, WC_BUTTONW, L"-",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_START_DECREASE), instance_, nullptr);

    controls_.startTrack = ::CreateWindowExW(
        0, kRangeSliderClassName, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_START_TRACK), instance_, this);

    controls_.fileDurationValue = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE | SS_RIGHT | SS_CENTERIMAGE,
        0, 0, 0, 0, hwnd_, nullptr, instance_, nullptr);

    controls_.startValue = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_START_VALUE), instance_, nullptr);

    controls_.startIncrease = ::CreateWindowExW(
        0, WC_BUTTONW, L"+",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_START_INCREASE), instance_, nullptr);

    controls_.cutDurationValue = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE,
        0, 0, 0, 0, hwnd_, nullptr, instance_, nullptr);

    controls_.endLabel = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_END_LABEL), instance_, nullptr);

    controls_.endTrack = nullptr;

    controls_.endDecrease = ::CreateWindowExW(
        0, WC_BUTTONW, L"-",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_END_DECREASE), instance_, nullptr);

    controls_.endValue = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_END_VALUE), instance_, nullptr);

    controls_.endIncrease = ::CreateWindowExW(
        0, WC_BUTTONW, L"+",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_END_INCREASE), instance_, nullptr);

    controls_.runButton = ::CreateWindowExW(
        0, WC_BUTTONW, nullptr,
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_RUN_BUTTON), instance_, nullptr);

    controls_.clearListButton = ::CreateWindowExW(
        0, WC_BUTTONW, L"清空列表",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONCAT_CLEAR_BUTTON), instance_, nullptr);

    controls_.openFolderButton = ::CreateWindowExW(
        0, WC_BUTTONW, nullptr,
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_OPEN_FOLDER_BUTTON), instance_, nullptr);

    controls_.concatList = ::CreateWindowExW(
        WS_EX_CLIENTEDGE, WC_LISTVIEWW, nullptr,
        WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONCAT_LIST), instance_, nullptr);

    controls_.convertModeVideoButton = ::CreateWindowExW(
        0, WC_BUTTONW, L"视频转 MP4",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_MODE_VIDEO), instance_, nullptr);

    controls_.convertModeAudioButton = ::CreateWindowExW(
        0, WC_BUTTONW, L"M4A 转 MP3",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_MODE_AUDIO), instance_, nullptr);

    controls_.convertHint = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_HINT), instance_, nullptr);

    controls_.convertInfoLabelFile = ::CreateWindowExW(
        0, WC_STATICW, L"文件名",
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_LABEL_FILE), instance_, nullptr);
    controls_.convertInfoValueFile = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_VALUE_FILE), instance_, nullptr);
    controls_.convertInfoLabelFormat = ::CreateWindowExW(
        0, WC_STATICW, L"格式",
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_LABEL_FORMAT), instance_, nullptr);
    controls_.convertInfoValueFormat = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_VALUE_FORMAT), instance_, nullptr);
    controls_.convertInfoLabelVideoCodec = ::CreateWindowExW(
        0, WC_STATICW, L"视频编码",
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_LABEL_VCODEC), instance_, nullptr);
    controls_.convertInfoValueVideoCodec = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_VALUE_VCODEC), instance_, nullptr);
    controls_.convertInfoLabelVideoBitrate = ::CreateWindowExW(
        0, WC_STATICW, L"视频码率",
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_LABEL_VBITRATE), instance_, nullptr);
    controls_.convertInfoValueVideoBitrate = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_VALUE_VBITRATE), instance_, nullptr);
    controls_.convertInfoLabelResolution = ::CreateWindowExW(
        0, WC_STATICW, L"分辨率",
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_LABEL_RESOLUTION), instance_, nullptr);
    controls_.convertInfoValueResolution = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_VALUE_RESOLUTION), instance_, nullptr);
    controls_.convertInfoLabelAudioCodec = ::CreateWindowExW(
        0, WC_STATICW, L"音频编码",
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_LABEL_ACODEC), instance_, nullptr);
    controls_.convertInfoValueAudioCodec = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_VALUE_ACODEC), instance_, nullptr);
    controls_.convertInfoLabelAudioBitrate = ::CreateWindowExW(
        0, WC_STATICW, L"音频码率",
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_LABEL_ABITRATE), instance_, nullptr);
    controls_.convertInfoValueAudioBitrate = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_INFO_VALUE_ABITRATE), instance_, nullptr);
    controls_.convertToMp3Button = ::CreateWindowExW(
        0, WC_BUTTONW, L"转为 MP3",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_TO_MP3_BUTTON), instance_, nullptr);
    controls_.convertToMp4Button = ::CreateWindowExW(
        0, WC_BUTTONW, L"转为 MP4",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CONVERT_TO_MP4_BUTTON), instance_, nullptr);

    controls_.fadeInTrack = ::CreateWindowExW(
        0, TRACKBAR_CLASSW, nullptr,
        WS_CHILD | TBS_AUTOTICKS,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_FADE_IN_TRACK), instance_, nullptr);

    controls_.fadeOutTrack = ::CreateWindowExW(
        0, TRACKBAR_CLASSW, nullptr,
        WS_CHILD | TBS_AUTOTICKS,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_FADE_OUT_TRACK), instance_, nullptr);

    controls_.logEdit = ::CreateWindowExW(
        WS_EX_CLIENTEDGE, WC_EDITW, nullptr,
        WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
        0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(IDC_CUT_LOG_EDIT), instance_, nullptr);

    controls_.placeholderStatic = ::CreateWindowExW(
        0, WC_STATICW, nullptr,
        WS_CHILD,
        0, 0, 0, 0, hwnd_, nullptr, instance_, nullptr);

    ::SendMessageW(controls_.logEdit, EM_SETLIMITTEXT, kMaxLogChars * 2, 0);
    ::SendMessageW(controls_.fadeInTrack, TBM_SETRANGE, TRUE, MAKELPARAM(0, 30));
    ::SendMessageW(controls_.fadeInTrack, TBM_SETPAGESIZE, 0, 1);
    ::SendMessageW(controls_.fadeInTrack, TBM_SETTICFREQ, 1, 0);
    ::SendMessageW(controls_.fadeInTrack, TBM_SETPOS, TRUE, fadeInMilliseconds_ / 1000);
    ::SendMessageW(controls_.fadeOutTrack, TBM_SETRANGE, TRUE, MAKELPARAM(0, 30));
    ::SendMessageW(controls_.fadeOutTrack, TBM_SETPAGESIZE, 0, 1);
    ::SendMessageW(controls_.fadeOutTrack, TBM_SETTICFREQ, 1, 0);
    ::SendMessageW(controls_.fadeOutTrack, TBM_SETPOS, TRUE, fadeOutMilliseconds_ / 1000);
    InitializeConcatList();

    ApplyFonts();
    RefreshText();
    ResetDurationState();
    UpdateTimeDisplays();
    UpdateLogPanelState();
    RefreshModule();
    UpdatePrimaryActionState();
}

void MainWindow::ApplyFonts() {
    LOGFONTW uiFont{};
    uiFont.lfHeight = -24;
    uiFont.lfWeight = FW_MEDIUM;
    lstrcpyW(uiFont.lfFaceName, L"Segoe UI");
    uiFont_ = ::CreateFontIndirectW(&uiFont);

    LOGFONTW navFont = uiFont;
    navFont.lfHeight = -26;
    navFont.lfWeight = FW_SEMIBOLD;
    navFont_ = ::CreateFontIndirectW(&navFont);

    LOGFONTW totalFont = uiFont;
    totalFont.lfHeight = -28;
    totalFont.lfWeight = FW_SEMIBOLD;
    totalTimeFont_ = ::CreateFontIndirectW(&totalFont);

    const HWND controls[] = {
        controls_.sidebar, controls_.navCutButton, controls_.navConcatButton, controls_.navFadeButton,
        controls_.navConvertButton, controls_.settingsButton, controls_.contentTitle, controls_.inputLabel,
        controls_.inputEdit, controls_.inputBrowse, controls_.startLabel, controls_.startDecrease,
        controls_.startTrack, controls_.fileDurationValue, controls_.startValue, controls_.startIncrease, controls_.cutDurationValue,
        controls_.endLabel, controls_.endDecrease, controls_.endValue, controls_.endIncrease,
        controls_.runButton, controls_.clearListButton, controls_.openFolderButton,
        controls_.concatList,
        controls_.convertModeVideoButton, controls_.convertModeAudioButton, controls_.convertHint,
        controls_.convertInfoLabelFile, controls_.convertInfoValueFile,
        controls_.convertInfoLabelFormat, controls_.convertInfoValueFormat,
        controls_.convertInfoLabelVideoCodec, controls_.convertInfoValueVideoCodec,
        controls_.convertInfoLabelVideoBitrate, controls_.convertInfoValueVideoBitrate,
        controls_.convertInfoLabelResolution, controls_.convertInfoValueResolution,
        controls_.convertInfoLabelAudioCodec, controls_.convertInfoValueAudioCodec,
        controls_.convertInfoLabelAudioBitrate, controls_.convertInfoValueAudioBitrate,
        controls_.convertToMp3Button, controls_.convertToMp4Button,
        controls_.fadeInTrack, controls_.fadeOutTrack,
        controls_.logEdit,
        controls_.placeholderStatic
    };

    for (HWND control : controls) {
        if (control != nullptr) {
            ::SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(uiFont_), TRUE);
        }
    }
    if (controls_.fileDurationValue != nullptr && totalTimeFont_ != nullptr) {
        ::SendMessageW(controls_.fileDurationValue, WM_SETFONT, reinterpret_cast<WPARAM>(totalTimeFont_), TRUE);
    }

    const HWND navButtons[] = {
        controls_.navCutButton, controls_.navConcatButton, controls_.navFadeButton, controls_.navConvertButton
    };
    for (HWND button : navButtons) {
        if (button != nullptr && navFont_ != nullptr) {
            ::SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(navFont_), TRUE);
        }
    }
}

void MainWindow::InitializeConcatList() {
    if (controls_.concatList == nullptr) {
        return;
    }

    ::SendMessageW(controls_.concatList, LVM_SETEXTENDEDLISTVIEWSTYLE, 0,
                   LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP);
    ::SendMessageW(controls_.concatList, LVM_SETBKCOLOR, 0, RGB(50, 53, 58));
    ::SendMessageW(controls_.concatList, LVM_SETTEXTBKCOLOR, 0, RGB(50, 53, 58));
    ::SendMessageW(controls_.concatList, LVM_SETTEXTCOLOR, 0, RGB(226, 229, 233));
    ConfigureMediaListColumns();
}

void MainWindow::ConfigureMediaListColumns() {
    if (controls_.concatList == nullptr) {
        return;
    }

    while (::SendMessageW(controls_.concatList, LVM_DELETECOLUMN, 0, 0)) {
    }

    const struct ColumnDef {
        const wchar_t* title;
        int width;
    } *columns = nullptr;
    int columnCount = 0;

    static const ColumnDef concatColumns[] = {
        {L"文件名", 300},
        {L"时长", 120},
        {L"分辨率", 150},
        {L"视频编码", 120},
        {L"视频码率", 150},
        {L"音频编码", 120},
        {L"音频码率", 110},
    };
    static const ColumnDef fadeColumns[] = {
        {L"文件名", 420},
        {L"时长", 150},
        {L"音频编码", 160},
        {L"音频码率", 150},
        {L"结果", 120},
    };
    static const ColumnDef convertColumns[] = {
        {L"文件名", 360},
        {L"时长", 140},
        {L"视频编码", 140},
        {L"音频编码", 140},
        {L"结果", 160},
    };
    if (selectedModule_ == 2) {
        columns = fadeColumns;
        columnCount = static_cast<int>(std::size(fadeColumns));
    } else if (selectedModule_ == 3) {
        columns = convertColumns;
        columnCount = static_cast<int>(std::size(convertColumns));
    } else {
        columns = concatColumns;
        columnCount = static_cast<int>(std::size(concatColumns));
    }

    for (int index = 0; index < columnCount; ++index) {
        LVCOLUMNW column{};
        column.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        column.pszText = const_cast<LPWSTR>(columns[index].title);
        column.cx = columns[index].width;
        column.iSubItem = index;
        ::SendMessageW(controls_.concatList, LVM_INSERTCOLUMNW, index, reinterpret_cast<LPARAM>(&column));
    }
}

void MainWindow::LayoutControls(int width, int height) {
    const int sidebarLeft = 12;
    const int sidebarTop = 12;
    const int sidebarWidth = 236;
    const int settingsHeight = 40;
    const int sidebarHeight = std::max(120, height - sidebarTop - 12 - settingsHeight - 8);
    ::MoveWindow(controls_.sidebar, sidebarLeft, sidebarTop, sidebarWidth, sidebarHeight, TRUE);
    const int navLeft = sidebarLeft + 12;
    const int navWidth = sidebarWidth - 24;
    const int navTop = sidebarTop + 28;
    const int navHeight = 44;
    const int navGap = 14;
    ::MoveWindow(controls_.navCutButton, navLeft, navTop, navWidth, navHeight, TRUE);
    ::MoveWindow(controls_.navConcatButton, navLeft, navTop + (navHeight + navGap), navWidth, navHeight, TRUE);
    ::MoveWindow(controls_.navFadeButton, navLeft, navTop + ((navHeight + navGap) * 2), navWidth, navHeight, TRUE);
    ::MoveWindow(controls_.navConvertButton, navLeft, navTop + ((navHeight + navGap) * 3), navWidth, navHeight, TRUE);
    ::MoveWindow(controls_.settingsButton, sidebarLeft, sidebarTop + sidebarHeight + 8, sidebarWidth, settingsHeight, TRUE);

    const int contentLeft = sidebarLeft + sidebarWidth + 12;
    const int contentTop = 12;
    const int contentWidth = std::max(300, width - contentLeft - 12);
    const int contentHeight = std::max(300, height - contentTop - 12);
    ::ShowWindow(controls_.contentTitle, SW_HIDE);
    ::MoveWindow(controls_.cutGroup, contentLeft, contentTop, contentWidth, contentHeight, TRUE);

    const int innerLeft = contentLeft + 24;
    const int innerTop = contentTop + 30;
    const int innerWidth = contentWidth - 48;
    const int labelWidth = 118;
    const int buttonWidth = 168;
    const int smallButtonWidth = 40;
    const int valueWidth = 170;
    const int rowGap = 20;
    const int trackHeight = 96;

    int y = innerTop;
    ::MoveWindow(controls_.inputLabel, innerLeft, y + 6, labelWidth, 28, TRUE);
    ::ShowWindow(controls_.inputLabel, SW_SHOW);
    ::MoveWindow(controls_.inputEdit, innerLeft + labelWidth, y, innerWidth - labelWidth - buttonWidth - 12, 36, TRUE);
    ::MoveWindow(controls_.inputBrowse, innerLeft + innerWidth - buttonWidth, y, buttonWidth, 36, TRUE);

    if (selectedModule_ == 1) {
        ::MoveWindow(controls_.inputLabel, innerLeft, y + 6, 0, 0, TRUE);
        ::MoveWindow(controls_.inputEdit, innerLeft, y, 0, 0, TRUE);
        ::MoveWindow(controls_.inputBrowse, innerLeft, y, 156, 38, TRUE);
        ::MoveWindow(controls_.runButton, innerLeft + 172, y, 156, 38, TRUE);
        ::MoveWindow(controls_.clearListButton, innerLeft + 344, y, 156, 38, TRUE);
        ::MoveWindow(controls_.openFolderButton, innerLeft + 516, y, 236, 38, TRUE);

        y += 38 + rowGap;
        const int remainingHeight = std::max(240, contentTop + contentHeight - y - 16);
        const int logHeight = std::max(110, remainingHeight / 3);
        const int listHeight = std::max(160, remainingHeight - logHeight - 14);
        ::MoveWindow(controls_.concatList, innerLeft, y, innerWidth, listHeight, TRUE);

        y += listHeight + 14;
        ::MoveWindow(controls_.logEdit, innerLeft, y, innerWidth, logHeight, TRUE);
        ::MoveWindow(controls_.placeholderStatic, innerLeft, innerTop, innerWidth, 120, TRUE);
        return;
    }
    if (selectedModule_ == 2) {
        const int fadeGap = 18;
        const int fadeColumnWidth = (innerWidth - fadeGap) / 2;
        const int fadeRowHeight = 36;
        ::MoveWindow(controls_.inputLabel, innerLeft, y + 6, 0, 0, TRUE);
        ::MoveWindow(controls_.inputEdit, innerLeft, y, 0, 0, TRUE);
        ::MoveWindow(controls_.inputBrowse, innerLeft, y, 156, 38, TRUE);
        ::MoveWindow(controls_.runButton, innerLeft + 172, y, 156, 38, TRUE);
        ::MoveWindow(controls_.clearListButton, innerLeft + 344, y, 156, 38, TRUE);
        ::MoveWindow(controls_.openFolderButton, innerLeft + 516, y, 236, 38, TRUE);
        y += 38 + 16;
        ::MoveWindow(controls_.fadeInTrack, innerLeft, y, fadeColumnWidth, 36, TRUE);
        ::MoveWindow(controls_.fadeOutTrack, innerLeft + fadeColumnWidth + fadeGap, y, fadeColumnWidth, 36, TRUE);

        y += 44;
        ::MoveWindow(controls_.startLabel, innerLeft, y, 138, fadeRowHeight, TRUE);
        ::MoveWindow(controls_.startValue, innerLeft + 148, y, fadeColumnWidth - 148, fadeRowHeight, TRUE);
        ::MoveWindow(controls_.endLabel, innerLeft + fadeColumnWidth + fadeGap, y, 138, fadeRowHeight, TRUE);
        ::MoveWindow(controls_.endValue, innerLeft + fadeColumnWidth + fadeGap + 148, y, fadeColumnWidth - 148, fadeRowHeight, TRUE);

        y += fadeRowHeight + rowGap;
        const int remainingHeight = std::max(240, contentTop + contentHeight - y - 16);
        const int logHeight = std::max(110, remainingHeight / 3);
        const int listHeight = std::max(160, remainingHeight - logHeight - 14);
        ::MoveWindow(controls_.concatList, innerLeft, y, innerWidth, listHeight, TRUE);

        y += listHeight + 14;
        ::MoveWindow(controls_.logEdit, innerLeft, y, innerWidth, logHeight, TRUE);
        ::MoveWindow(controls_.placeholderStatic, innerLeft, innerTop, innerWidth, 120, TRUE);
        return;
    }
    if (selectedModule_ == 3) {
        const int infoLabelWidth = 120;
        const int infoValueWidth = std::max(300, innerWidth - infoLabelWidth - 16);
        const int infoRowHeight = 30;
        ::MoveWindow(controls_.inputLabel, innerLeft, y + 6, 0, 0, TRUE);
        ::MoveWindow(controls_.inputEdit, innerLeft, y, 0, 0, TRUE);
        ::MoveWindow(controls_.inputBrowse, innerLeft, y, 156, 38, TRUE);
        ::MoveWindow(controls_.convertToMp3Button, innerLeft + 172, y, 156, 38, TRUE);
        ::MoveWindow(controls_.convertToMp4Button, innerLeft + 344, y, 156, 38, TRUE);
        ::MoveWindow(controls_.openFolderButton, innerLeft + 516, y, 236, 38, TRUE);
        ::MoveWindow(controls_.convertHint, innerLeft, y, 0, 0, TRUE);
        y += 52;
        auto placeInfoRow = [&](HWND label, HWND value) {
            ::MoveWindow(label, innerLeft, y, infoLabelWidth, infoRowHeight, TRUE);
            ::MoveWindow(value, innerLeft + infoLabelWidth + 10, y, infoValueWidth, infoRowHeight, TRUE);
            y += infoRowHeight + 8;
        };
        placeInfoRow(controls_.convertInfoLabelFile, controls_.convertInfoValueFile);
        placeInfoRow(controls_.convertInfoLabelFormat, controls_.convertInfoValueFormat);
        placeInfoRow(controls_.convertInfoLabelVideoCodec, controls_.convertInfoValueVideoCodec);
        placeInfoRow(controls_.convertInfoLabelVideoBitrate, controls_.convertInfoValueVideoBitrate);
        placeInfoRow(controls_.convertInfoLabelResolution, controls_.convertInfoValueResolution);
        placeInfoRow(controls_.convertInfoLabelAudioCodec, controls_.convertInfoValueAudioCodec);
        placeInfoRow(controls_.convertInfoLabelAudioBitrate, controls_.convertInfoValueAudioBitrate);

        y += 8;
        const int logHeight = std::max(140, contentTop + contentHeight - y - 16);
        ::MoveWindow(controls_.logEdit, innerLeft, y, innerWidth, logHeight, TRUE);
        ::MoveWindow(controls_.placeholderStatic, innerLeft, innerTop, innerWidth, 120, TRUE);
        return;
    }

    y += 36 + rowGap + 8;
    ::MoveWindow(controls_.startTrack, innerLeft, y, innerWidth, trackHeight, TRUE);
    ::MoveWindow(controls_.fileDurationValue, innerLeft + std::max(0, innerWidth - 320), y + trackHeight, 320, 28, TRUE);
    ::ShowWindow(controls_.cutDurationValue, SW_SHOW);

    y += trackHeight + 52;
    const int cutLabelWidth = 110;
    ::MoveWindow(controls_.startLabel, innerLeft, y + 4, cutLabelWidth, 28, TRUE);
    ::MoveWindow(controls_.startValue, innerLeft + cutLabelWidth, y, valueWidth, 32, TRUE);
    ::MoveWindow(controls_.startDecrease, innerLeft + cutLabelWidth + valueWidth + 10, y, smallButtonWidth, 32, TRUE);
    ::MoveWindow(controls_.startIncrease, innerLeft + cutLabelWidth + valueWidth + 10 + smallButtonWidth + 6, y, smallButtonWidth, 32, TRUE);

    const int endBlockLeft = innerLeft + innerWidth - (cutLabelWidth + valueWidth + 10 + smallButtonWidth + 6 + smallButtonWidth);
    const int centerLeft = innerLeft + cutLabelWidth + valueWidth + 10 + smallButtonWidth + 6 + smallButtonWidth + 20;
    const int centerRight = endBlockLeft - 20;
    const int centerWidth = std::max(180, centerRight - centerLeft);
    ::MoveWindow(controls_.cutDurationValue, centerLeft, y, centerWidth, 32, TRUE);
    ::MoveWindow(controls_.endLabel, endBlockLeft, y + 4, cutLabelWidth, 28, TRUE);
    ::MoveWindow(controls_.endValue, endBlockLeft + cutLabelWidth, y, valueWidth, 32, TRUE);
    ::MoveWindow(controls_.endDecrease, endBlockLeft + cutLabelWidth + valueWidth + 10, y, smallButtonWidth, 32, TRUE);
    ::MoveWindow(controls_.endIncrease, endBlockLeft + cutLabelWidth + valueWidth + 10 + smallButtonWidth + 6, y, smallButtonWidth, 32, TRUE);

    y += 32 + rowGap + 8;
    ::MoveWindow(controls_.runButton, innerLeft, y, 156, 38, TRUE);
    ::MoveWindow(controls_.clearListButton, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.openFolderButton, innerLeft + 172, y, 236, 38, TRUE);
    ::MoveWindow(controls_.convertModeVideoButton, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertModeAudioButton, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertHint, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoLabelFile, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoValueFile, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoLabelFormat, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoValueFormat, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoLabelVideoCodec, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoValueVideoCodec, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoLabelVideoBitrate, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoValueVideoBitrate, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoLabelResolution, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoValueResolution, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoLabelAudioCodec, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoValueAudioCodec, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoLabelAudioBitrate, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertInfoValueAudioBitrate, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertToMp3Button, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.convertToMp4Button, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.fadeInTrack, innerLeft, y, 0, 0, TRUE);
    ::MoveWindow(controls_.fadeOutTrack, innerLeft, y, 0, 0, TRUE);

    y += 38 + rowGap;
    const int logHeight = std::max(88, contentTop + contentHeight - y - 56);
    ::MoveWindow(controls_.concatList, innerLeft, y, innerWidth, 0, TRUE);
    ::MoveWindow(controls_.logEdit, innerLeft, y, innerWidth, logHeight, TRUE);
    ::MoveWindow(controls_.placeholderStatic, innerLeft, innerTop, innerWidth, 120, TRUE);
}

void MainWindow::RefreshText() {
    ::SetWindowTextW(hwnd_, LoadStringResource(IDS_APP_TITLE).c_str());
    ::SetWindowTextW(controls_.settingsButton, LoadStringResource(IDS_SETTINGS).c_str());
    ::SetWindowTextW(controls_.navCutButton, LoadStringResource(IDS_MODULE_CUT).c_str());
    ::SetWindowTextW(controls_.navConcatButton, LoadStringResource(IDS_MODULE_CONCAT).c_str());
    ::SetWindowTextW(controls_.navFadeButton, LoadStringResource(IDS_MODULE_FADE).c_str());
    ::SetWindowTextW(controls_.navConvertButton, LoadStringResource(IDS_MODULE_CONVERT).c_str());
    ::SetWindowTextW(controls_.contentTitle, LoadStringResource(IDS_CUT_TITLE).c_str());
    ::SetWindowTextW(controls_.cutGroup, LoadStringResource(IDS_CUT_TITLE).c_str());
    ::SetWindowTextW(controls_.inputLabel, LoadStringResource(IDS_LABEL_INPUT).c_str());
    ::SetWindowTextW(controls_.inputBrowse, LoadStringResource(IDS_BUTTON_OPEN_FILE).c_str());
    ::SetWindowTextW(controls_.startLabel, LoadStringResource(IDS_LABEL_START).c_str());
    ::SetWindowTextW(controls_.endLabel, LoadStringResource(IDS_LABEL_END).c_str());
    ::SetWindowTextW(controls_.runButton, LoadStringResource(IDS_BUTTON_RUN_CUT).c_str());
    ::SetWindowTextW(controls_.openFolderButton, LoadStringResource(IDS_BUTTON_OPEN_FOLDER).c_str());
    ::SetWindowTextW(controls_.fileDurationValue, L"");
    ::SetWindowTextW(controls_.cutDurationValue, L"");
    ::SetWindowTextW(controls_.logEdit, L"");
    UpdateTimeDisplays();
}

void MainWindow::RefreshModule() {
    const bool cutVisible = selectedModule_ == 0;
    const bool concatVisible = selectedModule_ == 1;
    const bool fadeVisible = selectedModule_ == 2;
    const bool convertVisible = selectedModule_ == 3;
    ::EnableWindow(controls_.navCutButton, selectedModule_ != 0);
    ::EnableWindow(controls_.navConcatButton, selectedModule_ != 1);
    ::EnableWindow(controls_.navFadeButton, selectedModule_ != 2);
    ::EnableWindow(controls_.navConvertButton, selectedModule_ != 3);

    const HWND sharedControls[] = {
        controls_.cutGroup, controls_.inputLabel, controls_.inputEdit, controls_.inputBrowse,
        controls_.openFolderButton, controls_.logEdit
    };

    const HWND concatOnlyControls[] = {
    };

    const HWND cutOnlyControls[] = {
        controls_.startDecrease, controls_.startTrack, controls_.fileDurationValue, controls_.startIncrease,
        controls_.cutDurationValue,
        controls_.endDecrease, controls_.endIncrease
    };
    const HWND fadeOnlyControls[] = {
        controls_.fadeInTrack, controls_.fadeOutTrack
    };
    const HWND convertOnlyControls[] = {
        controls_.convertHint,
        controls_.convertInfoLabelFile, controls_.convertInfoValueFile,
        controls_.convertInfoLabelFormat, controls_.convertInfoValueFormat,
        controls_.convertInfoLabelVideoCodec, controls_.convertInfoValueVideoCodec,
        controls_.convertInfoLabelVideoBitrate, controls_.convertInfoValueVideoBitrate,
        controls_.convertInfoLabelResolution, controls_.convertInfoValueResolution,
        controls_.convertInfoLabelAudioCodec, controls_.convertInfoValueAudioCodec,
        controls_.convertInfoLabelAudioBitrate, controls_.convertInfoValueAudioBitrate,
        controls_.convertToMp3Button, controls_.convertToMp4Button
    };

    for (HWND control : sharedControls) {
        ::ShowWindow(control, (cutVisible || concatVisible || fadeVisible || convertVisible) ? SW_SHOW : SW_HIDE);
    }
    for (HWND control : cutOnlyControls) {
        ::ShowWindow(control, cutVisible ? SW_SHOW : SW_HIDE);
    }
    for (HWND control : concatOnlyControls) {
        ::ShowWindow(control, concatVisible ? SW_SHOW : SW_HIDE);
    }
    for (HWND control : fadeOnlyControls) {
        ::ShowWindow(control, fadeVisible ? SW_SHOW : SW_HIDE);
    }
    for (HWND control : convertOnlyControls) {
        ::ShowWindow(control, convertVisible ? SW_SHOW : SW_HIDE);
    }
    ::ShowWindow(controls_.startLabel, (cutVisible || fadeVisible) ? SW_SHOW : SW_HIDE);
    ::ShowWindow(controls_.startValue, (cutVisible || fadeVisible) ? SW_SHOW : SW_HIDE);
    ::ShowWindow(controls_.endLabel, (cutVisible || fadeVisible) ? SW_SHOW : SW_HIDE);
    ::ShowWindow(controls_.endValue, (cutVisible || fadeVisible) ? SW_SHOW : SW_HIDE);
    ::ShowWindow(controls_.runButton, (cutVisible || concatVisible || fadeVisible) ? SW_SHOW : SW_HIDE);
    ::ShowWindow(controls_.concatList, (concatVisible || fadeVisible) ? SW_SHOW : SW_HIDE);
    ::ShowWindow(controls_.clearListButton, (concatVisible || fadeVisible) ? SW_SHOW : SW_HIDE);
    ::ShowWindow(controls_.convertModeVideoButton, SW_HIDE);
    ::ShowWindow(controls_.convertModeAudioButton, SW_HIDE);

    if (cutVisible) {
        ConfigureMediaListColumns();
        ::SetWindowTextW(controls_.contentTitle, LoadStringResource(IDS_CUT_TITLE).c_str());
        ::SetWindowTextW(controls_.cutGroup, LoadStringResource(IDS_CUT_TITLE).c_str());
        ::SetWindowTextW(controls_.inputLabel, LoadStringResource(IDS_LABEL_INPUT).c_str());
        ::SetWindowTextW(controls_.inputBrowse, LoadStringResource(IDS_BUTTON_OPEN_FILE).c_str());
        ::SetWindowTextW(controls_.startLabel, LoadStringResource(IDS_LABEL_START).c_str());
        ::SetWindowTextW(controls_.endLabel, LoadStringResource(IDS_LABEL_END).c_str());
        if (!cutSucceeded_) {
            ::SetWindowTextW(controls_.runButton, LoadStringResource(IDS_BUTTON_RUN_CUT).c_str());
        }
        ::SetWindowTextW(controls_.clearListButton, L"");
        ::SetWindowTextW(controls_.openFolderButton, LoadStringResource(IDS_BUTTON_OPEN_FOLDER).c_str());
        ::SetWindowTextW(controls_.inputEdit,
                         inputFilePath_.empty() ? L"" : std::filesystem::path(inputFilePath_).filename().wstring().c_str());
        ::ShowWindow(controls_.placeholderStatic, SW_HIDE);
        UpdatePrimaryActionState();
        UpdateLogPanelState();
        return;
    }

    if (concatVisible) {
        ConfigureMediaListColumns();
        ::SetWindowTextW(controls_.contentTitle, L"音频拼接");
        ::SetWindowTextW(controls_.cutGroup, L"音频拼接");
        ::SetWindowTextW(controls_.inputLabel, L"");
        ::SetWindowTextW(controls_.inputBrowse, L"选择文件");
        ::SetWindowTextW(controls_.inputEdit, L"");
        if (!cutSucceeded_) {
            ::SetWindowTextW(controls_.runButton, L"开始拼接");
        }
        ::SetWindowTextW(controls_.clearListButton, L"清空列表");
        ::SetWindowTextW(controls_.openFolderButton, L"打开输出文件夹");
        RefreshConcatList();
        ::ShowWindow(controls_.placeholderStatic, SW_HIDE);
        UpdatePrimaryActionState();
        UpdateLogPanelState();
        return;
    }
    if (fadeVisible) {
        ConfigureMediaListColumns();
        ::SetWindowTextW(controls_.contentTitle, L"淡入淡出");
        ::SetWindowTextW(controls_.cutGroup, L"淡入淡出");
        ::SetWindowTextW(controls_.inputLabel, L"");
        ::SetWindowTextW(controls_.inputBrowse, L"选择文件");
        ::SetWindowTextW(controls_.inputEdit, L"");
        ::SetWindowTextW(controls_.startLabel, L"开头淡入");
        ::SetWindowTextW(controls_.endLabel, L"结尾淡出");
        if (!cutSucceeded_) {
            ::SetWindowTextW(controls_.runButton, L"开始处理");
        }
        ::SetWindowTextW(controls_.clearListButton, L"清空列表");
        ::SetWindowTextW(controls_.openFolderButton, L"打开输出文件夹");
        ::SetWindowTextW(controls_.startValue, FormatMilliseconds(fadeInMilliseconds_).c_str());
        ::SetWindowTextW(controls_.endValue, FormatMilliseconds(fadeOutMilliseconds_).c_str());
        RefreshConcatList();
        ::ShowWindow(controls_.placeholderStatic, SW_HIDE);
        UpdatePrimaryActionState();
        UpdateLogPanelState();
        return;
    }
    if (convertVisible) {
        ::SetWindowTextW(controls_.contentTitle, L"格式转换");
        ::SetWindowTextW(controls_.cutGroup, L"格式转换");
        ::SetWindowTextW(controls_.inputLabel, L"");
        ::SetWindowTextW(controls_.inputBrowse, L"选择单个文件");
        ::SetWindowTextW(controls_.inputEdit, L"");
        ::SetWindowTextW(controls_.convertHint, L"");
        ::SetWindowTextW(controls_.openFolderButton, L"打开输出文件夹");
        RefreshConvertInfo();
        ::ShowWindow(controls_.placeholderStatic, SW_HIDE);
        UpdateLogPanelState();
        return;
    }

    unsigned int placeholderId = IDS_PLACEHOLDER_CONCAT;
    unsigned int titleId = IDS_MODULE_CONCAT;
    if (selectedModule_ == 2) {
        placeholderId = IDS_PLACEHOLDER_FADE;
        titleId = IDS_MODULE_FADE;
    } else if (selectedModule_ == 3) {
        placeholderId = IDS_PLACEHOLDER_CONVERT;
        titleId = IDS_MODULE_CONVERT;
    }

    ::SetWindowTextW(controls_.contentTitle, LoadStringResource(titleId).c_str());
    ::SetWindowTextW(controls_.placeholderStatic, LoadStringResource(placeholderId).c_str());
    ::ShowWindow(controls_.placeholderStatic, SW_SHOW);
}

void MainWindow::AppendLog(const std::wstring& line) {
    std::wstring appended = line;
    if (!appended.empty() && appended.back() != L'\n') {
        appended += L"\r\n";
    }
    std::wstring lower = line;
    std::transform(lower.begin(), lower.end(), lower.begin(), [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    const bool expandForErrors =
        lower.find(L"error") != std::wstring::npos ||
        lower.find(L"failed") != std::wstring::npos ||
        lower.find(L"invalid") != std::wstring::npos;
    AppendLogText(appended, expandForErrors);
}

void MainWindow::AppendLogLines(const std::vector<std::wstring>& lines) {
    if (lines.empty()) {
        return;
    }

    std::wstring appended;
    size_t totalLength = 0;
    bool expandForErrors = false;
    for (const auto& line : lines) {
        totalLength += line.size() + 2;
        std::wstring lower = line;
        std::transform(lower.begin(), lower.end(), lower.begin(), [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
        if (lower.find(L"error") != std::wstring::npos ||
            lower.find(L"failed") != std::wstring::npos ||
            lower.find(L"invalid") != std::wstring::npos) {
            expandForErrors = true;
        }
    }
    appended.reserve(totalLength);
    for (const auto& line : lines) {
        appended += line;
        if (appended.empty() || appended.back() != L'\n') {
            appended += L"\r\n";
        }
    }
    AppendLogText(appended, expandForErrors);
}

void MainWindow::AppendLogText(const std::wstring& text, bool expandForErrors) {
    if (controls_.logEdit == nullptr || text.empty()) {
        return;
    }

    int textLength = ::GetWindowTextLengthW(controls_.logEdit);
    if (textLength > kMaxLogChars) {
        const int keepFrom = std::max(0, textLength - kTrimLogChars);
        ::SendMessageW(controls_.logEdit, EM_SETSEL, 0, keepFrom);
        ::SendMessageW(controls_.logEdit, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(L""));
        textLength = ::GetWindowTextLengthW(controls_.logEdit);
    }

    ::SendMessageW(controls_.logEdit, EM_SETSEL, textLength, textLength);
    ::SendMessageW(controls_.logEdit, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(text.c_str()));
    if (expandForErrors) {
        SetLogExpanded(true);
    }
}

void MainWindow::ClearLog() {
    ::SetWindowTextW(controls_.logEdit, L"");
}

void MainWindow::OpenInputFileDialog() {
    OPENFILENAMEW dialog{};
    wchar_t filePath[MAX_PATH] = {};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = hwnd_;
    dialog.lpstrFile = filePath;
    dialog.nMaxFile = MAX_PATH;
    dialog.lpstrFilter = L"Media Files\0*.mp4;*.mkv;*.mov;*.mp3;*.wav;*.m4a;*.aac;*.flac\0All Files\0*.*\0";
    dialog.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    dialog.lpstrTitle = LoadStringResource(IDS_DIALOG_SELECT_MEDIA).c_str();

    if (::GetOpenFileNameW(&dialog)) {
        inputFilePath_ = filePath;
        lastOutputPath_.clear();
        ResetCutResultState();
        ::SetWindowTextW(controls_.inputEdit, std::filesystem::path(filePath).filename().wstring().c_str());
        UpdatePrimaryActionState();
        AutoDetectInputDuration();
    }
}

void MainWindow::OpenConcatFilesDialog() {
    std::vector<wchar_t> buffer(65536, L'\0');
    OPENFILENAMEW dialog{};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = hwnd_;
    dialog.lpstrFile = buffer.data();
    dialog.nMaxFile = static_cast<DWORD>(buffer.size());
    dialog.lpstrFilter = L"Media Files\0*.mp4;*.mkv;*.mov;*.mp3;*.wav;*.flac;*.m4a;*.aac\0All Files\0*.*\0";
    dialog.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_EXPLORER;
    dialog.lpstrTitle = L"选择要拼接的音频文件";

    if (!::GetOpenFileNameW(&dialog)) {
        return;
    }

    std::vector<std::wstring> selectedFiles;
    const wchar_t* basePath = buffer.data();
    const wchar_t* cursor = basePath + wcslen(basePath) + 1;
    if (*cursor == L'\0') {
        selectedFiles.emplace_back(basePath);
    } else {
        const std::filesystem::path folder(basePath);
        while (*cursor != L'\0') {
            selectedFiles.push_back((folder / cursor).wstring());
            cursor += wcslen(cursor) + 1;
        }
    }
    if (selectedFiles.empty()) {
        return;
    }

    concatInputPaths_ = std::move(selectedFiles);
    concatItems_.clear();
    concatItems_.reserve(concatInputPaths_.size());
    for (const auto& path : concatInputPaths_) {
        ConcatListItem item{};
        if (!TryProbeConcatItem(path, item)) {
            item.filePath = path;
            item.fileName = std::filesystem::path(path).filename().wstring();
            item.durationText = L"-";
            if (item.resolutionText.empty()) {
                item.resolutionText = L"-";
            }
            if (item.videoCodec.empty()) {
                item.videoCodec = L"-";
            }
            if (item.videoBitrate.empty()) {
                item.videoBitrate = L"-";
            }
            if (item.audioCodec.empty()) {
                item.audioCodec = L"错误：无法识别音频编码";
            }
            if (item.audioBitrate.empty()) {
                item.audioBitrate = L"-";
            }
            item.hasError = true;
        }
        item.resultText = L"-";
        concatItems_.push_back(std::move(item));
    }
    lastOutputPath_.clear();
    ResetCutResultState();
    const std::wstring summary = std::filesystem::path(concatInputPaths_.front()).filename().wstring() +
                                 L" 等 " + std::to_wstring(concatInputPaths_.size()) + L" 个文件";
    ::SetWindowTextW(controls_.inputEdit, summary.c_str());
    UpdatePrimaryActionState();
    RefreshConcatList();
    StartMediaProbe(1, concatInputPaths_, false);
}

void MainWindow::OpenFadeFilesDialog() {
    std::vector<wchar_t> buffer(65536, L'\0');
    OPENFILENAMEW dialog{};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = hwnd_;
    dialog.lpstrFile = buffer.data();
    dialog.nMaxFile = static_cast<DWORD>(buffer.size());
    dialog.lpstrFilter = L"Audio Files\0*.mp3;*.wav;*.flac;*.m4a;*.aac\0All Files\0*.*\0";
    dialog.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_EXPLORER;
    dialog.lpstrTitle = L"选择要处理的音频文件";

    if (!::GetOpenFileNameW(&dialog)) {
        return;
    }

    std::vector<std::wstring> selectedFiles;
    const wchar_t* basePath = buffer.data();
    const wchar_t* cursor = basePath + wcslen(basePath) + 1;
    if (*cursor == L'\0') {
        selectedFiles.emplace_back(basePath);
    } else {
        const std::filesystem::path folder(basePath);
        while (*cursor != L'\0') {
            selectedFiles.push_back((folder / cursor).wstring());
            cursor += wcslen(cursor) + 1;
        }
    }
    if (selectedFiles.empty()) {
        return;
    }

    fadeInputPaths_ = std::move(selectedFiles);
    fadeItems_.clear();
    fadeItems_.reserve(fadeInputPaths_.size());
    for (const auto& path : fadeInputPaths_) {
        ConcatListItem item{};
        if (!TryProbeConcatItem(path, item)) {
            item.filePath = path;
            item.fileName = std::filesystem::path(path).filename().wstring();
            item.durationText = L"-";
            if (item.audioCodec.empty()) {
                item.audioCodec = L"错误：无法识别音频编码";
            }
            if (item.audioBitrate.empty()) {
                item.audioBitrate = L"-";
            }
            item.hasError = true;
        }
        item.resultText = L"待处理";
        fadeItems_.push_back(std::move(item));
    }
    lastOutputPath_.clear();
    ResetCutResultState();
    const std::wstring summary = std::filesystem::path(fadeInputPaths_.front()).filename().wstring() +
                                 L" 等 " + std::to_wstring(fadeInputPaths_.size()) + L" 个文件";
    ::SetWindowTextW(controls_.inputEdit, summary.c_str());
    UpdatePrimaryActionState();
    RefreshConcatList();
    StartMediaProbe(2, fadeInputPaths_, false);
}

void MainWindow::OpenConvertFilesDialog() {
    wchar_t filePath[MAX_PATH] = {};
    OPENFILENAMEW dialog{};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = hwnd_;
    dialog.lpstrFile = filePath;
    dialog.nMaxFile = MAX_PATH;
    dialog.lpstrFilter = L"Supported Files\0*.m4a;*.mp4;*.mkv;*.mov;*.avi;*.flv;*.wmv;*.webm\0All Files\0*.*\0";
    dialog.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    dialog.lpstrTitle = L"选择要转换的单个文件";

    if (!::GetOpenFileNameW(&dialog)) {
        return;
    }

    convertInputPaths_.clear();
    convertInputPaths_.push_back(filePath);
    convertItems_.clear();
    convertItem_ = {};
    if (!TryProbeConcatItem(filePath, convertItem_)) {
        convertItem_.filePath = filePath;
        convertItem_.fileName = std::filesystem::path(filePath).filename().wstring();
        convertItem_.durationText = L"-";
        if (convertItem_.videoCodec.empty()) {
            convertItem_.videoCodec = L"-";
        }
        if (convertItem_.videoBitrate.empty()) {
            convertItem_.videoBitrate = L"-";
        }
        if (convertItem_.resolutionText.empty()) {
            convertItem_.resolutionText = L"-";
        }
        if (convertItem_.audioCodec.empty()) {
            convertItem_.audioCodec = L"错误：无法识别媒体信息";
        }
        if (convertItem_.audioBitrate.empty()) {
            convertItem_.audioBitrate = L"-";
        }
        convertItem_.hasError = true;
    }
    lastOutputPath_.clear();
    ResetCutResultState();
    ::SetWindowTextW(controls_.inputEdit, std::filesystem::path(filePath).filename().wstring().c_str());
    RefreshConvertInfo();
}

void MainWindow::StartMediaProbe(int module, const std::vector<std::wstring>& inputPaths, bool singleItem) {
    const int requestId = ++mediaProbeRequestId_;
    const HWND target = hwnd_;
    SpawnBackgroundTask([this, target, requestId, module, inputPaths, singleItem]() {
        for (int index = 0; index < static_cast<int>(inputPaths.size()); ++index) {
            if (shuttingDown_.load()) {
                return;
            }
            ConcatListItem item{};
            if (!TryProbeConcatItem(inputPaths[index], item)) {
                item.filePath = inputPaths[index];
                item.fileName = std::filesystem::path(inputPaths[index]).filename().wstring();
                if (item.durationText.empty()) item.durationText = L"-";
                if (item.resolutionText.empty()) item.resolutionText = L"-";
                if (item.videoCodec.empty()) item.videoCodec = L"-";
                if (item.videoBitrate.empty()) item.videoBitrate = L"-";
                if (item.audioCodec.empty()) item.audioCodec = L"错误：无法识别媒体信息";
                if (item.audioBitrate.empty()) item.audioBitrate = L"-";
                item.hasError = true;
            }
            if (module == 2 && item.resultText.empty()) {
                item.resultText = L"待处理";
            } else if (item.resultText.empty()) {
                item.resultText = L"-";
            }
            auto message = std::make_unique<MediaProbeItemMessage>();
            message->requestId = requestId;
            message->module = module;
            message->index = index;
            message->singleItem = singleItem;
            message->item = std::move(item);
            if (!PostOwnedMessage(target, WM_APP_MEDIA_PROBE_ITEM, 0, std::move(message))) {
                return;
            }
        }
        auto finished = std::make_unique<MediaProbeFinishedMessage>();
        finished->requestId = requestId;
        finished->module = module;
        PostOwnedMessage(target, WM_APP_MEDIA_PROBE_FINISHED, 0, std::move(finished));
    });
}

bool MainWindow::PromptOutputFolder(std::wstring& folderPath) {
    BROWSEINFOW browseInfo{};
    browseInfo.hwndOwner = hwnd_;
    browseInfo.lpszTitle = LoadStringResource(IDS_DIALOG_SAVE_MEDIA).c_str();
    browseInfo.ulFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI | BIF_NEWDIALOGSTYLE;

    PIDLIST_ABSOLUTE itemIdList = ::SHBrowseForFolderW(&browseInfo);
    if (itemIdList == nullptr) {
        return false;
    }

    wchar_t selectedPath[MAX_PATH] = {};
    const BOOL ok = ::SHGetPathFromIDListW(itemIdList, selectedPath);
    ::CoTaskMemFree(itemIdList);
    if (!ok || selectedPath[0] == L'\0') {
        return false;
    }

    folderPath = selectedPath;
    return true;
}

void MainWindow::OpenOutputFolder() {
    std::filesystem::path folder;
    if (!lastOutputPath_.empty()) {
        folder = std::filesystem::path(lastOutputPath_).parent_path();
    } else if (!fadeInputPaths_.empty()) {
        folder = std::filesystem::path(fadeInputPaths_.front()).parent_path();
    } else if (!convertInputPaths_.empty()) {
        folder = std::filesystem::path(convertInputPaths_.front()).parent_path();
    } else if (!concatInputPaths_.empty()) {
        folder = std::filesystem::path(concatInputPaths_.front()).parent_path();
    } else if (!inputFilePath_.empty()) {
        folder = std::filesystem::path(inputFilePath_).parent_path();
    }
    if (folder.empty() || !std::filesystem::exists(folder)) {
        return;
    }

    const std::wstring folderText = folder.wstring();
    ::ShellExecuteW(hwnd_, L"open", folderText.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
}

void MainWindow::OpenSettingsDialog() {
    WNDCLASSW wc{};
    wc.lpfnWndProc = MainWindow::SettingsProc;
    wc.hInstance = instance_;
    wc.lpszClassName = L"VideoAudioCutSettingsDialog";
    wc.hCursor = ::LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = panelBrush_;
    ::RegisterClassW(&wc);

    HWND dialog = ::CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        wc.lpszClassName,
        LoadStringResource(IDS_SETTINGS_TITLE).c_str(),
        WS_POPUP | WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 760, 460,
        hwnd_, nullptr, instance_, this);

    if (dialog == nullptr) {
        return;
    }

    CenterWindowToScreen(dialog, 760, 460);
    BOOL darkMode = TRUE;
    ::DwmSetWindowAttribute(dialog, DWMWA_USE_IMMERSIVE_DARK_MODE, &darkMode, sizeof(darkMode));

    ::EnableWindow(hwnd_, FALSE);
    ::ShowWindow(dialog, SW_SHOW);

    MSG message{};
    while (::IsWindow(dialog) && ::GetMessageW(&message, nullptr, 0, 0)) {
        if (!::IsDialogMessageW(dialog, &message)) {
            ::TranslateMessage(&message);
            ::DispatchMessageW(&message);
        }
    }

    ::EnableWindow(hwnd_, TRUE);
    ::SetActiveWindow(hwnd_);
}

void MainWindow::StartCutTask() {
    if (taskRunning_ || runner_.IsRunning()) {
        return;
    }

    if (!ValidateFfmpegPath()) {
        ::MessageBoxW(hwnd_, LoadStringResource(IDS_STATUS_INVALID_FFMPEG).c_str(),
                      LoadStringResource(IDS_SETTINGS_TITLE).c_str(), MB_OK | MB_ICONWARNING);
        return;
    }

    const std::wstring inputPath = inputFilePath_;

    if (inputPath.empty()) {
        ::MessageBoxW(hwnd_, LoadStringResource(IDS_STATUS_INPUT_REQUIRED).c_str(),
                      LoadStringResource(IDS_CUT_TITLE).c_str(), MB_OK | MB_ICONWARNING);
        return;
    }
    if (!durationAvailable_ || mediaDurationMilliseconds_ <= 0) {
        ::MessageBoxW(hwnd_, L"\x8BF7\x5148\x6DFB\x52A0\x53EF\x89E3\x6790\x65F6\x957F\x7684\x97F3\x9891\x6216\x89C6\x9891\x6587\x4EF6",
                      LoadStringResource(IDS_CUT_TITLE).c_str(), MB_OK | MB_ICONWARNING);
        return;
    }

    std::wstring outputPath;
    if (config_.saveToSourceFolder) {
        outputPath = BuildDefaultOutputPath(inputPath);
    } else {
        std::wstring outputFolder;
        if (!PromptOutputFolder(outputFolder)) {
            return;
        }
        outputPath = BuildOutputPathInFolder(inputPath, outputFolder);
    }

    if (outputPath.empty()) {
        ::MessageBoxW(hwnd_, LoadStringResource(IDS_STATUS_OUTPUT_REQUIRED).c_str(),
                      LoadStringResource(IDS_CUT_TITLE).c_str(), MB_OK | MB_ICONWARNING);
        return;
    }
    lastOutputPath_ = outputPath;

    const std::wstring startTime = FormatMilliseconds(rangeStartMilliseconds_);
    const std::wstring endTime = FormatMilliseconds(rangeEndMilliseconds_);
    const std::wstring totalCutTime = FormatMilliseconds(std::max(0, rangeEndMilliseconds_ - rangeStartMilliseconds_));

    std::wstringstream args;
    args << L"-hide_banner -ss " << EscapeArgument(startTime)
         << L" -i " << EscapeArgument(inputPath)
         << L" -t " << EscapeArgument(totalCutTime)
         << L" -c copy -avoid_negative_ts make_zero -y "
         << EscapeArgument(outputPath);
    const std::wstring fullCommand = L"\"" + config_.ffmpegPath + L"\" " + args.str();

    ClearLog();
    AppendLog(L"\x5F00\x59CB\x65F6\x95F4\xFF1A" + startTime);
    AppendLog(L"\x7ED3\x675F\x65F6\x95F4\xFF1A" + endTime);
    AppendLog(L"\x526A\x5207\x603B\x65F6\x95F4\xFF1A" + totalCutTime);
    AppendLog(L"FFmpeg \x547D\x4EE4\xFF1A");
    AppendLog(fullCommand);
    AppendLog(LoadStringResource(IDS_STATUS_RUNNING));
    ::EnableWindow(controls_.runButton, FALSE);
    taskRunning_ = true;
    activeTaskModule_ = 0;

    auto pendingLines = std::make_shared<std::vector<std::wstring>>();
    if (!runner_.Start(
        config_.ffmpegPath,
        args.str(),
        [target = hwnd_, pendingLines](const std::wstring& line) {
            QueueLogLine(target, *pendingLines, line);
        },
        [target = hwnd_, pendingLines](int exitCode) {
            FlushPendingLogLines(target, *pendingLines);
            ::PostMessageW(target, WM_APP_RUN_FINISHED, static_cast<WPARAM>(exitCode), 0);
        })) {
        taskRunning_ = false;
        ::EnableWindow(controls_.runButton, TRUE);
    }
}

void MainWindow::StartConcatTask() {
    if (taskRunning_ || runner_.IsRunning()) {
        return;
    }

    if (!ValidateFfmpegPath()) {
        ::MessageBoxW(hwnd_, LoadStringResource(IDS_STATUS_INVALID_FFMPEG).c_str(),
                      LoadStringResource(IDS_SETTINGS_TITLE).c_str(), MB_OK | MB_ICONWARNING);
        return;
    }

    if (concatInputPaths_.empty()) {
        ::MessageBoxW(hwnd_, L"请先选择要拼接的音频文件",
                      L"音频拼接", MB_OK | MB_ICONWARNING);
        return;
    }

    std::wstring listFilePath;
    int fileCount = 0;
    if (!BuildConcatListFile(concatInputPaths_, listFilePath, fileCount)) {
        ::MessageBoxW(hwnd_, L"无法生成拼接文件列表，请确认所选文件存在且支持：mp3、wav、flac、m4a、aac",
                      L"音频拼接", MB_OK | MB_ICONWARNING);
        return;
    }

    const std::wstring outputPath = BuildConcatOutputPath(concatInputPaths_);
    if (outputPath.empty()) {
        ::DeleteFileW(listFilePath.c_str());
        ::MessageBoxW(hwnd_, L"无法生成拼接输出文件路径", L"音频拼接", MB_OK | MB_ICONWARNING);
        return;
    }

    concatListFilePath_ = listFilePath;
    lastOutputPath_ = outputPath;

    std::wstringstream args;
    args << L"-hide_banner -f concat -safe 0 -i " << EscapeArgument(listFilePath)
         << L" -c copy -y " << EscapeArgument(outputPath);
    const std::wstring fullCommand = L"\"" + config_.ffmpegPath + L"\" " + args.str();

    ClearLog();
    AppendLog(L"拼接文件数量：" + std::to_wstring(fileCount));
    for (const auto& path : concatInputPaths_) {
        AppendLog(L"输入文件：" + path);
    }
    AppendLog(L"输出文件：" + outputPath);
    AppendLog(L"FFmpeg 命令：");
    AppendLog(fullCommand);
    AppendLog(L"正在执行拼接...");
    ::EnableWindow(controls_.runButton, FALSE);
    taskRunning_ = true;
    activeTaskModule_ = 1;

    auto pendingLines = std::make_shared<std::vector<std::wstring>>();
    if (!runner_.Start(
        config_.ffmpegPath,
        args.str(),
        [target = hwnd_, pendingLines](const std::wstring& line) {
            QueueLogLine(target, *pendingLines, line);
        },
        [target = hwnd_, listFilePath, pendingLines](int exitCode) {
            FlushPendingLogLines(target, *pendingLines);
            if (!listFilePath.empty()) {
                std::error_code errorCode;
                std::filesystem::remove(std::filesystem::path(listFilePath), errorCode);
            }
            ::PostMessageW(target, WM_APP_RUN_FINISHED, static_cast<WPARAM>(exitCode), 0);
        })) {
        taskRunning_ = false;
        ::EnableWindow(controls_.runButton, TRUE);
    }
}

void MainWindow::StartFadeTask() {
    if (taskRunning_ || runner_.IsRunning()) {
        return;
    }
    if (!ValidateFfmpegPath()) {
        ::MessageBoxW(hwnd_, LoadStringResource(IDS_STATUS_INVALID_FFMPEG).c_str(),
                      LoadStringResource(IDS_SETTINGS_TITLE).c_str(), MB_OK | MB_ICONWARNING);
        return;
    }
    if (fadeInputPaths_.empty()) {
        ::MessageBoxW(hwnd_, L"请先选择要处理的音频文件",
                      L"淡入淡出", MB_OK | MB_ICONWARNING);
        return;
    }

    ClearLog();
    AppendLog(L"开头淡入：" + FormatMilliseconds(fadeInMilliseconds_));
    AppendLog(L"结尾淡出：" + FormatMilliseconds(fadeOutMilliseconds_));
    AppendLog(L"处理文件数量：" + std::to_wstring(fadeInputPaths_.size()));
    ::EnableWindow(controls_.runButton, FALSE);
    taskRunning_ = true;
    activeTaskModule_ = 2;

    const HWND target = hwnd_;
    const std::wstring ffmpegPath = config_.ffmpegPath;
    const std::vector<std::wstring> inputPaths = fadeInputPaths_;
    const int fadeInMilliseconds = fadeInMilliseconds_;
    const int fadeOutMilliseconds = fadeOutMilliseconds_;
    SpawnBackgroundTask([this, target, ffmpegPath, inputPaths, fadeInMilliseconds, fadeOutMilliseconds]() {
        int finalExitCode = 0;
        std::vector<std::wstring> pendingLines;
        pendingLines.reserve(kLogBatchSize);
        for (int index = 0; index < static_cast<int>(inputPaths.size()); ++index) {
            const auto& inputPath = inputPaths[index];
            ConcatListItem info{};
            TryProbeConcatItem(inputPath, info);

            int durationMilliseconds = 0;
            TryProbeDurationMilliseconds(inputPath, durationMilliseconds);
            const int localFadeIn = std::max(0, std::min(fadeInMilliseconds, durationMilliseconds > 0 ? durationMilliseconds : fadeInMilliseconds));
            const int localFadeOut = std::max(0, std::min(fadeOutMilliseconds, durationMilliseconds > 0 ? durationMilliseconds : fadeOutMilliseconds));
            const int fadeOutStart = std::max(0, durationMilliseconds - localFadeOut);
            const std::wstring outputPath = BuildFadeOutputPath(inputPath);

            std::wstringstream filter;
            filter << L"afade=t=in:st=0:d=" << (localFadeIn / 1000.0)
                   << L",afade=t=out:st=" << (fadeOutStart / 1000.0)
                   << L":d=" << (localFadeOut / 1000.0);

            std::wstringstream args;
            args << L"-hide_banner -y -i " << EscapeArgument(inputPath)
                 << L" -af " << EscapeArgument(filter.str())
                 << L" -c:a " << EscapeArgument(info.audioCodec == L"-" ? L"aac" : info.audioCodec);
            if (!info.audioBitrate.empty() && info.audioBitrate != L"-") {
                std::wstring bitrateValue = info.audioBitrate;
                const size_t spacePos = bitrateValue.find(L' ');
                if (spacePos != std::wstring::npos) {
                    bitrateValue = bitrateValue.substr(0, spacePos);
                }
                if (!bitrateValue.empty() &&
                    bitrateValue.find_first_not_of(L"0123456789.") == std::wstring::npos) {
                    bitrateValue += L"k";
                }
                args << L" -b:a " << EscapeArgument(bitrateValue);
            }
            args << L" " << EscapeArgument(outputPath);

            const std::wstring fullCommand = L"\"" + ffmpegPath + L"\" " + args.str();
            QueueLogLine(target, pendingLines, L"输入文件：" + inputPath);
            QueueLogLine(target, pendingLines, L"输出文件：" + outputPath);
            QueueLogLine(target, pendingLines, L"FFmpeg 命令：");
            QueueLogLine(target, pendingLines, fullCommand);

            std::wstring lastErrorLine;
            const int exitCode = RunProcessCapture(
                ffmpegPath,
                args.str(),
                [target, &lastErrorLine](const std::wstring& line) {
                    std::wstring trimmed = line;
                    while (!trimmed.empty() && (trimmed.back() == L'\r' || trimmed.back() == L'\n' || trimmed.back() == L' ')) {
                        trimmed.pop_back();
                    }
                    if (!trimmed.empty()) {
                        lastErrorLine = trimmed;
                    }
                    PostOwnedMessage(target, WM_APP_RUN_LOG, 0, std::make_unique<RunLogMessage>(RunLogMessage{line}));
                });
            if (exitCode != 0) {
                finalExitCode = exitCode;
                auto* status = new FadeItemStatusMessage{};
                status->index = index;
                status->success = false;
                status->errorText = L"文件处理失败：\n" + inputPath + (lastErrorLine.empty() ? L"" : L"\n\n" + lastErrorLine);
                PostOwnedMessage(target, WM_APP_FADE_ITEM_STATUS, 0, std::unique_ptr<FadeItemStatusMessage>(status));
                break;
            }
            auto statusOk = std::make_unique<FadeItemStatusMessage>();
            statusOk->index = index;
            statusOk->success = true;
            statusOk->applyLastOutputPath = true;
            statusOk->outputPath = outputPath;
            PostOwnedMessage(target, WM_APP_FADE_ITEM_STATUS, 0, std::move(statusOk));
        }
        ::PostMessageW(target, WM_APP_RUN_FINISHED, static_cast<WPARAM>(finalExitCode), 0);
    });
}

void MainWindow::StartConvertTask() {
    if (taskRunning_ || runner_.IsRunning()) {
        return;
    }
    if (!ValidateFfmpegPath()) {
        ::MessageBoxW(hwnd_, LoadStringResource(IDS_STATUS_INVALID_FFMPEG).c_str(),
                      LoadStringResource(IDS_SETTINGS_TITLE).c_str(), MB_OK | MB_ICONWARNING);
        return;
    }
    if (convertInputPaths_.empty()) {
        ::MessageBoxW(hwnd_, L"请先选择要转换的文件",
                      L"格式转换", MB_OK | MB_ICONWARNING);
        return;
    }

    ClearLog();
    AppendLog(L"处理文件数量：" + std::to_wstring(convertInputPaths_.size()));
    ::EnableWindow(controls_.runButton, FALSE);
    taskRunning_ = true;
    activeTaskModule_ = 3;

    const HWND target = hwnd_;
    const std::wstring ffmpegPath = config_.ffmpegPath;
    const std::vector<std::wstring> inputPaths = convertInputPaths_;
    SpawnBackgroundTask([this, target, ffmpegPath, inputPaths]() {
        int finalExitCode = 0;
        for (int index = 0; index < static_cast<int>(inputPaths.size()); ++index) {
            const std::wstring& inputPath = inputPaths[index];
            const std::filesystem::path inputFile(inputPath);
            std::wstring extension = inputFile.extension().wstring();
            std::transform(extension.begin(), extension.end(), extension.begin(),
                           [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });

            std::wstring outputPath = BuildConvertOutputPath(inputPath);
            std::wstringstream args;
            if (convertMode_ == 1) {
                if (extension != L".m4a") {
                    finalExitCode = 1;
                    auto status = std::make_unique<FadeItemStatusMessage>();
                    status->index = index;
                    status->success = false;
                    status->errorText = L"文件不是 m4a，无法执行 M4A 转 MP3：\n" + inputPath;
                    PostOwnedMessage(target, WM_APP_FADE_ITEM_STATUS, 0, std::move(status));
                    break;
                }
                ConcatListItem info{};
                TryProbeConcatItem(inputPath, info);
                std::wstring bitrateValue = info.audioBitrate;
                if (bitrateValue.empty() || bitrateValue == L"-") {
                    bitrateValue = L"192k";
                } else {
                    const size_t spacePos = bitrateValue.find(L' ');
                    if (spacePos != std::wstring::npos) {
                        bitrateValue = bitrateValue.substr(0, spacePos);
                    }
                    if (!bitrateValue.empty() &&
                        bitrateValue.find_first_not_of(L"0123456789.") == std::wstring::npos) {
                        bitrateValue += L"k";
                    }
                }
                args << L"-hide_banner -y -i " << EscapeArgument(inputPath)
                     << L" -c:a libmp3lame -b:a " << EscapeArgument(bitrateValue)
                     << L" -map_metadata 0 -id3v2_version 3 "
                     << EscapeArgument(outputPath);
            } else {
                args << L"-hide_banner -y -i " << EscapeArgument(inputPath)
                     << L" -c:v copy -c:a copy "
                     << EscapeArgument(outputPath);
            }

            const std::wstring fullCommand = L"\"" + ffmpegPath + L"\" " + args.str();
            PostOwnedMessage(target, WM_APP_RUN_LOG, 0, std::make_unique<RunLogMessage>(RunLogMessage{L"\x8F93\x5165\x6587\x4EF6\xFF1A" + inputPath}));
            PostOwnedMessage(target, WM_APP_RUN_LOG, 0, std::make_unique<RunLogMessage>(RunLogMessage{L"\x8F93\x51FA\x6587\x4EF6\xFF1A" + outputPath}));
            PostOwnedMessage(target, WM_APP_RUN_LOG, 0, std::make_unique<RunLogMessage>(RunLogMessage{L"FFmpeg \x547D\x4EE4\xFF1A"}));
            PostOwnedMessage(target, WM_APP_RUN_LOG, 0, std::make_unique<RunLogMessage>(RunLogMessage{fullCommand}));


            std::wstring lastErrorLine;
            const int exitCode = RunProcessCapture(
                ffmpegPath,
                args.str(),
                [target, &lastErrorLine](const std::wstring& line) {
                    std::wstring trimmed = line;
                    while (!trimmed.empty() && (trimmed.back() == L'\r' || trimmed.back() == L'\n' || trimmed.back() == L' ')) {
                        trimmed.pop_back();
                    }
                    if (!trimmed.empty()) {
                        lastErrorLine = trimmed;
                    }
                    PostOwnedMessage(target, WM_APP_RUN_LOG, 0, std::make_unique<RunLogMessage>(RunLogMessage{line}));
                });

            if (exitCode != 0) {
                finalExitCode = exitCode;
                auto status = std::make_unique<FadeItemStatusMessage>();
                status->index = index;
                status->success = false;
                status->errorText = L"文件转换失败：\n" + inputPath + (lastErrorLine.empty() ? L"" : L"\n\n" + lastErrorLine);
                PostOwnedMessage(target, WM_APP_FADE_ITEM_STATUS, 0, std::move(status));
                break;
            }

            auto statusOk = std::make_unique<FadeItemStatusMessage>();
            statusOk->index = index;
            statusOk->success = true;
            statusOk->applyLastOutputPath = true;
            statusOk->outputPath = outputPath;
            PostOwnedMessage(target, WM_APP_FADE_ITEM_STATUS, 0, std::move(statusOk));
        }
        ::PostMessageW(target, WM_APP_RUN_FINISHED, static_cast<WPARAM>(finalExitCode), 0);
    });
}

void MainWindow::StartConvertToMp3Task() {
    if (taskRunning_ || runner_.IsRunning()) {
        return;
    }
    if (!ValidateFfmpegPath()) {
        ::MessageBoxW(hwnd_, LoadStringResource(IDS_STATUS_INVALID_FFMPEG).c_str(),
                      LoadStringResource(IDS_SETTINGS_TITLE).c_str(), MB_OK | MB_ICONWARNING);
        return;
    }
    if (convertInputPaths_.empty()) {
        ::MessageBoxW(hwnd_, L"请先选择要转换的文件", L"格式转换", MB_OK | MB_ICONWARNING);
        return;
    }
    const std::wstring inputPath = convertInputPaths_.front();
    std::wstring bitrateValue = convertItem_.audioBitrate;
    if (bitrateValue.empty() || bitrateValue == L"-") {
        bitrateValue = L"192k";
    } else {
        const size_t spacePos = bitrateValue.find(L' ');
        if (spacePos != std::wstring::npos) {
            bitrateValue = bitrateValue.substr(0, spacePos);
        }
        if (!bitrateValue.empty() &&
            bitrateValue.find_first_not_of(L"0123456789.") == std::wstring::npos) {
            bitrateValue += L"k";
        }
    }
    const std::wstring outputPath = BuildConvertMp3OutputPath(inputPath);
    std::wstringstream args;
    args << L"-hide_banner -y -i " << EscapeArgument(inputPath)
         << L" -c:a libmp3lame -b:a " << EscapeArgument(bitrateValue)
         << L" -map_metadata 0 -id3v2_version 3 "
         << EscapeArgument(outputPath);
    const std::wstring fullCommand = L"\"" + config_.ffmpegPath + L"\" " + args.str();
    ClearLog();
    AppendLog(L"输入文件：" + inputPath);
    AppendLog(L"输出文件：" + outputPath);
    AppendLog(L"FFmpeg 命令：");
    AppendLog(fullCommand);
    ::EnableWindow(controls_.convertToMp3Button, FALSE);
    ::EnableWindow(controls_.convertToMp4Button, FALSE);
    ::SetWindowTextW(controls_.convertToMp3Button, L"\x8F6C\x6362\x4E2D...");
    ::SetWindowTextW(controls_.convertToMp4Button, L"\x8F6C\x6362\x4E2D...");
    taskRunning_ = true;
    activeTaskModule_ = 3;

    const HWND target = hwnd_;
    SpawnBackgroundTask([this, target, inputPath, outputPath, args = args.str()]() {
        std::wstring lastErrorLine;
        const int exitCode = RunProcessCapture(
            config_.ffmpegPath,
            args,
            [target, &lastErrorLine](const std::wstring& line) {
                std::wstring trimmed = line;
                while (!trimmed.empty() && (trimmed.back() == L'\r' || trimmed.back() == L'\n' || trimmed.back() == L' ')) {
                    trimmed.pop_back();
                }
                if (!trimmed.empty()) {
                    lastErrorLine = trimmed;
                }
                PostOwnedMessage(target, WM_APP_RUN_LOG, 0, std::make_unique<RunLogMessage>(RunLogMessage{line}));
            });
        auto status = std::make_unique<FadeItemStatusMessage>();
        status->index = 0;
        status->success = exitCode == 0;
        status->applyLastOutputPath = exitCode == 0;
        status->outputPath = outputPath;
        status->applyConvertItemResult = true;
        status->convertResultText = exitCode == 0 ? L"成功" : L"失败";
        status->convertHasError = exitCode != 0;
        if (exitCode != 0) {
            status->errorText = L"文件转换失败：\n" + inputPath + (lastErrorLine.empty() ? L"" : L"\n\n" + lastErrorLine);
        }
        PostOwnedMessage(target, WM_APP_FADE_ITEM_STATUS, 0, std::move(status));
        ::PostMessageW(target, WM_APP_RUN_FINISHED, static_cast<WPARAM>(exitCode), 0);
    });
}

void MainWindow::StartConvertToMp4Task() {
    if (taskRunning_ || runner_.IsRunning()) {
        return;
    }
    if (!ValidateFfmpegPath()) {
        ::MessageBoxW(hwnd_, LoadStringResource(IDS_STATUS_INVALID_FFMPEG).c_str(),
                      LoadStringResource(IDS_SETTINGS_TITLE).c_str(), MB_OK | MB_ICONWARNING);
        return;
    }
    if (convertInputPaths_.empty()) {
        ::MessageBoxW(hwnd_, L"请先选择要转换的文件", L"格式转换", MB_OK | MB_ICONWARNING);
        return;
    }
    if (convertItem_.videoCodec.empty() || convertItem_.videoCodec == L"-") {
        ::MessageBoxW(hwnd_, L"当前文件没有视频轨，不能转为 MP4", L"格式转换", MB_OK | MB_ICONWARNING);
        return;
    }
    const std::wstring inputPath = convertInputPaths_.front();
    const std::wstring outputPath = BuildConvertMp4OutputPath(inputPath);
    std::wstringstream args;
    args << L"-hide_banner -y -i " << EscapeArgument(inputPath)
         << L" -c:v copy -c:a copy "
         << EscapeArgument(outputPath);
    const std::wstring fullCommand = L"\"" + config_.ffmpegPath + L"\" " + args.str();
    ClearLog();
    AppendLog(L"输入文件：" + inputPath);
    AppendLog(L"输出文件：" + outputPath);
    AppendLog(L"FFmpeg 命令：");
    AppendLog(fullCommand);
    ::EnableWindow(controls_.convertToMp3Button, FALSE);
    ::EnableWindow(controls_.convertToMp4Button, FALSE);
    ::SetWindowTextW(controls_.convertToMp3Button, L"\x8F6C\x6362\x4E2D...");
    ::SetWindowTextW(controls_.convertToMp4Button, L"\x8F6C\x6362\x4E2D...");
    taskRunning_ = true;
    activeTaskModule_ = 3;

    const HWND target = hwnd_;
    SpawnBackgroundTask([this, target, inputPath, outputPath, args = args.str()]() {
        std::wstring lastErrorLine;
        const int exitCode = RunProcessCapture(
            config_.ffmpegPath,
            args,
            [target, &lastErrorLine](const std::wstring& line) {
                std::wstring trimmed = line;
                while (!trimmed.empty() && (trimmed.back() == L'\r' || trimmed.back() == L'\n' || trimmed.back() == L' ')) {
                    trimmed.pop_back();
                }
                if (!trimmed.empty()) {
                    lastErrorLine = trimmed;
                }
                PostOwnedMessage(target, WM_APP_RUN_LOG, 0, std::make_unique<RunLogMessage>(RunLogMessage{line}));
            });
        auto status = std::make_unique<FadeItemStatusMessage>();
        status->index = 0;
        status->success = exitCode == 0;
        status->applyLastOutputPath = exitCode == 0;
        status->outputPath = outputPath;
        status->applyConvertItemResult = true;
        status->convertResultText = exitCode == 0 ? L"成功" : L"失败";
        status->convertHasError = exitCode != 0;
        if (exitCode != 0) {
            status->errorText = L"文件转换失败：\n" + inputPath + (lastErrorLine.empty() ? L"" : L"\n\n" + lastErrorLine);
        }
        PostOwnedMessage(target, WM_APP_FADE_ITEM_STATUS, 0, std::move(status));
        ::PostMessageW(target, WM_APP_RUN_FINISHED, static_cast<WPARAM>(exitCode), 0);
    });
}

void MainWindow::AutoDetectInputDuration() {
    const std::wstring inputPath = inputFilePath_;
    if (inputPath.empty() || !FileExists(inputPath)) {
        ResetDurationState();
        return;
    }

    const int requestId = ++durationProbeRequestId_;
    ResetDurationState();
    if (controls_.fileDurationValue != nullptr) {
        ::SetWindowTextW(controls_.fileDurationValue, L"\x6587\x4EF6\x603B\x65F6\x957F\xFF1A\x89E3\x6790\x4E2D...");
    }

    const HWND target = hwnd_;
    SpawnBackgroundTask([this, target, inputPath, requestId]() {
        int durationMilliseconds = 0;
        const bool success = TryProbeDurationMilliseconds(inputPath, durationMilliseconds);
        PostOwnedMessage(target, WM_APP_DURATION_PROBED, 0,
                         std::make_unique<DurationProbeResult>(DurationProbeResult{requestId, success, durationMilliseconds}));
    });
}

void MainWindow::ResetDurationState() {
    durationAvailable_ = false;
    mediaDurationMilliseconds_ = 0;
    rangeStartMilliseconds_ = 0;
    rangeEndMilliseconds_ = 0;
    activeRangeThumb_ = kRangeThumbNone;
    ResetCutResultState();
    UpdateTrackbarPositions();
    UpdateFileDurationDisplay();
    UpdateTimeDisplays();
}

void MainWindow::ResetCutResultState() {
    if (!cutSucceeded_) {
        return;
    }
    cutSucceeded_ = false;
    if (controls_.runButton != nullptr) {
        if (selectedModule_ == 1) {
            ::SetWindowTextW(controls_.runButton, L"开始拼接");
        } else if (selectedModule_ == 2) {
            ::SetWindowTextW(controls_.runButton, L"开始处理");
        } else if (selectedModule_ == 3) {
            ::SetWindowTextW(controls_.runButton, convertMode_ == 0 ? L"开始转 MP4" : L"开始转 MP3");
        } else {
            ::SetWindowTextW(controls_.runButton, LoadStringResource(IDS_BUTTON_RUN_CUT).c_str());
        }
        ::InvalidateRect(controls_.runButton, nullptr, TRUE);
    }
}

void MainWindow::UpdateTrackbarPositions() {
    if (controls_.startTrack != nullptr) {
        ::InvalidateRect(controls_.startTrack, nullptr, TRUE);
    }
}

void MainWindow::UpdateFileDurationDisplay() {
    if (controls_.fileDurationValue == nullptr) {
        return;
    }
    const std::wstring fileTotalDuration = L"\x6587\x4EF6\x603B\x65F6\x957F\xFF1A" +
        (durationAvailable_ ? FormatClockTextNoMilliseconds(mediaDurationMilliseconds_) : L"--:--:--");
    ::SetWindowTextW(controls_.fileDurationValue, fileTotalDuration.c_str());
}

void MainWindow::UpdateTimeDisplays() {
    UpdateTrackbarPositions();
    const std::wstring startText = durationAvailable_ ? FormatMilliseconds(rangeStartMilliseconds_) : L"--:--:--.---";
    const std::wstring endText = durationAvailable_ ? FormatMilliseconds(rangeEndMilliseconds_) : L"--:--:--.---";
    ::SetWindowTextW(controls_.startValue, startText.c_str());
    ::SetWindowTextW(controls_.endValue, endText.c_str());
    const std::wstring cutTotalDuration = L"\x526A\x5207\x603B\x65F6\x95F4\xFF1A" +
        FormatClockTextNoMilliseconds(std::max(0, rangeEndMilliseconds_ - rangeStartMilliseconds_));
    ::SetWindowTextW(controls_.cutDurationValue, cutTotalDuration.c_str());
}

void MainWindow::SetLogExpanded(bool) {
    logExpanded_ = true;
    UpdateLogPanelState();
}

void MainWindow::SpawnBackgroundTask(std::function<void()> task) {
    backgroundThreads_.emplace_back([task = std::move(task)]() mutable {
        task();
    });
}

void MainWindow::RequestBackgroundStop() {
    std::lock_guard<std::mutex> lock(backgroundProcessMutex_);
    for (HANDLE processHandle : backgroundProcesses_) {
        if (processHandle != nullptr) {
            ::TerminateProcess(processHandle, 1);
        }
    }
}

void MainWindow::RegisterBackgroundProcess(HANDLE processHandle) const {
    if (processHandle == nullptr) {
        return;
    }
    std::lock_guard<std::mutex> lock(backgroundProcessMutex_);
    backgroundProcesses_.push_back(processHandle);
}

void MainWindow::UnregisterBackgroundProcess(HANDLE processHandle) const {
    if (processHandle == nullptr) {
        return;
    }
    std::lock_guard<std::mutex> lock(backgroundProcessMutex_);
    auto it = std::remove(backgroundProcesses_.begin(), backgroundProcesses_.end(), processHandle);
    backgroundProcesses_.erase(it, backgroundProcesses_.end());
}

void MainWindow::UpdateLogPanelState() {
    if (hwnd_ != nullptr) {
        RECT rect{};
        ::GetClientRect(hwnd_, &rect);
        LayoutControls(rect.right - rect.left, rect.bottom - rect.top);
    }
}

void MainWindow::UpdatePrimaryActionState() {
    if (controls_.runButton == nullptr) {
        return;
    }
    bool enabled = false;
    if (!taskRunning_) {
        if (selectedModule_ == 0) {
            enabled = !inputFilePath_.empty();
        } else if (selectedModule_ == 1) {
            enabled = !concatInputPaths_.empty();
        } else if (selectedModule_ == 2) {
            enabled = !fadeInputPaths_.empty();
        }
    }
    ::EnableWindow(controls_.runButton, enabled ? TRUE : FALSE);
    ::InvalidateRect(controls_.runButton, nullptr, TRUE);
}

void MainWindow::RefreshConcatList() {
    if (controls_.concatList == nullptr) {
        return;
    }

    const std::vector<ConcatListItem>* items = &concatItems_;
    if (selectedModule_ == 2) {
        items = &fadeItems_;
    } else if (selectedModule_ == 3) {
        items = &convertItems_;
    }
    ::SendMessageW(controls_.concatList, WM_SETREDRAW, FALSE, 0);
    ::SendMessageW(controls_.concatList, LVM_DELETEALLITEMS, 0, 0);
    for (int index = 0; index < static_cast<int>(items->size()); ++index) {
        const ConcatListItem& item = (*items)[index];

        LVITEMW listItem{};
        listItem.mask = LVIF_TEXT;
        listItem.iItem = index;
        listItem.iSubItem = 0;
        listItem.pszText = const_cast<LPWSTR>(item.fileName.c_str());
        ::SendMessageW(controls_.concatList, LVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&listItem));

        const std::vector<std::wstring> columns =
            selectedModule_ == 2
                ? std::vector<std::wstring>{item.durationText, item.audioCodec, item.audioBitrate, item.resultText}
            : selectedModule_ == 3
                ? std::vector<std::wstring>{item.durationText, item.videoCodec, item.audioCodec, item.resultText}
                : std::vector<std::wstring>{item.durationText, item.resolutionText, item.videoCodec,
                                            item.videoBitrate, item.audioCodec, item.audioBitrate};
        for (int subItem = 0; subItem < static_cast<int>(columns.size()); ++subItem) {
            LVITEMW subItemData{};
            subItemData.iItem = index;
            subItemData.iSubItem = subItem + 1;
            subItemData.mask = LVIF_TEXT;
            subItemData.pszText = const_cast<LPWSTR>(columns[subItem].c_str());
            ::SendMessageW(controls_.concatList, LVM_SETITEMW, 0, reinterpret_cast<LPARAM>(&subItemData));
        }
    }
    ::SendMessageW(controls_.concatList, WM_SETREDRAW, TRUE, 0);
    ::InvalidateRect(controls_.concatList, nullptr, TRUE);
}

void MainWindow::UpdateConcatListItem(int index) {
    if (controls_.concatList == nullptr || index < 0) {
        return;
    }

    const std::vector<ConcatListItem>* items = &concatItems_;
    if (selectedModule_ == 2) {
        items = &fadeItems_;
    } else if (selectedModule_ == 3) {
        items = &convertItems_;
    }
    if (index >= static_cast<int>(items->size())) {
        return;
    }

    const ConcatListItem& item = (*items)[index];
    LVITEMW listItem{};
    listItem.mask = LVIF_TEXT;
    listItem.iItem = index;
    listItem.iSubItem = 0;
    listItem.pszText = const_cast<LPWSTR>(item.fileName.c_str());
    ::SendMessageW(controls_.concatList, LVM_SETITEMW, 0, reinterpret_cast<LPARAM>(&listItem));

    const std::vector<std::wstring> columns =
        selectedModule_ == 2
            ? std::vector<std::wstring>{item.durationText, item.audioCodec, item.audioBitrate, item.resultText}
        : selectedModule_ == 3
            ? std::vector<std::wstring>{item.durationText, item.videoCodec, item.audioCodec, item.resultText}
            : std::vector<std::wstring>{item.durationText, item.resolutionText, item.videoCodec,
                                        item.videoBitrate, item.audioCodec, item.audioBitrate};
    for (int subItem = 0; subItem < static_cast<int>(columns.size()); ++subItem) {
        LVITEMW subItemData{};
        subItemData.iItem = index;
        subItemData.iSubItem = subItem + 1;
        subItemData.mask = LVIF_TEXT;
        subItemData.pszText = const_cast<LPWSTR>(columns[subItem].c_str());
        ::SendMessageW(controls_.concatList, LVM_SETITEMW, 0, reinterpret_cast<LPARAM>(&subItemData));
    }
}

void MainWindow::RefreshConvertInfo() {
    const bool hasFile = !convertInputPaths_.empty();
    const ConcatListItem& item = convertItem_;
    ::SetWindowTextW(controls_.convertToMp3Button,
                     taskRunning_ && activeTaskModule_ == 3 ? L"\x8F6C\x6362\x4E2D..." : L"\x8F6C\x4E3A MP3");
    ::SetWindowTextW(controls_.convertToMp4Button,
                     taskRunning_ && activeTaskModule_ == 3 ? L"\x8F6C\x6362\x4E2D..." : L"\x8F6C\x4E3A MP4");
    ::SetWindowTextW(controls_.convertInfoValueFile, hasFile ? item.fileName.c_str() : L"-");
    if (hasFile) {
        const std::wstring extension = std::filesystem::path(convertInputPaths_.front()).extension().wstring();
        ::SetWindowTextW(controls_.convertInfoValueFormat, extension.empty() ? L"-" : extension.c_str());
    } else {
        ::SetWindowTextW(controls_.convertInfoValueFormat, L"-");
    }
    ::SetWindowTextW(controls_.convertInfoValueVideoCodec, hasFile ? item.videoCodec.c_str() : L"-");
    ::SetWindowTextW(controls_.convertInfoValueVideoBitrate, hasFile ? item.videoBitrate.c_str() : L"-");
    ::SetWindowTextW(controls_.convertInfoValueResolution, hasFile ? item.resolutionText.c_str() : L"-");
    ::SetWindowTextW(controls_.convertInfoValueAudioCodec, hasFile ? item.audioCodec.c_str() : L"-");
    ::SetWindowTextW(controls_.convertInfoValueAudioBitrate, hasFile ? item.audioBitrate.c_str() : L"-");

    const bool hasVideo = hasFile && !item.videoCodec.empty() && item.videoCodec != L"-";
    convertCanToMp3_ = hasFile && !taskRunning_;
    convertCanToMp4_ = hasVideo && !taskRunning_;
    ::EnableWindow(controls_.convertToMp3Button, convertCanToMp3_);
    ::EnableWindow(controls_.convertToMp4Button, convertCanToMp4_);
    ::InvalidateRect(controls_.convertToMp3Button, nullptr, TRUE);
    ::InvalidateRect(controls_.convertToMp4Button, nullptr, TRUE);
}

void MainWindow::AdjustRangeValueByMilliseconds(bool adjustStart, int deltaMilliseconds) {
    if (!durationAvailable_ || mediaDurationMilliseconds_ <= 0) {
        return;
    }

    ResetCutResultState();
    if (adjustStart) {
        rangeStartMilliseconds_ = std::clamp(rangeStartMilliseconds_ + deltaMilliseconds, 0, rangeEndMilliseconds_);
    } else {
        rangeEndMilliseconds_ = std::clamp(rangeEndMilliseconds_ + deltaMilliseconds, rangeStartMilliseconds_, mediaDurationMilliseconds_);
    }
    UpdateTimeDisplays();
}

std::wstring MainWindow::FormatMilliseconds(int totalMilliseconds) const {
    if (totalMilliseconds < 0) {
        totalMilliseconds = 0;
    }

    const int hours = totalMilliseconds / 3600000;
    const int minutes = (totalMilliseconds % 3600000) / 60000;
    const int seconds = (totalMilliseconds % 60000) / 1000;
    const int milliseconds = totalMilliseconds % 1000;

    wchar_t buffer[24] = {};
    swprintf_s(buffer, L"%02d:%02d:%02d.%03d", hours, minutes, seconds, milliseconds);
    return buffer;
}

std::wstring MainWindow::FormatClockTextNoMilliseconds(int totalMilliseconds) const {
    if (totalMilliseconds < 0) {
        totalMilliseconds = 0;
    }
    const int totalSeconds = totalMilliseconds / 1000;
    const int hours = totalSeconds / 3600;
    const int minutes = (totalSeconds % 3600) / 60;
    const int seconds = totalSeconds % 60;
    wchar_t buffer[16] = {};
    swprintf_s(buffer, L"%02d:%02d:%02d", hours, minutes, seconds);
    return buffer;
}

std::wstring MainWindow::FormatBitrateText(long long bitrate) const {
    if (bitrate <= 0) {
        return L"-";
    }
    wchar_t buffer[32] = {};
    if (bitrate >= 1000000) {
        swprintf_s(buffer, L"%.2f Mbps", static_cast<double>(bitrate) / 1000000.0);
    } else {
        swprintf_s(buffer, L"%.0f kbps", static_cast<double>(bitrate) / 1000.0);
    }
    return buffer;
}

void MainWindow::SaveWindowPlacement() {
    if (::IsZoomed(hwnd_) || ::IsIconic(hwnd_)) {
        return;
    }

    RECT rect{};
    if (!::GetWindowRect(hwnd_, &rect)) {
        return;
    }

    config_.windowWidth = rect.right - rect.left;
    config_.windowHeight = rect.bottom - rect.top;
    configManager_.SaveIfChanged(config_);
}

std::wstring MainWindow::BuildDefaultOutputPath(const std::wstring& inputPath) const {
    SYSTEMTIME localTime{};
    ::GetLocalTime(&localTime);
    wchar_t timeSuffix[32] = {};
    swprintf_s(timeSuffix, L"_cut_%02d.%02d.%02d",
               localTime.wHour, localTime.wMinute, localTime.wSecond);
    const std::wstring suffix = timeSuffix;
    const size_t slashPos = inputPath.find_last_of(L"\\/");
    const size_t dotPos = inputPath.find_last_of(L'.');

    if (slashPos == std::wstring::npos || dotPos == std::wstring::npos || dotPos < slashPos) {
        return inputPath + suffix;
    }

    return inputPath.substr(0, dotPos) + suffix + inputPath.substr(dotPos);
}

std::wstring MainWindow::BuildOutputPathInFolder(const std::wstring& inputPath, const std::wstring& folderPath) const {
    if (inputPath.empty() || folderPath.empty()) {
        return L"";
    }

    std::filesystem::path input(inputPath);
    std::filesystem::path folder(folderPath);
    SYSTEMTIME localTime{};
    ::GetLocalTime(&localTime);
    wchar_t timeSuffix[32] = {};
    swprintf_s(timeSuffix, L"_cut_%02d.%02d.%02d",
               localTime.wHour, localTime.wMinute, localTime.wSecond);
    const std::wstring suffix = timeSuffix;
    const std::wstring fileName = input.stem().wstring() + suffix + input.extension().wstring();
    return (folder / fileName).wstring();
}

std::wstring MainWindow::BuildConcatOutputPath(const std::vector<std::wstring>& inputPaths) const {
    if (inputPaths.empty()) {
        return L"";
    }
    SYSTEMTIME localTime{};
    ::GetLocalTime(&localTime);
    wchar_t timeSuffix[32] = {};
    swprintf_s(timeSuffix, L"_%02d.%02d.%02d",
               localTime.wHour, localTime.wMinute, localTime.wSecond);
    const std::filesystem::path firstPath(inputPaths.front());
    const std::wstring extension = firstPath.extension().wstring().empty() ? L".mp4" : firstPath.extension().wstring();
    return (firstPath.parent_path() / (L"merged_output" + std::wstring(timeSuffix) + extension)).wstring();
}

std::wstring MainWindow::BuildFadeOutputPath(const std::wstring& inputPath) const {
    if (inputPath.empty()) {
        return L"";
    }
    SYSTEMTIME localTime{};
    ::GetLocalTime(&localTime);
    wchar_t timeSuffix[32] = {};
    swprintf_s(timeSuffix, L"_faded_%02d.%02d.%02d",
               localTime.wHour, localTime.wMinute, localTime.wSecond);
    const std::filesystem::path input(inputPath);
    return (input.parent_path() / (input.stem().wstring() + std::wstring(timeSuffix) + input.extension().wstring())).wstring();
}

std::wstring MainWindow::BuildConvertOutputPath(const std::wstring& inputPath) const {
    if (inputPath.empty()) {
        return L"";
    }
    return BuildConvertMp4OutputPath(inputPath);
}

std::wstring MainWindow::BuildConvertMp3OutputPath(const std::wstring& inputPath) const {
    if (inputPath.empty()) {
        return L"";
    }
    SYSTEMTIME localTime{};
    ::GetLocalTime(&localTime);
    wchar_t timeSuffix[32] = {};
    const std::filesystem::path input(inputPath);
    swprintf_s(timeSuffix, L"_mp3_%02d.%02d.%02d",
               localTime.wHour, localTime.wMinute, localTime.wSecond);
    return (input.parent_path() / (input.stem().wstring() + std::wstring(timeSuffix) + L".mp3")).wstring();
}

std::wstring MainWindow::BuildConvertMp4OutputPath(const std::wstring& inputPath) const {
    if (inputPath.empty()) {
        return L"";
    }
    SYSTEMTIME localTime{};
    ::GetLocalTime(&localTime);
    wchar_t timeSuffix[32] = {};
    const std::filesystem::path input(inputPath);
    swprintf_s(timeSuffix, L"_mp4_%02d.%02d.%02d",
               localTime.wHour, localTime.wMinute, localTime.wSecond);
    return (input.parent_path() / (input.stem().wstring() + std::wstring(timeSuffix) + L".mp4")).wstring();
}

std::wstring MainWindow::BuildFfprobePath() const {
    return config_.ffprobePath;
}

bool MainWindow::TryProbeDurationMilliseconds(const std::wstring& inputPath, int& durationMilliseconds) const {
    durationMilliseconds = 0;

    const std::wstring ffprobePath = BuildFfprobePath();
    if (ffprobePath.empty() || !FileExists(ffprobePath)) {
        return false;
    }

    SECURITY_ATTRIBUTES sa{};
    sa.nLength = sizeof(sa);
    sa.bInheritHandle = TRUE;

    HANDLE readPipe = nullptr;
    HANDLE writePipe = nullptr;
    if (!::CreatePipe(&readPipe, &writePipe, &sa, 0)) {
        return false;
    }
    ::SetHandleInformation(readPipe, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW si{};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    si.wShowWindow = SW_HIDE;
    si.hStdOutput = writePipe;
    si.hStdError = writePipe;

    PROCESS_INFORMATION pi{};
    std::wstring commandLine = L"\"" + ffprobePath +
        L"\" -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 -i " +
        EscapeArgument(inputPath);
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
        &si,
        &pi);

    ::CloseHandle(writePipe);
    if (!created) {
        ::CloseHandle(readPipe);
        return false;
    }

    std::string output;
    std::array<char, 256> buffer{};
    DWORD readBytes = 0;
    while (::ReadFile(readPipe, buffer.data(), static_cast<DWORD>(buffer.size() - 1), &readBytes, nullptr) && readBytes > 0) {
        buffer[readBytes] = '\0';
        output.append(buffer.data(), readBytes);
    }

    ::WaitForSingleObject(pi.hProcess, INFINITE);
    DWORD processExitCode = 1;
    ::GetExitCodeProcess(pi.hProcess, &processExitCode);

    ::CloseHandle(pi.hThread);
    ::CloseHandle(pi.hProcess);
    ::CloseHandle(readPipe);

    if (processExitCode != 0 || output.empty()) {
        return false;
    }

    double bestSeconds = 0.0;
    std::stringstream ss(output);
    std::string token;
    while (ss >> token) {
        char* endPtr = nullptr;
        const double value = std::strtod(token.c_str(), &endPtr);
        if (endPtr != token.c_str() && value > bestSeconds) {
            bestSeconds = value;
        }
    }

    if (bestSeconds <= 0.0) {
        return false;
    }

    durationMilliseconds = static_cast<int>((bestSeconds * 1000.0) + 0.5);
    return durationMilliseconds > 0;
}

bool MainWindow::TryProbeConcatItem(const std::wstring& inputPath, ConcatListItem& item) const {
    item = {};
    item.filePath = inputPath;
    item.fileName = std::filesystem::path(inputPath).filename().wstring();
    item.durationText = L"-";
    item.resolutionText = L"-";
    item.videoCodec = L"-";
    item.videoBitrate = L"-";
    item.audioCodec = L"-";
    item.audioBitrate = L"-";
    item.hasError = false;

    const std::wstring ffprobePath = BuildFfprobePath();
    if (ffprobePath.empty() || !FileExists(ffprobePath) || inputPath.empty() || !FileExists(inputPath)) {
        item.audioCodec = L"错误：FFprobe 无效";
        item.hasError = true;
        return false;
    }

    SECURITY_ATTRIBUTES sa{};
    sa.nLength = sizeof(sa);
    sa.bInheritHandle = TRUE;

    HANDLE readPipe = nullptr;
    HANDLE writePipe = nullptr;
    if (!::CreatePipe(&readPipe, &writePipe, &sa, 0)) {
        item.audioCodec = L"错误：创建探测进程失败";
        item.hasError = true;
        return false;
    }
    ::SetHandleInformation(readPipe, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW si{};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    si.wShowWindow = SW_HIDE;
    si.hStdOutput = writePipe;
    si.hStdError = writePipe;

    PROCESS_INFORMATION pi{};
    std::wstring commandLine = L"\"" + ffprobePath +
        L"\" -v error -show_entries stream=codec_type,codec_name,bit_rate,width,height:stream_disposition=attached_pic:format=duration,bit_rate "
        L"-of default=noprint_wrappers=0 " + EscapeArgument(inputPath);
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
        &si,
        &pi);

    ::CloseHandle(writePipe);
    if (!created) {
        ::CloseHandle(readPipe);
        item.audioCodec = L"错误：启动 FFprobe 失败";
        item.hasError = true;
        return false;
    }

    std::string output;
    std::array<char, 512> buffer{};
    DWORD readBytes = 0;
    while (::ReadFile(readPipe, buffer.data(), static_cast<DWORD>(buffer.size() - 1), &readBytes, nullptr) && readBytes > 0) {
        buffer[readBytes] = '\0';
        output.append(buffer.data(), readBytes);
    }

    ::WaitForSingleObject(pi.hProcess, INFINITE);
    DWORD processExitCode = 1;
    ::GetExitCodeProcess(pi.hProcess, &processExitCode);
    ::CloseHandle(pi.hThread);
    ::CloseHandle(pi.hProcess);
    ::CloseHandle(readPipe);
    if (processExitCode != 0 || output.empty()) {
        item.audioCodec = L"错误：无法读取媒体信息";
        item.hasError = true;
        return false;
    }

    auto trim = [](std::string value) {
        const auto first = value.find_first_not_of(" \t\r\n");
        if (first == std::string::npos) {
            return std::string();
        }
        const auto last = value.find_last_not_of(" \t\r\n");
        return value.substr(first, last - first + 1);
    };
    auto toWide = [](const std::string& value) {
        if (value.empty()) {
            return std::wstring();
        }
        const int length = ::MultiByteToWideChar(CP_UTF8, 0, value.c_str(), -1, nullptr, 0);
        if (length > 0) {
            std::wstring wide(length - 1, L'\0');
            ::MultiByteToWideChar(CP_UTF8, 0, value.c_str(), -1, wide.data(), length);
            return wide;
        }
        const int ansiLength = ::MultiByteToWideChar(CP_ACP, 0, value.c_str(), -1, nullptr, 0);
        std::wstring wide(ansiLength - 1, L'\0');
        ::MultiByteToWideChar(CP_ACP, 0, value.c_str(), -1, wide.data(), ansiLength);
        return wide;
    };

    std::stringstream stream(output);
    std::string line;
    std::string currentSection;
    std::string currentStreamType;
    std::string pendingCodecName;
    bool pendingAttachedPic = false;
    long long pendingBitrate = 0;
    int pendingWidth = 0;
    int pendingHeight = 0;
    long long videoBitrate = 0;
    long long audioBitrate = 0;
    int videoWidth = 0;
    int videoHeight = 0;
    double durationSeconds = 0.0;
    long long formatBitrate = 0;
    while (std::getline(stream, line)) {
        line = trim(line);
        if (line.empty()) {
            continue;
        }
        if (line == "[STREAM]") {
            currentSection = "stream";
            currentStreamType.clear();
            pendingCodecName.clear();
            pendingAttachedPic = false;
            pendingBitrate = 0;
            pendingWidth = 0;
            pendingHeight = 0;
            continue;
        }
        if (line == "[/STREAM]") {
            if (currentStreamType == "video" && !pendingAttachedPic) {
                if (!pendingCodecName.empty() && item.videoCodec == L"-") {
                    item.videoCodec = toWide(pendingCodecName);
                }
                if (pendingBitrate > 0 && videoBitrate <= 0) {
                    videoBitrate = pendingBitrate;
                }
                if (pendingWidth > 0 && pendingHeight > 0 && videoWidth <= 0 && videoHeight <= 0) {
                    videoWidth = pendingWidth;
                    videoHeight = pendingHeight;
                }
            } else if (currentStreamType == "audio") {
                if (!pendingCodecName.empty() && item.audioCodec == L"-") {
                    item.audioCodec = toWide(pendingCodecName);
                }
                if (pendingBitrate > 0 && audioBitrate <= 0) {
                    audioBitrate = pendingBitrate;
                }
            }
            currentSection.clear();
            currentStreamType.clear();
            pendingCodecName.clear();
            pendingAttachedPic = false;
            pendingBitrate = 0;
            pendingWidth = 0;
            pendingHeight = 0;
            continue;
        }
        if (line == "[/FORMAT]") {
            currentSection.clear();
            currentStreamType.clear();
            continue;
        }
        if (line == "[FORMAT]") {
            currentSection = "format";
            continue;
        }

        const std::size_t pos = line.find('=');
        if (pos == std::string::npos) {
            continue;
        }
        const std::string key = line.substr(0, pos);
        const std::string value = trim(line.substr(pos + 1));

        if (currentSection == "stream") {
            if (key == "codec_type") {
                currentStreamType = value;
            } else if (key == "codec_name") {
                pendingCodecName = value;
            } else if (key.find("attached_pic") != std::string::npos) {
                pendingAttachedPic = value == "1";
            } else if (key == "bit_rate") {
                pendingBitrate = _strtoi64(value.c_str(), nullptr, 10);
            } else if (key == "width") {
                pendingWidth = std::atoi(value.c_str());
            } else if (key == "height") {
                pendingHeight = std::atoi(value.c_str());
            }
        } else if (currentSection == "format") {
            if (key == "duration" && durationSeconds <= 0.0) {
                durationSeconds = std::strtod(value.c_str(), nullptr);
            } else if (key == "bit_rate" && formatBitrate <= 0) {
                formatBitrate = _strtoi64(value.c_str(), nullptr, 10);
            }
        }
    }

    if (durationSeconds > 0.0) {
        item.durationText = FormatClockTextNoMilliseconds(static_cast<int>((durationSeconds * 1000.0) + 0.5));
    }
    if (videoWidth > 0 && videoHeight > 0) {
        item.resolutionText = std::to_wstring(videoWidth) + L"x" + std::to_wstring(videoHeight);
    }
    if (videoBitrate <= 0 && formatBitrate > 0 && item.videoCodec != L"-" && item.audioCodec == L"-") {
        videoBitrate = formatBitrate;
    }
    if (audioBitrate <= 0 && formatBitrate > 0 && item.audioCodec != L"-" && item.videoCodec == L"-") {
        audioBitrate = formatBitrate;
    }
    item.videoBitrate = FormatBitrateText(videoBitrate);
    item.audioBitrate = FormatBitrateText(audioBitrate);
    if (item.audioCodec == L"-") {
        item.audioCodec = L"错误：无法识别音频编码";
        item.hasError = true;
        return false;
    }
    return true;
}

int MainWindow::RunProcessCapture(const std::wstring& executablePath,
                                  const std::wstring& arguments,
                                  const std::function<void(const std::wstring&)>& logCallback) const {
    HANDLE readPipe = nullptr;
    PROCESS_INFORMATION processInfo{};
    if (!ProcessUtils::CreateRedirectedProcess(executablePath, arguments, processInfo, readPipe)) {
        return -1;
    }
    RegisterBackgroundProcess(processInfo.hProcess);

    ProcessUtils::DrainProcessOutput(readPipe, [&](const std::wstring& line) {
        if (!cancelBackgroundWork_.load()) {
            logCallback(line);
        }
    });

    ::WaitForSingleObject(processInfo.hProcess, INFINITE);
    DWORD exitCode = 1;
    ::GetExitCodeProcess(processInfo.hProcess, &exitCode);
    UnregisterBackgroundProcess(processInfo.hProcess);
    ::CloseHandle(processInfo.hThread);
    ::CloseHandle(processInfo.hProcess);
    ::CloseHandle(readPipe);
    return static_cast<int>(exitCode);
}

bool MainWindow::BuildConcatListFile(const std::vector<std::wstring>& inputPaths,
                                     std::wstring& listFilePath,
                                     int& fileCount) const {
    listFilePath.clear();
    fileCount = 0;
    if (inputPaths.empty()) {
        return false;
    }

    const std::array<std::wstring, 8> extensions = {L".mp3", L".wav", L".flac", L".m4a", L".aac", L".mp4", L".mkv", L".mov"};
    std::vector<std::filesystem::path> files;
    files.reserve(inputPaths.size());
    for (const auto& inputPath : inputPaths) {
        if (inputPath.empty()) {
            continue;
        }
        std::filesystem::path path(inputPath);
        if (!std::filesystem::is_regular_file(path)) {
            continue;
        }
        std::wstring extension = path.extension().wstring();
        std::transform(extension.begin(), extension.end(), extension.begin(),
                       [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
        if (std::find(extensions.begin(), extensions.end(), extension) != extensions.end()) {
            files.push_back(path);
        }
    }
    if (files.empty()) {
        return false;
    }

    const std::filesystem::path folder = files.front().parent_path();
    for (const auto& file : files) {
        if (file.parent_path() != folder) {
            return false;
        }
    }

    std::error_code errorCode;

    std::filesystem::path tempPath;
    for (int attempt = 0; attempt < 32; ++attempt) {
        std::wstringstream nameBuilder;
        nameBuilder << L"vac_concat_" << ::GetCurrentProcessId() << L"_"
                    << ::GetTickCount64() << L"_" << attempt << L".txt";
        tempPath = folder / nameBuilder.str();
        if (!std::filesystem::exists(tempPath, errorCode)) {
            break;
        }
        tempPath.clear();
    }
    if (tempPath.empty()) {
        return false;
    }

    std::ofstream output(tempPath, std::ios::binary | std::ios::trunc);
    if (!output.is_open()) {
        return false;
    }

    auto escapeForConcat = [](const std::filesystem::path& path) {
        std::string utf8 = path.u8string();
        std::string escaped;
        escaped.reserve(utf8.size());
        for (char ch : utf8) {
            if (ch == '\'') {
                escaped += "'\\''";
            } else {
                escaped += ch;
            }
        }
        return escaped;
    };

    for (const auto& file : files) {
        output << "file '" << escapeForConcat(file.filename()) << "'\n";
        ++fileCount;
    }
    output.close();
    if (!output.good()) {
        std::filesystem::remove(tempPath, errorCode);
        fileCount = 0;
        return false;
    }

    listFilePath = tempPath.wstring();
    return true;
}

bool MainWindow::ValidateFfmpegPath() const {
    return !config_.ffmpegPath.empty() && FileExists(config_.ffmpegPath);
}

void MainWindow::SetDarkTitleBar() const {
    BOOL darkMode = TRUE;
    ::DwmSetWindowAttribute(hwnd_, DWMWA_USE_IMMERSIVE_DARK_MODE, &darkMode, sizeof(darkMode));
}

std::wstring MainWindow::EscapeArgument(const std::wstring& value) {
    return ProcessUtils::QuoteCommandLineArgument(value);
}

LRESULT CALLBACK MainWindow::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    MainWindow* self = nullptr;

    if (message == WM_NCCREATE) {
        auto* createStruct = reinterpret_cast<CREATESTRUCTW*>(lParam);
        self = static_cast<MainWindow*>(createStruct->lpCreateParams);
        self->hwnd_ = hwnd;
        ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    } else {
        self = reinterpret_cast<MainWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self != nullptr) {
        return self->HandleMessage(message, wParam, lParam);
    }

    return ::DefWindowProcW(hwnd, message, wParam, lParam);
}

LRESULT CALLBACK MainWindow::SettingsProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    struct SettingsState {
        MainWindow* owner = nullptr;
        HWND ffmpegLabel = nullptr;
        HWND ffmpegEdit = nullptr;
        HWND ffmpegBrowseButton = nullptr;
        HWND ffprobeLabel = nullptr;
        HWND ffprobeEdit = nullptr;
        HWND ffprobeBrowseButton = nullptr;
        HWND suffixLabel = nullptr;
        HWND suffixEdit = nullptr;
        HWND saveToSourceCheck = nullptr;
        HWND saveButton = nullptr;
        HWND cancelButton = nullptr;
    };

    auto* state = reinterpret_cast<SettingsState*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    if (message == WM_NCCREATE) {
        auto* createStruct = reinterpret_cast<CREATESTRUCTW*>(lParam);
        auto* owner = static_cast<MainWindow*>(createStruct->lpCreateParams);
        state = new SettingsState();
        state->owner = owner;
        ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
    }

    if (state == nullptr) {
        return ::DefWindowProcW(hwnd, message, wParam, lParam);
    }

    switch (message) {
    case WM_CREATE: {
        state->ffmpegLabel = ::CreateWindowExW(0, WC_STATICW, LoadStringResource(IDS_SETTINGS_FFMPEG).c_str(),
                                             WS_CHILD | WS_VISIBLE, 24, 24, 680, 28, hwnd,
                                             reinterpret_cast<HMENU>(IDC_SETTINGS_PATH_LABEL), state->owner->instance_, nullptr);
        state->ffmpegEdit = ::CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDITW, state->owner->config_.ffmpegPath.c_str(),
                                            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 24, 58, 560, 44, hwnd,
                                            reinterpret_cast<HMENU>(IDC_SETTINGS_PATH_EDIT), state->owner->instance_, nullptr);
        state->ffmpegBrowseButton = ::CreateWindowExW(0, WC_BUTTONW, L"\x6D4F\x89C8",
                                                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 598, 58, 120, 36, hwnd,
                                                reinterpret_cast<HMENU>(IDC_SETTINGS_BROWSE), state->owner->instance_, nullptr);
        state->ffprobeLabel = ::CreateWindowExW(0, WC_STATICW, LoadStringResource(IDS_SETTINGS_FFPROBE).c_str(),
                                             WS_CHILD | WS_VISIBLE, 24, 122, 680, 28, hwnd,
                                             reinterpret_cast<HMENU>(IDC_SETTINGS_PROBE_LABEL), state->owner->instance_, nullptr);
        state->ffprobeEdit = ::CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDITW, state->owner->config_.ffprobePath.c_str(),
                                            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 24, 156, 560, 44, hwnd,
                                            reinterpret_cast<HMENU>(IDC_SETTINGS_PROBE_EDIT), state->owner->instance_, nullptr);
        state->ffprobeBrowseButton = ::CreateWindowExW(0, WC_BUTTONW, L"\x6D4F\x89C8",
                                                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 598, 156, 120, 36, hwnd,
                                                reinterpret_cast<HMENU>(IDC_SETTINGS_PROBE_BROWSE), state->owner->instance_, nullptr);
        state->suffixLabel = ::CreateWindowExW(0, WC_STATICW, LoadStringResource(IDS_SETTINGS_SUFFIX).c_str(),
                                             WS_CHILD | WS_VISIBLE, 24, 220, 680, 28, hwnd,
                                             reinterpret_cast<HMENU>(IDC_SETTINGS_SUFFIX_LABEL), state->owner->instance_, nullptr);
        state->suffixEdit = ::CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDITW, state->owner->config_.cutSuffix.c_str(),
                                            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 24, 254, 240, 40, hwnd,
                                            reinterpret_cast<HMENU>(IDC_SETTINGS_SUFFIX_EDIT), state->owner->instance_, nullptr);
        state->saveToSourceCheck = ::CreateWindowExW(0, WC_BUTTONW, LoadStringResource(IDS_CHECK_AUTO_OUTPUT).c_str(),
                                            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 24, 312, 380, 32, hwnd,
                                            reinterpret_cast<HMENU>(IDC_SETTINGS_SAVE_SOURCE_CHECK), state->owner->instance_, nullptr);
        ::SendMessageW(state->saveToSourceCheck, BM_SETCHECK,
                       state->owner->config_.saveToSourceFolder ? BST_CHECKED : BST_UNCHECKED, 0);
        state->saveButton = ::CreateWindowExW(0, WC_BUTTONW, LoadStringResource(IDS_BUTTON_SAVE).c_str(),
                                              WS_CHILD | WS_VISIBLE | BS_OWNERDRAW, 488, 372, 108, 38, hwnd,
                                              reinterpret_cast<HMENU>(IDC_SETTINGS_OK), state->owner->instance_, nullptr);
        state->cancelButton = ::CreateWindowExW(0, WC_BUTTONW, LoadStringResource(IDS_BUTTON_CANCEL).c_str(),
                                                WS_CHILD | WS_VISIBLE | BS_OWNERDRAW, 610, 372, 108, 38, hwnd,
                                                reinterpret_cast<HMENU>(IDC_SETTINGS_CANCEL), state->owner->instance_, nullptr);

        const HWND dialogControls[] = {state->ffmpegLabel, state->ffmpegEdit, state->ffmpegBrowseButton,
            state->ffprobeLabel, state->ffprobeEdit, state->ffprobeBrowseButton,
            state->suffixLabel, state->suffixEdit, state->saveToSourceCheck,
            state->saveButton, state->cancelButton};
        for (HWND control : dialogControls) {
            ::SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(state->owner->uiFont_), TRUE);
        }
        const HFONT defaultGuiFont = static_cast<HFONT>(::GetStockObject(DEFAULT_GUI_FONT));
        ::SendMessageW(state->ffmpegBrowseButton, WM_SETFONT, reinterpret_cast<WPARAM>(defaultGuiFont), TRUE);
        ::SendMessageW(state->ffprobeBrowseButton, WM_SETFONT, reinterpret_cast<WPARAM>(defaultGuiFont), TRUE);
        return 0;
    }
    case WM_CTLCOLORDLG:
    case WM_CTLCOLORSTATIC: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        ::SetTextColor(hdc, state->owner->textColor_);
        ::SetBkColor(hdc, RGB(42, 45, 50));
        return reinterpret_cast<INT_PTR>(state->owner->panelBrush_);
    }
    case WM_CTLCOLOREDIT: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        ::SetTextColor(hdc, state->owner->textColor_);
        ::SetBkColor(hdc, RGB(50, 53, 58));
        return reinterpret_cast<INT_PTR>(state->owner->editBrush_);
    }
    case WM_CTLCOLORBTN: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        ::SetTextColor(hdc, state->owner->buttonTextColor_);
        ::SetBkColor(hdc, state->owner->accentColor_);
        return reinterpret_cast<INT_PTR>(state->owner->buttonBrush_);
    }
    case WM_DRAWITEM: {
        const auto* drawItem = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
        state->owner->DrawOwnerDrawButton(drawItem);
        return TRUE;
    }
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_SETTINGS_BROWSE: {
            OPENFILENAMEW dialog{};
            wchar_t filePath[MAX_PATH] = {};
            dialog.lStructSize = sizeof(dialog);
            dialog.hwndOwner = hwnd;
            dialog.lpstrFile = filePath;
            dialog.nMaxFile = MAX_PATH;
            dialog.lpstrFilter = L"ffmpeg.exe\0ffmpeg.exe\0Executable\0*.exe\0";
            dialog.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
            if (::GetOpenFileNameW(&dialog)) {
                ::SetWindowTextW(state->ffmpegEdit, filePath);
            }
            return 0;
        }
        case IDC_SETTINGS_PROBE_BROWSE: {
            OPENFILENAMEW dialog{};
            wchar_t filePath[MAX_PATH] = {};
            dialog.lStructSize = sizeof(dialog);
            dialog.hwndOwner = hwnd;
            dialog.lpstrFile = filePath;
            dialog.nMaxFile = MAX_PATH;
            dialog.lpstrFilter = L"ffprobe.exe\0ffprobe.exe\0Executable\0*.exe\0";
            dialog.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
            if (::GetOpenFileNameW(&dialog)) {
                ::SetWindowTextW(state->ffprobeEdit, filePath);
            }
            return 0;
        }
        case IDC_SETTINGS_OK: {
            state->owner->config_.ffmpegPath = GetWindowTextString(state->ffmpegEdit);
            state->owner->config_.ffprobePath = GetWindowTextString(state->ffprobeEdit);
            state->owner->config_.cutSuffix = GetWindowTextString(state->suffixEdit);
            state->owner->config_.saveToSourceFolder =
                ::SendMessageW(state->saveToSourceCheck, BM_GETCHECK, 0, 0) == BST_CHECKED;
            if (state->owner->config_.cutSuffix.empty()) {
                state->owner->config_.cutSuffix = L"_cut";
            }
            if (!state->owner->ValidateFfmpegPath()) {
                ::MessageBoxW(hwnd, LoadStringResource(IDS_STATUS_INVALID_FFMPEG).c_str(),
                              LoadStringResource(IDS_SETTINGS_TITLE).c_str(), MB_OK | MB_ICONWARNING);
                return 0;
            }
            if (state->owner->config_.ffprobePath.empty() || !FileExists(state->owner->config_.ffprobePath)) {
                ::MessageBoxW(hwnd, LoadStringResource(IDS_STATUS_INVALID_FFPROBE).c_str(),
                              LoadStringResource(IDS_SETTINGS_TITLE).c_str(), MB_OK | MB_ICONWARNING);
                return 0;
            }
            const bool saved = state->owner->configManager_.SaveIfChanged(state->owner->config_);
            ::MessageBoxW(hwnd,
                          LoadStringResource(saved ? IDS_STATUS_SAVE_OK : IDS_STATUS_SAVE_FAIL).c_str(),
                          LoadStringResource(IDS_SETTINGS_TITLE).c_str(),
                          MB_OK | (saved ? MB_ICONINFORMATION : MB_ICONERROR));
            if (saved) {
                ::DestroyWindow(hwnd);
            }
            return 0;
        }
        case IDC_SETTINGS_CANCEL:
            ::DestroyWindow(hwnd);
            return 0;
        default:
            break;
        }
        break;
    case WM_CLOSE:
        ::DestroyWindow(hwnd);
        return 0;
    case WM_NCDESTROY:
        delete state;
        ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
        return 0;
    default:
        break;
    }

    return ::DefWindowProcW(hwnd, message, wParam, lParam);
}

LRESULT MainWindow::HandleMessage(UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE:
        CreateControls();
        return 0;
    case WM_SIZE:
        LayoutControls(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
        ::RedrawWindow(hwnd_, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        return 0;
    case WM_GETMINMAXINFO: {
        auto* minMaxInfo = reinterpret_cast<MINMAXINFO*>(lParam);
        if (minMaxInfo != nullptr) {
            minMaxInfo->ptMinTrackSize.x = kMinimumWindowWidth;
            minMaxInfo->ptMinTrackSize.y = kMinimumWindowHeight;
            return 0;
        }
        break;
    }
    case WM_ERASEBKGND: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        RECT clientRect{};
        ::GetClientRect(hwnd_, &clientRect);
        ::FillRect(hdc, &clientRect, backgroundBrush_);
        return 1;
    }
    case WM_DRAWITEM: {
        auto* drawItem = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
        if (drawItem == nullptr) {
            return FALSE;
        }
        if (drawItem->CtlID == IDC_SIDEBAR) {
            ::FillRect(drawItem->hDC, &drawItem->rcItem, navBrush_);
            HPEN oldPen = reinterpret_cast<HPEN>(::SelectObject(drawItem->hDC, borderPen_));
            HGDIOBJ oldBrush = ::SelectObject(drawItem->hDC, navBrush_);
            ::RoundRect(drawItem->hDC, drawItem->rcItem.left, drawItem->rcItem.top,
                        drawItem->rcItem.right, drawItem->rcItem.bottom, 18, 18);
            ::SelectObject(drawItem->hDC, oldBrush);
            ::SelectObject(drawItem->hDC, oldPen);
            return TRUE;
        }
        if (drawItem->CtlID == IDC_CUT_GROUP) {
            ::FillRect(drawItem->hDC, &drawItem->rcItem, panelBrush_);
            HPEN oldPen = reinterpret_cast<HPEN>(::SelectObject(drawItem->hDC, borderPen_));
            HGDIOBJ oldBrush = ::SelectObject(drawItem->hDC, panelBrush_);
            ::RoundRect(drawItem->hDC, drawItem->rcItem.left, drawItem->rcItem.top,
                        drawItem->rcItem.right, drawItem->rcItem.bottom, 18, 18);
            ::SelectObject(drawItem->hDC, oldBrush);
            ::SelectObject(drawItem->hDC, oldPen);
            return TRUE;
        }
        DrawOwnerDrawButton(drawItem);
        return TRUE;
    }
    case WM_CTLCOLORLISTBOX:
    case WM_CTLCOLOREDIT: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        ::SetTextColor(hdc, textColor_);
        ::SetBkColor(hdc, RGB(50, 53, 58));
        return reinterpret_cast<INT_PTR>(editBrush_);
    }
    case WM_NOTIFY: {
        auto* header = reinterpret_cast<NMHDR*>(lParam);
        if (header == nullptr) {
            return 0;
        }

        const HWND concatHeader = controls_.concatList != nullptr ? ListView_GetHeader(controls_.concatList) : nullptr;
        if (header->hwndFrom == controls_.concatList) {
            if (header->code == NM_CUSTOMDRAW) {
                auto* draw = reinterpret_cast<LPNMLVCUSTOMDRAW>(lParam);
                switch (draw->nmcd.dwDrawStage) {
                case CDDS_PREPAINT:
                    return CDRF_NOTIFYITEMDRAW;
                case CDDS_ITEMPREPAINT:
                    if (selectedModule_ == 2) {
                        if (draw->nmcd.dwItemSpec < fadeItems_.size() && fadeItems_[draw->nmcd.dwItemSpec].hasError) {
                            draw->clrText = RGB(255, 110, 110);
                        } else {
                            draw->clrText = textColor_;
                        }
                    } else if (selectedModule_ == 3) {
                        if (draw->nmcd.dwItemSpec < convertItems_.size() && convertItems_[draw->nmcd.dwItemSpec].hasError) {
                            draw->clrText = RGB(255, 110, 110);
                        } else {
                            draw->clrText = textColor_;
                        }
                    } else {
                        if (draw->nmcd.dwItemSpec < concatItems_.size() && concatItems_[draw->nmcd.dwItemSpec].hasError) {
                            draw->clrText = RGB(255, 110, 110);
                        } else {
                            draw->clrText = textColor_;
                        }
                    }
                    draw->clrTextBk = RGB(50, 53, 58);
                    return CDRF_NEWFONT | CDRF_NOTIFYPOSTPAINT;
                case CDDS_ITEMPOSTPAINT: {
                    HDC hdc = draw->nmcd.hdc;
                    HPEN gridPen = ::CreatePen(PS_SOLID, 1, RGB(112, 122, 136));
                    HGDIOBJ oldPen = ::SelectObject(hdc, gridPen);

                    RECT rowRect = draw->nmcd.rc;
                    ::MoveToEx(hdc, rowRect.left, rowRect.bottom - 1, nullptr);
                    ::LineTo(hdc, rowRect.right, rowRect.bottom - 1);

                    const int columnCount = Header_GetItemCount(concatHeader);
                    for (int column = 0; column < columnCount - 1; ++column) {
                        RECT subItemRect{};
                        if (ListView_GetSubItemRect(controls_.concatList,
                                                    static_cast<int>(draw->nmcd.dwItemSpec),
                                                    column,
                                                    LVIR_BOUNDS,
                                                    &subItemRect)) {
                            ::MoveToEx(hdc, subItemRect.right - 1, rowRect.top, nullptr);
                            ::LineTo(hdc, subItemRect.right - 1, rowRect.bottom);
                        }
                    }

                    ::SelectObject(hdc, oldPen);
                    ::DeleteObject(gridPen);
                    return CDRF_DODEFAULT;
                }
                default:
                    break;
                }
            }
        } else if (concatHeader != nullptr && header->hwndFrom == concatHeader) {
            if (header->code == NM_CUSTOMDRAW) {
                auto* draw = reinterpret_cast<LPNMCUSTOMDRAW>(lParam);
                switch (draw->dwDrawStage) {
                case CDDS_PREPAINT:
                    return CDRF_NOTIFYITEMDRAW;
                case CDDS_ITEMPREPAINT: {
                    HDC hdc = draw->hdc;
                    RECT rect = draw->rc;
                    ::FillRect(hdc, &rect, panelBrush_);
                    HPEN pen = ::CreatePen(PS_SOLID, 1, borderColor_);
                    HGDIOBJ oldPen = ::SelectObject(hdc, pen);
                    HGDIOBJ oldBrush = ::SelectObject(hdc, ::GetStockObject(HOLLOW_BRUSH));
                    ::Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);
                    ::SelectObject(hdc, oldBrush);
                    ::SelectObject(hdc, oldPen);
                    ::DeleteObject(pen);

                    HDITEMW item{};
                    wchar_t text[128] = {};
                    item.mask = HDI_TEXT | HDI_FORMAT;
                    item.pszText = text;
                    item.cchTextMax = static_cast<int>(std::size(text));
                    ::SendMessageW(concatHeader, HDM_GETITEMW, static_cast<WPARAM>(draw->dwItemSpec),
                                   reinterpret_cast<LPARAM>(&item));

                    RECT textRect = rect;
                    textRect.left += 10;
                    textRect.right -= 10;
                    ::SetBkMode(hdc, TRANSPARENT);
                    ::SetTextColor(hdc, textColor_);
                    HFONT oldFont = nullptr;
                    if (uiFont_ != nullptr) {
                        oldFont = reinterpret_cast<HFONT>(::SelectObject(hdc, uiFont_));
                    }
                    ::DrawTextW(hdc, text, -1, &textRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
                    if (oldFont != nullptr) {
                        ::SelectObject(hdc, oldFont);
                    }
                    return CDRF_SKIPDEFAULT;
                }
                default:
                    break;
                }
            }
        }
        return 0;
    }
    case WM_HSCROLL: {
        const HWND source = reinterpret_cast<HWND>(lParam);
        if (selectedModule_ == 2 && (source == controls_.fadeInTrack || source == controls_.fadeOutTrack)) {
            ResetCutResultState();
            if (source == controls_.fadeInTrack) {
                fadeInMilliseconds_ = static_cast<int>(::SendMessageW(controls_.fadeInTrack, TBM_GETPOS, 0, 0)) * 1000;
                ::SetWindowTextW(controls_.startValue, FormatMilliseconds(fadeInMilliseconds_).c_str());
            } else {
                fadeOutMilliseconds_ = static_cast<int>(::SendMessageW(controls_.fadeOutTrack, TBM_GETPOS, 0, 0)) * 1000;
                ::SetWindowTextW(controls_.endValue, FormatMilliseconds(fadeOutMilliseconds_).c_str());
            }
            return 0;
        }
        break;
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        HWND control = reinterpret_cast<HWND>(lParam);
        ::SetTextColor(hdc, (control == controls_.startLabel || control == controls_.endLabel || control == controls_.inputLabel)
                                ? subtleTextColor_ : textColor_);
        ::SetBkColor(hdc, RGB(42, 45, 50));
        return reinterpret_cast<INT_PTR>(panelBrush_);
    }
    case WM_CTLCOLORBTN: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        ::SetTextColor(hdc, textColor_);
        ::SetBkColor(hdc, RGB(42, 45, 50));
        return reinterpret_cast<INT_PTR>(panelBrush_);
    }
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_SETTINGS_BUTTON:
            OpenSettingsDialog();
            return 0;
        case IDC_CUT_INPUT_BROWSE:
            if (selectedModule_ == 1) {
                OpenConcatFilesDialog();
            } else if (selectedModule_ == 2) {
                OpenFadeFilesDialog();
            } else if (selectedModule_ == 3) {
                OpenConvertFilesDialog();
            } else {
                OpenInputFileDialog();
            }
            return 0;
        case IDC_CUT_OPEN_FOLDER_BUTTON:
            OpenOutputFolder();
            return 0;
        case IDC_CUT_RUN_BUTTON:
            if (selectedModule_ == 1) {
                StartConcatTask();
            } else if (selectedModule_ == 2) {
                StartFadeTask();
            } else {
                StartCutTask();
            }
            return 0;
        case IDC_CONVERT_TO_MP3_BUTTON:
            if (selectedModule_ == 3) {
                StartConvertToMp3Task();
            }
            return 0;
        case IDC_CONVERT_TO_MP4_BUTTON:
            if (selectedModule_ == 3) {
                StartConvertToMp4Task();
            }
            return 0;
        case IDC_CONCAT_CLEAR_BUTTON:
            if (selectedModule_ == 2) {
                fadeInputPaths_.clear();
                fadeItems_.clear();
                lastOutputPath_.clear();
                cutSucceeded_ = false;
                ::SetWindowTextW(controls_.inputEdit, L"");
                ::SetWindowTextW(controls_.runButton, L"开始处理");
                ClearLog();
                UpdatePrimaryActionState();
                RefreshConcatList();
            } else if (selectedModule_ == 3) {
                convertInputPaths_.clear();
                convertItems_.clear();
                convertItem_ = {};
                lastOutputPath_.clear();
                cutSucceeded_ = false;
                ::SetWindowTextW(controls_.inputEdit, L"");
                ClearLog();
                RefreshConvertInfo();
            } else {
                concatInputPaths_.clear();
                concatItems_.clear();
                concatListFilePath_.clear();
                lastOutputPath_.clear();
                cutSucceeded_ = false;
                ::SetWindowTextW(controls_.inputEdit, L"");
                ::SetWindowTextW(controls_.runButton, L"开始拼接");
                ClearLog();
                UpdatePrimaryActionState();
                RefreshConcatList();
            }
            return 0;
        case IDC_CUT_START_DECREASE:
            AdjustRangeValueByMilliseconds(true, -1000);
            return 0;
        case IDC_CUT_START_INCREASE:
            AdjustRangeValueByMilliseconds(true, 1000);
            return 0;
        case IDC_CUT_END_DECREASE:
            AdjustRangeValueByMilliseconds(false, -1000);
            return 0;
        case IDC_CUT_END_INCREASE:
            AdjustRangeValueByMilliseconds(false, 1000);
            return 0;
        case IDC_NAV_CUT:
            cutSucceeded_ = false;
            selectedModule_ = 0;
            ClearLog();
            RefreshModule();
            return 0;
        case IDC_NAV_CONCAT:
            cutSucceeded_ = false;
            selectedModule_ = 1;
            ClearLog();
            RefreshModule();
            return 0;
        case IDC_NAV_FADE:
            cutSucceeded_ = false;
            selectedModule_ = 2;
            ClearLog();
            RefreshModule();
            return 0;
        case IDC_NAV_CONVERT:
            cutSucceeded_ = false;
            selectedModule_ = 3;
            ClearLog();
            RefreshModule();
            return 0;
        default:
            if (HIWORD(wParam) == EN_CHANGE && LOWORD(wParam) == IDC_CUT_INPUT_EDIT &&
                GetFocus() == controls_.inputEdit) {
                AutoDetectInputDuration();
            }
            return 0;
        }
    case WM_APP_RUN_LOG: {
        std::unique_ptr<RunLogMessage> runLog(reinterpret_cast<RunLogMessage*>(lParam));
        AppendLog(runLog->text);
        return 0;
    }
    case WM_APP_RUN_LOG_BATCH: {
        std::unique_ptr<RunLogBatchMessage> runLogBatch(reinterpret_cast<RunLogBatchMessage*>(lParam));
        if (runLogBatch == nullptr) {
            return 0;
        }
        AppendLogLines(runLogBatch->lines);
        return 0;
    }
    case WM_APP_FADE_ITEM_STATUS: {
        std::unique_ptr<FadeItemStatusMessage> status(reinterpret_cast<FadeItemStatusMessage*>(lParam));
        if (status == nullptr) {
            return 0;
        }
        if (status->applyLastOutputPath) {
            lastOutputPath_ = status->outputPath;
        }
        if (status->applyConvertItemResult) {
            convertItem_.resultText = status->convertResultText;
            convertItem_.hasError = status->convertHasError;
        }
        std::vector<ConcatListItem>* items = nullptr;
        if (activeTaskModule_ == 2) {
            items = &fadeItems_;
        } else if (activeTaskModule_ == 3) {
            items = &convertItems_;
        }
        if (items != nullptr && status->index >= 0 && status->index < static_cast<int>(items->size())) {
            (*items)[status->index].resultText = status->success ? L"成功" : L"失败";
            if (!status->success) {
                (*items)[status->index].hasError = true;
            }
            if ((selectedModule_ == 2 || selectedModule_ == 3)) {
                UpdateConcatListItem(status->index);
            }
        }
        if (!status->success && !status->errorText.empty()) {
            ::MessageBoxW(hwnd_, status->errorText.c_str(),
                          activeTaskModule_ == 3 ? L"格式转换失败" : L"淡入淡出失败",
                          MB_OK | MB_ICONERROR);
        }
        return 0;
    }
    case WM_APP_MEDIA_PROBE_ITEM: {
        std::unique_ptr<MediaProbeItemMessage> probe(reinterpret_cast<MediaProbeItemMessage*>(lParam));
        if (probe == nullptr || probe->requestId != mediaProbeRequestId_) {
            return 0;
        }
        if (probe->singleItem) {
            convertItem_ = std::move(probe->item);
            RefreshConvertInfo();
            UpdatePrimaryActionState();
            return 0;
        }
        std::vector<ConcatListItem>* items = nullptr;
        if (probe->module == 1) {
            items = &concatItems_;
        } else if (probe->module == 2) {
            items = &fadeItems_;
        } else if (probe->module == 3) {
            items = &convertItems_;
        }
        if (items != nullptr && probe->index >= 0 && probe->index < static_cast<int>(items->size())) {
            (*items)[probe->index] = std::move(probe->item);
            if ((probe->module == 1 && selectedModule_ == 1) || (probe->module == 2 && selectedModule_ == 2)) {
                UpdateConcatListItem(probe->index);
            }
        }
        return 0;
    }
    case WM_APP_MEDIA_PROBE_FINISHED: {
        std::unique_ptr<MediaProbeFinishedMessage> probe(reinterpret_cast<MediaProbeFinishedMessage*>(lParam));
        if (probe == nullptr || probe->requestId != mediaProbeRequestId_) {
            return 0;
        }
        if (probe->module == 3) {
            RefreshConvertInfo();
            UpdatePrimaryActionState();
        } else if ((probe->module == 1 && selectedModule_ == 1) || (probe->module == 2 && selectedModule_ == 2)) {
            RefreshConcatList();
        }
        return 0;
    }
    case WM_APP_DURATION_PROBED: {
        std::unique_ptr<DurationProbeResult> result(reinterpret_cast<DurationProbeResult*>(lParam));
        if (result == nullptr || result->requestId != durationProbeRequestId_) {
            return 0;
        }
        if (!result->success) {
            ResetDurationState();
            return 0;
        }
        const int durationMilliseconds = std::clamp(result->durationMilliseconds, kTrackMin, kTrackMax);
        durationAvailable_ = durationMilliseconds > 0;
        mediaDurationMilliseconds_ = durationMilliseconds;
        rangeStartMilliseconds_ = 0;
        rangeEndMilliseconds_ = durationMilliseconds;
        UpdateFileDurationDisplay();
        UpdateTrackbarPositions();
        UpdateTimeDisplays();
        return 0;
    }
    case WM_APP_RUN_FINISHED: {
        ::EnableWindow(controls_.runButton, TRUE);
        const int exitCode = static_cast<int>(wParam);
        taskRunning_ = false;
        if (activeTaskModule_ == 1) {
            AppendLog(exitCode == 0 ? L"拼接完成" : L"拼接失败，请检查日志");
            concatListFilePath_.clear();
        } else if (activeTaskModule_ == 2) {
            AppendLog(exitCode == 0 ? L"淡入淡出处理完成" : L"淡入淡出处理失败，请检查日志");
        } else if (activeTaskModule_ == 3) {
            AppendLog(exitCode == 0 ? L"格式转换完成" : L"格式转换失败，请检查日志");
            RefreshConvertInfo();
        } else {
            const unsigned int messageId = exitCode == 0 ? IDS_STATUS_SUCCESS : IDS_STATUS_FAILURE;
            AppendLog(LoadStringResource(messageId));
        }
        if (exitCode == 0) {
            cutSucceeded_ = true;
            ::SetWindowTextW(controls_.runButton, LoadStringResource(IDS_BUTTON_SUCCESS).c_str());
            ::InvalidateRect(controls_.runButton, nullptr, TRUE);
        }
        if (selectedModule_ == 3) {
            ::SetWindowTextW(controls_.convertToMp3Button, L"转为 MP3");
            ::SetWindowTextW(controls_.convertToMp4Button, L"转为 MP4");
            ::InvalidateRect(controls_.convertToMp3Button, nullptr, TRUE);
            ::InvalidateRect(controls_.convertToMp4Button, nullptr, TRUE);
        }
        activeTaskModule_ = -1;
        UpdatePrimaryActionState();
        return 0;
    }
    case WM_DESTROY:
        shuttingDown_.store(true);
        cancelBackgroundWork_.store(true);
        runner_.RequestStop();
        RequestBackgroundStop();
        SaveWindowPlacement();
        ::PostQuitMessage(0);
        return 0;
    default:
        break;
    }

    return ::DefWindowProcW(hwnd_, message, wParam, lParam);
}

void MainWindow::DrawOwnerDrawButton(const DRAWITEMSTRUCT* drawItem) const {
    if (drawItem == nullptr || drawItem->CtlType != ODT_BUTTON) {
        return;
    }

    const bool disabled = (drawItem->itemState & ODS_DISABLED) != 0;
    const bool selected = (drawItem->itemState & ODS_SELECTED) != 0;
    const bool focused = (drawItem->itemState & ODS_FOCUS) != 0;

    const bool primary = drawItem->CtlID == IDC_CUT_RUN_BUTTON || drawItem->CtlID == IDC_SETTINGS_OK;
    const bool successButton = drawItem->CtlID == IDC_CUT_RUN_BUTTON && cutSucceeded_;
    const bool convertActionButton =
        drawItem->CtlID == IDC_CONVERT_TO_MP3_BUTTON || drawItem->CtlID == IDC_CONVERT_TO_MP4_BUTTON;
    const bool greenRunButton =
        drawItem->CtlID == IDC_CUT_RUN_BUTTON &&
        !disabled &&
        ((selectedModule_ == 0 && !inputFilePath_.empty()) ||
         (selectedModule_ == 1 && !concatInputPaths_.empty()) ||
         (selectedModule_ == 2 && !fadeInputPaths_.empty()));
    const bool navButton = drawItem->CtlID == IDC_NAV_CUT || drawItem->CtlID == IDC_NAV_CONCAT ||
        drawItem->CtlID == IDC_NAV_FADE || drawItem->CtlID == IDC_NAV_CONVERT;
    const bool convertModeButton =
        drawItem->CtlID == IDC_CONVERT_MODE_VIDEO || drawItem->CtlID == IDC_CONVERT_MODE_AUDIO;
    const bool settingsBrowseButton =
        drawItem->CtlID == IDC_SETTINGS_BROWSE || drawItem->CtlID == IDC_SETTINGS_PROBE_BROWSE;
    const bool navActive =
        (drawItem->CtlID == IDC_NAV_CUT && selectedModule_ == 0) ||
        (drawItem->CtlID == IDC_NAV_CONCAT && selectedModule_ == 1) ||
        (drawItem->CtlID == IDC_NAV_FADE && selectedModule_ == 2) ||
        (drawItem->CtlID == IDC_NAV_CONVERT && selectedModule_ == 3);
    const bool convertModeActive =
        (drawItem->CtlID == IDC_CONVERT_MODE_VIDEO && convertMode_ == 0) ||
        (drawItem->CtlID == IDC_CONVERT_MODE_AUDIO && convertMode_ == 1);
    const bool convertActionActive =
        (drawItem->CtlID == IDC_CONVERT_TO_MP3_BUTTON && convertCanToMp3_) ||
        (drawItem->CtlID == IDC_CONVERT_TO_MP4_BUTTON && convertCanToMp4_);
    COLORREF fill = primary ? accentColor_ : RGB(40, 49, 68);
    COLORREF border = primary ? accentHotColor_ : RGB(76, 90, 116);
    COLORREF text = buttonTextColor_;
    if (navButton) {
        fill = navActive ? RGB(56, 60, 66) : RGB(44, 47, 52);
        border = navActive ? RGB(110, 136, 166) : RGB(66, 70, 76);
        text = navActive ? RGB(255, 255, 255) : RGB(170, 176, 184);
    }
    if (convertModeButton) {
        fill = convertModeActive ? RGB(76, 90, 108) : RGB(56, 60, 66);
        border = convertModeActive ? RGB(128, 150, 178) : RGB(78, 84, 92);
        text = RGB(228, 231, 235);
    }
    if (convertActionButton) {
        fill = convertActionActive ? RGB(34, 170, 92) : RGB(56, 60, 66);
        border = convertActionActive ? RGB(96, 214, 138) : RGB(78, 84, 92);
        text = RGB(228, 231, 235);
    }
    if (greenRunButton) {
        fill = RGB(34, 170, 92);
        border = RGB(96, 214, 138);
        text = RGB(255, 255, 255);
    }
    if (settingsBrowseButton) {
        fill = RGB(82, 88, 98);
        border = RGB(132, 144, 160);
        text = RGB(255, 255, 255);
    }
    if (successButton) {
        fill = RGB(34, 170, 92);
        border = RGB(96, 214, 138);
    }
    if (navButton && navActive) {
        fill = RGB(78, 84, 92);
        border = RGB(154, 176, 202);
        text = RGB(255, 255, 255);
    } else if (disabled) {
        fill = RGB(48, 52, 60);
        border = RGB(66, 70, 82);
        text = RGB(134, 142, 160);
    } else if (selected) {
        fill = successButton ? RGB(28, 148, 80) : (primary ? RGB(8, 112, 226) : RGB(54, 64, 86));
        border = successButton ? RGB(126, 226, 161) : (primary ? RGB(124, 206, 255) : RGB(102, 118, 148));
        if (navButton) {
            fill = navActive ? RGB(68, 74, 84) : RGB(56, 60, 66);
            border = navActive ? RGB(142, 164, 190) : RGB(92, 100, 112);
        }
        if (convertModeButton) {
            fill = convertModeActive ? RGB(88, 102, 122) : RGB(68, 73, 80);
            border = convertModeActive ? RGB(156, 178, 204) : RGB(108, 116, 126);
        }
        if (convertActionButton) {
            fill = convertActionActive ? RGB(28, 148, 80) : RGB(68, 73, 80);
            border = convertActionActive ? RGB(126, 226, 161) : RGB(108, 116, 126);
        }
        if (greenRunButton) {
            fill = RGB(28, 148, 80);
            border = RGB(126, 226, 161);
        }
        if (settingsBrowseButton) {
            fill = RGB(96, 104, 116);
            border = RGB(164, 176, 192);
        }
    }

    HBRUSH fillBrush = ::CreateSolidBrush(fill);
    HPEN borderPen = ::CreatePen(PS_SOLID, 1, border);
    HGDIOBJ oldBrush = ::SelectObject(drawItem->hDC, fillBrush);
    HGDIOBJ oldPen = ::SelectObject(drawItem->hDC, borderPen);
    ::RoundRect(drawItem->hDC, drawItem->rcItem.left, drawItem->rcItem.top,
                drawItem->rcItem.right, drawItem->rcItem.bottom, 14, 14);
    ::SelectObject(drawItem->hDC, oldPen);
    ::SelectObject(drawItem->hDC, oldBrush);
    ::DeleteObject(borderPen);
    ::DeleteObject(fillBrush);

    wchar_t textBuffer[128] = {};
    ::GetWindowTextW(drawItem->hwndItem, textBuffer, static_cast<int>(std::size(textBuffer)));

    RECT textRect = drawItem->rcItem;
    if (selected) {
        ::OffsetRect(&textRect, 0, 1);
    }
    ::SetBkMode(drawItem->hDC, TRANSPARENT);
    ::SetTextColor(drawItem->hDC, text);
    ::DrawTextW(drawItem->hDC, textBuffer, -1, &textRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    if (focused) {
        RECT focusRect = drawItem->rcItem;
        ::InflateRect(&focusRect, -4, -4);
        ::DrawFocusRect(drawItem->hDC, &focusRect);
    }
}

void MainWindow::DrawSidebarItem(const DRAWITEMSTRUCT* drawItem) const {
    if (drawItem == nullptr || drawItem->itemID == static_cast<UINT>(-1)) {
        return;
    }

    wchar_t textBuffer[128] = {};
    ::SendMessageW(drawItem->hwndItem, LB_GETTEXT, drawItem->itemID, reinterpret_cast<LPARAM>(textBuffer));

    RECT rect = drawItem->rcItem;
    const bool selected = (drawItem->itemState & ODS_SELECTED) != 0;
    const bool focused = (drawItem->itemState & ODS_FOCUS) != 0;

    HBRUSH baseBrush = ::CreateSolidBrush(RGB(20, 26, 36));
    ::FillRect(drawItem->hDC, &rect, baseBrush);
    ::DeleteObject(baseBrush);

    RECT buttonRect = rect;
    ::InflateRect(&buttonRect, -6, -10);

    COLORREF fill = selected ? accentColor_ : RGB(34, 42, 58);
    COLORREF border = selected ? accentHotColor_ : RGB(62, 76, 100);
    COLORREF text = selected ? buttonTextColor_ : RGB(198, 208, 226);

    HBRUSH fillBrush = ::CreateSolidBrush(fill);
    HPEN borderPen = ::CreatePen(PS_SOLID, 1, border);
    HGDIOBJ oldBrush = ::SelectObject(drawItem->hDC, fillBrush);
    HGDIOBJ oldPen = ::SelectObject(drawItem->hDC, borderPen);
    ::RoundRect(drawItem->hDC, buttonRect.left, buttonRect.top, buttonRect.right, buttonRect.bottom, 18, 18);
    ::SelectObject(drawItem->hDC, oldPen);
    ::SelectObject(drawItem->hDC, oldBrush);
    ::DeleteObject(borderPen);
    ::DeleteObject(fillBrush);

    if (selected) {
        RECT accentStrip = buttonRect;
        accentStrip.right = accentStrip.left + 4;
        HBRUSH accentBrush = ::CreateSolidBrush(RGB(182, 224, 255));
        ::FillRect(drawItem->hDC, &accentStrip, accentBrush);
        ::DeleteObject(accentBrush);
    }

    RECT textRect = buttonRect;
    textRect.left += 18;
    textRect.right -= 18;
    ::SetBkMode(drawItem->hDC, TRANSPARENT);
    ::SetTextColor(drawItem->hDC, text);
    ::DrawTextW(drawItem->hDC, textBuffer, -1, &textRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    if (focused) {
        RECT focusRect = buttonRect;
        ::InflateRect(&focusRect, -4, -4);
        ::DrawFocusRect(drawItem->hDC, &focusRect);
    }
}

void MainWindow::PaintFrame(HDC hdc) const {
    if (hdc == nullptr || hwnd_ == nullptr) {
        return;
    }

    RECT clientRect{};
    ::GetClientRect(hwnd_, &clientRect);
    ::FillRect(hdc, &clientRect, backgroundBrush_);

    RECT navRect = {kMargin, kMargin, kMargin + kSidebarWidth, clientRect.bottom - kMargin};
    HGDIOBJ oldBrush = ::SelectObject(hdc, navBrush_);
    HPEN oldPen = reinterpret_cast<HPEN>(::SelectObject(hdc, borderPen_));
    ::RoundRect(hdc, navRect.left, navRect.top, navRect.right, navRect.bottom, 24, 24);

    RECT contentRect = {kMargin * 2 + kSidebarWidth, kMargin, clientRect.right - kMargin, clientRect.bottom - kMargin};
    ::SelectObject(hdc, panelBrush_);
    ::RoundRect(hdc, contentRect.left, contentRect.top, contentRect.right, contentRect.bottom, 24, 24);

    ::SelectObject(hdc, oldPen);
    ::SelectObject(hdc, oldBrush);
}

void MainWindow::PaintRangeSlider(HDC hdc, const RECT& clientRect) const {
    ::FillRect(hdc, &clientRect, panelBrush_);

    RECT trackRect = clientRect;
    trackRect.left += kRangeTrackPadding;
    trackRect.right -= kRangeTrackPadding;
    trackRect.top = clientRect.top + 22;
    trackRect.bottom = trackRect.top + kRangeTrackHeight;

    HBRUSH railBrush = ::CreateSolidBrush(durationAvailable_ ? RGB(51, 66, 92) : RGB(38, 46, 62));
    ::FillRect(hdc, &trackRect, railBrush);
    ::DeleteObject(railBrush);

    HPEN railPen = ::CreatePen(PS_SOLID, 1, borderColor_);
    HGDIOBJ oldPen = ::SelectObject(hdc, railPen);
    HGDIOBJ oldBrush = ::SelectObject(hdc, ::GetStockObject(HOLLOW_BRUSH));
    ::Rectangle(hdc, trackRect.left, trackRect.top, trackRect.right, trackRect.bottom);
    ::SelectObject(hdc, oldBrush);
    ::SelectObject(hdc, oldPen);
    ::DeleteObject(railPen);

    int startX = trackRect.left;
    int endX = trackRect.left;
    if (durationAvailable_ && mediaDurationMilliseconds_ > 0) {
        const int usableWidth = std::max(1, static_cast<int>(trackRect.right - trackRect.left));
        startX = trackRect.left + static_cast<int>((static_cast<long long>(rangeStartMilliseconds_) * usableWidth) / mediaDurationMilliseconds_);
        endX = trackRect.left + static_cast<int>((static_cast<long long>(rangeEndMilliseconds_) * usableWidth) / mediaDurationMilliseconds_);
        endX = std::max(endX, startX);

        RECT fillRect = trackRect;
        fillRect.left = startX;
        fillRect.right = endX;
        HBRUSH fillBrush = ::CreateSolidBrush(accentColor_);
        ::FillRect(hdc, &fillRect, fillBrush);
        ::DeleteObject(fillBrush);
    }

    auto drawThumb = [&](int centerX) {
        RECT thumbRect = {
            centerX - kRangeThumbRadius,
            trackRect.top - (kRangeThumbRadius - 2),
            centerX + kRangeThumbRadius + 1,
            trackRect.top + kRangeThumbRadius + 3
        };
        HBRUSH thumbBrush = ::CreateSolidBrush(durationAvailable_ ? RGB(242, 247, 255) : RGB(108, 118, 136));
        HPEN thumbPen = ::CreatePen(PS_SOLID, 1, durationAvailable_ ? accentHotColor_ : RGB(84, 96, 120));
        HGDIOBJ localOldBrush = ::SelectObject(hdc, thumbBrush);
        HGDIOBJ localOldPen = ::SelectObject(hdc, thumbPen);
        ::Ellipse(hdc, thumbRect.left, thumbRect.top, thumbRect.right, thumbRect.bottom);
        ::SelectObject(hdc, localOldPen);
        ::SelectObject(hdc, localOldBrush);
        ::DeleteObject(thumbBrush);
        ::DeleteObject(thumbPen);
    };

    drawThumb(startX);
    drawThumb(durationAvailable_ && mediaDurationMilliseconds_ > 0 ? endX : trackRect.right);
}

LRESULT MainWindow::HandleRangeSliderMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_ERASEBKGND:
        return 1;
    case WM_PAINT: {
        PAINTSTRUCT paint{};
        HDC hdc = ::BeginPaint(hwnd, &paint);
        RECT clientRect{};
        ::GetClientRect(hwnd, &clientRect);
        PaintRangeSlider(hdc, clientRect);
        ::EndPaint(hwnd, &paint);
        return 0;
    }
    case WM_LBUTTONDOWN: {
        if (!durationAvailable_ || mediaDurationMilliseconds_ <= 0) {
            return 0;
        }
        ResetCutResultState();
        RECT clientRect{};
        ::GetClientRect(hwnd, &clientRect);
        RECT trackRect = clientRect;
        trackRect.left += kRangeTrackPadding;
        trackRect.right -= kRangeTrackPadding;
        const int usableWidth = std::max(1, static_cast<int>(trackRect.right - trackRect.left));
        const int mouseX = GET_X_LPARAM(lParam);
        const int startX = trackRect.left + static_cast<int>((static_cast<long long>(rangeStartMilliseconds_) * usableWidth) / mediaDurationMilliseconds_);
        const int endX = trackRect.left + static_cast<int>((static_cast<long long>(rangeEndMilliseconds_) * usableWidth) / mediaDurationMilliseconds_);
        activeRangeThumb_ = (std::abs(mouseX - startX) <= std::abs(mouseX - endX)) ? kRangeThumbStart : kRangeThumbEnd;
        ::SetCapture(hwnd);
        return HandleRangeSliderMessage(hwnd, WM_MOUSEMOVE, wParam, lParam);
    }
    case WM_MOUSEMOVE: {
        if (::GetCapture() != hwnd || activeRangeThumb_ == kRangeThumbNone || !durationAvailable_ || mediaDurationMilliseconds_ <= 0) {
            return 0;
        }
        RECT clientRect{};
        ::GetClientRect(hwnd, &clientRect);
        RECT trackRect = clientRect;
        trackRect.left += kRangeTrackPadding;
        trackRect.right -= kRangeTrackPadding;
        const int usableWidth = std::max(1, static_cast<int>(trackRect.right - trackRect.left));
        const int mouseX = std::clamp(GET_X_LPARAM(lParam), static_cast<int>(trackRect.left), static_cast<int>(trackRect.right));
        const int value = static_cast<int>(((static_cast<long long>(mouseX - trackRect.left) * mediaDurationMilliseconds_) + (usableWidth / 2)) / usableWidth);
        if (activeRangeThumb_ == kRangeThumbStart) {
            rangeStartMilliseconds_ = std::min(value, rangeEndMilliseconds_);
        } else {
            rangeEndMilliseconds_ = std::max(value, rangeStartMilliseconds_);
        }
        UpdateTimeDisplays();
        ::InvalidateRect(hwnd, nullptr, TRUE);
        return 0;
    }
    case WM_LBUTTONUP:
        if (::GetCapture() == hwnd) {
            ::ReleaseCapture();
        }
        activeRangeThumb_ = kRangeThumbNone;
        return 0;
    case WM_CAPTURECHANGED:
        activeRangeThumb_ = kRangeThumbNone;
        return 0;
    default:
        break;
    }
    return ::DefWindowProcW(hwnd, message, wParam, lParam);
}

LRESULT CALLBACK MainWindow::RangeSliderProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    MainWindow* self = reinterpret_cast<MainWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (message == WM_NCCREATE) {
        auto* createStruct = reinterpret_cast<CREATESTRUCTW*>(lParam);
        self = static_cast<MainWindow*>(createStruct->lpCreateParams);
        ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    }
    if (self != nullptr) {
        return self->HandleRangeSliderMessage(hwnd, message, wParam, lParam);
    }
    return ::DefWindowProcW(hwnd, message, wParam, lParam);
}

