#include "MainWindow.h"

#include <windows.h>

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int showCommand) {
    MainWindow window;
    if (!window.Create(instance, showCommand)) {
        return 1;
    }

    MSG message{};
    while (::GetMessageW(&message, nullptr, 0, 0)) {
        ::TranslateMessage(&message);
        ::DispatchMessageW(&message);
    }

    return static_cast<int>(message.wParam);
}
