#include <windows.h>
#include <dwmapi.h>
#include "CustomTabControl.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "dwmapi.lib")

// グローバル変数
CustomTabControl g_tabControl;
static HWND g_hMainWnd;

LRESULT CALLBACK MainWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE: {
        // メインウィンドウの非クライアント領域をダークモードに設定
        BOOL dark = TRUE;
        DwmSetWindowAttribute(hWnd, 20 /* DWMWA_USE_IMMERSIVE_DARK_MODE */, &dark, sizeof(dark));

        // カスタムタブコントロールを作成
        g_tabControl.Create(hWnd, 0, 0, 800, 40, 1000);

        // レイアウトを更新
        RECT rc;
        GetClientRect(hWnd, &rc);
        SendMessage(hWnd, WM_SIZE, 0, MAKELPARAM(rc.right, rc.bottom));
        return 0;
    }
    case WM_SIZE:
        // ウィンドウサイズ変更時にタブコントロールのサイズを調整
        if (g_tabControl.GetCurSel() != -1) {
            RECT rc;
            GetClientRect(hWnd, &rc);
            SetWindowPos(g_tabControl.GetHwnd(), NULL, 0, 0, rc.right, 40, SWP_NOZORDER);
        }
        return 0;
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hWnd, uMsg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {
    WNDCLASSEXW wc = { 0 };
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = CreateSolidBrush(RGB(32, 32, 32));
    wc.lpszClassName = L"CustomTabApp";
    if (!RegisterClassExW(&wc)) return 1;

    g_hMainWnd = CreateWindowExW(
        0, L"CustomTabApp", L"Custom Tab Control",
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
        NULL, NULL, hInstance, NULL
    );

    if (!g_hMainWnd) return 1;

    ShowWindow(g_hMainWnd, nCmdShow);
    UpdateWindow(g_hMainWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}