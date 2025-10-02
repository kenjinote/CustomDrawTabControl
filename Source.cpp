#include <windows.h>
#include <dwmapi.h>
#include <commctrl.h>
#include "CustomTabControl.h"
#include "CUtil.h"
#include "resource.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "dwmapi.lib")

// グローバル変数
CustomTabControl g_tabControl;
static HWND g_hMainWnd;

LRESULT CALLBACK MainWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE: {
        BOOL isDarkMode = CUtil::IsSystemInDarkTheme() ? TRUE : FALSE;
        DwmSetWindowAttribute(hWnd, 20 /* DWMWA_USE_IMMERSIVE_DARK_MODE */, &isDarkMode, sizeof(isDarkMode));
        g_tabControl.Create(hWnd, 0, 0, 800, 40, 1000, isDarkMode);

        // レイアウトを更新
        RECT rc;
        GetClientRect(hWnd, &rc);
        SendMessage(hWnd, WM_SIZE, 0, MAKELPARAM(rc.right, rc.bottom));
        return 0;
    }
    case WM_SIZE:
        {
		    HWND hTab = g_tabControl.GetHwnd();
            if (IsWindow(hTab)) {
                RECT rc;
                GetClientRect(hWnd, &rc);
                SetWindowPos(hTab, NULL, 0, 0, rc.right, 40, SWP_NOZORDER);
            }
        }
        return 0;
    case WM_COMMAND:
		if (LOWORD(wParam) == ID_ACCELERATOR40001) { // メニューからタブ追加
			static int tabIndex = 1;
			std::wstring title = L"Tab " + std::to_wstring(tabIndex++);
			g_tabControl.AddTab(title);
			g_tabControl.SetCurSel(g_tabControl.GetTabCount() - 1);
		}
        else if (LOWORD(wParam) == ID_ACCELERATOR40002) { // メニューからタブ選択
            // タブ削除
            int curSel = g_tabControl.GetCurSel();
            if (curSel != -1) {
                g_tabControl.SetCurSel(max(0, curSel - 1));
                g_tabControl.RemoveTab(curSel);
            }
        }
        else if (LOWORD(wParam) == ID_ACCELERATOR40005) { //右のタブに切り替える
            int curSel = g_tabControl.GetCurSel();
			if (curSel != -1 && curSel < g_tabControl.GetTabCount() - 1) {
				g_tabControl.SetCurSel(curSel + 1);
			}
        }
        else if (LOWORD(wParam) == ID_ACCELERATOR40006) { //左のタブに切り替える
			int curSel = g_tabControl.GetCurSel();
            if (curSel > 0) {
                g_tabControl.SetCurSel(curSel - 1);
            }
        }
        break;
    case WM_SETTINGCHANGE:
        if (lParam != NULL && wcscmp(L"ImmersiveColorSet", (LPCWSTR)lParam) == 0) {
			BOOL isDarkMode = CUtil::IsSystemInDarkTheme() ? TRUE : FALSE;
            DwmSetWindowAttribute(hWnd, 20 /* DWMWA_USE_IMMERSIVE_DARK_MODE */, &isDarkMode, sizeof(isDarkMode));
            SendMessage(g_tabControl.GetHwnd(), WM_APP, isDarkMode, 0);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hWnd, uMsg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {

    InitCommonControls();

    WNDCLASSEXW wc = { 0 };
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
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

    // アクセラレーター
	HACCEL hAccel = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));
    while (GetMessage(&msg, NULL, 0, 0)) {
		if (!TranslateAccelerator(g_hMainWnd, hAccel, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
    // アクセラレーター削除
	DestroyAcceleratorTable(hAccel);

    return (int)msg.wParam;
}