#include "CustomTabControl.h"
#include <dwmapi.h>
#include <uxtheme.h>
#include <commctrl.h>
#include <algorithm>
#include <math.h>
#include "CUtil.h"

#define TAB_PADDING_X 16
#define TAB_PADDING_Y 8
#define TAB_ROUND_RADIUS 8
#define FONT_SIZE 16

static const WCHAR s_szClassName[] = L"CustomTabControlClass";
static const WCHAR s_szDragClassName[] = L"CustomTabDragClass";
static const WCHAR s_szPopupClassName[] = L"CustomTabPopupClass";
static bool s_classRegistered = false;
static bool s_dragClassRegistered = false;
static bool s_popupClassRegistered = false;

LRESULT CALLBACK CustomTabControl::PopupWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    CustomTabControl* pThis = reinterpret_cast<CustomTabControl*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));
    if (pThis) {
        switch (uMsg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            RECT rcClient;
            GetClientRect(hWnd, &rcClient);

            // 背景の描画を修正
            HBRUSH hBrush = CreateSolidBrush(pThis->m_clrTooltipBg);
            FillRect(hdc, &rcClient, hBrush);
            DeleteObject(hBrush);

            // テキストの描画を修正
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, pThis->m_clrTooltipText);
            SelectObject(hdc, pThis->m_hFont);
            DrawTextW(hdc, pThis->m_popupText.c_str(), -1, &rcClient, DT_SINGLELINE | DT_CENTER | DT_VCENTER);

            EndPaint(hWnd, &ps);
            return 0;
        }
        case WM_DESTROY:
            pThis->m_hPopupWnd = NULL;
            break;
        }
    }
    return DefWindowProcW(hWnd, uMsg, wParam, lParam);
}

void CustomTabControl::RegisterPopupWindowClass(HINSTANCE hInstance) {
    if (!s_popupClassRegistered) {
        WNDCLASSEXW wc = { 0 };
        wc.cbSize = sizeof(WNDCLASSEXW);
        wc.lpfnWndProc = PopupWndProc;
        wc.hInstance = hInstance;
        wc.lpszClassName = s_szPopupClassName;
        RegisterClassExW(&wc);
        s_popupClassRegistered = true;
    }
}

CustomTabControl::CustomTabControl()
    : m_hWnd(NULL), m_hFont(NULL), m_dpi(96), m_selectedTab(0), m_hoveredTab(-1),
    m_hoveredCloseButtonTab(-1), m_pressedCloseButtonTab(-1),
    m_draggedTabIndex(-1), m_isDragging(false),
    m_scrollOffset(0), m_isScrollLeftHovered(false), m_isScrollRightHovered(false),
    m_totalTabsWidth(0), m_scrollButtonWidth(0), m_scrollButtonHeight(0),
    m_hDragWnd(NULL), m_hPopupWnd(NULL), m_isPopupVisible(false) {

    m_tabTitles.push_back(L"Tab 1");
    m_tabTitles.push_back(L"Tab 2");
    m_tabTitles.push_back(L"Tab 3");
    m_tabTitles.push_back(L"Long Tab Title 4");
    m_tabTitles.push_back(L"Another Tab");
    m_tabTitles.push_back(L"Final Tab 6");
    m_tabTitles.push_back(L"Tab 7");

    m_clrBg = RGB(32, 32, 32);
    m_clrText = RGB(220, 220, 220);
    m_clrActiveTab = RGB(50, 50, 50);
    m_clrSeparator = RGB(60, 60, 60);
    m_clrCloseHoverBg = RGB(200, 0, 0);
    m_clrCloseText = RGB(150, 150, 150);
}

CustomTabControl::~CustomTabControl() {
    if (m_hFont) {
        DeleteObject(m_hFont);
    }
    if (m_hPopupWnd) {
        DestroyWindow(m_hPopupWnd);
    }
    DestroyDragWindow();
}

void CustomTabControl::RegisterWindowClass(HINSTANCE hInstance) {
    if (!s_classRegistered) {
        WNDCLASSEXW wc = { 0 };
        wc.cbSize = sizeof(WNDCLASSEXW);
        wc.lpfnWndProc = WndProc;
        wc.hInstance = hInstance;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = CreateSolidBrush(RGB(32, 32, 32));
        wc.lpszClassName = s_szClassName;
        RegisterClassExW(&wc);
        s_classRegistered = true;
    }

    if (!s_dragClassRegistered) {
        WNDCLASSEXW wc = { 0 };
        wc.cbSize = sizeof(WNDCLASSEXW);
        wc.lpfnWndProc = DragWndProc;
        wc.hInstance = hInstance;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)GetStockObject(NULL_BRUSH);
        wc.lpszClassName = s_szDragClassName;
        RegisterClassExW(&wc);
        s_dragClassRegistered = true;
    }
}

HWND CustomTabControl::Create(HWND hParent, int x, int y, int width, int height, UINT_PTR uId, BOOL IsDarkMode) {
    RegisterWindowClass(GetModuleHandle(NULL));
    RegisterPopupWindowClass(GetModuleHandle(NULL));

    m_hWnd = CreateWindowExW(
        0, s_szClassName, 0,
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
        x, y, width, height,
        hParent, (HMENU)uId, GetModuleHandle(NULL), this
    );

    if (m_hWnd) {
        // 独自のポップアップウィンドウを作成
        m_hPopupWnd = CreateWindowExW(
            WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOACTIVATE,
            s_szPopupClassName,
            L"",
            WS_POPUP,
            0, 0, 0, 0,
            m_hWnd, // 親ウィンドウを設定
            NULL,
            GetModuleHandle(NULL),
            NULL
        );
        if (m_hPopupWnd) {
            // GWLP_USERDATAを設定して、ウィンドウプロシージャからCustomTabControlインスタンスにアクセス可能にする
            SetWindowLongPtr(m_hPopupWnd, GWLP_USERDATA, (LONG_PTR)this);
        }

        m_dpi = GetDpiForWindow(m_hWnd);
        int lfHeight = -MulDiv(FONT_SIZE, m_dpi, 72);
        m_hFont = CreateFontW(
            lfHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI"
        );
        SendMessage(m_hWnd, WM_SETFONT, (WPARAM)m_hFont, FALSE);

        TRACKMOUSEEVENT tme;
        tme.cbSize = sizeof(tme);
        tme.dwFlags = TME_HOVER | TME_LEAVE;
        tme.hwndTrack = m_hWnd;
        tme.dwHoverTime = HOVER_DEFAULT;
        _TrackMouseEvent(&tme);

        UpdateTheme(IsDarkMode);
    }
    return m_hWnd;
}

void CustomTabControl::AddTab(const std::wstring& title) {
    m_tabTitles.push_back(title);
    RecalculateTabPositions();
}

void CustomTabControl::RemoveTab(int index) {
    if (index >= 0 && index < (int)m_tabTitles.size()) {
        m_tabTitles.erase(m_tabTitles.begin() + index);
        if (m_selectedTab == index) {
            m_selectedTab = min((int)m_tabTitles.size() - 1, m_selectedTab);
        }
        else if (m_selectedTab > index) {
            m_selectedTab--;
        }
        RecalculateTabPositions();
    }
}

void CustomTabControl::RenameTab(int index, const std::wstring& newTitle) {
    if (index >= 0 && index < (int)m_tabTitles.size()) {
        m_tabTitles[index] = newTitle;
        RecalculateTabPositions();
    }
}

int CustomTabControl::GetCurSel() const {
    return m_selectedTab;
}

void CustomTabControl::SetCurSel(int index) {
    if (index >= 0 && index < (int)m_tabTitles.size()) {
        m_selectedTab = index;

        RECT rcClient;
        GetClientRect(m_hWnd, &rcClient);
        bool showScrollButtons = m_totalTabsWidth > rcClient.right;
        int effectiveClientWidth = showScrollButtons ? (rcClient.right - m_scrollButtonWidth * 2) : rcClient.right;

        int currentX = 0;
        int tabWidth = 0;
        for (size_t i = 0; i <= (size_t)index; ++i) {
            tabWidth = GetTabWidth(i);
            if (i < (size_t)index) {
                currentX += tabWidth;
            }
        }

        if (currentX < m_scrollOffset) {
            m_scrollOffset = currentX;
        }
        else if (currentX + tabWidth > m_scrollOffset + effectiveClientWidth) {
            m_scrollOffset = currentX + tabWidth - effectiveClientWidth;
        }

        InvalidateRect(m_hWnd, NULL, TRUE);
    }
}

int CustomTabControl::GetTabCount() const {
    return (int)m_tabTitles.size();
}

HWND CustomTabControl::GetHwnd() const {
    return m_hWnd;
}

void CustomTabControl::SwitchTabOrder(int index1, int index2) {
    if (index1 == index2 || index1 < 0 || index2 < 0 ||
        index1 >= (int)m_tabTitles.size() || index2 >= (int)m_tabTitles.size()) {
        return;
    }
    std::wstring draggedTabTitle = m_tabTitles[index1];
    m_tabTitles.erase(m_tabTitles.begin() + index1);
    if (index1 < index2) {
        m_tabTitles.insert(m_tabTitles.begin() + index2, draggedTabTitle);
    }
    else {
        m_tabTitles.insert(m_tabTitles.begin() + index2, draggedTabTitle);
    }
    if (m_selectedTab == index1) {
        if (index1 < index2) {
            m_selectedTab = index2 - 1;
        }
        else {
            m_selectedTab = index2;
        }
    }
    else if (m_selectedTab > index1 && m_selectedTab <= index2) {
        m_selectedTab--;
    }
    else if (m_selectedTab < index1&& m_selectedTab >= index2) {
        m_selectedTab++;
    }
    InvalidateRect(m_hWnd, NULL, TRUE);
}

int CustomTabControl::HitTest(int x, int y, bool* isCloseButton, bool* isScrollLeft, bool* isScrollRight) const {
    if (isCloseButton) *isCloseButton = false;
    if (isScrollLeft) *isScrollLeft = false;
    if (isScrollRight) *isScrollRight = false;

    RECT rcClient;
    GetClientRect(m_hWnd, &rcClient);
    bool showScrollButtons = m_totalTabsWidth > rcClient.right;
    int effectiveClientWidth = showScrollButtons ? (rcClient.right - m_scrollButtonWidth * 2) : rcClient.right;

    if (showScrollButtons) {
        if (x >= m_scrollRightRect.left && x <= m_scrollRightRect.right) {
            if (isScrollRight) *isScrollRight = true;
            return -1;
        }
        if (x >= m_scrollLeftRect.left && x <= m_scrollLeftRect.right) {
            if (isScrollLeft) *isScrollLeft = true;
            return -1;
        }
    }

    if (m_isDragging) {
        int totalWidth = 0;
        for (size_t i = 0; i < m_tabTitles.size(); ++i) {
            totalWidth += GetTabWidth(i);
        }

        if (x < -m_scrollOffset) {
            return 0;
        }
        if (x > totalWidth - m_scrollOffset) {
            return (int)m_tabTitles.size() - 1;
        }
    }

    if (x > effectiveClientWidth) {
        return -1;
    }

    int totalWidth = -m_scrollOffset;
    int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
    int closeBtnW = tabHeight;
    int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);

    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        int tabWidth = GetTabWidth(i);
        RECT tabRect = { totalWidth, 0, totalWidth + tabWidth, tabHeight };

        if (x >= tabRect.left && x < tabRect.right &&
            tabRect.right > 0 && tabRect.left < effectiveClientWidth) {
            if (isCloseButton) {
                int closeBtnX = tabRect.right - closeBtnW;
                *isCloseButton = (x >= closeBtnX);
            }
            return (int)i;
        }
        totalWidth += tabWidth;
    }
    return -1;
}

LRESULT CALLBACK CustomTabControl::WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    CustomTabControl* pThis = nullptr;
    if (uMsg == WM_NCCREATE) {
        LPCREATESTRUCTW pCS = reinterpret_cast<LPCREATESTRUCTW>(lParam);
        pThis = reinterpret_cast<CustomTabControl*>(pCS->lpCreateParams);
        SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pThis));
        pThis->m_hWnd = hWnd;
    }
    else {
        pThis = reinterpret_cast<CustomTabControl*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));
    }
    if (pThis) {
        switch (uMsg) {
        case WM_PAINT:
            pThis->OnPaint(hWnd);
            return 0;
        case WM_SIZE:
            pThis->OnSize(hWnd);
            break;
        case WM_LBUTTONDOWN:
            pThis->OnLButtonDown(hWnd, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            return 0;
        case WM_MOUSEMOVE:
            pThis->OnMouseMove(hWnd, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            break;
        case WM_MOUSEHOVER:
            pThis->OnMouseHover(hWnd, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            return 0;
        case WM_LBUTTONUP:
            pThis->OnLButtonUp(hWnd, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            return 0;
        case WM_MOUSELEAVE:
            pThis->OnMouseLeave(hWnd);
            return 0;
        case WM_DPICHANGED:
            pThis->OnDpiChanged(hWnd, LOWORD(wParam));
            return 0;
        case WM_APP:
            pThis->UpdateTheme((BOOL)wParam);
            return 0;
        case WM_DESTROY:
            pThis->m_hWnd = NULL;
            break;
        }
    }
    return DefWindowProcW(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK CustomTabControl::DragWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    CustomTabControl* pThis = reinterpret_cast<CustomTabControl*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));
    if (pThis) {
        switch (uMsg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            pThis->DrawDragWindow(hdc);
            EndPaint(hWnd, &ps);
            return 0;
        }
        case WM_DESTROY:
            pThis->m_hDragWnd = NULL;
            break;
        }
    }
    return DefWindowProcW(hWnd, uMsg, wParam, lParam);
}

void CustomTabControl::OnPaint(HWND hWnd) {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hWnd, &ps);

    RECT clientRect;
    GetClientRect(hWnd, &clientRect);

    HDC hdcMem = CreateCompatibleDC(hdc);
    HBITMAP hbmMem = CreateCompatibleBitmap(hdc, clientRect.right, clientRect.bottom);
    HBITMAP hbmOld = (HBITMAP)SelectObject(hdcMem, hbmMem);

    HBRUSH hBrush = CreateSolidBrush(m_clrBg);
    FillRect(hdcMem, &clientRect, hBrush);
    DeleteObject(hBrush);

    int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);

    m_scrollButtonWidth = MulDiv(30, m_dpi, 96);
    m_scrollButtonHeight = tabHeight;
    bool showScrollButtons = m_totalTabsWidth > clientRect.right;
    int effectiveClientWidth = clientRect.right;
    if (showScrollButtons) {
        effectiveClientWidth -= m_scrollButtonWidth * 2;
    }
    int maxScrollOffset = max(0, m_totalTabsWidth - effectiveClientWidth);
    m_scrollOffset = min(maxScrollOffset, max(0, m_scrollOffset));

    RECT tabsDrawingRect = clientRect;
    if (showScrollButtons) {
        tabsDrawingRect.right = clientRect.right - m_scrollButtonWidth * 2;
    }
    IntersectClipRect(hdcMem, tabsDrawingRect.left, tabsDrawingRect.top, tabsDrawingRect.right, tabsDrawingRect.bottom);

    int currentX = -m_scrollOffset;
    std::vector<RECT> tabRects(m_tabTitles.size());
    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        int xPos = currentX;
        int tabWidth = GetTabWidth(i);

        if (m_isDragging) {
            if ((int)i == m_draggedTabIndex) {
                currentX += tabWidth;
                continue;
            }
            else {
                int newPos = -1;
                int draggedTabWidth = GetTabWidth(m_draggedTabIndex);

                if (m_draggedTabIndex < m_hoveredTab) {
                    if ((int)i > m_draggedTabIndex && (int)i <= m_hoveredTab) {
                        xPos -= draggedTabWidth;
                    }
                }
                else if (m_draggedTabIndex > m_hoveredTab) {
                    if ((int)i >= m_hoveredTab && (int)i < m_draggedTabIndex) {
                        xPos += draggedTabWidth;
                    }
                }
            }
        }

        tabRects[i] = { xPos, 0, xPos + tabWidth, tabHeight };
        bool isActive = ((int)i == m_selectedTab);
        bool isHovered = ((int)i == m_hoveredTab);
        bool isCloseHovered = ((int)i == m_hoveredCloseButtonTab);

        DrawTab(hdcMem, i, tabRects[i], isActive, isHovered, isCloseHovered);

        currentX += tabWidth;
    }

    SelectClipRgn(hdcMem, NULL);

    if (showScrollButtons) {
        m_scrollLeftRect = { clientRect.right - m_scrollButtonWidth * 2, 0, clientRect.right - m_scrollButtonWidth, m_scrollButtonHeight };
        m_scrollRightRect = { clientRect.right - m_scrollButtonWidth, 0, clientRect.right, m_scrollButtonHeight };
        HBRUSH hScrollBrush = CreateSolidBrush(m_isScrollLeftHovered ? m_clrScrollButtonHoverBg : m_clrBg);
        FillRect(hdcMem, &m_scrollLeftRect, hScrollBrush);
        DeleteObject(hScrollBrush);
        POINT triangleLeft[] = { {m_scrollLeftRect.left + MulDiv(10, m_dpi, 96), m_scrollLeftRect.top + MulDiv(15, m_dpi, 96)},{m_scrollLeftRect.left + MulDiv(15, m_dpi, 96), m_scrollLeftRect.top + MulDiv(10, m_dpi, 96)},{m_scrollLeftRect.left + MulDiv(15, m_dpi, 96), m_scrollLeftRect.top + MulDiv(20, m_dpi, 96)} };
        HBRUSH hTriangleBrush = CreateSolidBrush(m_clrText);
        SelectObject(hdcMem, hTriangleBrush);
        Polygon(hdcMem, triangleLeft, 3);
        DeleteObject(hTriangleBrush);
        hScrollBrush = CreateSolidBrush(m_isScrollRightHovered ? m_clrScrollButtonHoverBg : m_clrBg);
        FillRect(hdcMem, &m_scrollRightRect, hScrollBrush);
        DeleteObject(hScrollBrush);
        POINT triangleRight[] = { {m_scrollRightRect.left + MulDiv(15, m_dpi, 96), m_scrollRightRect.top + MulDiv(20, m_dpi, 96)},{m_scrollRightRect.left + MulDiv(20, m_dpi, 96), m_scrollRightRect.top + MulDiv(15, m_dpi, 96)},{m_scrollRightRect.left + MulDiv(15, m_dpi, 96), m_scrollRightRect.top + MulDiv(10, m_dpi, 96)} };
        hTriangleBrush = CreateSolidBrush(m_clrText);
        SelectObject(hdcMem, hTriangleBrush);
        Polygon(hdcMem, triangleRight, 3);
        DeleteObject(hTriangleBrush);
    }

    BitBlt(hdc, 0, 0, clientRect.right, clientRect.bottom, hdcMem, 0, 0, SRCCOPY);

    SelectObject(hdcMem, hbmOld);
    DeleteObject(hbmMem);
    DeleteDC(hdcMem);

    EndPaint(hWnd, &ps);
}

void CustomTabControl::DrawTab(HDC hdc, int index, const RECT& rect, bool isActive, bool isHovered, bool isCloseHovered) {
    RECT rc = rect;
    COLORREF bgColor = isActive ? m_clrActiveTab : m_clrBg;
    if (isHovered && !isActive) {
        bgColor = m_clrHoverBg;
    }

    HBRUSH hBrush = CreateSolidBrush(bgColor);
    HPEN hPen = CreatePen(PS_SOLID, 1, m_clrSeparator);
    HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, hBrush);
    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);

    int radius = MulDiv(TAB_ROUND_RADIUS, m_dpi, 96);
    if (isActive) rc.bottom += 1;
    HRGN hRgn = CreateRoundRectRgn(rc.left, rc.top, rc.right + 1, rc.bottom + 1, radius, radius);
    HRGN hRectRgn = CreateRectRgn(rc.left, rc.top + radius, rc.right + 1, rc.bottom + 1);
    CombineRgn(hRgn, hRgn, hRectRgn, RGN_OR);
    DeleteObject(hRectRgn);

    SelectClipRgn(hdc, hRgn);
    FillRect(hdc, &rc, hBrush);
    SelectClipRgn(hdc, NULL);
    DeleteObject(hRgn);

    if (!isActive) {
        SelectObject(hdc, hPen);
        MoveToEx(hdc, rc.left + radius, rc.top, NULL);
        LineTo(hdc, rc.right - radius, rc.top);
        MoveToEx(hdc, rc.right - 1, rc.top + radius, NULL); LineTo(hdc, rc.right - 1, rc.bottom);
        MoveToEx(hdc, rc.left, rc.top + radius, NULL); LineTo(hdc, rc.left, rc.bottom);
    }

    SelectObject(hdc, hOldBrush);
    SelectObject(hdc, hOldPen);
    DeleteObject(hBrush);
    DeleteObject(hPen);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, m_clrText);
    SelectObject(hdc, m_hFont);
    RECT rcText = rect;
    rcText.left += MulDiv(TAB_PADDING_X, m_dpi, 96) / 2;
    int closeBtnW = rect.bottom - rect.top;
    rcText.right -= closeBtnW;
    DrawTextW(hdc, m_tabTitles[index].c_str(), -1, &rcText, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);

    int closeBtnX = rect.right - closeBtnW;
    RECT rcCloseRect = { closeBtnX, rect.top, rect.right, rect.bottom };

    bool isPressed = ((int)index == m_pressedCloseButtonTab);
    // ホバー時にm_clrCloseButtonHoverBgを使用
    HBRUSH hCloseBrush = (isCloseHovered || isPressed) ? CreateSolidBrush(m_clrCloseButtonHoverBg) : (HBRUSH)GetStockObject(NULL_BRUSH);

    if (isCloseHovered || isPressed) {
        FillRect(hdc, &rcCloseRect, hCloseBrush);
    }
    DeleteObject(hCloseBrush);

    COLORREF oldTextColor = SetTextColor(hdc, (isCloseHovered || isPressed) ? RGB(255, 255, 255) : m_clrCloseText);
    HPEN hClosePen = CreatePen(PS_SOLID, 1, (isCloseHovered || isPressed) ? m_clrText : m_clrCloseText); // ★ 修正: ホバー時のX印の色をm_clrTextに
    HPEN hOldClosePen = (HPEN)SelectObject(hdc, hClosePen);

    int crossPadding = MulDiv(8, m_dpi, 96);
    int x1 = rcCloseRect.left + crossPadding;
    int y1 = rcCloseRect.top + crossPadding;
    int x2 = rcCloseRect.right - crossPadding;
    int y2 = rcCloseRect.bottom - crossPadding;
    MoveToEx(hdc, x1, y1, NULL); LineTo(hdc, x2, y2);
    MoveToEx(hdc, x1, y2, NULL); LineTo(hdc, x2, y1);

    SetTextColor(hdc, oldTextColor);
    SelectObject(hdc, hOldClosePen);
    DeleteObject(hClosePen);
}

void CustomTabControl::OnSize(HWND hWnd) {
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);
    int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
    SetWindowPos(hWnd, NULL, 0, 0, rcClient.right, tabHeight, SWP_NOZORDER);
    RecalculateTabPositions();
}

void CustomTabControl::OnLButtonDown(HWND hWnd, int x, int y) {
    bool isClose = false;
    bool isScrollLeft = false;
    bool isScrollRight = false;
    int index = HitTest(x, y, &isClose, &isScrollLeft, &isScrollRight);

    HideCustomTooltip();

    if (isScrollLeft) {
        m_scrollOffset = max(0, m_scrollOffset - 50);
        InvalidateRect(hWnd, NULL, TRUE);
        return;
    }

    if (isScrollRight) {
        RECT rcClient;
        GetClientRect(hWnd, &rcClient);
        int effectiveClientWidth = rcClient.right;
        if (m_totalTabsWidth > rcClient.right) {
            effectiveClientWidth -= m_scrollButtonWidth * 2;
        }
        int maxScrollOffset = max(0, m_totalTabsWidth - effectiveClientWidth);
        m_scrollOffset = min(maxScrollOffset, m_scrollOffset + 50);
        InvalidateRect(hWnd, NULL, TRUE);
        return;
    }

    if (index != -1) {
        if (isClose) {
            m_pressedCloseButtonTab = index;
            InvalidateRect(hWnd, NULL, FALSE);
            SetCapture(hWnd);
        }
        else {
            m_selectedTab = index;
            m_draggedTabIndex = index;
            m_dragStartPos.x = x;
            m_dragStartPos.y = y;
            SetCapture(hWnd);

            InvalidateRect(hWnd, NULL, TRUE);
        }
    }
}

void CustomTabControl::OnMouseMove(HWND hWnd, int x, int y) {
    bool isClose = false;
    bool isScrollLeft = false;
    bool isScrollRight = false;
    int newHoveredTab = HitTest(x, y, &isClose, &isScrollLeft, &isScrollRight);

    // 新しいホバー状態を直接設定
    int oldHoveredTab = m_hoveredTab;
    int oldHoveredCloseButtonTab = m_hoveredCloseButtonTab;

    if (m_isDragging) {
        isClose = false;
    }
    m_hoveredTab = newHoveredTab;

    m_hoveredCloseButtonTab = isClose ? m_hoveredTab : -1;

    // ホバー状態が変化した場合のみ再描画
    if (oldHoveredTab != m_hoveredTab || oldHoveredCloseButtonTab != m_hoveredCloseButtonTab || isScrollLeft != m_isScrollLeftHovered || isScrollRight != m_isScrollRightHovered) {
        m_isScrollLeftHovered = isScrollLeft;
        m_isScrollRightHovered = isScrollRight;
        InvalidateRect(hWnd, NULL, FALSE);
    }

    if (m_draggedTabIndex != -1 && GetCapture() == hWnd && m_pressedCloseButtonTab == -1) {
        if (!m_isDragging) {
            if (abs(x - m_dragStartPos.x) > GetSystemMetrics(SM_CXDRAG) ||
                abs(y - m_dragStartPos.y) > GetSystemMetrics(SM_CYDRAG)) {
                m_isDragging = true;
                CreateDragWindow(m_draggedTabIndex);
            }
        }
        else {
            POINT pt;
            GetCursorPos(&pt);
            int tabWidth = GetTabWidth(m_draggedTabIndex);
            int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
            SetWindowPos(m_hDragWnd, NULL, pt.x - tabWidth / 2, pt.y - tabHeight / 2, tabWidth, tabHeight, SWP_NOZORDER | SWP_NOACTIVATE);

            InvalidateRect(hWnd, NULL, FALSE);
        }
    }
    else if (m_pressedCloseButtonTab != -1 && GetCapture() == hWnd) {
        bool isCloseBtnHoveredNow = isClose && (newHoveredTab == m_pressedCloseButtonTab);
        if (isCloseBtnHoveredNow != (m_hoveredCloseButtonTab != -1)) {
            m_hoveredCloseButtonTab = isCloseBtnHoveredNow ? m_pressedCloseButtonTab : -1;
            InvalidateRect(hWnd, NULL, FALSE);
        }
    }

    TRACKMOUSEEVENT tme;
    tme.cbSize = sizeof(tme);
    tme.dwFlags = TME_HOVER | TME_LEAVE;
    tme.hwndTrack = hWnd;
    tme.dwHoverTime = HOVER_DEFAULT;
    _TrackMouseEvent(&tme);
}

void CustomTabControl::OnMouseHover(HWND hWnd, int x, int y) {
    bool isClose = false;
    bool isScrollLeft = false;
    bool isScrollRight = false;
    int index = HitTest(x, y, &isClose, &isScrollLeft, &isScrollRight);

    if (index != -1 && !isClose && !m_isDragging) {
        ShowCustomTooltip(index, x, y);
    }
    else {
        HideCustomTooltip();
    }
}

void CustomTabControl::OnLButtonUp(HWND hWnd, int x, int y) {
    if (m_hDragWnd) {
        DestroyDragWindow();
    }

    ReleaseCapture();

    if (m_isDragging) {
        bool isClose = false;
        bool isScrollLeft = false;
        bool isScrollRight = false;
        int dropIndex = -1;
        RECT rcClient;
        GetClientRect(hWnd, &rcClient);
        int totalTabsWidth = 0;
        for (size_t i = 0; i < m_tabTitles.size(); ++i) {
            totalTabsWidth += GetTabWidth(i);
        }
        if (m_hoveredTab == -1 && x > totalTabsWidth - m_scrollOffset) {
            dropIndex = (int)m_tabTitles.size() - 1;
        }
        else {
            dropIndex = HitTest(x, y, &isClose, &isScrollLeft, &isScrollRight);
        }
        if (dropIndex == -1) {
            dropIndex = m_draggedTabIndex;
        }
        if (dropIndex != m_draggedTabIndex) {
            SwitchTabOrder(m_draggedTabIndex, dropIndex);
        }
        SetCurSel(dropIndex);
        RecalculateTabPositions();
    }
    else if (m_pressedCloseButtonTab != -1) {
        bool isClose = false;
        bool isScrollLeft = false;
        bool isScrollRight = false;
        int index = HitTest(x, y, &isClose, &isScrollLeft, &isScrollRight);

        if (index != -1 && isClose && index == m_pressedCloseButtonTab) {
            RemoveTab(index);
        }
    }

    m_draggedTabIndex = -1;
    m_isDragging = false;
    m_pressedCloseButtonTab = -1;
    InvalidateRect(hWnd, NULL, TRUE);
}

void CustomTabControl::OnMouseLeave(HWND hWnd) {
    if (m_hoveredTab != -1 || m_isScrollLeftHovered || m_isScrollRightHovered || m_hoveredCloseButtonTab != -1) {
        m_hoveredTab = -1;
        m_hoveredCloseButtonTab = -1;
        m_isScrollLeftHovered = false;
        m_isScrollRightHovered = false;
        HideCustomTooltip();
        InvalidateRect(hWnd, NULL, FALSE);
    }
}

void CustomTabControl::OnDpiChanged(HWND hWnd, int dpi) {
    m_dpi = dpi;
    if (m_hFont) {
        DeleteObject(m_hFont);
    }
    int lfHeight = -MulDiv(FONT_SIZE, m_dpi, 72);
    m_hFont = CreateFontW(
        lfHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI"
    );
    SendMessage(m_hWnd, WM_SETFONT, (WPARAM)m_hFont, FALSE);
    InvalidateRect(hWnd, NULL, TRUE);
}

void CustomTabControl::RecalculateTabPositions() {
    m_totalTabsWidth = 0;
    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        m_totalTabsWidth += GetTabWidth(i);
    }
    InvalidateRect(m_hWnd, NULL, TRUE);
}

int CustomTabControl::GetTabWidth(int index) const {
    if (index < 0 || index >= (int)m_tabTitles.size()) {
        return 0;
    }
    int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
    int closeBtnW = tabHeight;
    int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);

    HDC hdc = GetDC(m_hWnd);
    SelectObject(hdc, m_hFont);
    SIZE size;
    GetTextExtentPoint32W(hdc, m_tabTitles[index].c_str(), (int)m_tabTitles[index].length(), &size);
    ReleaseDC(m_hWnd, hdc);

    return size.cx + tabPaddingX + closeBtnW;
}

void CustomTabControl::CreateDragWindow(int tabIndex) {
    if (m_hDragWnd) {
        return;
    }

    POINT ptCursor;
    GetCursorPos(&ptCursor);

    int tabWidth = GetTabWidth(tabIndex);
    int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);

    m_hDragWnd = CreateWindowExW(
        WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
        s_szDragClassName, 0,
        WS_POPUP,
        ptCursor.x - tabWidth / 2, ptCursor.y - tabHeight / 2, tabWidth, tabHeight,
        NULL, NULL, GetModuleHandle(NULL), this
    );

    ShowWindow(m_hDragWnd, SW_SHOWNOACTIVATE);

    if (m_hDragWnd) {
        BLENDFUNCTION blend = { 0 };
        blend.BlendOp = AC_SRC_OVER;
        blend.SourceConstantAlpha = 180;
        blend.AlphaFormat = 0;

        HDC hdcScreen = GetDC(NULL);
        HDC hdcMem = CreateCompatibleDC(hdcScreen);
        HBITMAP hbmTab = CreateCompatibleBitmap(hdcScreen, tabWidth, tabHeight);
        HBITMAP hbmOld = (HBITMAP)SelectObject(hdcMem, hbmTab);

        RECT rc = { 0, 0, tabWidth, tabHeight };
        HBRUSH hClearBrush = CreateSolidBrush(m_clrBg);
        FillRect(hdcMem, &rc, hClearBrush);
        DeleteObject(hClearBrush);

        DrawTab(hdcMem, tabIndex, rc, true, false, false);

        POINT ptZero = { 0, 0 };
        SIZE sizeTab = { tabWidth, tabHeight };
        UpdateLayeredWindow(m_hDragWnd, hdcScreen, NULL, &sizeTab, hdcMem, &ptZero, 0, &blend, ULW_ALPHA);

        SelectObject(hdcMem, hbmOld);
        DeleteObject(hbmTab);
        DeleteDC(hdcMem);
        ReleaseDC(NULL, hdcScreen);
    }
}

void CustomTabControl::DestroyDragWindow() {
    if (m_hDragWnd) {
        DestroyWindow(m_hDragWnd);
        m_hDragWnd = NULL;
    }
}

void CustomTabControl::DrawDragWindow(HDC hdc) {
    if (m_draggedTabIndex < 0 || m_draggedTabIndex >= m_tabTitles.size()) {
        return;
    }

    int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
    int closeBtnW = tabHeight;
    int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);

    RECT rc;
    GetClientRect(m_hDragWnd, &rc);
    int tabWidth = rc.right;

    HBRUSH hBrush = CreateSolidBrush(m_clrActiveTab);
    FillRect(hdc, &rc, hBrush);
    DeleteObject(hBrush);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, m_clrText);
    SelectObject(hdc, m_hFont);

    RECT rcText = rc;
    rcText.left += MulDiv(TAB_PADDING_X, m_dpi, 96) / 2;
    rcText.right -= closeBtnW;
    DrawTextW(hdc, m_tabTitles[m_draggedTabIndex].c_str(), -1, &rcText, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);

    int closeBtnX = tabWidth - closeBtnW;
    RECT rcCloseRect = { closeBtnX, rc.top, rc.right, rc.bottom };
    HBRUSH hCloseBrush = CreateSolidBrush(RGB(96, 96, 96));
    FillRect(hdc, &rcCloseRect, hCloseBrush);
    DeleteObject(hCloseBrush);

    COLORREF oldTextColor = SetTextColor(hdc, RGB(255, 255, 255));
    HPEN hClosePen = CreatePen(PS_SOLID, 1, RGB(255, 255, 255));
    HPEN hOldClosePen = (HPEN)SelectObject(hdc, hClosePen);

    int crossPadding = MulDiv(8, m_dpi, 96);
    int x1 = rcCloseRect.left + crossPadding;
    int y1 = rcCloseRect.top + crossPadding;
    int x2 = rcCloseRect.right - crossPadding;
    int y2 = rcCloseRect.bottom - crossPadding;
    MoveToEx(hdc, x1, y1, NULL); LineTo(hdc, x2, y2);
    MoveToEx(hdc, x1, y2, NULL); LineTo(hdc, x2, y1);

    SetTextColor(hdc, oldTextColor);
    SelectObject(hdc, hOldClosePen);
    DeleteObject(hClosePen);
}

void CustomTabControl::ShowCustomTooltip(int index, int x, int y) {
    if (!m_hPopupWnd || index < 0 || index >= (int)m_tabTitles.size()) {
        return;
    }

    m_popupText = m_tabTitles[index];

    // テキストサイズを計算
    HDC hdc = GetDC(m_hPopupWnd);
    SelectObject(hdc, m_hFont);
    SIZE size;
    GetTextExtentPoint32W(hdc, m_popupText.c_str(), (int)m_popupText.length(), &size);
    ReleaseDC(m_hPopupWnd, hdc);

    m_popupWidth = size.cx + 20; // 左右に10ピクセルのパディング
    m_popupHeight = size.cy + 10; // 上下に5ピクセルのパディング

    // ポップアップ位置を計算（クライアント座標からスクリーン座標へ）
    POINT pt = { x, y };
    ClientToScreen(m_hWnd, &pt);

    // ポップアップウィンドウの表示と位置設定
    SetWindowPos(m_hPopupWnd, HWND_TOPMOST,
        pt.x + 10, pt.y + 20, m_popupWidth, m_popupHeight,
        SWP_NOACTIVATE | SWP_SHOWWINDOW);

    m_isPopupVisible = true;

    // 再描画を強制
    InvalidateRect(m_hPopupWnd, NULL, TRUE);
}

void CustomTabControl::HideCustomTooltip() {
    if (m_hPopupWnd && m_isPopupVisible) {
        ShowWindow(m_hPopupWnd, SW_HIDE);
        m_isPopupVisible = false;
    }
}

// テーマの変更に応じて色を更新する関数
void CustomTabControl::UpdateTheme(BOOL bIsDarkMode) {
    if (bIsDarkMode) {
        m_clrBg = RGB(32, 32, 32);
        m_clrText = RGB(220, 220, 220);
        m_clrActiveTab = RGB(50, 50, 50);
        m_clrSeparator = RGB(60, 60, 60);
        m_clrCloseText = RGB(150, 150, 150);
        m_clrHoverBg = RGB(45, 45, 45);
        m_clrCloseButtonHoverBg = RGB(96, 96, 96);
        m_clrScrollButtonHoverBg = RGB(60, 60, 60);
        m_clrTooltipBg = RGB(50, 50, 50);
        m_clrTooltipText = RGB(255, 255, 255);
    }
    else {
        m_clrBg = RGB(220, 220, 220);
        m_clrText = RGB(32, 32, 32);
        m_clrActiveTab = RGB(240, 240, 240);
        m_clrSeparator = RGB(200, 200, 200);
        m_clrCloseText = RGB(100, 100, 100);
        m_clrHoverBg = RGB(230, 230, 230);
        m_clrCloseButtonHoverBg = RGB(200, 200, 200);
        m_clrScrollButtonHoverBg = RGB(220, 220, 220);
        m_clrTooltipBg = RGB(250, 250, 250);
        m_clrTooltipText = RGB(32, 32, 32);
    }
    InvalidateRect(m_hWnd, NULL, TRUE);
    if (m_hPopupWnd) {
        InvalidateRect(m_hPopupWnd, NULL, TRUE);
    }
}