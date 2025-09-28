#include "CustomTabControl.h"
#include <dwmapi.h>
#include <uxtheme.h>
#include <commctrl.h>
#include <algorithm>
#include <math.h>

#define TAB_PADDING_X 16
#define TAB_PADDING_Y 8
#define TAB_ROUND_RADIUS 8
#define FONT_SIZE 16
#define TAB_ANIMATION_TIMER_ID 101

static const WCHAR s_szClassName[] = L"CustomTabControlClass";
static const WCHAR s_szDragClassName[] = L"CustomTabDragClass";
static bool s_classRegistered = false;
static bool s_dragClassRegistered = false;

CustomTabControl::CustomTabControl()
    : m_hWnd(NULL), m_hFont(NULL), m_dpi(96), m_selectedTab(0), m_hoveredTab(-1),
    m_hoveredCloseButtonTab(-1), m_draggedTabIndex(-1), m_isDragging(false),
    m_scrollOffset(0), m_isScrollLeftHovered(false), m_isScrollRightHovered(false),
    m_totalTabsWidth(0), m_scrollButtonWidth(0), m_scrollButtonHeight(0),
    m_animationTimerRunning(false), m_hDragWnd(NULL) {

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

HWND CustomTabControl::Create(HWND hParent, int x, int y, int width, int height, UINT_PTR uId) {
    RegisterWindowClass(GetModuleHandle(NULL));

    m_hWnd = CreateWindowExW(
        0, s_szClassName, L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
        x, y, width, height,
        hParent, (HMENU)uId, GetModuleHandle(NULL), this
    );

    if (m_hWnd) {
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
        tme.dwFlags = TME_LEAVE;
        tme.hwndTrack = m_hWnd;
        _TrackMouseEvent(&tme);
    }
    return m_hWnd;
}

void CustomTabControl::AddTab(const std::wstring& title) {
    m_tabTitles.push_back(title);
    InvalidateRect(m_hWnd, NULL, TRUE);
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
        InvalidateRect(m_hWnd, NULL, TRUE);
    }
}

int CustomTabControl::GetCurSel() const {
    return m_selectedTab;
}

void CustomTabControl::SetCurSel(int index) {
    if (index >= 0 && index < (int)m_tabTitles.size()) {
        m_selectedTab = index;
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
    std::swap(m_tabTitles[index1], m_tabTitles[index2]);

    if (m_selectedTab == index1) {
        m_selectedTab = index2;
    }
    else if (m_selectedTab == index2) {
        m_selectedTab = index1;
    }

    InvalidateRect(m_hWnd, NULL, TRUE);
}

int CustomTabControl::HitTest(int x, int y, bool* isCloseButton, bool* isScrollLeft, bool* isScrollRight) const {
    if (isCloseButton) *isCloseButton = false;
    if (isScrollLeft) *isScrollLeft = false;
    if (isScrollRight) *isScrollRight = false;

    RECT rcClient;
    GetClientRect(m_hWnd, &rcClient);

    if (m_totalTabsWidth > rcClient.right && x >= rcClient.right - m_scrollButtonWidth * 2) {
        if (x >= m_scrollLeftRect.left && x <= m_scrollLeftRect.right && y >= m_scrollLeftRect.top && y <= m_scrollLeftRect.bottom) {
            if (isScrollLeft) *isScrollLeft = true;
            return -1;
        }
        if (x >= m_scrollRightRect.left && x <= m_scrollRightRect.right && y >= m_scrollRightRect.top && y <= m_scrollRightRect.bottom) {
            if (isScrollRight) *isScrollRight = true;
            return -1;
        }
    }

    int totalWidth = -m_scrollOffset;
    int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
    int closeBtnW = tabHeight;
    int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);

    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        HDC hdc = GetDC(m_hWnd);
        SelectObject(hdc, m_hFont);
        SIZE size;
        GetTextExtentPoint32W(hdc, m_tabTitles[i].c_str(), (int)m_tabTitles[i].length(), &size);
        ReleaseDC(m_hWnd, hdc);

        int tabWidth = size.cx + tabPaddingX + closeBtnW;
        RECT tabRect = { totalWidth, 0, totalWidth + tabWidth, tabHeight };

        if (x >= tabRect.left && x < tabRect.right && y >= tabRect.top && y < tabRect.bottom) {
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
        case WM_LBUTTONUP:
            pThis->OnLButtonUp(hWnd, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            return 0;
        case WM_MOUSELEAVE:
            pThis->OnMouseLeave(hWnd);
            return 0;
        case WM_DPICHANGED:
            pThis->OnDpiChanged(hWnd, LOWORD(wParam));
            return 0;
        case WM_TIMER:
            if (wParam == TAB_ANIMATION_TIMER_ID) {
                pThis->OnTimer();
                return 0;
            }
            break;
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
    HBRUSH hBrush = CreateSolidBrush(m_clrBg);
    FillRect(hdc, &clientRect, hBrush);
    DeleteObject(hBrush);

    m_totalTabsWidth = 0;
    int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
    int closeBtnW = tabHeight;
    int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);
    std::vector<int> tabWidths(m_tabTitles.size());
    std::vector<int> tabXPositions(m_tabTitles.size());

    HDC hdcTemp = GetDC(hWnd);
    SelectObject(hdcTemp, m_hFont);
    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        SIZE size;
        GetTextExtentPoint32W(hdcTemp, m_tabTitles[i].c_str(), (int)m_tabTitles[i].length(), &size);
        tabWidths[i] = size.cx + tabPaddingX + closeBtnW;
        m_totalTabsWidth += tabWidths[i];
    }
    ReleaseDC(hWnd, hdcTemp);

    m_scrollButtonWidth = MulDiv(30, m_dpi, 96);
    m_scrollButtonHeight = tabHeight;
    bool showScrollButtons = m_totalTabsWidth > clientRect.right;
    int maxScrollOffset = max(0, m_totalTabsWidth - (clientRect.right - (showScrollButtons ? m_scrollButtonWidth * 2 : 0)));
    m_scrollOffset = min(maxScrollOffset, max(0, m_scrollOffset));

    RECT tabsDrawingRect = clientRect;
    if (showScrollButtons) {
        tabsDrawingRect.right -= m_scrollButtonWidth * 2;
    }
    IntersectClipRect(hdc, tabsDrawingRect.left, tabsDrawingRect.top, tabsDrawingRect.right, tabsDrawingRect.bottom);

    int currentX = -m_scrollOffset;
    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        tabXPositions[i] = currentX;
        currentX += tabWidths[i];
    }

    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        if (m_isDragging && i == m_draggedTabIndex) {
            continue;
        }

        int xPos = tabXPositions[i];
        if (m_isDragging && i < m_draggedTabsCurrentPositions.size()) {
            xPos = m_draggedTabsCurrentPositions[i] + (int)((m_draggedTabsTargetPositions[i] - m_draggedTabsCurrentPositions[i]) * 0.2);
            m_draggedTabsCurrentPositions[i] = xPos;
        }

        RECT rect = { xPos, 0, xPos + tabWidths[i], tabHeight };
        bool isActive = (i == m_selectedTab);
        bool isHovered = (i == m_hoveredTab);
        bool isCloseHovered = (i == m_hoveredCloseButtonTab);

        if (m_isDragging && m_hoveredTab != -1 && i == m_hoveredTab) {
            HPEN hPenIndicator = CreatePen(PS_SOLID, 2, RGB(0, 120, 215));
            HPEN hOldPen = (HPEN)SelectObject(hdc, hPenIndicator);
            int dropX = (m_draggedTabIndex < m_hoveredTab) ? rect.right : rect.left;
            MoveToEx(hdc, dropX, rect.top + MulDiv(5, m_dpi, 96), NULL);
            LineTo(hdc, dropX, rect.bottom - MulDiv(5, m_dpi, 96));
            SelectObject(hdc, hOldPen);
            DeleteObject(hPenIndicator);
        }

        DrawTab(hdc, i, rect, isActive, isHovered, isCloseHovered);
    }

    SelectClipRgn(hdc, NULL);

    if (showScrollButtons) {
        m_scrollLeftRect = { clientRect.right - m_scrollButtonWidth * 2, 0, clientRect.right - m_scrollButtonWidth, m_scrollButtonHeight };
        m_scrollRightRect = { clientRect.right - m_scrollButtonWidth, 0, clientRect.right, m_scrollButtonHeight };
        HBRUSH hScrollBrush = CreateSolidBrush(m_isScrollLeftHovered ? RGB(60, 60, 60) : m_clrBg);
        FillRect(hdc, &m_scrollLeftRect, hScrollBrush);
        DeleteObject(hScrollBrush);
        POINT triangleLeft[] = { {m_scrollLeftRect.left + MulDiv(10, m_dpi, 96), m_scrollLeftRect.top + MulDiv(15, m_dpi, 96)},{m_scrollLeftRect.left + MulDiv(15, m_dpi, 96), m_scrollLeftRect.top + MulDiv(10, m_dpi, 96)},{m_scrollLeftRect.left + MulDiv(15, m_dpi, 96), m_scrollLeftRect.top + MulDiv(20, m_dpi, 96)} };
        HBRUSH hTriangleBrush = CreateSolidBrush(m_clrText);
        SelectObject(hdc, hTriangleBrush);
        Polygon(hdc, triangleLeft, 3);
        DeleteObject(hTriangleBrush);
        hScrollBrush = CreateSolidBrush(m_isScrollRightHovered ? RGB(60, 60, 60) : m_clrBg);
        FillRect(hdc, &m_scrollRightRect, hScrollBrush);
        DeleteObject(hScrollBrush);
        POINT triangleRight[] = { {m_scrollRightRect.left + MulDiv(15, m_dpi, 96), m_scrollRightRect.top + MulDiv(10, m_dpi, 96)},{m_scrollRightRect.left + MulDiv(20, m_dpi, 96), m_scrollRightRect.top + MulDiv(15, m_dpi, 96)},{m_scrollRightRect.left + MulDiv(15, m_dpi, 96), m_scrollRightRect.top + MulDiv(20, m_dpi, 96)} };
        hTriangleBrush = CreateSolidBrush(m_clrText);
        SelectObject(hdc, hTriangleBrush);
        Polygon(hdc, triangleRight, 3);
        DeleteObject(hTriangleBrush);
    }
    EndPaint(hWnd, &ps);
}

void CustomTabControl::DrawTab(HDC hdc, int index, const RECT& rect, bool isActive, bool isHovered, bool isCloseHovered) {
    RECT rc = rect;
    COLORREF bgColor = isActive ? m_clrActiveTab : m_clrBg;
    if (isHovered && !isActive) {
        bgColor = RGB(45, 45, 45);
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

    HBRUSH hCloseBrush = isCloseHovered ? CreateSolidBrush(RGB(96, 96, 96)) : (HBRUSH)GetStockObject(NULL_BRUSH);
    if (isCloseHovered) {
        FillRect(hdc, &rcCloseRect, hCloseBrush);
    }
    DeleteObject(hCloseBrush);

    COLORREF oldTextColor = SetTextColor(hdc, isCloseHovered ? RGB(255, 255, 255) : m_clrCloseText);
    HPEN hClosePen = CreatePen(PS_SOLID, 1, isCloseHovered ? RGB(255, 255, 255) : m_clrCloseText);
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
    InvalidateRect(hWnd, NULL, TRUE);
}

void CustomTabControl::OnLButtonDown(HWND hWnd, int x, int y) {
    bool isClose = false;
    bool isScrollLeft = false;
    bool isScrollRight = false;
    int index = HitTest(x, y, &isClose, &isScrollLeft, &isScrollRight);

    if (isScrollLeft) {
        m_scrollOffset = max(0, m_scrollOffset - 50);
        InvalidateRect(hWnd, NULL, TRUE);
        return;
    }

    if (isScrollRight) {
        RECT rcClient;
        GetClientRect(hWnd, &rcClient);
        m_scrollOffset = min(m_totalTabsWidth - rcClient.right, m_scrollOffset + 50);
        InvalidateRect(hWnd, NULL, TRUE);
        return;
    }

    if (index != -1) {
        if (isClose) {
            RemoveTab(index);
        }
        else {
            m_selectedTab = index;
            m_draggedTabIndex = index;
            m_dragStartPos.x = x;
            m_dragStartPos.y = y;
            SetCapture(hWnd);

            CreateDragWindow(index);

            InvalidateRect(hWnd, NULL, TRUE);
        }
    }
}

void CustomTabControl::OnMouseMove(HWND hWnd, int x, int y) {
    bool isClose = false;
    bool isScrollLeft = false;
    bool isScrollRight = false;
    int newHoveredTab = HitTest(x, y, &isClose, &isScrollLeft, &isScrollRight);

    if (m_isScrollLeftHovered != isScrollLeft || m_isScrollRightHovered != isScrollRight) {
        m_isScrollLeftHovered = isScrollLeft;
        m_isScrollRightHovered = isScrollRight;
        InvalidateRect(hWnd, NULL, FALSE);
    }

    int newHoveredCloseButtonTab = isClose ? newHoveredTab : -1;

    if (newHoveredTab != m_hoveredTab || newHoveredCloseButtonTab != m_hoveredCloseButtonTab) {
        m_hoveredTab = newHoveredTab;
        m_hoveredCloseButtonTab = newHoveredCloseButtonTab;
        InvalidateRect(hWnd, NULL, FALSE);
    }

    if (m_draggedTabIndex != -1 && GetCapture() == hWnd) {
        if (!m_isDragging) {
            if (abs(x - m_dragStartPos.x) > GetSystemMetrics(SM_CXDRAG) ||
                abs(y - m_dragStartPos.y) > GetSystemMetrics(SM_CYDRAG)) {
                m_isDragging = true;

                int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
                int closeBtnW = tabHeight;
                int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);

                // アニメーションの初期位置と目標位置を計算
                m_draggedTabsCurrentPositions.clear();
                m_draggedTabsTargetPositions.clear();
                int currentX = -m_scrollOffset;
                for (size_t i = 0; i < m_tabTitles.size(); ++i) {
                    HDC hdc = GetDC(m_hWnd);
                    SelectObject(hdc, m_hFont);
                    SIZE s;
                    GetTextExtentPoint32W(hdc, m_tabTitles[i].c_str(), (int)m_tabTitles[i].length(), &s);
                    ReleaseDC(m_hWnd, hdc);
                    int w = s.cx + tabPaddingX + closeBtnW;

                    if (i == m_draggedTabIndex) {
                        m_draggedTabsCurrentPositions.push_back(currentX); // ダミー位置
                        m_draggedTabsTargetPositions.push_back(currentX); // ダミー位置
                        currentX += w;
                        continue;
                    }

                    m_draggedTabsCurrentPositions.push_back(currentX);
                    m_draggedTabsTargetPositions.push_back(currentX);
                    currentX += w;
                }

                SetTimer(hWnd, TAB_ANIMATION_TIMER_ID, 10, NULL);
                m_animationTimerRunning = true;
            }
        }

        if (m_isDragging) {
            POINT pt;
            GetCursorPos(&pt);

            int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
            int closeBtnW = tabHeight;
            int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);
            HDC hdc = GetDC(m_hWnd);
            SelectObject(hdc, m_hFont);
            SIZE size;
            GetTextExtentPoint32W(hdc, m_tabTitles[m_draggedTabIndex].c_str(), (int)m_tabTitles[m_draggedTabIndex].length(), &size);
            ReleaseDC(m_hWnd, hdc);
            int tabWidth = size.cx + tabPaddingX + closeBtnW;

            SetWindowPos(m_hDragWnd, NULL, pt.x - tabWidth / 2, pt.y - tabHeight / 2, tabWidth, tabHeight, SWP_NOZORDER | SWP_NOACTIVATE);

            int hoveredIndex = HitTest(x, y, nullptr, nullptr, nullptr);
            if (hoveredIndex != -1 && hoveredIndex != m_draggedTabIndex) {
                if (m_hoveredTab != hoveredIndex) {
                    m_hoveredTab = hoveredIndex;
                    std::vector<int> newPositions;
                    int currentX = -m_scrollOffset;

                    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
                        HDC hdc = GetDC(m_hWnd);
                        SelectObject(hdc, m_hFont);
                        SIZE s;
                        GetTextExtentPoint32W(hdc, m_tabTitles[i].c_str(), (int)m_tabTitles[i].length(), &s);
                        ReleaseDC(m_hWnd, hdc);
                        int w = s.cx + tabPaddingX + closeBtnW;

                        if (i == m_draggedTabIndex) {
                            // ドラッグ中のタブは位置計算から除外
                            continue;
                        }

                        // 目標位置を再計算
                        if (m_draggedTabIndex < hoveredIndex) {
                            if (i > m_draggedTabIndex && i <= hoveredIndex) {
                                newPositions.push_back(currentX - w);
                            }
                            else {
                                newPositions.push_back(currentX);
                            }
                        }
                        else {
                            if (i < m_draggedTabIndex && i >= hoveredIndex) {
                                newPositions.push_back(currentX + tabWidth);
                            }
                            else {
                                newPositions.push_back(currentX);
                            }
                        }

                        currentX += w;
                    }
                    m_draggedTabsTargetPositions = newPositions;
                }
            }
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

void CustomTabControl::OnLButtonUp(HWND hWnd, int x, int y) {
    if (m_isDragging) {
        ReleaseCapture();
        DestroyDragWindow();

        if (m_animationTimerRunning) {
            KillTimer(hWnd, TAB_ANIMATION_TIMER_ID);
            m_animationTimerRunning = false;
        }

        bool isClose = false;
        bool isScrollLeft = false;
        bool isScrollRight = false;
        int dropIndex = HitTest(x, y, &isClose, &isScrollLeft, &isScrollRight);
        if (dropIndex != -1 && dropIndex != m_draggedTabIndex) {
            SwitchTabOrder(m_draggedTabIndex, dropIndex);
        }
    }
    m_draggedTabIndex = -1;
    m_isDragging = false;
    InvalidateRect(hWnd, NULL, TRUE);
}

void CustomTabControl::OnMouseLeave(HWND hWnd) {
    if (m_hoveredTab != -1 || m_isScrollLeftHovered || m_isScrollRightHovered) {
        m_hoveredTab = -1;
        m_hoveredCloseButtonTab = -1;
        m_isScrollLeftHovered = false;
        m_isScrollRightHovered = false;
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

void CustomTabControl::OnTimer() {
    bool animationComplete = true;
    for (size_t i = 0; i < m_draggedTabsCurrentPositions.size(); ++i) {
        if (i < m_draggedTabsTargetPositions.size() && m_draggedTabsCurrentPositions[i] != m_draggedTabsTargetPositions[i]) {
            int delta = (m_draggedTabsTargetPositions[i] - m_draggedTabsCurrentPositions[i]) / 5;
            if (abs(delta) < 2) {
                delta = m_draggedTabsTargetPositions[i] - m_draggedTabsCurrentPositions[i];
            }
            m_draggedTabsCurrentPositions[i] += delta;
            animationComplete = false;
        }
    }
    if (animationComplete) {
        KillTimer(m_hWnd, TAB_ANIMATION_TIMER_ID);
        m_animationTimerRunning = false;
    }
    InvalidateRect(m_hWnd, NULL, FALSE);
}

void CustomTabControl::CreateDragWindow(int tabIndex) {
    if (m_hDragWnd) {
        return;
    }

    POINT ptCursor;
    GetCursorPos(&ptCursor);

    int tabHeight = MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96);
    int closeBtnW = tabHeight;
    int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);
    HDC hdc = GetDC(m_hWnd);
    SelectObject(hdc, m_hFont);
    SIZE size;
    GetTextExtentPoint32W(hdc, m_tabTitles[tabIndex].c_str(), (int)m_tabTitles[tabIndex].length(), &size);
    ReleaseDC(m_hWnd, hdc);
    int tabWidth = size.cx + tabPaddingX + closeBtnW;

    m_hDragWnd = CreateWindowExW(
        WS_EX_LAYERED | WS_EX_TOOLWINDOW, s_szDragClassName, L"",
        WS_POPUP | WS_VISIBLE,
        ptCursor.x - tabWidth / 2, ptCursor.y - tabHeight / 2, tabWidth, tabHeight,
        NULL, NULL, GetModuleHandle(NULL), this
    );

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
        SIZE  sizeTab = { tabWidth, tabHeight };
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
    RECT rcCloseRect = { closeBtnX, 0, tabWidth, tabHeight };
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