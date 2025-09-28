#include "CustomTabControl.h"
#include <dwmapi.h>
#include <uxtheme.h>
#include <commctrl.h>
#include <algorithm>
#include <math.h>

#define TAB_CLOSE_BUTTON_WIDTH 24
#define TAB_CLOSE_BUTTON_HEIGHT 16
#define TAB_PADDING_X 16
#define TAB_PADDING_Y 8
#define TAB_ROUND_RADIUS 8
#define FONT_SIZE 16

static const WCHAR s_szClassName[] = L"CustomTabControlClass";
static bool s_classRegistered = false;

CustomTabControl::CustomTabControl()
    : m_hWnd(NULL), m_hFont(NULL), m_dpi(96), m_selectedTab(0), m_hoveredTab(-1),
    m_hoveredCloseButtonTab(-1), m_draggedTabIndex(-1), m_isDragging(false),
    m_scrollOffset(0), m_isScrollLeftHovered(false), m_isScrollRightHovered(false),
    m_totalTabsWidth(0), m_scrollButtonWidth(0), m_scrollButtonHeight(0) {

    // デフォルトのタブを追加
    m_tabTitles.push_back(L"Tab 1");
    m_tabTitles.push_back(L"Tab 2");
    m_tabTitles.push_back(L"Tab 3");
    m_tabTitles.push_back(L"Long Tab Title 4");
    m_tabTitles.push_back(L"Another Tab");
    m_tabTitles.push_back(L"Final Tab 6");
    m_tabTitles.push_back(L"Tab 7");

    // ダークモードをデフォルトとする
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
}

void CustomTabControl::RegisterWindowClass(HINSTANCE hInstance) {
    if (s_classRegistered) return;

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
        // フォントとレイアウトを初期化
        int lfHeight = -MulDiv(FONT_SIZE, m_dpi, 72);
        m_hFont = CreateFontW(
            lfHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI"
        );
        SendMessage(m_hWnd, WM_SETFONT, (WPARAM)m_hFont, FALSE);

        // トラッキングを開始
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

    // スクロールボタンのヒットテスト
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
    int closeBtnW = MulDiv(TAB_CLOSE_BUTTON_WIDTH, m_dpi, 96);
    int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);

    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        HDC hdc = GetDC(m_hWnd);
        SelectObject(hdc, m_hFont);
        SIZE size;
        GetTextExtentPoint32W(hdc, m_tabTitles[i].c_str(), (int)m_tabTitles[i].length(), &size);
        ReleaseDC(m_hWnd, hdc);

        int tabWidth = size.cx + tabPaddingX + closeBtnW;
        RECT tabRect = { totalWidth, 0, totalWidth + tabWidth, MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96) };

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
        case WM_DESTROY:
            // オブジェクトのクリーンアップは所有者が行う
            pThis->m_hWnd = NULL;
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
    int closeBtnW = MulDiv(TAB_CLOSE_BUTTON_WIDTH, m_dpi, 96);
    int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);

    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        HDC hdcTemp = GetDC(hWnd);
        SelectObject(hdcTemp, m_hFont);
        SIZE size;
        GetTextExtentPoint32W(hdcTemp, m_tabTitles[i].c_str(), (int)m_tabTitles[i].length(), &size);
        ReleaseDC(hWnd, hdcTemp);
        m_totalTabsWidth += size.cx + tabPaddingX + closeBtnW;
    }

    m_scrollButtonWidth = MulDiv(30, m_dpi, 96);
    m_scrollButtonHeight = tabHeight;
    bool showScrollButtons = m_totalTabsWidth > clientRect.right;

    // スクロール可能な最大オフセットを計算
    int maxScrollOffset = max(0, m_totalTabsWidth - (clientRect.right - (showScrollButtons ? m_scrollButtonWidth * 2 : 0)));
    m_scrollOffset = min(maxScrollOffset, max(0, m_scrollOffset));

    // タブの描画領域を設定
    RECT tabsDrawingRect = clientRect;
    if (showScrollButtons) {
        tabsDrawingRect.right -= m_scrollButtonWidth * 2;
    }
    IntersectClipRect(hdc, tabsDrawingRect.left, tabsDrawingRect.top, tabsDrawingRect.right, tabsDrawingRect.bottom);

    // タブの描画
    int currentX = -m_scrollOffset;
    for (size_t i = 0; i < m_tabTitles.size(); ++i) {
        HDC hdcTemp = GetDC(hWnd);
        SelectObject(hdcTemp, m_hFont);
        SIZE size;
        GetTextExtentPoint32W(hdcTemp, m_tabTitles[i].c_str(), (int)m_tabTitles[i].length(), &size);
        ReleaseDC(hWnd, hdcTemp);

        int tabWidth = size.cx + tabPaddingX + closeBtnW;
        RECT rect = { currentX, 0, currentX + tabWidth, tabHeight };

        bool isActive = (i == m_selectedTab);
        bool isHovered = (i == m_hoveredTab);
        bool isCloseHovered = (i == m_hoveredCloseButtonTab);

        if (m_isDragging && i == m_draggedTabIndex) {
            currentX += tabWidth;
            continue;
        }

        if (m_isDragging && m_hoveredTab != -1 && m_draggedTabIndex != -1 && m_draggedTabIndex != m_hoveredTab && i == m_hoveredTab) {
            HPEN hPenIndicator = CreatePen(PS_SOLID, 2, RGB(0, 120, 215));
            HPEN hOldPen = (HPEN)SelectObject(hdc, hPenIndicator);
            int dropX = (m_draggedTabIndex < m_hoveredTab) ? rect.right : rect.left;
            MoveToEx(hdc, dropX, rect.top + MulDiv(5, m_dpi, 96), NULL);
            LineTo(hdc, dropX, rect.bottom - MulDiv(5, m_dpi, 96));
            SelectObject(hdc, hOldPen);
            DeleteObject(hPenIndicator);
        }

        DrawTab(hdc, i, rect, isActive, isHovered, isCloseHovered);
        currentX += tabWidth;
    }

    if (m_isDragging && m_draggedTabIndex != -1) {
        DrawTab(hdc, m_draggedTabIndex, m_draggedTabRect, true, true, false);
    }

    // クリップ領域をリセット
    SelectClipRgn(hdc, NULL);

    // スクロールボタンの描画
    if (showScrollButtons) {
        m_scrollLeftRect = { clientRect.right - m_scrollButtonWidth * 2, 0, clientRect.right - m_scrollButtonWidth, m_scrollButtonHeight };
        m_scrollRightRect = { clientRect.right - m_scrollButtonWidth, 0, clientRect.right, m_scrollButtonHeight };

        HBRUSH hScrollBrush = CreateSolidBrush(m_isScrollLeftHovered ? RGB(60, 60, 60) : m_clrBg);
        FillRect(hdc, &m_scrollLeftRect, hScrollBrush);
        DeleteObject(hScrollBrush);
        POINT triangleLeft[] = {
            {m_scrollLeftRect.left + MulDiv(10, m_dpi, 96), m_scrollLeftRect.top + MulDiv(15, m_dpi, 96)},
            {m_scrollLeftRect.left + MulDiv(15, m_dpi, 96), m_scrollLeftRect.top + MulDiv(10, m_dpi, 96)},
            {m_scrollLeftRect.left + MulDiv(15, m_dpi, 96), m_scrollLeftRect.top + MulDiv(20, m_dpi, 96)}
        };
        HBRUSH hTriangleBrush = CreateSolidBrush(m_clrText);
        SelectObject(hdc, hTriangleBrush);
        Polygon(hdc, triangleLeft, 3);
        DeleteObject(hTriangleBrush);

        hScrollBrush = CreateSolidBrush(m_isScrollRightHovered ? RGB(60, 60, 60) : m_clrBg);
        FillRect(hdc, &m_scrollRightRect, hScrollBrush);
        DeleteObject(hScrollBrush);
        POINT triangleRight[] = {
            {m_scrollRightRect.left + MulDiv(15, m_dpi, 96), m_scrollRightRect.top + MulDiv(10, m_dpi, 96)},
            {m_scrollRightRect.left + MulDiv(20, m_dpi, 96), m_scrollRightRect.top + MulDiv(15, m_dpi, 96)},
            {m_scrollRightRect.left + MulDiv(15, m_dpi, 96), m_scrollRightRect.top + MulDiv(20, m_dpi, 96)}
        };
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
    int closeBtnW = MulDiv(TAB_CLOSE_BUTTON_WIDTH, m_dpi, 96);
    rcText.right -= closeBtnW;
    DrawTextW(hdc, m_tabTitles[index].c_str(), -1, &rcText, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);

    int closeBtnX = rect.right - MulDiv(TAB_PADDING_X, m_dpi, 96) / 2 - closeBtnW;
    int closeBtnY = rect.top + (rect.bottom - rect.top - MulDiv(TAB_CLOSE_BUTTON_HEIGHT, m_dpi, 96)) / 2;
    RECT rcClose = { closeBtnX, closeBtnY, closeBtnX + closeBtnW, closeBtnY + MulDiv(TAB_CLOSE_BUTTON_HEIGHT, m_dpi, 96) };

    HBRUSH hCloseBrush = isCloseHovered ? CreateSolidBrush(m_clrCloseHoverBg) : (HBRUSH)GetStockObject(NULL_BRUSH);
    HPEN hClosePen = CreatePen(PS_SOLID, 1, isCloseHovered ? RGB(255, 255, 255) : m_clrCloseText);
    HBRUSH hOldCloseBrush = (HBRUSH)SelectObject(hdc, hCloseBrush);
    HPEN hOldClosePen = (HPEN)SelectObject(hdc, hClosePen);

    Ellipse(hdc, rcClose.left, rcClose.top, rcClose.right, rcClose.bottom);

    COLORREF oldTextColor = SetTextColor(hdc, isCloseHovered ? RGB(255, 255, 255) : m_clrCloseText);
    int x1 = rcClose.left + roundf(closeBtnW * 0.3f);
    int y1 = rcClose.top + roundf(MulDiv(TAB_CLOSE_BUTTON_HEIGHT, m_dpi, 96) * 0.3f);
    int x2 = rcClose.right - roundf(closeBtnW * 0.3f);
    int y2 = rcClose.bottom - roundf(MulDiv(TAB_CLOSE_BUTTON_HEIGHT, m_dpi, 96) * 0.3f);
    MoveToEx(hdc, x1, y1, NULL); LineTo(hdc, x2, y2);
    MoveToEx(hdc, x1, y2, NULL); LineTo(hdc, x2, y1);
    SetTextColor(hdc, oldTextColor);

    SelectObject(hdc, hOldCloseBrush);
    SelectObject(hdc, hOldClosePen);
    DeleteObject(hCloseBrush);
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
            if (m_tabTitles.size() > 1) {
                m_tabTitles.erase(m_tabTitles.begin() + index);
                if (m_selectedTab == index) {
                    m_selectedTab = min((int)m_tabTitles.size() - 1, m_selectedTab);
                }
                else if (m_selectedTab > index) {
                    m_selectedTab--;
                }
                InvalidateRect(hWnd, NULL, TRUE);
            }
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
                int totalWidth = -m_scrollOffset;
                int closeBtnW = MulDiv(TAB_CLOSE_BUTTON_WIDTH, m_dpi, 96);
                int tabPaddingX = MulDiv(TAB_PADDING_X, m_dpi, 96);
                for (int i = 0; i < m_draggedTabIndex; ++i) {
                    HDC hdc = GetDC(m_hWnd);
                    SelectObject(hdc, m_hFont);
                    SIZE size;
                    GetTextExtentPoint32W(hdc, m_tabTitles[i].c_str(), (int)m_tabTitles[i].length(), &size);
                    ReleaseDC(m_hWnd, hdc);
                    totalWidth += size.cx + tabPaddingX + closeBtnW;
                }
                m_draggedTabRect = { totalWidth, 0, totalWidth + (int)m_tabTitles[m_draggedTabIndex].length() + tabPaddingX + closeBtnW, (int)MulDiv(FONT_SIZE, m_dpi, 72) + MulDiv(TAB_PADDING_Y * 2, m_dpi, 96) };
            }
        }

        if (m_isDragging) {
            m_draggedTabRect.left += (x - m_dragStartPos.x);
            m_draggedTabRect.right += (x - m_dragStartPos.x);
            m_dragStartPos.x = x;
            m_dragStartPos.y = y;
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