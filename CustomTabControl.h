#pragma once

#include <Windows.h>
#include <Windowsx.h>
#include <vector>
#include <string>

class CustomTabControl {
public:
    CustomTabControl();
    ~CustomTabControl();

    static void RegisterWindowClass(HINSTANCE hInstance);
    HWND Create(HWND hParent, int x, int y, int width, int height, UINT_PTR uId, BOOL IsDarkMode);

    void AddTab(const std::wstring& title);
    void RemoveTab(int index);
    void RenameTab(int index, const std::wstring& newTitle);
    int GetCurSel() const;
    void SetCurSel(int index);
    int GetTabCount() const;
    HWND GetHwnd() const;
    void SwitchTabOrder(int index1, int index2);

private:
    static LRESULT CALLBACK WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK DragWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK PopupWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static void RegisterPopupWindowClass(HINSTANCE hInstance);

    void OnPaint(HWND hWnd);
    void OnSize(HWND hWnd);
    void OnLButtonDown(HWND hWnd, int x, int y);
    void OnMouseMove(HWND hWnd, int x, int y);
    void OnMouseHover(HWND hWnd, int x, int y);
    void OnLButtonUp(HWND hWnd, int x, int y);
    void OnMouseLeave(HWND hWnd);
    void OnDpiChanged(HWND hWnd, int dpi);

    void RecalculateTabPositions();
    int GetTabWidth(int index) const;
    int HitTest(int x, int y, bool* isCloseButton, bool* isScrollLeft, bool* isScrollRight) const;
    void DrawTab(HDC hdc, int index, const RECT& rect, bool isActive, bool isHovered, bool isCloseHovered);

    void CreateDragWindow(int tabIndex);
    void DestroyDragWindow();
    void DrawDragWindow(HDC hdc);

    void ShowCustomTooltip(int index, int x, int y);
    void HideCustomTooltip();

    void UpdateTheme(BOOL bIsDarkMode);

    HWND m_hWnd;
    HFONT m_hFont;
    int m_dpi;
    std::vector<std::wstring> m_tabTitles;
    int m_selectedTab;
    int m_hoveredTab;
    int m_hoveredCloseButtonTab;
    int m_pressedCloseButtonTab;
    int m_draggedTabIndex;
    bool m_isDragging;
    POINT m_dragStartPos;

    int m_scrollOffset;
    bool m_isScrollLeftHovered;
    bool m_isScrollRightHovered;
    int m_totalTabsWidth;
    int m_scrollButtonWidth;
    int m_scrollButtonHeight;
    RECT m_scrollLeftRect;
    RECT m_scrollRightRect;

    COLORREF m_clrBg;
    COLORREF m_clrText;
    COLORREF m_clrActiveTab;
    COLORREF m_clrSeparator;
    COLORREF m_clrCloseHoverBg;
    COLORREF m_clrCloseText;
    COLORREF m_clrHoverBg; // ★ 追加: タブのホバー色
    COLORREF m_clrCloseButtonHoverBg;
    COLORREF m_clrScrollButtonHoverBg;
    COLORREF m_clrTooltipBg;
    COLORREF m_clrTooltipText;

    // 独自ツールチップ用のメンバ変数
    HWND m_hDragWnd;
    HWND m_hPopupWnd; // 独自のポップアップウィンドウハンドル
    bool m_isPopupVisible;
    std::wstring m_popupText;
    int m_popupWidth;
    int m_popupHeight;
};