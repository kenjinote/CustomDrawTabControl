#pragma once

#include <windows.h>
#include <windowsx.h>
#include <string>
#include <vector>
#include <memory>

class CustomTabControl {
public:
    CustomTabControl();
    ~CustomTabControl();

    HWND Create(HWND hParent, int x, int y, int width, int height, UINT_PTR uId);
    void AddTab(const std::wstring& title);
    int GetCurSel() const;
    void SetCurSel(int index);
    int GetTabCount() const;
    HWND GetHwnd() const;

private:
    static LRESULT CALLBACK WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static void RegisterWindowClass(HINSTANCE hInstance);

    void OnPaint(HWND hWnd);
    void OnSize(HWND hWnd);
    void OnLButtonDown(HWND hWnd, int x, int y);
    void OnMouseMove(HWND hWnd, int x, int y);
    void OnLButtonUp(HWND hWnd, int x, int y);
    void OnMouseLeave(HWND hWnd);
    void OnDpiChanged(HWND hWnd, int dpi);
    void DrawTab(HDC hdc, int index, const RECT& rect, bool isActive, bool isHovered, bool isCloseHovered);
    void SwitchTabOrder(int index1, int index2);

    int HitTest(int x, int y, bool* isCloseButton = nullptr, bool* isScrollLeft = nullptr, bool* isScrollRight = nullptr) const;

private:
    HWND m_hWnd;
    HFONT m_hFont;
    int m_dpi;
    std::vector<std::wstring> m_tabTitles;
    int m_selectedTab;
    int m_hoveredTab;
    int m_hoveredCloseButtonTab;
    int m_draggedTabIndex;
    POINT m_dragStartPos;
    bool m_isDragging;
    RECT m_draggedTabRect;

    // 新しいメンバー変数
    int m_scrollOffset;
    bool m_isScrollLeftHovered;
    bool m_isScrollRightHovered;
    int m_totalTabsWidth;
    int m_scrollButtonWidth;
    int m_scrollButtonHeight;
    RECT m_scrollLeftRect;
    RECT m_scrollRightRect;


    // テーマ色
    COLORREF m_clrBg;
    COLORREF m_clrText;
    COLORREF m_clrActiveTab;
    COLORREF m_clrSeparator;
    COLORREF m_clrCloseHoverBg;
    COLORREF m_clrCloseText;
};