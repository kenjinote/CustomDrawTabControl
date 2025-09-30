#pragma once

#include <windows.h>
#include <windowsx.h>
#include <vector>
#include <string>

class CustomTabControl {
public:
	CustomTabControl();
	~CustomTabControl();

	static void RegisterWindowClass(HINSTANCE hInstance);
	HWND Create(HWND hParent, int x, int y, int width, int height, UINT_PTR uId);
	void AddTab(const std::wstring& title);
	void RemoveTab(int index);
	int GetCurSel() const;
	void SetCurSel(int index);
	int GetTabCount() const;
	HWND GetHwnd() const;
	void SwitchTabOrder(int index1, int index2);

protected:
	static LRESULT CALLBACK WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LRESULT CALLBACK DragWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
	void OnPaint(HWND hWnd);
	void OnSize(HWND hWnd);
	void OnLButtonDown(HWND hWnd, int x, int y);
	void OnLButtonUp(HWND hWnd, int x, int y);
	void OnMouseMove(HWND hWnd, int x, int y);
	void OnMouseLeave(HWND hWnd);
	void OnDpiChanged(HWND hWnd, int dpi);
	int HitTest(int x, int y, bool* isCloseButton, bool* isScrollLeft, bool* isScrollRight) const;
	void DrawTab(HDC hdc, int index, const RECT& rect, bool isActive, bool isHovered, bool isCloseHovered);
	void CreateDragWindow(int tabIndex);
	void DestroyDragWindow();
	void DrawDragWindow(HDC hdc);

private:
	HWND m_hWnd;
	HWND m_hDragWnd;
	HFONT m_hFont;
	int m_dpi;

	std::vector<std::wstring> m_tabTitles;
	int m_selectedTab;
	int m_hoveredTab;
	int m_hoveredCloseButtonTab;
	int m_pressedCloseButtonTab; // êVÇµÇ¢ÉÅÉìÉoÅ[ïœêî

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
};