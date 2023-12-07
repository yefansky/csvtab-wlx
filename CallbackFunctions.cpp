#include "stdafx.h"

LRESULT CALLBACK CallHotKeyProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LRESULT nResult = 0;

	if (uMsg == WM_KEYDOWN && SendMessage(GetMainWindow(hWnd), WMU_HOT_KEYS, wParam, lParam))
	{
		nResult = 0;
		goto Exit0;
	}
	// Prevent beep
	if (uMsg == WM_CHAR && SendMessage(GetMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
	{
		nResult = 0;
		goto Exit0;
	}
	nResult = CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, uMsg, wParam, lParam);

Exit0:
	return nResult;
}

LRESULT CALLBACK CallNewHeaderProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (uMsg == WM_CTLCOLOREDIT) 
	{
		HWND hMainWnd = GetMainWindow(hWnd);
		SetBkColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
		SetTextColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR")));
		return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK CallNewFilterEditProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch (uMsg) 
	{
		case WM_PAINT:
		{
			cbDefault(hWnd, uMsg, wParam, lParam);

			RECT rc;
			GetClientRect(hWnd, &rc);
			HWND hMainWnd = GetMainWindow(hWnd);
			BOOL bisDark = *(int*)GetProp(hMainWnd, TEXT("DARKTHEME"));

			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
			HPEN oldPen = (HPEN)SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			LineTo(hDC, rc.right - 1, rc.bottom - 1);

			if (bisDark)
			{
				DeleteObject(hPen);
				hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
				SelectObject(hDC, hPen);

				MoveToEx(hDC, 0, 0, 0);
				LineTo(hDC, 0, rc.bottom);
				MoveToEx(hDC, 0, rc.bottom - 1, 0);
				LineTo(hDC, rc.right, rc.bottom - 1);
				MoveToEx(hDC, 0, rc.bottom - 2, 0);
				LineTo(hDC, rc.right, rc.bottom - 2);
			}

			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);

			return 0;
		}
		break;

	case WM_SETFOCUS:
		SetProp(GetMainWindow(hWnd), TEXT("LASTFOCUS"), hWnd);
		break;

	case WM_KEYDOWN:
		{
			HWND hMainWnd = GetMainWindow(hWnd);
			if (wParam == VK_RETURN)
			{
				SendMessage(hMainWnd, WMU_UPDATE_RESULTSET, 0, 0);
				return 0;
			}

			if (SendMessage(hMainWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;
		}
		break;

	// Prevent beep
	case WM_CHAR:
		{
			if (SendMessage(GetMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
				return 0;
		}
		break;

	case WM_DESTROY:
		RemoveProp(hWnd, TEXT("WNDPROC"));
		break;
	}

	return CallWindowProc(cbDefault, hWnd, uMsg, wParam, lParam);
}

int CALLBACK CallEnumTabStopChildrenProc(HWND hWnd, LPARAM lParam)
{
	if (GetWindowLong(hWnd, GWL_STYLE) & WS_TABSTOP && IsWindowVisible(hWnd))
	{
		int nNo = 0;
		HWND* pWnds = (HWND*)lParam;
		while (pWnds[nNo])
			nNo++;
		pWnds[nNo] = hWnd;
	}

	return TRUE;
}