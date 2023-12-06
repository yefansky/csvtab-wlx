
// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "exportDef.h"

BOOL APIENTRY DllMain (HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	if (ul_reason_for_call == DLL_PROCESS_ATTACH && iniPath[0] == 0) {
		TCHAR path[MAX_PATH + 1] = {0};
		GetModuleFileName((HMODULE)hModule, path, MAX_PATH);
		TCHAR* dot = _tcsrchr(path, TEXT('.'));
		_tcsncpy(dot, TEXT(".ini"), 5);
		if (_taccess(path, 0) == 0)
			_tcscpy(iniPath, path);	
	}
	return TRUE;
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen) {
	TCHAR* detectString16 = getStoredString(TEXT("detect-string"), TEXT("MULTIMEDIA & (ext=\"CSV\" | ext=\"TAB\" | ext=\"TSV\")"));
	char* detectString8 = utf16to8(detectString16);
	snprintf(DetectString, maxlen, detectString8);
	free(detectString16);
	free(detectString8);
}

void __stdcall ListSetDefaultParams(ListDefaultParamStruct* dps) {
	if (iniPath[0] == 0) {
		DWORD size = MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, iniPath, size);
	}
}

int __stdcall ListSearchTextW(HWND hWnd, TCHAR* searchString, int searchParameter) {
	HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
	HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);	
	
	TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
	int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
	int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
	int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
	if (!resultset || rowCount == 0)
		return 0;

	BOOL isFindFirst = searchParameter & LCS_FINDFIRST;		
	BOOL isBackward = searchParameter & LCS_BACKWARDS;
	BOOL isMatchCase = searchParameter & LCS_MATCHCASE;
	BOOL isWholeWords = searchParameter & LCS_WHOLEWORDS;	

	if (isFindFirst) {
		*(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) = 0;
		*(int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")) = 0;	
		*(int*)GetProp(hWnd, TEXT("CURRENTROWNO")) = isBackward ? rowCount - 1 : 0;
	}	
		
	int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
	int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
	int *pStartPos = (int*)GetProp(hWnd, TEXT("SEARCHCELLPOS"));	
	rowNo = rowNo == -1 || rowNo >= rowCount ? 0 : rowNo;
	colNo = colNo == -1 || colNo >= colCount ? 0 : colNo;	
			
	int pos = -1;
	do {
		for (; (pos == -1) && colNo < colCount; colNo++) {
			pos = findString(cache[resultset[rowNo]][colNo] + *pStartPos, searchString, isMatchCase, isWholeWords);
			if (pos != -1) 
				pos += *pStartPos;
			*pStartPos = pos == -1 ? 0 : pos + _tcslen(searchString);
		}
		colNo = pos != -1 ? colNo - 1 : 0;
		rowNo += pos != -1 ? 0 : isBackward ? -1 : 1; 	
	} while ((pos == -1) && (isBackward ? rowNo > 0 : rowNo < rowCount));
	ListView_SetItemState(hGridWnd, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);

	TCHAR buf[256] = {0};
	if (pos != -1) {
		ListView_EnsureVisible(hGridWnd, rowNo, FALSE);
		ListView_SetItemState(hGridWnd, rowNo, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
		
		TCHAR* val = cache[resultset[rowNo]][colNo];
		int len = _tcslen(searchString);
		_sntprintf(buf, 255, TEXT("%ls%.*ls%ls"),
			pos > 0 ? TEXT("...") : TEXT(""), 
			len + pos + 10, val + pos,
			_tcslen(val + pos + len) > 10 ? TEXT("...") : TEXT(""));
		SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
	} else { 
		MessageBox(hWnd, searchString, TEXT("Not found:"), MB_OK);
	}
	SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)buf);	
	SetFocus(hGridWnd);

	return 0;
}

int __stdcall ListSearchText(HWND hWnd, char* searchString, int searchParameter) {
	DWORD len = MultiByteToWideChar(CP_ACP, 0, searchString, -1, NULL, 0);
	TCHAR* searchString16 = (TCHAR*)calloc (len, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, searchString, -1, searchString16, len);
	int rc = ListSearchTextW(hWnd, searchString16, searchParameter);
	free(searchString16);
	return rc;
}	

HWND APIENTRY ListLoadW (HWND hListerWnd, const TCHAR* fileToLoad, int showFlags) {		
	int size = _tcslen(fileToLoad);
	TCHAR* filepath = (TCHAR*)calloc(size + 1, sizeof(TCHAR));
	_tcsncpy(filepath, fileToLoad, size);
		
	int maxFileSize = getStoredValue(TEXT("max-file-size"), 10000000);	
	struct _stat st = {0};
	if (_tstat(filepath, &st) != 0 || st.st_size == 0/* || maxFileSize > 0 && st.st_size > maxFileSize*/)
		return 0;

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&icex);
	
	setlocale(LC_CTYPE, "");

	BOOL isStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindow(WC_STATIC, APP_NAME, WS_CHILD | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&mainWndProc));
	SetProp(hMainWnd, TEXT("FILEPATH"), filepath);
	SetProp(hMainWnd, TEXT("FILESIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("DELIMITER"), calloc(1, sizeof(TCHAR)));
	SetProp(hMainWnd, TEXT("CODEPAGE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERROW"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("HEADERROW"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SKIPCOMMENTS"), calloc(1, sizeof(int)));		
	SetProp(hMainWnd, TEXT("CACHE"), 0);
	SetProp(hMainWnd, TEXT("RESULTSET"), 0);
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("COLCOUNT"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("CURRENTROWNO"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CURRENTCOLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SEARCHCELLPOS"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTFAMILY"), getStoredString(TEXT("font"), TEXT("Arial")));
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERALIGN"), calloc(1, sizeof(int)));	
	
	SetProp(hMainWnd, TEXT("DARKTHEME"), calloc(1, sizeof(int)));			
	SetProp(hMainWnd, TEXT("TEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR2"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("FILTERTEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERBACKCOLOR"), calloc(1, sizeof(int)));		
	SetProp(hMainWnd, TEXT("CURRENTCELLCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SELECTIONTEXTCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SELECTIONBACKCOLOR"), calloc(1, sizeof(int)));		
	
	*(int*)GetProp(hMainWnd, TEXT("FILESIZE")) = st.st_size;
	*(int*)GetProp(hMainWnd, TEXT("CODEPAGE")) = -1;
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);	
	*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) = getStoredValue(TEXT("filter-row"), 1);
	*(int*)GetProp(hMainWnd, TEXT("HEADERROW")) = getStoredValue(TEXT("header-row"), 1);	
	*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) = getStoredValue(TEXT("dark-theme"), 0);	
	*(int*)GetProp(hMainWnd, TEXT("SKIPCOMMENTS")) = getStoredValue(TEXT("skip-comments"), 0);
	*(int*)GetProp(hMainWnd, TEXT("FILTERALIGN")) = getStoredValue(TEXT("filter-align"), 0);
		
	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE | SBT_TOOLTIPS | (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	HDC hDC = GetDC(hMainWnd);
	float z = GetDeviceCaps(hDC, LOGPIXELSX) / 96.0; // 96 = 100%, 120 = 125%, 144 = 150%
	ReleaseDC(hMainWnd, hDC);	
	int sizes[7] = {35 * z, 95 * z, 125 * z, 235 * z, 355 * z, 435 * z, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 7, (LPARAM)&sizes);
	TCHAR buf[32];
	_sntprintf(buf, 32, TEXT(" %ls"), APP_VERSION);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VERSION, (LPARAM)buf);
	SendMessage(hStatusWnd, SB_SETTIPTEXT, SB_COMMENTS, (LPARAM)TEXT("How to process lines starting with #. Check Wiki to get details."));

	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hMainWnd, (HMENU)IDC_GRID, GetModuleHandle(0), NULL);
		
	int noLines = getStoredValue(TEXT("disable-grid-lines"), 0);	
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | (noLines ? 0 : LVS_EX_GRIDLINES) | LVS_EX_LABELTIP | LVS_EX_HEADERDRAGDROP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));

	HWND hHeader = ListView_GetHeader(hGridWnd);
	SetWindowTheme(hHeader, TEXT(" "), TEXT(" "));
	SetProp(hHeader, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hHeader, GWLP_WNDPROC, (LONG_PTR)cbNewHeader));	 

	HMENU hGridMenu = CreatePopupMenu();
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_ROWS, TEXT("Copy row(s)"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_COLUMN, TEXT("Copy column"));	
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);
	AppendMenu(hGridMenu, MF_STRING, IDM_HIDE_COLUMN, TEXT("Hide column"));	
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);	
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_FILTER_ROW, TEXT("Filters"));		
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("HEADERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_HEADER_ROW, TEXT("Header row"));	
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_DARK_THEME, TEXT("Dark theme"));		
	SetProp(hMainWnd, TEXT("GRIDMENU"), hGridMenu);

	HMENU hCodepageMenu = CreatePopupMenu();
	AppendMenu(hCodepageMenu, MF_STRING, IDM_ANSI, TEXT("ANSI"));
	AppendMenu(hCodepageMenu, MF_STRING, IDM_UTF8, TEXT("UTF-8"));
	AppendMenu(hCodepageMenu, MF_STRING, IDM_UTF16LE, TEXT("UTF-16LE"));
	AppendMenu(hCodepageMenu, MF_STRING, IDM_UTF16BE, TEXT("UTF-16BE"));	
	SetProp(hMainWnd, TEXT("CODEPAGEMENU"), hCodepageMenu);

	HMENU hDelimiterMenu = CreatePopupMenu();
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_COMMA, TEXT(","));
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_SEMICOLON, TEXT(";"));
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_VBAR, TEXT("|"));
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_TAB, TEXT("TAB"));	
	AppendMenu(hDelimiterMenu, MF_STRING, IDM_COLON, TEXT(":"));	
	SetProp(hMainWnd, TEXT("DELIMITERMENU"), hDelimiterMenu);
	
	HMENU hCommentMenu = CreatePopupMenu();
	AppendMenu(hCommentMenu, MF_STRING, IDM_DEFAULT, TEXT("#0 - Parse as usual"));
	AppendMenu(hCommentMenu, MF_STRING, IDM_NO_PARSE, TEXT("#1 - Don't parse"));
	AppendMenu(hCommentMenu, MF_STRING, IDM_NO_SHOW, TEXT("#2 - Ignore (hide)"));
	SetProp(hMainWnd, TEXT("COMMENTMENU"), hCommentMenu);	

	SendMessage(hMainWnd, WMU_SET_FONT, 0, 0);
	SendMessage(hMainWnd, WMU_SET_THEME, 0, 0);	
	SendMessage(hMainWnd, WMU_UPDATE_CACHE, 0, 0);
	ShowWindow(hMainWnd, SW_SHOW);
	SetFocus(hGridWnd);
	
	return hMainWnd;
}

HWND APIENTRY ListLoad (HWND hListerWnd, const char* fileToLoad, int showFlags) {
	DWORD size = MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, NULL, 0);
	TCHAR* fileToLoadW = (TCHAR*)calloc (size, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, fileToLoadW, size);
	HWND hWnd = ListLoadW(hListerWnd, fileToLoadW, showFlags);
	free(fileToLoadW);
	return hWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	setStoredValue(TEXT("filter-row"), *(int*)GetProp(hWnd, TEXT("FILTERROW")));		
	setStoredValue(TEXT("header-row"), *(int*)GetProp(hWnd, TEXT("HEADERROW")));	
	setStoredValue(TEXT("dark-theme"), *(int*)GetProp(hWnd, TEXT("DARKTHEME")));
	setStoredValue(TEXT("skip-comments"), *(int*)GetProp(hWnd, TEXT("SKIPCOMMENTS")));
	
	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
	free((TCHAR*)GetProp(hWnd, TEXT("FILEPATH")));
	free((int*)GetProp(hWnd, TEXT("FILESIZE")));	
	free((TCHAR*)GetProp(hWnd, TEXT("DELIMITER")));
	free((int*)GetProp(hWnd, TEXT("CODEPAGE")));
	free((int*)GetProp(hWnd, TEXT("FILTERROW")));
	free((int*)GetProp(hWnd, TEXT("HEADERROW")));	
	free((int*)GetProp(hWnd, TEXT("DARKTHEME")));		
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((int*)GetProp(hWnd, TEXT("COLCOUNT")));	
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("FONTSIZE")));
	free((int*)GetProp(hWnd, TEXT("CURRENTROWNO")));				
	free((int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));
	free((int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")));	
	free((TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
	free((int*)GetProp(hWnd, TEXT("FILTERALIGN")));	
		
	free((int*)GetProp(hWnd, TEXT("TEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR2")));	
	free((int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")));
	free((int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("SELECTIONBACKCOLOR")));		

	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));	
	DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
	DestroyMenu((HMENU)GetProp(hWnd, TEXT("GRIDMENU")));
	DestroyMenu((HMENU)GetProp(hWnd, TEXT("CODEPAGEMENU")));
	DestroyMenu((HMENU)GetProp(hWnd, TEXT("DELIMITERMENU")));
	DestroyMenu((HMENU)GetProp(hWnd, TEXT("COMMENTMENU")));

	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("CACHE"));
	RemoveProp(hWnd, TEXT("RESULTSET"));
	RemoveProp(hWnd, TEXT("FILEPATH"));
	RemoveProp(hWnd, TEXT("FILESIZE"));
	RemoveProp(hWnd, TEXT("DELIMITER"));
	RemoveProp(hWnd, TEXT("CODEPAGE"));
	RemoveProp(hWnd, TEXT("FILTERROW"));	
	RemoveProp(hWnd, TEXT("HEADERROW"));	
	RemoveProp(hWnd, TEXT("DARKTHEME"));	
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("COLCOUNT"));
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));
	RemoveProp(hWnd, TEXT("CURRENTROWNO"));	
	RemoveProp(hWnd, TEXT("CURRENTCOLNO"));
	RemoveProp(hWnd, TEXT("SEARCHCELLPOS"));	
	RemoveProp(hWnd, TEXT("FILTERALIGN"));	
	RemoveProp(hWnd, TEXT("LASTFOCUS"));	
	
	RemoveProp(hWnd, TEXT("TEXTCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR2"));	
	RemoveProp(hWnd, TEXT("FILTERTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("FILTERBACKCOLOR"));
	RemoveProp(hWnd, TEXT("CURRENTCELLCOLOR"));
	RemoveProp(hWnd, TEXT("SELECTIONTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("SELECTIONBACKCOLOR"));
	RemoveProp(hWnd, TEXT("BACKBRUSH"));
	RemoveProp(hWnd, TEXT("FILTERBACKBRUSH"));	
	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("FONTFAMILY"));
	RemoveProp(hWnd, TEXT("FONTSIZE"));	
	RemoveProp(hWnd, TEXT("GRIDMENU"));
	RemoveProp(hWnd, TEXT("CODEPAGEMENU"));	
	RemoveProp(hWnd, TEXT("DELIMITERMENU"));	
	RemoveProp(hWnd, TEXT("COMMENTMENU"));	

	DestroyWindow(hWnd);
}

LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_KEYDOWN && SendMessage(getMainWindow(hWnd), WMU_HOT_KEYS, wParam, lParam))
		return 0;

	// Prevent beep
	if (msg == WM_CHAR && SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
		return 0;	

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_CTLCOLOREDIT) {
		HWND hMainWnd = getMainWindow(hWnd);
		SetBkColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
		SetTextColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR")));
		return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));	
	}
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetClientRect(hWnd, &rc);
			HWND hMainWnd = getMainWindow(hWnd);
			BOOL isDark = *(int*)GetProp(hMainWnd, TEXT("DARKTHEME")); 

			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
			HPEN oldPen = (HPEN)SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			LineTo(hDC, rc.right - 1, rc.bottom - 1);
			
			if (isDark) {
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
		
		case WM_SETFOCUS: {
			SetProp(getMainWindow(hWnd), TEXT("LASTFOCUS"), hWnd);
		}
		break;

		case WM_KEYDOWN: {
			HWND hMainWnd = getMainWindow(hWnd);
			if (wParam == VK_RETURN) {
				SendMessage(hMainWnd, WMU_UPDATE_RESULTSET, 0, 0);
				return 0;			
			}
			
			if (SendMessage(hMainWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;
		}
		break;
	
		// Prevent beep
		case WM_CHAR: {
			if (SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
				return 0;	
		}
		break;
		
		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;
	}

	return CallWindowProc(cbDefault, hWnd, msg, wParam, lParam);
}

int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam) {
	if (GetWindowLong(hWnd, GWL_STYLE) & WS_TABSTOP && IsWindowVisible(hWnd)) {
		int no = 0;
		HWND* wnds = (HWND*)lParam;
		while (wnds[no])
			no++;
		wnds[no] = hWnd;
	}

	return TRUE;
}

int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt) {
	int colNo = Header_GetItemCount(ListView_GetHeader(hListWnd));
	LVCOLUMN lvc = {0};
	lvc.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM | LVCF_FMT;
	lvc.iSubItem = colNo;
	lvc.pszText = colName;
	lvc.cchTextMax = _tcslen(colName) + 1;
	lvc.cx = 100;
	lvc.fmt = fmt;
	return ListView_InsertColumn(hListWnd, colNo, &lvc);
}

int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, const int cchTextMax) {
	if (i < 0)
		return FALSE;

	TCHAR* buf = new TCHAR[cchTextMax];

	HDITEM hdi = {0};
	hdi.mask = HDI_TEXT;
	hdi.pszText = buf;
	hdi.cchTextMax = cchTextMax;
	int rc = Header_GetItem(hWnd, i, &hdi);

	_tcsncpy(pszText, buf, cchTextMax);

	delete [] buf;
	return rc;
}

void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState) {
	MENUITEMINFO mii = {0};
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_STATE;
	mii.fState = fState;
	SetMenuItemInfo(hMenu, wID, FALSE, &mii);
}