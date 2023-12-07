
// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "exportDef.h"

TCHAR g_szIniPath[MAX_PATH] = { 0 };

BOOL APIENTRY DllMain (HANDLE hModule, DWORD dwUlReasonForCall, LPVOID pvReserved) 
{
	if (dwUlReasonForCall == DLL_PROCESS_ATTACH && g_szIniPath[0] == 0) 
	{
		TCHAR path[MAX_PATH + 1] = {0};
		GetModuleFileName((HMODULE)hModule, path, MAX_PATH);
		TCHAR* dot = _tcsrchr(path, TEXT('.'));
		_tcsncpy(dot, TEXT(".ini"), 5);
		if (_taccess(path, 0) == 0)
			_tcscpy(g_szIniPath, path);	
	}
	return TRUE;
}

void __stdcall ListGetDetectString(char* pszDetectString, int nMaxlen) 
{
	TCHAR* pszDetectString16 = getStoredString(TEXT("detect-string"), TEXT("MULTIMEDIA & (ext=\"CSV\" | ext=\"TAB\" | ext=\"TSV\")"));
	char* pszDetectString8 = Utf16to8(pszDetectString16);
	snprintf(pszDetectString, nMaxlen, pszDetectString8);
	free(pszDetectString16);
	free(pszDetectString8);
}

void __stdcall ListSetDefaultParams(ListDefaultParamStruct* pDps)
{
	if (g_szIniPath[0] == 0) 
	{
		DWORD size = MultiByteToWideChar(CP_ACP, 0, pDps->szDefaultIniName, -1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, pDps->szDefaultIniName, -1, g_szIniPath, size);
	}
}

int __stdcall ListSearchTextW(HWND hWnd, TCHAR* pszSearchString, int nSearchParameter) 
{
	HWND hGridWnd	= GetDlgItem(hWnd, IDC_GRID);
	HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);	
	
	TCHAR***	pppCache	= (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
	int*		pnResultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
	int			nRowCount	= *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
	int			nColCount	= Header_GetItemCount(ListView_GetHeader(hGridWnd));

	if (!pnResultset || nRowCount == 0)
		return 0;

	BOOL bisFindFirst	= nSearchParameter & LCS_FINDFIRST;		
	BOOL bisBackward	= nSearchParameter & LCS_BACKWARDS;
	BOOL bIsMatchCase	= nSearchParameter & LCS_MATCHCASE;
	BOOL bIsWholeWords	= nSearchParameter & LCS_WHOLEWORDS;	

	if (bisFindFirst) {
		*(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) = 0;
		*(int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")) = 0;	
		*(int*)GetProp(hWnd, TEXT("CURRENTROWNO")) = bisBackward ? nRowCount - 1 : 0;
	}	
		
	int		nRowNo		= *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
	int		nColNo		= *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
	int *	pnStartPos	= (int*)GetProp(hWnd, TEXT("SEARCHCELLPOS"));	
	nRowNo = nRowNo == -1 || nRowNo >= nRowCount ? 0 : nRowNo;
	nColNo = nColNo == -1 || nColNo >= nColCount ? 0 : nColNo;	
			
	int nPos = -1;
	do {
		for (; (nPos == -1) && nColNo < nColCount; nColNo++) {
			nPos = findString(pppCache[pnResultset[nRowNo]][nColNo] + *pnStartPos, pszSearchString, bIsMatchCase, bIsWholeWords);
			if (nPos != -1) 
				nPos += *pnStartPos;
			*pnStartPos = (int)(nPos == -1 ? 0 : nPos + _tcslen(pszSearchString));
		}
		nColNo = nPos != -1 ? nColNo - 1 : 0;
		nRowNo += nPos != -1 ? 0 : bisBackward ? -1 : 1; 	
	} while ((nPos == -1) && (bisBackward ? nRowNo > 0 : nRowNo < nRowCount));
	ListView_SetItemState(hGridWnd, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);

	TCHAR szBuf[256] = {0};
	if (nPos != -1)
	{
		ListView_EnsureVisible(hGridWnd, nRowNo, FALSE);
		ListView_SetItemState(hGridWnd, nRowNo, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
		
		TCHAR* pszVal = pppCache[pnResultset[nRowNo]][nColNo];
		int nLen = (int)_tcslen(pszSearchString);
		_sntprintf(szBuf, 255, TEXT("%ls%.*ls%ls"),
			nPos > 0 ? TEXT("...") : TEXT(""), 
			nLen + nPos + 10, pszVal + nPos,
			_tcslen(pszVal + nPos + nLen) > 10 ? TEXT("...") : TEXT(""));
		SendMessage(hWnd, WMU_SET_CURRENT_CELL, nRowNo, nColNo);
	} 
	else 
	{ 
		MessageBox(hWnd, pszSearchString, TEXT("Not found:"), MB_OK);
	}
	SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)szBuf);	
	SetFocus(hGridWnd);

	return 0;
}

int __stdcall ListSearchText(HWND hWnd, char* pszSearchString, int nSearchParameter) {
	DWORD nLen = MultiByteToWideChar(CP_ACP, 0, pszSearchString, -1, NULL, 0);
	TCHAR* pszSearchString16 = (TCHAR*)calloc (nLen, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, pszSearchString, -1, pszSearchString16, nLen);
	int nResult = ListSearchTextW(hWnd, pszSearchString16, nSearchParameter);
	free(pszSearchString16);
	return nResult;
}	

HWND APIENTRY ListLoadW (HWND hListerWnd, const TCHAR* cpszFileToLoad, int nShowFlags) {		
	int nSize = (int)_tcslen(cpszFileToLoad);
	TCHAR* pszFilepath = (TCHAR*)calloc(nSize + 1, sizeof(TCHAR));
	_tcsncpy(pszFilepath, cpszFileToLoad, nSize);
		
	int nMaxFileSize = getStoredValue(TEXT("max-file-nSize"), 10000000);	
	struct _stat st = {0};
	if (_tstat(pszFilepath, &st) != 0 || st.st_size == 0/* || maxFileSize > 0 && st.st_size > maxFileSize*/)
		return 0;

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&icex);
	
	setlocale(LC_CTYPE, "");

	BOOL bisStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindow(WC_STATIC, APP_NAME, WS_CHILD | (bisStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&MainWndProc));
	SetProp(hMainWnd, TEXT("FILEPATH"), pszFilepath);
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
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-nSize"), 16);	
	*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) = getStoredValue(TEXT("filter-row"), 1);
	*(int*)GetProp(hMainWnd, TEXT("HEADERROW")) = getStoredValue(TEXT("header-row"), 1);	
	*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) = getStoredValue(TEXT("dark-theme"), 0);	
	*(int*)GetProp(hMainWnd, TEXT("SKIPCOMMENTS")) = getStoredValue(TEXT("skip-comments"), 0);
	*(int*)GetProp(hMainWnd, TEXT("FILTERALIGN")) = getStoredValue(TEXT("filter-align"), 0);
		
	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE | SBT_TOOLTIPS | (bisStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	HDC hDC = GetDC(hMainWnd);
	float fZ = (float)(GetDeviceCaps(hDC, LOGPIXELSX) / 96.0); // 96 = 100%, 120 = 125%, 144 = 150%
	ReleaseDC(hMainWnd, hDC);	
	int nSizes[7] = {(int)(35 * fZ), (int)(95 * fZ), (int)(125 * fZ), (int)(235 * fZ), (int)(355 * fZ), (int)(435 * fZ), -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 7, (LPARAM)&nSizes);
	TCHAR szBuf[32];
	_sntprintf(szBuf, 32, TEXT(" %ls"), APP_VERSION);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VERSION, (LPARAM)szBuf);
	SendMessage(hStatusWnd, SB_SETTIPTEXT, SB_COMMENTS, (LPARAM)TEXT("How to process lines starting with #. Check Wiki to get details."));

	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hMainWnd, (HMENU)IDC_GRID, GetModuleHandle(0), NULL);
		
	int nNoLines = getStoredValue(TEXT("disable-grid-lines"), 0);	
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | (nNoLines ? 0 : LVS_EX_GRIDLINES) | LVS_EX_LABELTIP | LVS_EX_HEADERDRAGDROP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)CallHotKeyProc));

	HWND hHeader = ListView_GetHeader(hGridWnd);
	SetWindowTheme(hHeader, TEXT(" "), TEXT(" "));
	SetProp(hHeader, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hHeader, GWLP_WNDPROC, (LONG_PTR)CallNewHeaderProc));	 

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

HWND APIENTRY ListLoad (HWND hListerWnd, const char* cpszFileToLoad, int nShowFlags) {
	DWORD nSize = MultiByteToWideChar(CP_ACP, 0, cpszFileToLoad, -1, NULL, 0);
	TCHAR* pszFileToLoadW = (TCHAR*)calloc (nSize, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, cpszFileToLoad, -1, pszFileToLoadW, nSize);
	HWND hWnd = ListLoadW(hListerWnd, pszFileToLoadW, nShowFlags);
	free(pszFileToLoadW);
	return hWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("font-nSize"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
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

int ListView_AddColumn(HWND hListWnd, TCHAR* pszColName, int nFmt) 
{
	int nColNo = Header_GetItemCount(ListView_GetHeader(hListWnd));
	LVCOLUMN lvc = {0};
	lvc.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM | LVCF_FMT;
	lvc.iSubItem = nColNo;
	lvc.pszText = pszColName;
	lvc.cchTextMax = (int)_tcslen(pszColName) + 1;
	lvc.cx = 100;
	lvc.fmt = nFmt;
	return ListView_InsertColumn(hListWnd, nColNo, &lvc);
}

int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, const int cnCchTextMax)
{
	if (i < 0)
		return FALSE;

	TCHAR* pszBuf = new TCHAR[cnCchTextMax];

	HDITEM hdi = {0};
	hdi.mask = HDI_TEXT;
	hdi.pszText = pszBuf;
	hdi.cchTextMax = cnCchTextMax;
	int nResult = Header_GetItem(hWnd, i, &hdi);

	_tcsncpy(pszText, pszBuf, cnCchTextMax);

	delete [] pszBuf;
	return nResult;
}

void Menu_SetItemState(HMENU hMenu, UINT wID, UINT uState) 
{
	MENUITEMINFO mii = {0};
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_STATE;
	mii.fState = uState;
	SetMenuItemInfo(hMenu, wID, FALSE, &mii);
}