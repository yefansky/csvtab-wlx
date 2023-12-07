#include "stdafx.h"

HWND GetMainWindow(HWND hWnd) {
	HWND hMainWnd = hWnd;
	while (hMainWnd && GetDlgCtrlID(hMainWnd) != IDC_MAIN)
		hMainWnd = GetParent(hMainWnd);
	return hMainWnd;
}

int OnMsgCommand(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int nResult = -1;
	WORD wCmd = LOWORD(wParam);
	if (wCmd == IDM_COPY_CELL || wCmd == IDM_COPY_ROWS || wCmd == IDM_COPY_COLUMN)
	{
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		HWND hHeader = ListView_GetHeader(hGridWnd);
		int nRowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
		int nColNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));

		int nColCount = Header_GetItemCount(hHeader);
		int nRowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
		int nSelCount = ListView_GetSelectedCount(hGridWnd);

		if (nRowNo == -1 ||
			nRowNo >= nRowCount ||
			nColCount == 0 ||
			wCmd == IDM_COPY_CELL && nColNo == -1 ||
			wCmd == IDM_COPY_CELL && nColNo >= nColCount ||
			wCmd == IDM_COPY_COLUMN && nColNo == -1 ||
			wCmd == IDM_COPY_COLUMN && nColNo >= nColCount ||
			wCmd == IDM_COPY_ROWS && nSelCount == 0) 
		{
			setClipboardText(TEXT(""));

			nResult = 0;
			goto Exit0;
		}

		TCHAR*** pppszCache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
		int* pnResultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
		if (!pnResultset)
		{
			nResult = 0;
			goto Exit0;
		}

		TCHAR* pszDelimiter = getStoredString(TEXT("column-pszDelimiter"), TEXT("\t"));

		int nLen = 0;
		if (wCmd == IDM_COPY_CELL)
			nLen = (int)_tcslen(pppszCache[pnResultset[nRowNo]][nColNo]);

		if (wCmd == IDM_COPY_ROWS) 
		{
			int nRowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
			while (nRowNo != -1) 
			{
				for (int colNo = 0; colNo < nColCount; colNo++) 
				{
					if (ListView_GetColumnWidth(hGridWnd, colNo))
						nLen += (int)_tcslen(pppszCache[pnResultset[nRowNo]][colNo]) + 1; /* column delimiter */
				}

				nLen++; /* \n */
				nRowNo = ListView_GetNextItem(hGridWnd, nRowNo, LVNI_SELECTED);
			}
		}

		if (wCmd == IDM_COPY_COLUMN) 
		{
			int nRowNo = nSelCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
			while (nRowNo != -1 && nRowNo < nRowCount) 
			{
				nLen += (int)_tcslen(pppszCache[pnResultset[nRowNo]][nColNo]) + 1 /* \n */;
				nRowNo = nSelCount < 2 ? nRowNo + 1 : ListView_GetNextItem(hGridWnd, nRowNo, LVNI_SELECTED);
			}
		}

		TCHAR* pszBuf = (TCHAR*)calloc(nLen + 1, sizeof(TCHAR));
		if (wCmd == IDM_COPY_CELL)
			_tcscat(pszBuf, pppszCache[pnResultset[nRowNo]][nColNo]);

		if (wCmd == IDM_COPY_ROWS) {
			int nPos = 0;
			int nRowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);

			int* pnColOrder = (int*)calloc(nColCount, sizeof(int));
			Header_GetOrderArray(hHeader, nColCount, pnColOrder);

			while (nRowNo != -1) 
			{
				for (int idx = 0; idx < nColCount; idx++) 
				{
					int colNo = pnColOrder[idx];
					if (ListView_GetColumnWidth(hGridWnd, colNo)) 
					{
						int len = (int)_tcslen(pppszCache[pnResultset[nRowNo]][colNo]);
						_tcsncpy(pszBuf + nPos, pppszCache[pnResultset[nRowNo]][colNo], len);
						pszBuf[nPos + len] = pszDelimiter[0];
						nPos += len + 1;
					}
				}

				pszBuf[nPos - (nPos > 0)] = TEXT('\n');
				nRowNo = ListView_GetNextItem(hGridWnd, nRowNo, LVNI_SELECTED);
			}
			pszBuf[nPos - 1] = 0; // remove last \n

			free(pnColOrder);
		}

		if (wCmd == IDM_COPY_COLUMN)
		{
			int nPos = 0;
			int nRowNo = nSelCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
			while (nRowNo != -1 && nRowNo < nRowCount) 
			{
				int nLen = (int)_tcslen(pppszCache[pnResultset[nRowNo]][nColNo]);
				_tcsncpy(pszBuf + nPos, pppszCache[pnResultset[nRowNo]][nColNo], nLen);
				nRowNo = nSelCount < 2 ? nRowNo + 1 : ListView_GetNextItem(hGridWnd, nRowNo, LVNI_SELECTED);
				if (nRowNo != -1 && nRowNo < nRowCount)
					pszBuf[nPos + nLen] = TEXT('\n');
				nPos += nLen + 1;
			}
		}

		setClipboardText(pszBuf);
		free(pszBuf);
		free(pszDelimiter);
	}

	if (wCmd == IDM_HIDE_COLUMN) {
		int nColNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
		SendMessage(hWnd, WMU_HIDE_COLUMN, nColNo, 0);
	}

	if (wCmd == IDM_FILTER_ROW || wCmd == IDM_HEADER_ROW || wCmd == IDM_DARK_THEME)
	{
		HMENU hMenu = (HMENU)GetProp(hWnd, TEXT("GRIDMENU"));
		int* pnOpt = (int*)GetProp(hWnd, wCmd == IDM_FILTER_ROW ? TEXT("FILTERROW") : wCmd == IDM_HEADER_ROW ? TEXT("HEADERROW" : TEXT("DARKTHEME")));
		*pnOpt = (*pnOpt + 1) % 2;
		Menu_SetItemState(hMenu, wCmd, *pnOpt ? MFS_CHECKED : 0);

		UINT uMsg = wCmd == IDM_FILTER_ROW ? WMU_SET_HEADER_FILTERS : wCmd == IDM_HEADER_ROW ? WMU_UPDATE_GRID : WMU_SET_THEME;
		SendMessage(hWnd, uMsg, 0, 0);
	}

	if (wCmd == IDM_ANSI || wCmd == IDM_UTF8 || wCmd == IDM_UTF16LE || wCmd == IDM_UTF16BE) 
	{
		int* pnCodepage = (int*)GetProp(hWnd, TEXT("CODEPAGE"));
		*pnCodepage = wCmd == IDM_ANSI ? CP_ACP : wCmd == IDM_UTF8 ? CP_UTF8 : wCmd == IDM_UTF16LE ? CP_UTF16LE : CP_UTF16BE;
		SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
	}

	if (wCmd == IDM_COMMA || wCmd == IDM_SEMICOLON || wCmd == IDM_VBAR || wCmd == IDM_TAB || wCmd == IDM_COLON)
	{
		TCHAR* pszDelimiter = (TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
		*pszDelimiter = wCmd == IDM_COMMA ? TEXT(',') :
			wCmd == IDM_SEMICOLON ? TEXT(';') :
			wCmd == IDM_TAB ? TEXT('\t') :
			wCmd == IDM_VBAR ? TEXT('|') :
			wCmd == IDM_COLON ? TEXT(':') :
			0;
		SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
	}

	if (wCmd == IDM_DEFAULT || wCmd == IDM_NO_PARSE || wCmd == IDM_NO_SHOW) 
	{
		int* pnSkipComments = (int*)GetProp(hWnd, TEXT("SKIPCOMMENTS"));
		*pnSkipComments = wCmd == IDM_DEFAULT ? 0 : wCmd == IDM_NO_PARSE ? 1 : 2;
		SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
	}
Exit0:
	return nResult;
}

LRESULT OnMsgNotify(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LRESULT nResult = -1;
	NMHDR* pHdr = (LPNMHDR)lParam;
	if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_GETDISPINFO) 
	{
		LV_DISPINFO* pDispInfo = (LV_DISPINFO*)lParam;
		LV_ITEM* pItem = &(pDispInfo)->item;

		TCHAR*** pppszCache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
		int* pnResultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
		if (pnResultset && pItem->mask & LVIF_TEXT) 
		{
			int rowNo = pnResultset[pItem->iItem];
			pItem->pszText = pppszCache[rowNo][pItem->iSubItem];
		}
	}

	if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_COLUMNCLICK) 
	{
		NMLISTVIEW* pLv = (NMLISTVIEW*)lParam;
		nResult = SendMessage(hWnd, HIWORD(GetKeyState(VK_CONTROL)) ? WMU_HIDE_COLUMN : WMU_SORT_COLUMN, pLv->iSubItem, 0);
		goto Exit0;
	}

	if (pHdr->idFrom == IDC_GRID && (pHdr->code == (DWORD)NM_CLICK || pHdr->code == (DWORD)NM_RCLICK)) 
	{
		NMITEMACTIVATE* ia = (LPNMITEMACTIVATE)lParam;
		SendMessage(hWnd, WMU_SET_CURRENT_CELL, ia->iItem, ia->iSubItem);
	}

	if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_CLICK && HIWORD(GetKeyState(VK_MENU))) 
	{
		NMITEMACTIVATE* pIa = (LPNMITEMACTIVATE)lParam;
		TCHAR*** pppszcache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
		int* pnResultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));

		TCHAR* pszUrl = extractUrl(pppszcache[pnResultset[pIa->iItem]][pIa->iSubItem]);
		ShellExecute(0, TEXT("open"), pszUrl, 0, 0, SW_SHOW);
		free(pszUrl);
	}

	if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_RCLICK)
	{
		POINT p;
		GetCursorPos(&p);
		TrackPopupMenu((HMENU)GetProp(hWnd, TEXT("GRIDMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
	}

	if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
		NMLISTVIEW* pLv = (NMLISTVIEW*)lParam;
		if (pLv->uOldState != pLv->uNewState && (pLv->uNewState & LVIS_SELECTED))
			SendMessage(hWnd, WMU_SET_CURRENT_CELL, pLv->iItem, *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));
	}

	if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) 
	{
		NMLVKEYDOWN* pKd = (LPNMLVKEYDOWN)lParam;
		if (pKd->wVKey == 0x43) 
		{ // C
			BOOL bisCtrl = HIWORD(GetKeyState(VK_CONTROL));
			BOOL bisShift = HIWORD(GetKeyState(VK_SHIFT));
			BOOL bisCopyColumn = getStoredValue(TEXT("copy-column"), 0) && ListView_GetSelectedCount(pHdr->hwndFrom) > 1;
			if (!bisCtrl && !bisShift)
			{
				nResult = FALSE;
				goto Exit0;
			}

			int nAction = !bisShift && !bisCopyColumn ? IDM_COPY_CELL : bisCtrl || bisCopyColumn ? IDM_COPY_COLUMN : IDM_COPY_ROWS;
			SendMessage(hWnd, WM_COMMAND, nAction, 0);
			nResult = TRUE;
			goto Exit0;
		}

		if (pKd->wVKey == 0x41 && HIWORD(GetKeyState(VK_CONTROL))) 
		{ // Ctrl + A
			HWND hGridWnd = pHdr->hwndFrom;
			SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
			int nRowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
			int nColNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
			ListView_SetItemState(hGridWnd, -1, LVIS_SELECTED, LVIS_SELECTED | LVIS_FOCUSED);
			SendMessage(hWnd, WMU_SET_CURRENT_CELL, nRowNo, nColNo);
			SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hGridWnd, NULL, TRUE);
		}

		if (pKd->wVKey == 0x20 && HIWORD(GetKeyState(VK_CONTROL))) 
		{ // Ctrl + Space				
			SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);
			nResult = TRUE;
			goto Exit0;
		}

		if (pKd->wVKey == VK_LEFT || pKd->wVKey == VK_RIGHT) 
		{
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);

			int nColCount = Header_GetItemCount(ListView_GetHeader(pHdr->hwndFrom));
			int nColNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));

			int* pnColOrder = (int*)calloc(nColCount, sizeof(int));
			Header_GetOrderArray(hHeader, nColCount, pnColOrder);

			int nDir = pKd->wVKey == VK_RIGHT ? 1 : -1;
			int idx = 0;
			for (idx; pnColOrder[idx] != nColNo; idx++)
				NULL;

			do {
				idx = (nColCount + idx + nDir) % nColCount;
			} while (ListView_GetColumnWidth(hGridWnd, pnColOrder[idx]) == 0);

			nColNo = pnColOrder[idx];
			free(pnColOrder);

			SendMessage(hWnd, WMU_SET_CURRENT_CELL, *(int*)GetProp(hWnd, TEXT("CURRENTROWNO")), nColNo);
			nResult = TRUE;
			goto Exit0;
		}
	}

	if ((pHdr->code == HDN_ITEMCHANGED || pHdr->code == HDN_ENDDRAG) && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(hWnd, IDC_GRID)))
		PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);

	if (pHdr->idFrom == IDC_STATUSBAR && (pHdr->code == NM_CLICK || pHdr->code == NM_RCLICK)) 
	{
		NMMOUSE* pM = (NMMOUSE*)lParam;
		ptrdiff_t nID = pM->dwItemSpec;
		if (nID != SB_CODEPAGE && nID != SB_DELIMITER && nID != SB_COMMENTS)
		{
			nResult = 0;
			goto Exit0;
		}
		
		RECT rc, rc2;
		GetWindowRect(pHdr->hwndFrom, &rc);
		SendMessage(pHdr->hwndFrom, SB_GETRECT, nID, (LPARAM)&rc2);
		HMENU hMenu = (HMENU)GetProp(hWnd, nID == SB_CODEPAGE ? TEXT("CODEPAGEMENU") : nID == SB_DELIMITER ? TEXT("DELIMITERMENU") : TEXT("COMMENTMENU"));

		if (nID == SB_CODEPAGE) {
			int codepage = *(int*)GetProp(hWnd, TEXT("CODEPAGE"));
			Menu_SetItemState(hMenu, IDM_ANSI, codepage == CP_ACP ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_UTF8, codepage == CP_UTF8 ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_UTF16LE, codepage == CP_UTF16LE ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_UTF16BE, codepage == CP_UTF16BE ? MFS_CHECKED : 0);
		}

		if (nID == SB_DELIMITER) {
			TCHAR delimiter = *(TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
			Menu_SetItemState(hMenu, IDM_COMMA, delimiter == TEXT(',') ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_SEMICOLON, delimiter == TEXT(';') ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_VBAR, delimiter == TEXT('|') ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_TAB, delimiter == TEXT('\t') ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_COLON, delimiter == TEXT(':') ? MFS_CHECKED : 0);
		}

		if (nID == SB_COMMENTS) {
			int skipComments = *(int*)GetProp(hWnd, TEXT("SKIPCOMMENTS"));
			Menu_SetItemState(hMenu, IDM_DEFAULT, skipComments == 0 ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_NO_PARSE, skipComments == 1 ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_NO_SHOW, skipComments == 2 ? MFS_CHECKED : 0);
		}

		POINT p = { rc.left + rc2.left, rc.top };
		TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_BOTTOMALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
	}

	if (pHdr->code == (UINT)NM_SETFOCUS)
		SetProp(hWnd, TEXT("LASTFOCUS"), pHdr->hwndFrom);

	if (pHdr->idFrom == IDC_GRID && pHdr->code == (UINT)NM_CUSTOMDRAW) 
	{
		int result = CDRF_DODEFAULT;

		NMLVCUSTOMDRAW* pCustomDraw = (LPNMLVCUSTOMDRAW)lParam;
		if (pCustomDraw->nmcd.dwDrawStage == CDDS_PREPAINT)
			result = CDRF_NOTIFYITEMDRAW;

		if (pCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
		{
			if (ListView_GetItemState(pHdr->hwndFrom, pCustomDraw->nmcd.dwItemSpec, LVIS_SELECTED)) 
			{
				pCustomDraw->nmcd.uItemState &= ~CDIS_SELECTED;
				result = CDRF_NOTIFYSUBITEMDRAW;
			}
			else {
				pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, pCustomDraw->nmcd.dwItemSpec % 2 == 0 ? TEXT("BACKCOLOR") : TEXT("BACKCOLOR2"));
			}
		}

		if (pCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM)) 
		{
			int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
			int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
			BOOL isCurrCell = (pCustomDraw->nmcd.dwItemSpec == (DWORD)rowNo) && (pCustomDraw->iSubItem == colNo);
			pCustomDraw->clrText = *(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR"));
			pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, isCurrCell ? TEXT("CURRENTCELLCOLOR") : TEXT("SELECTIONBACKCOLOR"));
		}

		nResult = result;
		goto Exit0;
	}
Exit0:
	return nResult;
}

int OnMsgUpdateCache(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int			nResult			= -1;
	TCHAR*		pszFilepath		= (TCHAR*)GetProp(hWnd, TEXT("FILEPATH"));
	int			nFilesize		= *(int*)GetProp(hWnd, TEXT("FILESIZE"));
	TCHAR		chDelimiter		= *(TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
	int			nCodepage		= *(int*)GetProp(hWnd, TEXT("CODEPAGE"));
	int			nSkipComments	= *(int*)GetProp(hWnd, TEXT("SKIPCOMMENTS"));
	int			nColCount		= 0;
	int			nRowNo			= 0;
	int			nLen			= 0;
	int			nCacheSize		= 0;
	TCHAR***	pppszCache		= nullptr;
	BOOL		bisTrim			= false;
	TCHAR*		pszData			= nullptr;

	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);

	char* pszRawdata = (char*)calloc(nFilesize + 2, sizeof(char)); // + 2!
	FILE* pfFile = _tfopen(pszFilepath, TEXT("rb"));
	fread(pszRawdata, sizeof(char), nFilesize, pfFile);
	fclose(pfFile);
	pszRawdata[nFilesize] = 0;
	pszRawdata[nFilesize + 1] = 0;

	int nLeadZeros = 0;
	for (int i = 0; nLeadZeros == i && i < nFilesize; i++)
		nLeadZeros += pszRawdata[i] == 0;

	if (nLeadZeros == nFilesize) {
		PostMessage(GetParent(hWnd), WM_KEYDOWN, 0x33, 0x20001); // Switch to Hex-mode

		// A special TC-command doesn't work under TC x32. 
		// https://flint-inc.ru/tcinfo/all_cmd.ru.htm#Misc
		// PostMessage(GetAncestor(hWnd, GA_ROOT), WM_USER + 51, 4005, 0);
		keybd_event(VK_TAB, VK_TAB, KEYEVENTF_EXTENDEDKEY, 0);

		nResult = FALSE;
		goto Exit0;
	}

	if (nCodepage == -1)
		nCodepage = DetectCodePage(pszRawdata, nFilesize);

	// Fix unexpected zeros
	if (nCodepage == CP_UTF16BE || nCodepage == CP_UTF16LE) {
		for (int i = 0; i < nFilesize / 2; i++) {
			if (pszRawdata[2 * i] == 0 && pszRawdata[2 * i + 1] == 0)
				pszRawdata[2 * i + nCodepage == CP_UTF16LE] = ' ';
		}
	}

	if (nCodepage == CP_UTF8 || nCodepage == CP_ACP) {
		for (int i = 0; i < nFilesize; i++) {
			if (pszRawdata[i] == 0)
				pszRawdata[i] = ' ';
		}
	}
	// end fix

	if (nCodepage == CP_UTF16BE) {
		for (int i = 0; i < nFilesize / 2; i++) {
			int c = pszRawdata[2 * i];
			pszRawdata[2 * i] = pszRawdata[2 * i + 1];
			pszRawdata[2 * i + 1] = c;
		}
	}

	if (nCodepage == CP_UTF16LE || nCodepage == CP_UTF16BE)
		pszData = (TCHAR*)(pszRawdata + nLeadZeros);

	if (nCodepage == CP_UTF8) {
		pszData = Utf8to16(pszRawdata + nLeadZeros);
		free(pszRawdata);
	}

	if (nCodepage == CP_ACP) {
		DWORD dwLen = MultiByteToWideChar(CP_ACP, 0, pszRawdata, -1, NULL, 0);
		pszData = (TCHAR*)calloc(dwLen, sizeof(TCHAR));
		if (!MultiByteToWideChar(CP_ACP, 0, pszRawdata + nLeadZeros, -1, pszData, dwLen))
			nCodepage = -1;
		free(pszRawdata);
	}

	if (nCodepage == -1) {
		MessageBox(hWnd, TEXT("Can't detect nCodepage"), NULL, MB_OK);
		nResult = 0;
		goto Exit0;
	}

	if (!chDelimiter) {
		TCHAR* pszStr = getStoredString(TEXT("default-column-pszDelimiter"), TEXT(""));
		chDelimiter = _tcslen(pszStr) ? pszStr[0] : detectDelimiter(pszData, nSkipComments);
		free(pszStr);
	}

	nColCount = 1;
	nRowNo = -1;
	nLen = (int)_tcslen(pszData);
	nCacheSize = nLen / 100 + 1;
	pppszCache = (TCHAR***)calloc(nCacheSize, sizeof(TCHAR**));
	bisTrim = getStoredValue(TEXT("trim-values"), 1);

	// Two step parsing: 0 - count columns, 1 - fill cache
	for (int stepNo = 0; stepNo < 2; stepNo++) {
		nRowNo = -1;
		int nStart = 0;
		for (int pos = 0; pos < nLen; pos++) {
			nRowNo++;
			if (stepNo == 1) {
				if (nRowNo >= nCacheSize) {
					nCacheSize += 100;
					pppszCache = (TCHAR***)realloc(pppszCache, nCacheSize * sizeof(TCHAR**));
				}

				pppszCache[nRowNo] = (TCHAR**)calloc(nColCount, sizeof(TCHAR*));
			}

			int nColNo = 0;
			nStart = pos;
			for (; pos < nLen && nColNo < MAX_COLUMN_COUNT; pos++) {
				TCHAR c = pszData[pos];

				while (nStart == pos && (stepNo == 0 && nSkipComments == 1 || nSkipComments == 2) && c == TEXT('#')) {
					while (pszData[pos] && !IsEOL(pszData[pos]))
						pos++;
					while (pos < nLen && IsEOL(pszData[pos]))
						pos++;
					c = pszData[pos];
					nStart = pos;
				}

				if (c == chDelimiter && !(nSkipComments && pszData[nStart] == TEXT('#')) || IsEOL(c) || pos >= nLen - 1) {
					int nvLen = pos - nStart + (pos >= nLen - 1);
					int nqPos = -1;
					for (int i = 0; i < nvLen; i++) {
						TCHAR c = pszData[nStart + i];
						if (c != TEXT(' ') && c != TEXT('\t')) {
							nqPos = c == TEXT('"') ? i : -1;
							break;
						}
					}

					if (nqPos != -1) {
						int nqCount = 0;
						for (int i = nqPos; i <= nvLen; i++)
							nqCount += pszData[nStart + i] == TEXT('"');

						if (pos < nLen && nqCount % 2)
							continue;

						while (nvLen > 0 && pszData[nStart + nvLen] != TEXT('"'))
							nvLen--;

						nvLen -= nqPos + 1;
					}

					if (stepNo == 1) {
						if (bisTrim) {
							int l = 0, r = 0;
							if (nqPos == -1) {
								for (int i = 0; i < nvLen && (pszData[nStart + i] == TEXT(' ') || pszData[nStart + i] == TEXT('\t')); i++)
									l++;
							}
							for (int i = nvLen - 1; i > 0 && (pszData[nStart + i] == TEXT(' ') || pszData[nStart + i] == TEXT('\t')); i--)
								r++;

							nStart += l;
							nvLen -= l + r;
						}

						TCHAR* value = (TCHAR*)calloc(nvLen + 1, sizeof(TCHAR));
						if (nvLen > 0) {
							_tcsncpy(value, pszData + nStart + nqPos + 1, nvLen);

							// replace "" by " in quoted value
							if (nqPos != -1) {
								int qCount = 0, j = 0;
								for (int i = 0; i < nvLen; i++) {
									BOOL isQ = value[i] == TEXT('"');
									qCount += isQ;
									value[j] = value[i];
									j += isQ && qCount % 2 || !isQ;
								}
								value[j] = 0;
							}
						}

						pppszCache[nRowNo][nColNo] = value;
					}

					nStart = pos + 1;
					nColNo++;
				}

				if (IsEOL(pszData[pos])) {
					while (IsEOL(pszData[pos + 1]))
						pos++;
					break;
				}
			}

			// Truncate 127+ columns
			if (!IsEOL(pszData[pos])) {
				while (pos < nLen && !IsEOL(pszData[pos + 1]))
					pos++;
				while (pos < nLen && IsEOL(pszData[pos + 1]))
					pos++;
			}

			while (stepNo == 1 && nColNo < nColCount) {
				pppszCache[nRowNo][nColNo] = (TCHAR*)calloc(1, sizeof(TCHAR));
				nColNo++;
			}

			if (stepNo == 0 && nColCount < nColNo)
				nColCount = nColNo;
		}

		if (stepNo == 1) {
			pppszCache = (TCHAR***)realloc(pppszCache, (nRowNo + 1) * sizeof(TCHAR**));
			if (nCodepage == CP_UTF16LE || nCodepage == CP_UTF16BE)
				free(pszRawdata);
			else
				free(pszData);
		}
	}

	if (nColCount > MAX_COLUMN_COUNT) 
	{
		HWND hListerWnd = GetParent(hWnd);
		TCHAR szMsg[255];
		_sntprintf(szMsg, 255, TEXT("Column count is overflow.\nFound: %i, max: %i"), nColCount, MAX_COLUMN_COUNT);
		MessageBox(hWnd, szMsg, NULL, 0);
		SendMessage(hListerWnd, WM_CLOSE, 0, 0);
		nResult = 0;
		goto Exit0;
	}

	{
		HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
		SendMessage(hStatusWnd, SB_SETTEXT, SB_CODEPAGE, (LPARAM)(
			nCodepage == CP_UTF8 ? TEXT("    UTF-8") :
			nCodepage == CP_UTF16LE ? TEXT(" UTF-16LE") :
			nCodepage == CP_UTF16BE ? TEXT(" UTF-16BE") : TEXT("     ANSI")
			));

		TCHAR szBuf[32] = { 0 };
		if (chDelimiter != TEXT('\t'))
			_sntprintf(szBuf, 32, TEXT(" %c"), chDelimiter);
		else
			_sntprintf(szBuf, 32, TEXT(" TAB"));
		SendMessage(hStatusWnd, SB_SETTEXT, SB_DELIMITER, (LPARAM)szBuf);

		_sntprintf(szBuf, 32, TEXT(" Comment mode: #%i     "), nSkipComments); // Long tail for a tip
		SendMessage(hStatusWnd, SB_SETTEXT, SB_COMMENTS, (LPARAM)szBuf);
	}

	SetProp(hWnd, TEXT("CACHE"), pppszCache);
	*(int*)GetProp(hWnd, TEXT("COLCOUNT")) = nColCount;
	*(int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")) = nRowNo + 1;
	*(int*)GetProp(hWnd, TEXT("CODEPAGE")) = nCodepage;
	*(TCHAR*)GetProp(hWnd, TEXT("DELIMITER")) = chDelimiter;
	*(int*)GetProp(hWnd, TEXT("ORDERBY")) = 0;

	SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);
	nResult = TRUE;
	goto Exit0;
	
Exit0:
	return nResult;
}

void OnMsgUpdagteGrid(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	HWND		hGridWnd		= GetDlgItem(hWnd, IDC_GRID);
	HWND		hStatusWnd		= GetDlgItem(hWnd, IDC_STATUSBAR);
	HWND		hHeader			= ListView_GetHeader(hGridWnd);
	TCHAR***	pppszCache		= (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
	int			nRowCount		= *(int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
	int			nisHeaderRow	= *(int*)GetProp(hWnd, TEXT("HEADERROW"));
	int			nFilterAlign	= *(int*)GetProp(hWnd, TEXT("FILTERALIGN"));

	int			nColCount		= Header_GetItemCount(hHeader);
	
	for (int colNo = 0; colNo < nColCount; colNo++)
		DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

	for (int colNo = 0; colNo < nColCount; colNo++)
		ListView_DeleteColumn(hGridWnd, nColCount - colNo - 1);

	nColCount = *(int*)GetProp(hWnd, TEXT("COLCOUNT"));
	for (int colNo = 0; colNo < nColCount; colNo++) {
		int nNum = 0;
		int nCount = nRowCount < 6 ? nRowCount : 6;
		for (int rowNo = 1; rowNo < nCount; rowNo++)
			nNum += pppszCache[rowNo][colNo] && _tcslen(pppszCache[rowNo][colNo]) && IsNumber(pppszCache[rowNo][colNo]);

		int fmt = nNum > nCount / 2 ? LVCFMT_RIGHT : LVCFMT_LEFT;
		TCHAR szColName[64];
		_sntprintf(szColName, 64, TEXT("Column #%i"), colNo + 1);
		ListView_AddColumn(hGridWnd, nisHeaderRow && pppszCache[0][colNo] && _tcslen(pppszCache[0][colNo]) > 0 ? pppszCache[0][colNo] : szColName, fmt);
	}

	int nAlign = nFilterAlign == -1 ? ES_LEFT : nFilterAlign == 1 ? ES_RIGHT : ES_CENTER;
	for (int nColNo = 0; nColNo < nColCount; nColNo++) {
		// Use WS_BORDER to vertical text aligment
		HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, nAlign | ES_AUTOHSCROLL | WS_CHILD | WS_TABSTOP | WS_BORDER,
			0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + nColNo), GetModuleHandle(0), NULL);
		SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
		SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)CallNewFilterEditProc));
	}

	SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
	SendMessage(hWnd, WMU_SET_HEADER_FILTERS, 0, 0);
	PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
}

int OnMsgUpdateResultset(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int		nResult		= -1;

	HWND	hGridWnd	= GetDlgItem(hWnd, IDC_GRID);
	HWND	hStatusWnd	= GetDlgItem(hWnd, IDC_STATUSBAR);
	HWND	hHeader		= ListView_GetHeader(hGridWnd);

	ListView_SetItemCount(hGridWnd, 0);
	TCHAR*** pppszCache			= (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
	int*	pnTotalRowCount		= (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
	int*	pnRowCount			= (int*)GetProp(hWnd, TEXT("ROWCOUNT"));
	int*	pnOrderBy			= (int*)GetProp(hWnd, TEXT("ORDERBY"));
	int		nisHeaderRow		= *(int*)GetProp(hWnd, TEXT("HEADERROW"));
	int*	pnResultset			= (int*)GetProp(hWnd, TEXT("RESULTSET"));
	BOOL	bisCaseSensitive	= getStoredValue(TEXT("szFilter-case-sensitive"), 0);
	BOOL*	pbResultset			= nullptr;
	int		nRowCount			= 0;
	int		nColCount			= 0;
	TCHAR	szBuf[255];

	if (pnResultset)
		free(pnResultset);

	if (!pppszCache)
	{
		nResult = 1;
		goto Exit0;
	}

	if (*pnTotalRowCount == 0)
	{
		nResult = 1;
		goto Exit0;
	}

	nColCount = Header_GetItemCount(hHeader);
	if (nColCount == 0)
	{
		nResult = 1;
		goto Exit0;
	}

	pbResultset = (BOOL*)calloc(*pnTotalRowCount, sizeof(BOOL));
	pbResultset[0] = !nisHeaderRow;
	for (int rowNo = 1; rowNo < *pnTotalRowCount; rowNo++)
		pbResultset[rowNo] = TRUE;

	for (int colNo = 0; colNo < nColCount; colNo++) 
	{
		HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
		TCHAR szFilter[MAX_FILTER_LENGTH];
		GetWindowText(hEdit, szFilter, MAX_FILTER_LENGTH);
		int nLen = (int)_tcslen(szFilter);
		if (nLen == 0)
			continue;

		for (int nRowNo = 0; nRowNo < *pnTotalRowCount; nRowNo++) 
		{
			if (!pbResultset[nRowNo])
				continue;

			TCHAR* pszValue = pppszCache[nRowNo][colNo];
			if (nLen > 1 && (szFilter[0] == TEXT('<') || szFilter[0] == TEXT('>')) && IsNumber(szFilter + 1)) {
				TCHAR* pszEnd = 0;
				double fF = _tcstod(szFilter + 1, &pszEnd);
				double fV = _tcstod(pszValue, &pszEnd);
				pbResultset[nRowNo] = (szFilter[0] == TEXT('<') && fV < fF) || (szFilter[0] == TEXT('>') && fV > fF);
			}
			else {
				pbResultset[nRowNo] = nLen == 1 ? hasString(pszValue, szFilter, bisCaseSensitive) :
					szFilter[0] == TEXT('=') && bisCaseSensitive ? _tcscmp(pszValue, szFilter + 1) == 0 :
					szFilter[0] == TEXT('=') && !bisCaseSensitive ? _tcsicmp(pszValue, szFilter + 1) == 0 :
					szFilter[0] == TEXT('!') ? !hasString(pszValue, szFilter + 1, bisCaseSensitive) :
					szFilter[0] == TEXT('<') ? _tcscmp(pszValue, szFilter + 1) < 0 :
					szFilter[0] == TEXT('>') ? _tcscmp(pszValue, szFilter + 1) > 0 :
					hasString(pszValue, szFilter, bisCaseSensitive);
			}
		}
	}

	pnResultset = (int*)calloc(*pnTotalRowCount, sizeof(int));
	for (int nRowNo = 0; nRowNo < *pnTotalRowCount; nRowNo++) 
	{
		if (!pbResultset[nRowNo])
			continue;

		pnResultset[nRowCount] = nRowNo;
		nRowCount++;
	}
	free(pbResultset);

	if (nRowCount > 0) 
	{
		if (nRowCount > *pnTotalRowCount)
			MessageBeep(0);
		pnResultset = (int*)realloc(pnResultset, nRowCount * sizeof(int));
		SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)pnResultset);
		int nOrderBy = *pnOrderBy;

		if (nOrderBy) 
		{
			int nColNo = nOrderBy > 0 ? nOrderBy - 1 : -nOrderBy - 1;
			BOOL bisBackward = nOrderBy < 0;

			BOOL bisNum = TRUE;
			for (int i = nisHeaderRow; i < *pnTotalRowCount && i <= 5; i++)
				bisNum = bisNum && IsNumber(pppszCache[i][nColNo]);

			if (bisNum) 
			{
				double* nums = (double*)calloc(*pnTotalRowCount, sizeof(double));
				for (int i = 0; i < nRowCount; i++)
					nums[pnResultset[i]] = _tcstod(pppszCache[pnResultset[i]][nColNo], NULL);

				MergeSort(pnResultset, (void*)nums, 0, nRowCount - 1, bisBackward, bisNum);
				free(nums);
			}
			else 
			{
				TCHAR** strings = (TCHAR**)calloc(*pnTotalRowCount, sizeof(TCHAR*));
				for (int i = 0; i < nRowCount; i++)
					strings[pnResultset[i]] = pppszCache[pnResultset[i]][nColNo];
				MergeSort(pnResultset, (void*)strings, 0, nRowCount - 1, bisBackward, bisNum);
				free(strings);
			}
		}
	}
	else {
		SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)0);
		free(pnResultset);
	}

	*pnRowCount = nRowCount;
	ListView_SetItemCount(hGridWnd, nRowCount);
	InvalidateRect(hGridWnd, NULL, TRUE);

	_sntprintf(szBuf, 255, TEXT(" Rows: %i/%i"), nRowCount, *pnTotalRowCount - nisHeaderRow);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)szBuf);

	PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);

Exit0:
	return nResult;
}

void OnMsgUpdateFilterSize(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
	HWND hHeader = ListView_GetHeader(hGridWnd);
	int nColCount = Header_GetItemCount(hHeader);
	SendMessage(hHeader, WM_SIZE, 0, 0);

	int* pnColOrder = (int*)calloc(nColCount, sizeof(int));
	Header_GetOrderArray(hHeader, nColCount, pnColOrder);

	for (int idx = 0; idx < nColCount; idx++)
	{
		int nColNo = pnColOrder[idx];
		RECT rc;
		Header_GetItemRect(hHeader, nColNo, &rc);
		int nH2 = (int)round((rc.bottom - rc.top) / 2);
		SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + nColNo), 0, rc.left, nH2, rc.right - rc.left, nH2 + 1, SWP_NOZORDER);
	}

	free(pnColOrder);
}

void OnMsAutoColumnSize(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
	SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
	HWND hHeader = ListView_GetHeader(hGridWnd);
	int nColCount = Header_GetItemCount(hHeader);

	for (int colNo = 0; colNo < nColCount - 1; colNo++)
		ListView_SetColumnWidth(hGridWnd, colNo, colNo < nColCount - 1 ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE);

	if (nColCount == 1 && ListView_GetColumnWidth(hGridWnd, 0) < 100)
		ListView_SetColumnWidth(hGridWnd, 0, 100);

	int nMaxWidth = getStoredValue(TEXT("max-column-width"), 300);
	if (nColCount > 1) {
		for (int colNo = 0; colNo < nColCount; colNo++) {
			if (ListView_GetColumnWidth(hGridWnd, colNo) > nMaxWidth)
				ListView_SetColumnWidth(hGridWnd, colNo, nMaxWidth);
		}
	}

	// Fix last column				
	if (nColCount > 1) {
		int nColNo = nColCount - 1;
		ListView_SetColumnWidth(hGridWnd, nColNo, LVSCW_AUTOSIZE);
		TCHAR szName16[MAX_COLUMN_LENGTH + 1];
		Header_GetItemText(hHeader, nColNo, szName16, MAX_COLUMN_LENGTH);

		SIZE s = { 0 };
		HDC hDC = GetDC(hHeader);
		HFONT hOldFont = (HFONT)SelectObject(hDC, (HFONT)GetProp(hWnd, TEXT("FONT")));
		GetTextExtentPoint32(hDC, szName16, (int)_tcslen(szName16), &s);
		SelectObject(hDC, hOldFont);
		ReleaseDC(hHeader, hDC);

		int nW = s.cx + 12;
		if (ListView_GetColumnWidth(hGridWnd, nColNo) < nW)
			ListView_SetColumnWidth(hGridWnd, nColNo, nW);

		if (ListView_GetColumnWidth(hGridWnd, nColNo) > nMaxWidth)
			ListView_SetColumnWidth(hGridWnd, nColNo, nMaxWidth);
	}

	SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
	InvalidateRect(hGridWnd, NULL, TRUE);

	PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
}

void OnMsgSetHeaderFilters(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
	HWND hHeader = ListView_GetHeader(hGridWnd);
	int nisFilterRow = *(int*)GetProp(hWnd, TEXT("FILTERROW"));
	int nColCount = Header_GetItemCount(hHeader);

	SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
	LONG_PTR pStyles = GetWindowLongPtr(hHeader, GWL_STYLE);
	pStyles = nisFilterRow ? pStyles | HDS_FILTERBAR : pStyles & (~HDS_FILTERBAR);
	SetWindowLongPtr(hHeader, GWL_STYLE, pStyles);

	for (int nColNo = 0; nColNo < nColCount; nColNo++)
		ShowWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + nColNo), nisFilterRow ? SW_SHOW : SW_HIDE);

	// Bug fix: force Windows to redraw header
	SetWindowPos(hGridWnd, 0, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOMOVE);
	SendMessage(GetMainWindow(hWnd), WM_SIZE, 0, 0);

	if (nisFilterRow)
		SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);

	SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
	InvalidateRect(hWnd, NULL, TRUE);
}

int OnMsgSetCurrentCell(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int		nResult		= -1;
	HWND	hGridWnd	= GetDlgItem(hWnd, IDC_GRID);
	HWND	hHeader		= ListView_GetHeader(hGridWnd);
	HWND	hStatusWnd	= GetDlgItem(hWnd, IDC_STATUSBAR);
	int		nW			= 0;
	int		nDx			 = 0;
	TCHAR	szBuf[32] = { 0 };

	SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)0);

	int* pnRowNo = (int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
	int* pnColNo = (int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
	if (*pnRowNo == wParam && *pnColNo == lParam)
	{
		nResult = 0;
		goto Exit0;
	}

	RECT rc, rc2;
	ListView_GetSubItemRect(hGridWnd, *pnRowNo, *pnColNo, LVIR_BOUNDS, &rc);
	if (*pnColNo == 0)
	rc.right = ListView_GetColumnWidth(hGridWnd, *pnColNo);
	InvalidateRect(hGridWnd, &rc, TRUE);

	*pnRowNo = (int)wParam;
	*pnColNo = (int)lParam;
	ListView_GetSubItemRect(hGridWnd, *pnRowNo, *pnColNo, LVIR_BOUNDS, &rc);
	if (*pnColNo == 0)
	rc.right = ListView_GetColumnWidth(hGridWnd, *pnColNo);
	InvalidateRect(hGridWnd, &rc, FALSE);

	GetClientRect(hGridWnd, &rc2);
	nW = rc.right - rc.left;
	nDx = rc2.right < rc.right ? rc.left - rc2.right + nW : rc.left < 0 ? rc.left : 0;

	ListView_Scroll(hGridWnd, nDx, 0);

	if (*pnColNo != -1 && *pnRowNo != -1)
	_sntprintf(szBuf, 32, TEXT(" %i:%i"), *pnRowNo + 1, *pnColNo + 1);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_CELL, (LPARAM)szBuf);
Exit0:
	return nResult;
}

void OnMsgResetCache(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	HWND		hGridWnd	= GetDlgItem(hWnd, IDC_GRID);
	TCHAR***	pppszCache	= (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
	int*		pnRowCount	= (int*)GetProp(hWnd, TEXT("ROWCOUNT"));

	int nColCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
	if (nColCount > 0 && pppszCache != 0) {
		for (int rowNo = 0; rowNo < *pnRowCount; rowNo++) {
			if (pppszCache[rowNo]) {
				for (int colNo = 0; colNo < nColCount; colNo++)
					if (pppszCache[rowNo][colNo])
						free(pppszCache[rowNo][colNo]);

				free(pppszCache[rowNo]);
			}
			pppszCache[rowNo] = 0;
		}
		free(pppszCache);
	}

	SetProp(hWnd, TEXT("CACHE"), 0);
	*pnRowCount = 0;
}

int OnMsgSetFont(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) 
{
	int		nResult		= -1;
	int*	pFontSize	= (int*)GetProp(hWnd, TEXT("FONTSIZE"));
	HFONT	hFont		= nullptr;
	HWND	hGridWnd	= nullptr;
	HWND	hHeader		= nullptr;

	if (*pFontSize + wParam < 10 || *pFontSize + wParam > 48)
	{
		nResult = 0;
		goto Exit0;
	}
	*pFontSize += (int)wParam;
	DeleteFont(GetProp(hWnd, TEXT("FONT")));

	hFont = CreateFont(*pFontSize, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, (TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
	hGridWnd = GetDlgItem(hWnd, IDC_GRID);
	SendMessage(hGridWnd, WM_SETFONT, (LPARAM)hFont, TRUE);

	hHeader = ListView_GetHeader(hGridWnd);
	for (int nColNo = 0; nColNo < Header_GetItemCount(hHeader); nColNo++)
		SendMessage(GetDlgItem(hHeader, IDC_HEADER_EDIT + nColNo), WM_SETFONT, (LPARAM)hFont, TRUE);

	SetProp(hWnd, TEXT("FONT"), hFont);
	PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);

Exit0:
	return nResult;
}

void OnMsgSetThemt(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) 
{
	HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
	BOOL bisDark = *(int*)GetProp(hWnd, TEXT("DARKTHEME"));

	int nTextColor = !bisDark ? getStoredValue(TEXT("text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("text-color-dark"), RGB(220, 220, 220));
	int nBackColor = !bisDark ? getStoredValue(TEXT("back-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("back-color-dark"), RGB(32, 32, 32));
	int nBackColor2 = !bisDark ? getStoredValue(TEXT("back-color2"), RGB(240, 240, 240)) : getStoredValue(TEXT("back-color2-dark"), RGB(52, 52, 52));
	int nFilterTextColor = !bisDark ? getStoredValue(TEXT("szFilter-text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("szFilter-text-color-dark"), RGB(255, 255, 255));
	int nFilterBackColor = !bisDark ? getStoredValue(TEXT("szFilter-back-color"), RGB(240, 240, 240)) : getStoredValue(TEXT("szFilter-back-color-dark"), RGB(60, 60, 60));
	int nCurrCellColor = !bisDark ? getStoredValue(TEXT("current-cell-back-color"), RGB(70, 96, 166)) : getStoredValue(TEXT("current-cell-back-color-dark"), RGB(32, 62, 62));
	int nSelectionTextColor = !bisDark ? getStoredValue(TEXT("selection-text-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("selection-text-color-dark"), RGB(220, 220, 220));
	int nSelectionBackColor = !bisDark ? getStoredValue(TEXT("selection-back-color"), RGB(10, 36, 106)) : getStoredValue(TEXT("selection-back-color-dark"), RGB(72, 102, 102));

	*(int*)GetProp(hWnd, TEXT("TEXTCOLOR")) = nTextColor;
	*(int*)GetProp(hWnd, TEXT("BACKCOLOR")) = nBackColor;
	*(int*)GetProp(hWnd, TEXT("BACKCOLOR2")) = nBackColor2;
	*(int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")) = nFilterTextColor;
	*(int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")) = nFilterBackColor;
	*(int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")) = nCurrCellColor;
	*(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")) = nSelectionTextColor;
	*(int*)GetProp(hWnd, TEXT("SELECTIONBACKCOLOR")) = nSelectionBackColor;

	DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
	SetProp(hWnd, TEXT("FILTERBACKBRUSH"), CreateSolidBrush(nFilterBackColor));

	ListView_SetTextColor(hGridWnd, nTextColor);
	ListView_SetBkColor(hGridWnd, nBackColor);
	ListView_SetTextBkColor(hGridWnd, nBackColor);
	InvalidateRect(hWnd, NULL, TRUE);
}

int OnMsgHotKeys(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) 
{
	int nResult = -1;
	BOOL bisCtrl = HIWORD(GetKeyState(VK_CONTROL));
	if (wParam == VK_TAB) {
		HWND hFocus = GetFocus();
		HWND wnds[1000] = { 0 };
		EnumChildWindows(hWnd, (WNDENUMPROC)CallEnumTabStopChildrenProc, (LPARAM)wnds);

		int nNo = 0;
		while (wnds[nNo] && wnds[nNo] != hFocus)
			nNo++;

		int nCnt = nNo;
		while (wnds[nCnt])
			nCnt++;

		nNo += bisCtrl ? -1 : 1;
		SetFocus(wnds[nNo] && nNo >= 0 ? wnds[nNo] : (bisCtrl ? wnds[nCnt - 1] : wnds[0]));
	}

	if (wParam == VK_F1) {
		ShellExecute(0, 0, TEXT("https://github.com/little-brother/csvtab-wlx/wiki"), 0, 0, SW_SHOW);
		nResult = TRUE;
		goto Exit0;
	}

	if (wParam == 0x20 && bisCtrl) { // Ctrl + Space
		SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);
		nResult = TRUE;
		goto Exit0;
	}

	if (wParam == VK_ESCAPE || wParam == VK_F11 ||
		wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7 || (bisCtrl && wParam == 0x46) || // Ctrl + F
		((wParam >= 0x31 && wParam <= 0x38) && !getStoredValue(TEXT("disable-num-keys"), 0) && !bisCtrl || // 1 - 8
			(wParam == 0x4E || wParam == 0x50) && !getStoredValue(TEXT("disable-np-keys"), 0) || // N, P
			wParam == 0x51 && getStoredValue(TEXT("exit-by-q"), 0)) && // Q
		GetDlgCtrlID(GetFocus()) / 100 * 100 != IDC_HEADER_EDIT) {
		SetFocus(GetParent(hWnd));
		keybd_event((BYTE)wParam, (BYTE)wParam, KEYEVENTF_EXTENDEDKEY, 0);

		nResult = TRUE;
		goto Exit0;
	}

	if (bisCtrl && wParam >= 0x30 && wParam <= 0x39 && !getStoredValue(TEXT("disable-num-keys"), 0)) {// 0-9
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		HWND hHeader = ListView_GetHeader(hGridWnd);
		int nColCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));

		BOOL bisCurrent = wParam == 0x30;
		int nColNo = (int)(bisCurrent ? *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) : wParam - 0x30 - 1);
		if (nColNo < 0 || nColNo > nColCount - 1 || bisCurrent && ListView_GetColumnWidth(hGridWnd, nColNo) == 0)
		{
			nResult = FALSE;
			goto Exit0;
		}

		if (!bisCurrent) {
			int* colOrder = (int*)calloc(nColCount, sizeof(int));
			Header_GetOrderArray(hHeader, nColCount, colOrder);

			int hiddenColCount = 0;
			for (int idx = 0; (idx < nColCount) && (idx - hiddenColCount <= nColNo); idx++)
				hiddenColCount += ListView_GetColumnWidth(hGridWnd, colOrder[idx]) == 0;

			nColNo = colOrder[nColNo + hiddenColCount];
			free(colOrder);
		}

		SendMessage(hWnd, WMU_SORT_COLUMN, nColNo, 0);
		nResult = TRUE;
		goto Exit0;
	}

	nResult = FALSE;
Exit0:
	return nResult;
}

int OnMsgHotChars(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	BOOL bisCtrl = HIWORD(GetKeyState(VK_CONTROL));

	unsigned char chScanCode = ((unsigned char*)&lParam)[2];
	UINT uKey = MapVirtualKey(chScanCode, MAPVK_VSC_TO_VK);

	return !_istprint((int)wParam) && (
		wParam == VK_ESCAPE || wParam == VK_F11 || wParam == VK_F1 ||
		wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7) ||
		wParam == VK_TAB || wParam == VK_RETURN ||
		bisCtrl && uKey == 0x46 || // Ctrl + F
		getStoredValue(TEXT("exit-by-q"), 0) && uKey == 0x51 && GetDlgCtrlID(GetFocus()) / 100 * 100 != IDC_HEADER_EDIT; // Q
}

LRESULT CALLBACK MainWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) 
{
	LRESULT nResult = -1;
	LRESULT nRetCode = 0;

	switch (uMsg)
	{
	case WM_SIZE:
		{
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, WM_SIZE, 0, 0);
			RECT rc;
			GetClientRect(hStatusWnd, &rc);
			int nStatusH = rc.bottom;

			GetClientRect(hWnd, &rc);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			SetWindowPos(hGridWnd, 0, 0, 0, rc.right, rc.bottom - nStatusH, SWP_NOZORDER);
		}
		break;

	// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
	case WM_NCHITTEST:
		nResult = 1;
		goto Exit0;

	case WM_SETCURSOR:
		{
			SetCursor(LoadCursor(0, IDC_ARROW));
			nResult = TRUE;
			goto Exit0;
		}
		break;

	case WM_SETFOCUS:
		{
			SetFocus((HWND)GetProp(hWnd, TEXT("LASTFOCUS")));
		}
		break;

	case WM_MOUSEWHEEL:
		{
			if (LOWORD(wParam) == MK_CONTROL) {
				SendMessage(hWnd, WMU_SET_FONT, GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? 1 : -1, 0);
				nResult = 1;
				goto Exit0;
			}
		}
		break;

	case WM_KEYDOWN:
		{
			if (SendMessage(hWnd, WMU_HOT_KEYS, wParam, lParam))
			{
				nResult = 0;
				goto Exit0;
			}
		}
		break;

	case WM_COMMAND:
		nRetCode = OnMsgCommand(hWnd, uMsg, wParam, lParam);
		KG_PROCESS_ERROR_RET_CODE(nRetCode != -1, nRetCode);
		break;

	case WM_NOTIFY:
		nRetCode = OnMsgNotify(hWnd, uMsg, wParam, lParam);
		KG_PROCESS_ERROR_RET_CODE(nRetCode != -1, nRetCode);
		break;

		// wParam = colNo
	case WMU_HIDE_COLUMN:
		{
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int nColNo = (int)wParam;

			HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + nColNo);
			SetWindowLongPtr(hEdit, GWLP_USERDATA, (LONG_PTR)ListView_GetColumnWidth(hGridWnd, nColNo));
			ListView_SetColumnWidth(hGridWnd, nColNo, 0);
			InvalidateRect(hHeader, NULL, TRUE);
		}
		break;

	case WMU_SHOW_COLUMNS:
		{
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int nColCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			for (int nColNo = 0; nColNo < nColCount; nColNo++) {
				if (ListView_GetColumnWidth(hGridWnd, nColNo) == 0) {
					HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + nColNo);
					ListView_SetColumnWidth(hGridWnd, nColNo, (int)GetWindowLongPtr(hEdit, GWLP_USERDATA));
				}
			}

			InvalidateRect(hGridWnd, NULL, TRUE);
		}
		break;

	// wParam = colNo
	case WMU_SORT_COLUMN:
		{
			int nColNo = (int)wParam + 1;
			KG_PROCESS_ERROR_RET_CODE(nColNo > 0, FALSE);

			int* pnOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
			int nOrderBy = *pnOrderBy;
			*pnOrderBy = nColNo == nOrderBy || nColNo == -nOrderBy ? -nOrderBy : nColNo;
			SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
		}
		break;

	case WMU_UPDATE_CACHE:
		nRetCode = OnMsgUpdateCache(hWnd, uMsg, wParam, lParam);
		KG_PROCESS_ERROR_RET_CODE(nRetCode != -1, nRetCode);
		break;

	case WMU_UPDATE_GRID:
		OnMsgUpdagteGrid(hWnd, uMsg, wParam, lParam);
		break;

	case WMU_UPDATE_RESULTSET:
		OnMsgUpdateResultset(hWnd, uMsg, wParam, lParam);
		break;

	case WMU_UPDATE_FILTER_SIZE:
		OnMsgUpdateFilterSize(hWnd, uMsg, wParam, lParam);
		break;

	case WMU_SET_HEADER_FILTERS:
		OnMsgSetHeaderFilters(hWnd, uMsg, wParam, lParam);
		break;

	case WMU_AUTO_COLUMN_SIZE:
		OnMsAutoColumnSize(hWnd, uMsg, wParam, lParam);
		break;

		// wParam = rowNo, lParam = colNo
	case WMU_SET_CURRENT_CELL:
		nRetCode = OnMsgSetCurrentCell(hWnd, uMsg, wParam, lParam);
		KG_PROCESS_ERROR_RET_CODE(nRetCode != -1, nRetCode);
		break;

	case WMU_RESET_CACHE:
		OnMsgResetCache(hWnd, uMsg, wParam, lParam);
		break;

		// wParam - size delta
	case WMU_SET_FONT:
		nRetCode = OnMsgSetFont(hWnd, uMsg, wParam, lParam);
		KG_PROCESS_ERROR_RET_CODE(nRetCode != -1, nRetCode);
		break;

	case WMU_SET_THEME:
		OnMsgSetThemt(hWnd, uMsg, wParam, lParam);
		break;

	case WMU_HOT_KEYS:
		nRetCode = OnMsgHotKeys(hWnd, uMsg, wParam, lParam);
		KG_PROCESS_ERROR_RET_CODE(nRetCode != -1, nRetCode);
		break;

	case WMU_HOT_CHARS:
		nRetCode = OnMsgHotChars(hWnd, uMsg, wParam, lParam);
		KG_PROCESS_ERROR_RET_CODE(nRetCode != -1, nRetCode);
		break;
	}

	nResult = CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, uMsg, wParam, lParam);
Exit0:
	return nResult;
}
