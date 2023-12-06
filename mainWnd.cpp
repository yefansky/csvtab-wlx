#include "stdafx.h"

HWND getMainWindow(HWND hWnd) {
	HWND hMainWnd = hWnd;
	while (hMainWnd && GetDlgCtrlID(hMainWnd) != IDC_MAIN)
		hMainWnd = GetParent(hMainWnd);
	return hMainWnd;
}

LRESULT CALLBACK mainWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
	case WM_SIZE: {
		HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
		SendMessage(hStatusWnd, WM_SIZE, 0, 0);
		RECT rc;
		GetClientRect(hStatusWnd, &rc);
		int statusH = rc.bottom;

		GetClientRect(hWnd, &rc);
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		SetWindowPos(hGridWnd, 0, 0, 0, rc.right, rc.bottom - statusH, SWP_NOZORDER);
	}
	break;

				// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
	case WM_NCHITTEST: {
		return 1;
	}
	break;

	case WM_SETCURSOR: {
		SetCursor(LoadCursor(0, IDC_ARROW));
		return TRUE;
	}
	break;

	case WM_SETFOCUS: {
		SetFocus((HWND)GetProp(hWnd, TEXT("LASTFOCUS")));
	}
	break;

	case WM_MOUSEWHEEL: {
		if (LOWORD(wParam) == MK_CONTROL) {
			SendMessage(hWnd, WMU_SET_FONT, GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? 1 : -1, 0);
			return 1;
		}
	}
	break;

	case WM_KEYDOWN: {
		if (SendMessage(hWnd, WMU_HOT_KEYS, wParam, lParam))
			return 0;
	}
	break;

	case WM_COMMAND: {
		WORD cmd = LOWORD(wParam);
		if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROWS || cmd == IDM_COPY_COLUMN) {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
			int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));

			int colCount = Header_GetItemCount(hHeader);
			int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
			int selCount = ListView_GetSelectedCount(hGridWnd);

			if (rowNo == -1 ||
				rowNo >= rowCount ||
				colCount == 0 ||
				cmd == IDM_COPY_CELL && colNo == -1 ||
				cmd == IDM_COPY_CELL && colNo >= colCount ||
				cmd == IDM_COPY_COLUMN && colNo == -1 ||
				cmd == IDM_COPY_COLUMN && colNo >= colCount ||
				cmd == IDM_COPY_ROWS && selCount == 0) {
				setClipboardText(TEXT(""));
				return 0;
			}

			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
			if (!resultset)
				return 0;

			TCHAR* delimiter = getStoredString(TEXT("column-delimiter"), TEXT("\t"));

			int len = 0;
			if (cmd == IDM_COPY_CELL)
				len = _tcslen(cache[resultset[rowNo]][colNo]);

			if (cmd == IDM_COPY_ROWS) {
				int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
				while (rowNo != -1) {
					for (int colNo = 0; colNo < colCount; colNo++) {
						if (ListView_GetColumnWidth(hGridWnd, colNo))
							len += _tcslen(cache[resultset[rowNo]][colNo]) + 1; /* column delimiter */
					}

					len++; /* \n */
					rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
				}
			}

			if (cmd == IDM_COPY_COLUMN) {
				int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
				while (rowNo != -1 && rowNo < rowCount) {
					len += _tcslen(cache[resultset[rowNo]][colNo]) + 1 /* \n */;
					rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
				}
			}

			TCHAR* buf = (TCHAR*)calloc(len + 1, sizeof(TCHAR));
			if (cmd == IDM_COPY_CELL)
				_tcscat(buf, cache[resultset[rowNo]][colNo]);

			if (cmd == IDM_COPY_ROWS) {
				int pos = 0;
				int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);

				int* colOrder = (int*)calloc(colCount, sizeof(int));
				Header_GetOrderArray(hHeader, colCount, colOrder);

				while (rowNo != -1) {
					for (int idx = 0; idx < colCount; idx++) {
						int colNo = colOrder[idx];
						if (ListView_GetColumnWidth(hGridWnd, colNo)) {
							int len = _tcslen(cache[resultset[rowNo]][colNo]);
							_tcsncpy(buf + pos, cache[resultset[rowNo]][colNo], len);
							buf[pos + len] = delimiter[0];
							pos += len + 1;
						}
					}

					buf[pos - (pos > 0)] = TEXT('\n');
					rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
				}
				buf[pos - 1] = 0; // remove last \n

				free(colOrder);
			}

			if (cmd == IDM_COPY_COLUMN) {
				int pos = 0;
				int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
				while (rowNo != -1 && rowNo < rowCount) {
					int len = _tcslen(cache[resultset[rowNo]][colNo]);
					_tcsncpy(buf + pos, cache[resultset[rowNo]][colNo], len);
					rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
					if (rowNo != -1 && rowNo < rowCount)
						buf[pos + len] = TEXT('\n');
					pos += len + 1;
				}
			}

			setClipboardText(buf);
			free(buf);
			free(delimiter);
		}

		if (cmd == IDM_HIDE_COLUMN) {
			int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
			SendMessage(hWnd, WMU_HIDE_COLUMN, colNo, 0);
		}

		if (cmd == IDM_FILTER_ROW || cmd == IDM_HEADER_ROW || cmd == IDM_DARK_THEME) {
			HMENU hMenu = (HMENU)GetProp(hWnd, TEXT("GRIDMENU"));
			int* pOpt = (int*)GetProp(hWnd, cmd == IDM_FILTER_ROW ? TEXT("FILTERROW") : cmd == IDM_HEADER_ROW ? TEXT("HEADERROW" : TEXT("DARKTHEME")));
			*pOpt = (*pOpt + 1) % 2;
			Menu_SetItemState(hMenu, cmd, *pOpt ? MFS_CHECKED : 0);

			UINT msg = cmd == IDM_FILTER_ROW ? WMU_SET_HEADER_FILTERS : cmd == IDM_HEADER_ROW ? WMU_UPDATE_GRID : WMU_SET_THEME;
			SendMessage(hWnd, msg, 0, 0);
		}

		if (cmd == IDM_ANSI || cmd == IDM_UTF8 || cmd == IDM_UTF16LE || cmd == IDM_UTF16BE) {
			int* pCodepage = (int*)GetProp(hWnd, TEXT("CODEPAGE"));
			*pCodepage = cmd == IDM_ANSI ? CP_ACP : cmd == IDM_UTF8 ? CP_UTF8 : cmd == IDM_UTF16LE ? CP_UTF16LE : CP_UTF16BE;
			SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
		}

		if (cmd == IDM_COMMA || cmd == IDM_SEMICOLON || cmd == IDM_VBAR || cmd == IDM_TAB || cmd == IDM_COLON) {
			TCHAR* pDelimiter = (TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
			*pDelimiter = cmd == IDM_COMMA ? TEXT(',') :
				cmd == IDM_SEMICOLON ? TEXT(';') :
				cmd == IDM_TAB ? TEXT('\t') :
				cmd == IDM_VBAR ? TEXT('|') :
				cmd == IDM_COLON ? TEXT(':') :
				0;
			SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
		}

		if (cmd == IDM_DEFAULT || cmd == IDM_NO_PARSE || cmd == IDM_NO_SHOW) {
			int* pSkipComments = (int*)GetProp(hWnd, TEXT("SKIPCOMMENTS"));
			*pSkipComments = cmd == IDM_DEFAULT ? 0 : cmd == IDM_NO_PARSE ? 1 : 2;
			SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
		}
	}
	break;

	case WM_NOTIFY: {
		NMHDR* pHdr = (LPNMHDR)lParam;
		if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_GETDISPINFO) {
			LV_DISPINFO* pDispInfo = (LV_DISPINFO*)lParam;
			LV_ITEM* pItem = &(pDispInfo)->item;

			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
			if (resultset && pItem->mask & LVIF_TEXT) {
				int rowNo = resultset[pItem->iItem];
				pItem->pszText = cache[rowNo][pItem->iSubItem];
			}
		}

		if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_COLUMNCLICK) {
			NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
			return SendMessage(hWnd, HIWORD(GetKeyState(VK_CONTROL)) ? WMU_HIDE_COLUMN : WMU_SORT_COLUMN, lv->iSubItem, 0);
		}

		if (pHdr->idFrom == IDC_GRID && (pHdr->code == (DWORD)NM_CLICK || pHdr->code == (DWORD)NM_RCLICK)) {
			NMITEMACTIVATE* ia = (LPNMITEMACTIVATE)lParam;
			SendMessage(hWnd, WMU_SET_CURRENT_CELL, ia->iItem, ia->iSubItem);
		}

		if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_CLICK && HIWORD(GetKeyState(VK_MENU))) {
			NMITEMACTIVATE* ia = (LPNMITEMACTIVATE)lParam;
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));

			TCHAR* url = extractUrl(cache[resultset[ia->iItem]][ia->iSubItem]);
			ShellExecute(0, TEXT("open"), url, 0, 0, SW_SHOW);
			free(url);
		}

		if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_RCLICK) {
			POINT p;
			GetCursorPos(&p);
			TrackPopupMenu((HMENU)GetProp(hWnd, TEXT("GRIDMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
		}

		if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
			NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
			if (lv->uOldState != lv->uNewState && (lv->uNewState & LVIS_SELECTED))
				SendMessage(hWnd, WMU_SET_CURRENT_CELL, lv->iItem, *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));
		}

		if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
			NMLVKEYDOWN* kd = (LPNMLVKEYDOWN)lParam;
			if (kd->wVKey == 0x43) { // C
				BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
				BOOL isShift = HIWORD(GetKeyState(VK_SHIFT));
				BOOL isCopyColumn = getStoredValue(TEXT("copy-column"), 0) && ListView_GetSelectedCount(pHdr->hwndFrom) > 1;
				if (!isCtrl && !isShift)
					return FALSE;

				int action = !isShift && !isCopyColumn ? IDM_COPY_CELL : isCtrl || isCopyColumn ? IDM_COPY_COLUMN : IDM_COPY_ROWS;
				SendMessage(hWnd, WM_COMMAND, action, 0);
				return TRUE;
			}

			if (kd->wVKey == 0x41 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + A
				HWND hGridWnd = pHdr->hwndFrom;
				SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
				int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				ListView_SetItemState(hGridWnd, -1, LVIS_SELECTED, LVIS_SELECTED | LVIS_FOCUSED);
				SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
				SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
				InvalidateRect(hGridWnd, NULL, TRUE);
			}

			if (kd->wVKey == 0x20 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + Space				
				SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);
				return TRUE;
			}

			if (kd->wVKey == VK_LEFT || kd->wVKey == VK_RIGHT) {
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);

				int colCount = Header_GetItemCount(ListView_GetHeader(pHdr->hwndFrom));
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));

				int* colOrder = (int*)calloc(colCount, sizeof(int));
				Header_GetOrderArray(hHeader, colCount, colOrder);

				int dir = kd->wVKey == VK_RIGHT ? 1 : -1;
				int idx = 0;
				for (idx; colOrder[idx] != colNo; idx++);
				do {
					idx = (colCount + idx + dir) % colCount;
				} while (ListView_GetColumnWidth(hGridWnd, colOrder[idx]) == 0);

				colNo = colOrder[idx];
				free(colOrder);

				SendMessage(hWnd, WMU_SET_CURRENT_CELL, *(int*)GetProp(hWnd, TEXT("CURRENTROWNO")), colNo);
				return TRUE;
			}
		}

		if ((pHdr->code == HDN_ITEMCHANGED || pHdr->code == HDN_ENDDRAG) && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(hWnd, IDC_GRID)))
			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);

		if (pHdr->idFrom == IDC_STATUSBAR && (pHdr->code == NM_CLICK || pHdr->code == NM_RCLICK)) {
			NMMOUSE* pm = (NMMOUSE*)lParam;
			int id = pm->dwItemSpec;
			if (id != SB_CODEPAGE && id != SB_DELIMITER && id != SB_COMMENTS)
				return 0;

			RECT rc, rc2;
			GetWindowRect(pHdr->hwndFrom, &rc);
			SendMessage(pHdr->hwndFrom, SB_GETRECT, id, (LPARAM)&rc2);
			HMENU hMenu = (HMENU)GetProp(hWnd, id == SB_CODEPAGE ? TEXT("CODEPAGEMENU") : id == SB_DELIMITER ? TEXT("DELIMITERMENU") : TEXT("COMMENTMENU"));

			if (id == SB_CODEPAGE) {
				int codepage = *(int*)GetProp(hWnd, TEXT("CODEPAGE"));
				Menu_SetItemState(hMenu, IDM_ANSI, codepage == CP_ACP ? MFS_CHECKED : 0);
				Menu_SetItemState(hMenu, IDM_UTF8, codepage == CP_UTF8 ? MFS_CHECKED : 0);
				Menu_SetItemState(hMenu, IDM_UTF16LE, codepage == CP_UTF16LE ? MFS_CHECKED : 0);
				Menu_SetItemState(hMenu, IDM_UTF16BE, codepage == CP_UTF16BE ? MFS_CHECKED : 0);
			}

			if (id == SB_DELIMITER) {
				TCHAR delimiter = *(TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
				Menu_SetItemState(hMenu, IDM_COMMA, delimiter == TEXT(',') ? MFS_CHECKED : 0);
				Menu_SetItemState(hMenu, IDM_SEMICOLON, delimiter == TEXT(';') ? MFS_CHECKED : 0);
				Menu_SetItemState(hMenu, IDM_VBAR, delimiter == TEXT('|') ? MFS_CHECKED : 0);
				Menu_SetItemState(hMenu, IDM_TAB, delimiter == TEXT('\t') ? MFS_CHECKED : 0);
				Menu_SetItemState(hMenu, IDM_COLON, delimiter == TEXT(':') ? MFS_CHECKED : 0);
			}

			if (id == SB_COMMENTS) {
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

		if (pHdr->idFrom == IDC_GRID && pHdr->code == (UINT)NM_CUSTOMDRAW) {
			int result = CDRF_DODEFAULT;

			NMLVCUSTOMDRAW* pCustomDraw = (LPNMLVCUSTOMDRAW)lParam;
			if (pCustomDraw->nmcd.dwDrawStage == CDDS_PREPAINT)
				result = CDRF_NOTIFYITEMDRAW;

			if (pCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
				if (ListView_GetItemState(pHdr->hwndFrom, pCustomDraw->nmcd.dwItemSpec, LVIS_SELECTED)) {
					pCustomDraw->nmcd.uItemState &= ~CDIS_SELECTED;
					result = CDRF_NOTIFYSUBITEMDRAW;
				}
				else {
					pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, pCustomDraw->nmcd.dwItemSpec % 2 == 0 ? TEXT("BACKCOLOR") : TEXT("BACKCOLOR2"));
				}
			}

			if (pCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM)) {
				int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				BOOL isCurrCell = (pCustomDraw->nmcd.dwItemSpec == (DWORD)rowNo) && (pCustomDraw->iSubItem == colNo);
				pCustomDraw->clrText = *(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR"));
				pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, isCurrCell ? TEXT("CURRENTCELLCOLOR") : TEXT("SELECTIONBACKCOLOR"));
			}

			return result;
		}
	}
	break;

				  // wParam = colNo
	case WMU_HIDE_COLUMN: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		HWND hHeader = ListView_GetHeader(hGridWnd);
		int colNo = (int)wParam;

		HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
		SetWindowLongPtr(hEdit, GWLP_USERDATA, (LONG_PTR)ListView_GetColumnWidth(hGridWnd, colNo));
		ListView_SetColumnWidth(hGridWnd, colNo, 0);
		InvalidateRect(hHeader, NULL, TRUE);
	}
	break;

	case WMU_SHOW_COLUMNS: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		HWND hHeader = ListView_GetHeader(hGridWnd);
		int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
		for (int colNo = 0; colNo < colCount; colNo++) {
			if (ListView_GetColumnWidth(hGridWnd, colNo) == 0) {
				HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
				ListView_SetColumnWidth(hGridWnd, colNo, (int)GetWindowLongPtr(hEdit, GWLP_USERDATA));
			}
		}

		InvalidateRect(hGridWnd, NULL, TRUE);
	}
	break;

						 // wParam = colNo
	case WMU_SORT_COLUMN: {
		int colNo = (int)wParam + 1;
		if (colNo <= 0)
			return FALSE;

		int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
		int orderBy = *pOrderBy;
		*pOrderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
		SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
	}
	break;

	case WMU_UPDATE_CACHE: {
		TCHAR* filepath = (TCHAR*)GetProp(hWnd, TEXT("FILEPATH"));
		int filesize = *(int*)GetProp(hWnd, TEXT("FILESIZE"));
		TCHAR delimiter = *(TCHAR*)GetProp(hWnd, TEXT("DELIMITER"));
		int codepage = *(int*)GetProp(hWnd, TEXT("CODEPAGE"));
		int skipComments = *(int*)GetProp(hWnd, TEXT("SKIPCOMMENTS"));

		SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);

		char* rawdata = (char*)calloc(filesize + 2, sizeof(char)); // + 2!
		FILE* f = _tfopen(filepath, TEXT("rb"));
		fread(rawdata, sizeof(char), filesize, f);
		fclose(f);
		rawdata[filesize] = 0;
		rawdata[filesize + 1] = 0;

		int leadZeros = 0;
		for (int i = 0; leadZeros == i && i < filesize; i++)
			leadZeros += rawdata[i] == 0;

		if (leadZeros == filesize) {
			PostMessage(GetParent(hWnd), WM_KEYDOWN, 0x33, 0x20001); // Switch to Hex-mode

			// A special TC-command doesn't work under TC x32. 
			// https://flint-inc.ru/tcinfo/all_cmd.ru.htm#Misc
			// PostMessage(GetAncestor(hWnd, GA_ROOT), WM_USER + 51, 4005, 0);
			keybd_event(VK_TAB, VK_TAB, KEYEVENTF_EXTENDEDKEY, 0);

			return FALSE;
		}

		if (codepage == -1)
			codepage = detectCodePage(rawdata, filesize);

		// Fix unexpected zeros
		if (codepage == CP_UTF16BE || codepage == CP_UTF16LE) {
			for (int i = 0; i < filesize / 2; i++) {
				if (rawdata[2 * i] == 0 && rawdata[2 * i + 1] == 0)
					rawdata[2 * i + codepage == CP_UTF16LE] = ' ';
			}
		}

		if (codepage == CP_UTF8 || codepage == CP_ACP) {
			for (int i = 0; i < filesize; i++) {
				if (rawdata[i] == 0)
					rawdata[i] = ' ';
			}
		}
		// end fix

		TCHAR* data = 0;
		if (codepage == CP_UTF16BE) {
			for (int i = 0; i < filesize / 2; i++) {
				int c = rawdata[2 * i];
				rawdata[2 * i] = rawdata[2 * i + 1];
				rawdata[2 * i + 1] = c;
			}
		}

		if (codepage == CP_UTF16LE || codepage == CP_UTF16BE)
			data = (TCHAR*)(rawdata + leadZeros);

		if (codepage == CP_UTF8) {
			data = utf8to16(rawdata + leadZeros);
			free(rawdata);
		}

		if (codepage == CP_ACP) {
			DWORD len = MultiByteToWideChar(CP_ACP, 0, rawdata, -1, NULL, 0);
			data = (TCHAR*)calloc(len, sizeof(TCHAR));
			if (!MultiByteToWideChar(CP_ACP, 0, rawdata + leadZeros, -1, data, len))
				codepage = -1;
			free(rawdata);
		}

		if (codepage == -1) {
			MessageBox(hWnd, TEXT("Can't detect codepage"), NULL, MB_OK);
			return 0;
		}

		if (!delimiter) {
			TCHAR* str = getStoredString(TEXT("default-column-delimiter"), TEXT(""));
			delimiter = _tcslen(str) ? str[0] : detectDelimiter(data, skipComments);
			free(str);
		}

		int colCount = 1;
		int rowNo = -1;
		int len = _tcslen(data);
		int cacheSize = len / 100 + 1;
		TCHAR*** cache = (TCHAR***)calloc(cacheSize, sizeof(TCHAR**));
		BOOL isTrim = getStoredValue(TEXT("trim-values"), 1);

		// Two step parsing: 0 - count columns, 1 - fill cache
		for (int stepNo = 0; stepNo < 2; stepNo++) {
			rowNo = -1;
			int start = 0;
			for (int pos = 0; pos < len; pos++) {
				rowNo++;
				if (stepNo == 1) {
					if (rowNo >= cacheSize) {
						cacheSize += 100;
						cache = (TCHAR***)realloc(cache, cacheSize * sizeof(TCHAR**));
					}

					cache[rowNo] = (TCHAR**)calloc(colCount, sizeof(TCHAR*));
				}

				int colNo = 0;
				start = pos;
				for (; pos < len && colNo < MAX_COLUMN_COUNT; pos++) {
					TCHAR c = data[pos];

					while (start == pos && (stepNo == 0 && skipComments == 1 || skipComments == 2) && c == TEXT('#')) {
						while (data[pos] && !isEOL(data[pos]))
							pos++;
						while (pos < len && isEOL(data[pos]))
							pos++;
						c = data[pos];
						start = pos;
					}

					if (c == delimiter && !(skipComments && data[start] == TEXT('#')) || isEOL(c) || pos >= len - 1) {
						int vLen = pos - start + (pos >= len - 1);
						int qPos = -1;
						for (int i = 0; i < vLen; i++) {
							TCHAR c = data[start + i];
							if (c != TEXT(' ') && c != TEXT('\t')) {
								qPos = c == TEXT('"') ? i : -1;
								break;
							}
						}

						if (qPos != -1) {
							int qCount = 0;
							for (int i = qPos; i <= vLen; i++)
								qCount += data[start + i] == TEXT('"');

							if (pos < len && qCount % 2)
								continue;

							while (vLen > 0 && data[start + vLen] != TEXT('"'))
								vLen--;

							vLen -= qPos + 1;
						}

						if (stepNo == 1) {
							if (isTrim) {
								int l = 0, r = 0;
								if (qPos == -1) {
									for (int i = 0; i < vLen && (data[start + i] == TEXT(' ') || data[start + i] == TEXT('\t')); i++)
										l++;
								}
								for (int i = vLen - 1; i > 0 && (data[start + i] == TEXT(' ') || data[start + i] == TEXT('\t')); i--)
									r++;

								start += l;
								vLen -= l + r;
							}

							TCHAR* value = (TCHAR*)calloc(vLen + 1, sizeof(TCHAR));
							if (vLen > 0) {
								_tcsncpy(value, data + start + qPos + 1, vLen);

								// replace "" by " in quoted value
								if (qPos != -1) {
									int qCount = 0, j = 0;
									for (int i = 0; i < vLen; i++) {
										BOOL isQ = value[i] == TEXT('"');
										qCount += isQ;
										value[j] = value[i];
										j += isQ && qCount % 2 || !isQ;
									}
									value[j] = 0;
								}
							}

							cache[rowNo][colNo] = value;
						}

						start = pos + 1;
						colNo++;
					}

					if (isEOL(data[pos])) {
						while (isEOL(data[pos + 1]))
							pos++;
						break;
					}
				}

				// Truncate 127+ columns
				if (!isEOL(data[pos])) {
					while (pos < len && !isEOL(data[pos + 1]))
						pos++;
					while (pos < len && isEOL(data[pos + 1]))
						pos++;
				}

				while (stepNo == 1 && colNo < colCount) {
					cache[rowNo][colNo] = (TCHAR*)calloc(1, sizeof(TCHAR));
					colNo++;
				}

				if (stepNo == 0 && colCount < colNo)
					colCount = colNo;
			}

			if (stepNo == 1) {
				cache = (TCHAR***)realloc(cache, (rowNo + 1) * sizeof(TCHAR**));
				if (codepage == CP_UTF16LE || codepage == CP_UTF16BE)
					free(rawdata);
				else
					free(data);
			}
		}

		if (colCount > MAX_COLUMN_COUNT) {
			HWND hListerWnd = GetParent(hWnd);
			TCHAR msg[255];
			_sntprintf(msg, 255, TEXT("Column count is overflow.\nFound: %i, max: %i"), colCount, MAX_COLUMN_COUNT);
			MessageBox(hWnd, msg, NULL, 0);
			SendMessage(hListerWnd, WM_CLOSE, 0, 0);
			return 0;
		}

		HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
		SendMessage(hStatusWnd, SB_SETTEXT, SB_CODEPAGE, (LPARAM)(
			codepage == CP_UTF8 ? TEXT("    UTF-8") :
			codepage == CP_UTF16LE ? TEXT(" UTF-16LE") :
			codepage == CP_UTF16BE ? TEXT(" UTF-16BE") : TEXT("     ANSI")
			));

		TCHAR buf[32] = { 0 };
		if (delimiter != TEXT('\t'))
			_sntprintf(buf, 32, TEXT(" %c"), delimiter);
		else
			_sntprintf(buf, 32, TEXT(" TAB"));
		SendMessage(hStatusWnd, SB_SETTEXT, SB_DELIMITER, (LPARAM)buf);

		_sntprintf(buf, 32, TEXT(" Comment mode: #%i     "), skipComments); // Long tail for a tip
		SendMessage(hStatusWnd, SB_SETTEXT, SB_COMMENTS, (LPARAM)buf);

		SetProp(hWnd, TEXT("CACHE"), cache);
		*(int*)GetProp(hWnd, TEXT("COLCOUNT")) = colCount;
		*(int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")) = rowNo + 1;
		*(int*)GetProp(hWnd, TEXT("CODEPAGE")) = codepage;
		*(TCHAR*)GetProp(hWnd, TEXT("DELIMITER")) = delimiter;
		*(int*)GetProp(hWnd, TEXT("ORDERBY")) = 0;

		SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);

		return TRUE;
	}
	break;

	case WMU_UPDATE_GRID: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
		HWND hHeader = ListView_GetHeader(hGridWnd);
		TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
		int rowCount = *(int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
		int isHeaderRow = *(int*)GetProp(hWnd, TEXT("HEADERROW"));
		int filterAlign = *(int*)GetProp(hWnd, TEXT("FILTERALIGN"));

		int colCount = Header_GetItemCount(hHeader);
		for (int colNo = 0; colNo < colCount; colNo++)
			DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

		for (int colNo = 0; colNo < colCount; colNo++)
			ListView_DeleteColumn(hGridWnd, colCount - colNo - 1);

		colCount = *(int*)GetProp(hWnd, TEXT("COLCOUNT"));
		for (int colNo = 0; colNo < colCount; colNo++) {
			int cNum = 0;
			int cCount = rowCount < 6 ? rowCount : 6;
			for (int rowNo = 1; rowNo < cCount; rowNo++)
				cNum += cache[rowNo][colNo] && _tcslen(cache[rowNo][colNo]) && isNumber(cache[rowNo][colNo]);

			int fmt = cNum > cCount / 2 ? LVCFMT_RIGHT : LVCFMT_LEFT;
			TCHAR colName[64];
			_sntprintf(colName, 64, TEXT("Column #%i"), colNo + 1);
			ListView_AddColumn(hGridWnd, isHeaderRow && cache[0][colNo] && _tcslen(cache[0][colNo]) > 0 ? cache[0][colNo] : colName, fmt);
		}

		int align = filterAlign == -1 ? ES_LEFT : filterAlign == 1 ? ES_RIGHT : ES_CENTER;
		for (int colNo = 0; colNo < colCount; colNo++) {
			// Use WS_BORDER to vertical text aligment
			HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, align | ES_AUTOHSCROLL | WS_CHILD | WS_TABSTOP | WS_BORDER,
				0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
			SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
			SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
		}

		SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
		SendMessage(hWnd, WMU_SET_HEADER_FILTERS, 0, 0);
		PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
	}
	break;

	case WMU_UPDATE_RESULTSET: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
		HWND hHeader = ListView_GetHeader(hGridWnd);

		ListView_SetItemCount(hGridWnd, 0);
		TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
		int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
		int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));
		int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
		int isHeaderRow = *(int*)GetProp(hWnd, TEXT("HEADERROW"));
		int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
		BOOL isCaseSensitive = getStoredValue(TEXT("filter-case-sensitive"), 0);

		if (resultset)
			free(resultset);

		if (!cache)
			return 1;

		if (*pTotalRowCount == 0)
			return 1;

		int colCount = Header_GetItemCount(hHeader);
		if (colCount == 0)
			return 1;

		BOOL* bResultset = (BOOL*)calloc(*pTotalRowCount, sizeof(BOOL));
		bResultset[0] = !isHeaderRow;
		for (int rowNo = 1; rowNo < *pTotalRowCount; rowNo++)
			bResultset[rowNo] = TRUE;

		for (int colNo = 0; colNo < colCount; colNo++) {
			HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
			TCHAR filter[MAX_FILTER_LENGTH];
			GetWindowText(hEdit, filter, MAX_FILTER_LENGTH);
			int len = _tcslen(filter);
			if (len == 0)
				continue;

			for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
				if (!bResultset[rowNo])
					continue;

				TCHAR* value = cache[rowNo][colNo];
				if (len > 1 && (filter[0] == TEXT('<') || filter[0] == TEXT('>')) && isNumber(filter + 1)) {
					TCHAR* end = 0;
					double df = _tcstod(filter + 1, &end);
					double dv = _tcstod(value, &end);
					bResultset[rowNo] = (filter[0] == TEXT('<') && dv < df) || (filter[0] == TEXT('>') && dv > df);
				}
				else {
					bResultset[rowNo] = len == 1 ? hasString(value, filter, isCaseSensitive) :
						filter[0] == TEXT('=') && isCaseSensitive ? _tcscmp(value, filter + 1) == 0 :
						filter[0] == TEXT('=') && !isCaseSensitive ? _tcsicmp(value, filter + 1) == 0 :
						filter[0] == TEXT('!') ? !hasString(value, filter + 1, isCaseSensitive) :
						filter[0] == TEXT('<') ? _tcscmp(value, filter + 1) < 0 :
						filter[0] == TEXT('>') ? _tcscmp(value, filter + 1) > 0 :
						hasString(value, filter, isCaseSensitive);
				}
			}
		}

		int rowCount = 0;
		resultset = (int*)calloc(*pTotalRowCount, sizeof(int));
		for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
			if (!bResultset[rowNo])
				continue;

			resultset[rowCount] = rowNo;
			rowCount++;
		}
		free(bResultset);

		if (rowCount > 0) {
			if (rowCount > *pTotalRowCount)
				MessageBeep(0);
			resultset = (int*)realloc(resultset, rowCount * sizeof(int));
			SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)resultset);
			int orderBy = *pOrderBy;

			if (orderBy) {
				int colNo = orderBy > 0 ? orderBy - 1 : -orderBy - 1;
				BOOL isBackward = orderBy < 0;

				BOOL isNum = TRUE;
				for (int i = isHeaderRow; i < *pTotalRowCount && i <= 5; i++)
					isNum = isNum && isNumber(cache[i][colNo]);

				if (isNum) {
					double* nums = (double*)calloc(*pTotalRowCount, sizeof(double));
					for (int i = 0; i < rowCount; i++)
						nums[resultset[i]] = _tcstod(cache[resultset[i]][colNo], NULL);

					mergeSort(resultset, (void*)nums, 0, rowCount - 1, isBackward, isNum);
					free(nums);
				}
				else {
					TCHAR** strings = (TCHAR**)calloc(*pTotalRowCount, sizeof(TCHAR*));
					for (int i = 0; i < rowCount; i++)
						strings[resultset[i]] = cache[resultset[i]][colNo];
					mergeSort(resultset, (void*)strings, 0, rowCount - 1, isBackward, isNum);
					free(strings);
				}
			}
		}
		else {
			SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)0);
			free(resultset);
		}

		*pRowCount = rowCount;
		ListView_SetItemCount(hGridWnd, rowCount);
		InvalidateRect(hGridWnd, NULL, TRUE);

		TCHAR buf[255];
		_sntprintf(buf, 255, TEXT(" Rows: %i/%i"), rowCount, *pTotalRowCount - isHeaderRow);
		SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)buf);

		PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
	}
	break;

	case WMU_UPDATE_FILTER_SIZE: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		HWND hHeader = ListView_GetHeader(hGridWnd);
		int colCount = Header_GetItemCount(hHeader);
		SendMessage(hHeader, WM_SIZE, 0, 0);

		int* colOrder = (int*)calloc(colCount, sizeof(int));
		Header_GetOrderArray(hHeader, colCount, colOrder);

		for (int idx = 0; idx < colCount; idx++) {
			int colNo = colOrder[idx];
			RECT rc;
			Header_GetItemRect(hHeader, colNo, &rc);
			int h2 = round((rc.bottom - rc.top) / 2);
			SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left, h2, rc.right - rc.left, h2 + 1, SWP_NOZORDER);
		}

		free(colOrder);
	}
	break;

	case WMU_SET_HEADER_FILTERS: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		HWND hHeader = ListView_GetHeader(hGridWnd);
		int isFilterRow = *(int*)GetProp(hWnd, TEXT("FILTERROW"));
		int colCount = Header_GetItemCount(hHeader);

		SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
		LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
		styles = isFilterRow ? styles | HDS_FILTERBAR : styles & (~HDS_FILTERBAR);
		SetWindowLongPtr(hHeader, GWL_STYLE, styles);

		for (int colNo = 0; colNo < colCount; colNo++)
			ShowWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), isFilterRow ? SW_SHOW : SW_HIDE);

		// Bug fix: force Windows to redraw header
		SetWindowPos(hGridWnd, 0, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOMOVE);
		SendMessage(getMainWindow(hWnd), WM_SIZE, 0, 0);

		if (isFilterRow)
			SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);

		SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
		InvalidateRect(hWnd, NULL, TRUE);
	}
	break;

	case WMU_AUTO_COLUMN_SIZE: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
		HWND hHeader = ListView_GetHeader(hGridWnd);
		int colCount = Header_GetItemCount(hHeader);

		for (int colNo = 0; colNo < colCount - 1; colNo++)
			ListView_SetColumnWidth(hGridWnd, colNo, colNo < colCount - 1 ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE);

		if (colCount == 1 && ListView_GetColumnWidth(hGridWnd, 0) < 100)
			ListView_SetColumnWidth(hGridWnd, 0, 100);

		int maxWidth = getStoredValue(TEXT("max-column-width"), 300);
		if (colCount > 1) {
			for (int colNo = 0; colNo < colCount; colNo++) {
				if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
					ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);
			}
		}

		// Fix last column				
		if (colCount > 1) {
			int colNo = colCount - 1;
			ListView_SetColumnWidth(hGridWnd, colNo, LVSCW_AUTOSIZE);
			TCHAR name16[MAX_COLUMN_LENGTH + 1];
			Header_GetItemText(hHeader, colNo, name16, MAX_COLUMN_LENGTH);

			SIZE s = { 0 };
			HDC hDC = GetDC(hHeader);
			HFONT hOldFont = (HFONT)SelectObject(hDC, (HFONT)GetProp(hWnd, TEXT("FONT")));
			GetTextExtentPoint32(hDC, name16, _tcslen(name16), &s);
			SelectObject(hDC, hOldFont);
			ReleaseDC(hHeader, hDC);

			int w = s.cx + 12;
			if (ListView_GetColumnWidth(hGridWnd, colNo) < w)
				ListView_SetColumnWidth(hGridWnd, colNo, w);

			if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
				ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);
		}

		SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
		InvalidateRect(hGridWnd, NULL, TRUE);

		PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
	}
	break;

							 // wParam = rowNo, lParam = colNo
	case WMU_SET_CURRENT_CELL: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		HWND hHeader = ListView_GetHeader(hGridWnd);
		HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
		SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)0);

		int* pRowNo = (int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
		int* pColNo = (int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
		if (*pRowNo == wParam && *pColNo == lParam)
			return 0;

		RECT rc, rc2;
		ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
		if (*pColNo == 0)
			rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);
		InvalidateRect(hGridWnd, &rc, TRUE);

		*pRowNo = wParam;
		*pColNo = lParam;
		ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
		if (*pColNo == 0)
			rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);
		InvalidateRect(hGridWnd, &rc, FALSE);

		GetClientRect(hGridWnd, &rc2);
		int w = rc.right - rc.left;
		int dx = rc2.right < rc.right ? rc.left - rc2.right + w : rc.left < 0 ? rc.left : 0;

		ListView_Scroll(hGridWnd, dx, 0);

		TCHAR buf[32] = { 0 };
		if (*pColNo != -1 && *pRowNo != -1)
			_sntprintf(buf, 32, TEXT(" %i:%i"), *pRowNo + 1, *pColNo + 1);
		SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_CELL, (LPARAM)buf);
	}
	break;

	case WMU_RESET_CACHE: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
		int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));

		int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
		if (colCount > 0 && cache != 0) {
			for (int rowNo = 0; rowNo < *pRowCount; rowNo++) {
				if (cache[rowNo]) {
					for (int colNo = 0; colNo < colCount; colNo++)
						if (cache[rowNo][colNo])
							free(cache[rowNo][colNo]);

					free(cache[rowNo]);
				}
				cache[rowNo] = 0;
			}
			free(cache);
		}

		SetProp(hWnd, TEXT("CACHE"), 0);
		*pRowCount = 0;
	}
	break;

						// wParam - size delta
	case WMU_SET_FONT: {
		int* pFontSize = (int*)GetProp(hWnd, TEXT("FONTSIZE"));
		if (*pFontSize + wParam < 10 || *pFontSize + wParam > 48)
			return 0;
		*pFontSize += wParam;
		DeleteFont(GetProp(hWnd, TEXT("FONT")));

		HFONT hFont = CreateFont(*pFontSize, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, (TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		SendMessage(hGridWnd, WM_SETFONT, (LPARAM)hFont, TRUE);

		HWND hHeader = ListView_GetHeader(hGridWnd);
		for (int colNo = 0; colNo < Header_GetItemCount(hHeader); colNo++)
			SendMessage(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), WM_SETFONT, (LPARAM)hFont, TRUE);

		SetProp(hWnd, TEXT("FONT"), hFont);
		PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
	}
	break;

	case WMU_SET_THEME: {
		HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
		BOOL isDark = *(int*)GetProp(hWnd, TEXT("DARKTHEME"));

		int textColor = !isDark ? getStoredValue(TEXT("text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("text-color-dark"), RGB(220, 220, 220));
		int backColor = !isDark ? getStoredValue(TEXT("back-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("back-color-dark"), RGB(32, 32, 32));
		int backColor2 = !isDark ? getStoredValue(TEXT("back-color2"), RGB(240, 240, 240)) : getStoredValue(TEXT("back-color2-dark"), RGB(52, 52, 52));
		int filterTextColor = !isDark ? getStoredValue(TEXT("filter-text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("filter-text-color-dark"), RGB(255, 255, 255));
		int filterBackColor = !isDark ? getStoredValue(TEXT("filter-back-color"), RGB(240, 240, 240)) : getStoredValue(TEXT("filter-back-color-dark"), RGB(60, 60, 60));
		int currCellColor = !isDark ? getStoredValue(TEXT("current-cell-back-color"), RGB(70, 96, 166)) : getStoredValue(TEXT("current-cell-back-color-dark"), RGB(32, 62, 62));
		int selectionTextColor = !isDark ? getStoredValue(TEXT("selection-text-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("selection-text-color-dark"), RGB(220, 220, 220));
		int selectionBackColor = !isDark ? getStoredValue(TEXT("selection-back-color"), RGB(10, 36, 106)) : getStoredValue(TEXT("selection-back-color-dark"), RGB(72, 102, 102));

		*(int*)GetProp(hWnd, TEXT("TEXTCOLOR")) = textColor;
		*(int*)GetProp(hWnd, TEXT("BACKCOLOR")) = backColor;
		*(int*)GetProp(hWnd, TEXT("BACKCOLOR2")) = backColor2;
		*(int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")) = filterTextColor;
		*(int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")) = filterBackColor;
		*(int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")) = currCellColor;
		*(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")) = selectionTextColor;
		*(int*)GetProp(hWnd, TEXT("SELECTIONBACKCOLOR")) = selectionBackColor;

		DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
		SetProp(hWnd, TEXT("FILTERBACKBRUSH"), CreateSolidBrush(filterBackColor));

		ListView_SetTextColor(hGridWnd, textColor);
		ListView_SetBkColor(hGridWnd, backColor);
		ListView_SetTextBkColor(hGridWnd, backColor);
		InvalidateRect(hWnd, NULL, TRUE);
	}
	break;

	case WMU_HOT_KEYS: {
		BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
		if (wParam == VK_TAB) {
			HWND hFocus = GetFocus();
			HWND wnds[1000] = { 0 };
			EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumTabStopChildren, (LPARAM)wnds);

			int no = 0;
			while (wnds[no] && wnds[no] != hFocus)
				no++;

			int cnt = no;
			while (wnds[cnt])
				cnt++;

			no += isCtrl ? -1 : 1;
			SetFocus(wnds[no] && no >= 0 ? wnds[no] : (isCtrl ? wnds[cnt - 1] : wnds[0]));
		}

		if (wParam == VK_F1) {
			ShellExecute(0, 0, TEXT("https://github.com/little-brother/csvtab-wlx/wiki"), 0, 0, SW_SHOW);
			return TRUE;
		}

		if (wParam == 0x20 && isCtrl) { // Ctrl + Space
			SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);
			return TRUE;
		}

		if (wParam == VK_ESCAPE || wParam == VK_F11 ||
			wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7 || (isCtrl && wParam == 0x46) || // Ctrl + F
			((wParam >= 0x31 && wParam <= 0x38) && !getStoredValue(TEXT("disable-num-keys"), 0) && !isCtrl || // 1 - 8
				(wParam == 0x4E || wParam == 0x50) && !getStoredValue(TEXT("disable-np-keys"), 0) || // N, P
				wParam == 0x51 && getStoredValue(TEXT("exit-by-q"), 0)) && // Q
			GetDlgCtrlID(GetFocus()) / 100 * 100 != IDC_HEADER_EDIT) {
			SetFocus(GetParent(hWnd));
			keybd_event(wParam, wParam, KEYEVENTF_EXTENDEDKEY, 0);

			return TRUE;
		}

		if (isCtrl && wParam >= 0x30 && wParam <= 0x39 && !getStoredValue(TEXT("disable-num-keys"), 0)) {// 0-9
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));

			BOOL isCurrent = wParam == 0x30;
			int colNo = isCurrent ? *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) : wParam - 0x30 - 1;
			if (colNo < 0 || colNo > colCount - 1 || isCurrent && ListView_GetColumnWidth(hGridWnd, colNo) == 0)
				return FALSE;

			if (!isCurrent) {
				int* colOrder = (int*)calloc(colCount, sizeof(int));
				Header_GetOrderArray(hHeader, colCount, colOrder);

				int hiddenColCount = 0;
				for (int idx = 0; (idx < colCount) && (idx - hiddenColCount <= colNo); idx++)
					hiddenColCount += ListView_GetColumnWidth(hGridWnd, colOrder[idx]) == 0;

				colNo = colOrder[colNo + hiddenColCount];
				free(colOrder);
			}

			SendMessage(hWnd, WMU_SORT_COLUMN, colNo, 0);
			return TRUE;
		}

		return FALSE;
	}
					 break;

	case WMU_HOT_CHARS: {
		BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));

		unsigned char scancode = ((unsigned char*)&lParam)[2];
		UINT key = MapVirtualKey(scancode, MAPVK_VSC_TO_VK);

		return !_istprint(wParam) && (
			wParam == VK_ESCAPE || wParam == VK_F11 || wParam == VK_F1 ||
			wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7) ||
			wParam == VK_TAB || wParam == VK_RETURN ||
			isCtrl && key == 0x46 || // Ctrl + F
			getStoredValue(TEXT("exit-by-q"), 0) && key == 0x51 && GetDlgCtrlID(GetFocus()) / 100 * 100 != IDC_HEADER_EDIT; // Q
		}
					  break;
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}
