#include "stdafx.h"

void setStoredValue(const TCHAR* name, int value) {
	TCHAR buf[128];
	_sntprintf(buf, 128, TEXT("%i"), value);
	WritePrivateProfileString(APP_NAME, name, buf, iniPath);
}

int getStoredValue(const TCHAR* name, int defValue) {
	TCHAR buf[128];
	return GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) ? _ttoi(buf) : defValue;
}

TCHAR* getStoredString(const TCHAR* name, const TCHAR* defValue) {
	TCHAR* buf = (TCHAR*)calloc(256, sizeof(TCHAR));
	if (0 == GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) && defValue)
		_tcsncpy(buf, defValue, 255);
	return buf;
}


TCHAR detectDelimiter(const TCHAR* data, BOOL skipComments) {
	const int rowCount = 5;
	const int delimCount = _countof(DELIMITERS) - 1;
	int colCount[rowCount][delimCount];
	int total[delimCount];
	int maxNo = 0;

	static_assert(delimCount == 5, "delimCount size error");

	memset(colCount, 0, sizeof(int) * (size_t)(rowCount * delimCount));

	int pos = 0;
	int len = _tcslen(data);
	BOOL inQuote = FALSE;
	int rowNo = 0;
	for (; rowNo < rowCount && pos < len;) {
		int total = 0;
		for (; pos < len; pos++) {
			TCHAR c = data[pos];

			while (!inQuote && skipComments == 2 && c == TEXT('#')) {
				while (data[pos] && !isEOL(data[pos]))
					pos++;
				while (pos < len && isEOL(data[pos]))
					pos++;
				c = data[pos];
			}

			if (!inQuote && c == TEXT('\n')) {
				break;
			}

			if (c == TEXT('"'))
				inQuote = !inQuote;

			for (int delimNo = 0; delimNo < delimCount && !inQuote; delimNo++) {
				colCount[rowNo][delimNo] += data[pos] == DELIMITERS[delimNo];
				total += colCount[rowNo][delimNo];
			}
		}

		rowNo += total > 0;
		pos++;
	}

	{
		int _rowCount = rowNo;

		memset(total, 0, sizeof(int) * (size_t)(delimCount));
		for (int delimNo = 0; delimNo < delimCount; delimNo++) {
			for (int i = 0; i < _rowCount; i++) {
				for (int j = 0; j < _rowCount; j++) {
					total[delimNo] += colCount[i][delimNo] == colCount[j][delimNo] && colCount[i][delimNo] > 0 ? 10 + colCount[i][delimNo] :
						colCount[i][delimNo] != 0 ? 5 :
						0;
				}
			}
		}
	}

	int maxCount = 0;
	for (int delimNo = 0; delimNo < delimCount; delimNo++) {
		if (maxCount < total[delimNo]) {
			maxNo = delimNo;
			maxCount = total[delimNo];
		}
	}

	return DELIMITERS[maxNo];
}

void setClipboardText(const TCHAR* text) {
	int len = (_tcslen(text) + 1) * sizeof(TCHAR);
	HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, len);
	memcpy(GlobalLock(hMem), text, len);
	GlobalUnlock(hMem);
	OpenClipboard(0);
	EmptyClipboard();
	SetClipboardData(CF_UNICODETEXT, hMem);
	CloseClipboard();
}

int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords) {
	if (!text || !word)
		return -1;

	int res = -1;
	int tlen = _tcslen(text);
	int wlen = _tcslen(word);
	if (!tlen || !wlen)
		return res;

	if (!isMatchCase) {
		TCHAR* ltext = (TCHAR*)calloc(tlen + 1, sizeof(TCHAR));
		_tcsncpy(ltext, text, tlen);
		text = _tcslwr(ltext);

		TCHAR* lword = (TCHAR*)calloc(wlen + 1, sizeof(TCHAR));
		_tcsncpy(lword, word, wlen);
		word = _tcslwr(lword);
	}

	if (isWholeWords) {
		for (int pos = 0; (res == -1) && (pos <= tlen - wlen); pos++)
			res = (pos == 0 || pos > 0 && !_istalnum(text[pos - 1])) &&
			!_istalnum(text[pos + wlen]) &&
			_tcsncmp(text + pos, word, wlen) == 0 ? pos : -1;
	}
	else {
		TCHAR* s = _tcsstr(text, word);
		res = s != NULL ? s - text : -1;
	}

	if (!isMatchCase) {
		free(text);
		free(word);
	}

	return res;
}

BOOL hasString(const TCHAR* str, const TCHAR* sub, BOOL isCaseSensitive) {
	BOOL res = FALSE;

	TCHAR* lstr = _tcsdup(str);
	_tcslwr(lstr);
	TCHAR* lsub = _tcsdup(sub);
	_tcslwr(lsub);
	res = isCaseSensitive ? _tcsstr(str, sub) != 0 : _tcsstr(lstr, lsub) != 0;
	free(lstr);
	free(lsub);

	return res;
};

TCHAR* extractUrl(TCHAR* data) {
	int len = data ? _tcslen(data) : 0;
	int start = len;
	int end = len;

	TCHAR* url = (TCHAR*)calloc(len + 10, sizeof(TCHAR));

	TCHAR* slashes = _tcsstr(data, TEXT("://"));
	if (slashes) {
		start = len - _tcslen(slashes);
		end = start + 3;
		for (; start > 0 && _istalpha(data[start - 1]); start--);
		for (; end < len && data[end] != TEXT(' ') && data[end] != TEXT('"') && data[end] != TEXT('\''); end++);
		_tcsncpy(url, data + start, end - start);

	}
	else if (_tcschr(data, TEXT('.'))) {
		_sntprintf(url, len + 10, TEXT("https://%ls"), data);
	}

	return url;
}
