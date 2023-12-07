#include "stdafx.h"

extern TCHAR g_szIniPath[MAX_PATH];

void setStoredValue(const TCHAR* cpszName, int nValue) 
{
	TCHAR szBuf[128];
	_sntprintf(szBuf, 128, TEXT("%i"), nValue);
	WritePrivateProfileString(APP_NAME, cpszName, szBuf, g_szIniPath);
}

int getStoredValue(const TCHAR* cpszName, int nDefValue) 
{
	TCHAR szBuf[128];
	return GetPrivateProfileString(APP_NAME, cpszName, NULL, szBuf, 128, g_szIniPath) ? _ttoi(szBuf) : nDefValue;
}

TCHAR* getStoredString(const TCHAR* cpszName, const TCHAR* cpszDefValue) 
{
	TCHAR* pszBuf = (TCHAR*)calloc(256, sizeof(TCHAR));
	if (0 == GetPrivateProfileString(APP_NAME, cpszName, NULL, pszBuf, 128, g_szIniPath) && cpszDefValue)
		_tcsncpy(pszBuf, cpszDefValue, 255);
	return pszBuf;
}


TCHAR detectDelimiter(const TCHAR* cpszData, BOOL bSkipComments) 
{
	const int nRowCount = 5;
	const int cnDelimCount = _countof(DELIMITERS) - 1;
	int nColCount[nRowCount][cnDelimCount];
	int nTotal[cnDelimCount];
	int nMaxNo = 0;

	static_assert(cnDelimCount == 5, "cnDelimCount size error");

	memset(nColCount, 0, sizeof(int) * (size_t)(nRowCount * cnDelimCount));

	int nPos = 0;
	int nLen = (int)_tcslen(cpszData);
	BOOL bInQuote = FALSE;
	int nRowNo = 0;
	for (; nRowNo < nRowCount && nPos < nLen;) 
	{
		int total = 0;
		for (; nPos < nLen; nPos++) 
		{
			TCHAR c = cpszData[nPos];

			while (!bInQuote && bSkipComments == 2 && c == TEXT('#')) 
			{
				while (cpszData[nPos] && !IsEOL(cpszData[nPos]))
					nPos++;
				while (nPos < nLen && IsEOL(cpszData[nPos]))
					nPos++;
				c = cpszData[nPos];
			}

			if (!bInQuote && c == TEXT('\n')) 
			{
				break;
			}

			if (c == TEXT('"'))
				bInQuote = !bInQuote;

			for (int delimNo = 0; delimNo < cnDelimCount && !bInQuote; delimNo++) 
			{
				nColCount[nRowNo][delimNo] += cpszData[nPos] == DELIMITERS[delimNo];
				total += nColCount[nRowNo][delimNo];
			}
		}

		nRowNo += total > 0;
		nPos++;
	}

	{
		int _rowCount = nRowNo;

		memset(nTotal, 0, sizeof(int) * (size_t)(cnDelimCount));
		for (int delimNo = 0; delimNo < cnDelimCount; delimNo++) 
		{
			for (int i = 0; i < _rowCount; i++) 
			{
				for (int j = 0; j < _rowCount; j++) 
				{
					nTotal[delimNo] += nColCount[i][delimNo] == nColCount[j][delimNo] && nColCount[i][delimNo] > 0 ? 10 + nColCount[i][delimNo] :
						nColCount[i][delimNo] != 0 ? 5 :
						0;
				}
			}
		}
	}

	int nMaxCount = 0;
	for (int nDelimNo = 0; nDelimNo < cnDelimCount; nDelimNo++)
	{
		if (nMaxCount < nTotal[nDelimNo]) 
		{
			nMaxNo = nDelimNo;
			nMaxCount = nTotal[nDelimNo];
		}
	}

	return DELIMITERS[nMaxNo];
}

void setClipboardText(const TCHAR* cpsText) 
{
	int nLen = (int)((_tcslen(cpsText) + 1) * sizeof(TCHAR));
	HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, nLen);
	memcpy(GlobalLock(hMem), cpsText, nLen);
	GlobalUnlock(hMem);
	OpenClipboard(0);
	EmptyClipboard();
	SetClipboardData(CF_UNICODETEXT, hMem);
	CloseClipboard();
}

int findString(TCHAR* pszText, TCHAR* pszWord, BOOL bisMatchCase, BOOL bisWholeWords) 
{
	if (!pszText || !pszWord)
		return -1;

	int nResult = -1;
	int ntlen = (int)_tcslen(pszText);
	int nwlen = (int)_tcslen(pszWord);
	if (!ntlen || !nwlen)
		return nResult;

	if (!bisMatchCase) 
	{
		TCHAR* ltext = (TCHAR*)calloc(ntlen + 1, sizeof(TCHAR));
		_tcsncpy(ltext, pszText, ntlen);
		pszText = _tcslwr(ltext);

		TCHAR* lword = (TCHAR*)calloc(nwlen + 1, sizeof(TCHAR));
		_tcsncpy(lword, pszWord, nwlen);
		pszWord = _tcslwr(lword);
	}

	if (bisWholeWords) 
	{
		for (int pos = 0; (nResult == -1) && (pos <= ntlen - nwlen); pos++)
			nResult = (pos == 0 || pos > 0 && !_istalnum(pszText[pos - 1])) &&
			!_istalnum(pszText[pos + nwlen]) &&
			_tcsncmp(pszText + pos, pszWord, nwlen) == 0 ? pos : -1;
	}
	else 
	{
		TCHAR* pS = _tcsstr(pszText, pszWord);
		nResult = (int)(pS != NULL ? pS - pszText : -1);
	}

	if (!bisMatchCase) {
		free(pszText);
		free(pszWord);
	}

	return nResult;
}

BOOL hasString(const TCHAR* cpszStr, const TCHAR* cpszSub, BOOL bisCaseSensitive) 
{
	BOOL bResult = FALSE;

	TCHAR* pszlstr = _tcsdup(cpszStr);
	_tcslwr(pszlstr);
	TCHAR* pszlsub = _tcsdup(cpszSub);
	_tcslwr(pszlsub);
	bResult = bisCaseSensitive ? _tcsstr(cpszStr, cpszSub) != 0 : _tcsstr(pszlstr, pszlsub) != 0;
	free(pszlstr);
	free(pszlsub);

	return bResult;
};

TCHAR* extractUrl(TCHAR* pszData) 
{
	int nLen = pszData ? (int)_tcslen(pszData) : 0;
	int nStart = nLen;
	int nEnd = nLen;

	TCHAR* pszUrl = (TCHAR*)calloc(nLen + 10, sizeof(TCHAR));

	TCHAR* pszSlashes = _tcsstr(pszData, TEXT("://"));
	if (pszSlashes) 
	{
		nStart = nLen - (int)_tcslen(pszSlashes);
		nEnd = nStart + 3;
		for (; nStart > 0 && _istalpha(pszData[nStart - 1]); nStart--);
		for (; nEnd < nLen && pszData[nEnd] != TEXT(' ') && pszData[nEnd] != TEXT('"') && pszData[nEnd] != TEXT('\''); nEnd++);
		_tcsncpy(pszUrl, pszData + nStart, nEnd - nStart);

	}
	else if (_tcschr(pszData, TEXT('.'))) 
	{
		_sntprintf(pszUrl, nLen + 10, TEXT("https://%ls"), pszData);
	}

	return pszUrl;
}
