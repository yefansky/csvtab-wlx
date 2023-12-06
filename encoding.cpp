#include "stdafx.h"

// https://stackoverflow.com/a/25023604/6121703
int detectCodePage(const char* cpszData, int nLen) {
	int nRetCode = 0;

	// BOM
	nRetCode = strncmp(cpszData, "\xEF\xBB\xBF", 3) == 0 ? CP_UTF8 :
		strncmp(cpszData, "\xFE\xFF", 2) == 0 ? CP_UTF16BE :
		strncmp(cpszData, "\xFF\xFE", 2) == 0 ? CP_UTF16LE :
		0;

	// By \n: UTF16LE - 0x0A 0x00, UTF16BE - 0x00 0x0A
	if (!nRetCode) 
	{
		int nCount = 0;
		int leadZeros = 0;
		int tailZeros = 0;
		for (int i = 0; i < nLen; i++) 
		{
			if (cpszData[i] != '\n')
				continue;

			nCount++;
			leadZeros += i > 0 && cpszData[i - 1] == 0;
			tailZeros += (i < nLen - 1) && cpszData[i + 1] == 0;
		}

		nRetCode = nCount == 0 ? 0 :
			nCount == tailZeros ? CP_UTF16LE :
			nCount == leadZeros ? CP_UTF16BE :
			0;
	}

	if (!nRetCode && isUtf8(cpszData))
		nRetCode = CP_UTF8;

	if (!nRetCode)
		nRetCode = CP_ACP;

	return nRetCode;
}

// https://stackoverflow.com/a/1031773/6121703
BOOL isUtf8(const char* cpszString) 
{
	if (!cpszString)
		return FALSE;

	const unsigned char* pszOffset = (const unsigned char*)cpszString;
	while (*pszOffset) 
	{
		if ((pszOffset[0] == 0x09 || pszOffset[0] == 0x0A || pszOffset[0] == 0x0D || (0x20 <= pszOffset[0] && pszOffset[0] <= 0x7E))) 
		{
			pszOffset += 1;
			continue;
		}

		if (((0xC2 <= pszOffset[0] && pszOffset[0] <= 0xDF) && (0x80 <= pszOffset[1] && pszOffset[1] <= 0xBF)))
		{
			pszOffset += 2;
			continue;
		}

		if ((pszOffset[0] == 0xE0 && (0xA0 <= pszOffset[1] && pszOffset[1] <= 0xBF) && (0x80 <= pszOffset[2] && pszOffset[2] <= 0xBF)) ||
			(((0xE1 <= pszOffset[0] && pszOffset[0] <= 0xEC) || pszOffset[0] == 0xEE || pszOffset[0] == 0xEF) && (0x80 <= pszOffset[1] && pszOffset[1] <= 0xBF) && (0x80 <= pszOffset[2] && pszOffset[2] <= 0xBF)) ||
			(pszOffset[0] == 0xED && (0x80 <= pszOffset[1] && pszOffset[1] <= 0x9F) && (0x80 <= pszOffset[2] && pszOffset[2] <= 0xBF))
			) 
		{
			pszOffset += 3;
			continue;
		}

		if ((pszOffset[0] == 0xF0 && (0x90 <= pszOffset[1] && pszOffset[1] <= 0xBF) && (0x80 <= pszOffset[2] && pszOffset[2] <= 0xBF) && (0x80 <= pszOffset[3] && pszOffset[3] <= 0xBF)) ||
			((0xF1 <= pszOffset[0] && pszOffset[0] <= 0xF3) && (0x80 <= pszOffset[1] && pszOffset[1] <= 0xBF) && (0x80 <= pszOffset[2] && pszOffset[2] <= 0xBF) && (0x80 <= pszOffset[3] && pszOffset[3] <= 0xBF)) ||
			(pszOffset[0] == 0xF4 && (0x80 <= pszOffset[1] && pszOffset[1] <= 0x8F) && (0x80 <= pszOffset[2] && pszOffset[2] <= 0xBF) && (0x80 <= pszOffset[3] && pszOffset[3] <= 0xBF))
			) 
		{
			pszOffset += 4;
			continue;
		}

		return FALSE;
	}

	return TRUE;
}

TCHAR* utf8to16(const char* cpszIn) 
{
	TCHAR* pszResult = nullptr;
	if (!cpszIn || strlen(cpszIn) == 0) 
	{
		pszResult = (TCHAR*)calloc(1, sizeof(TCHAR));
	}
	else 
	{
		DWORD size = MultiByteToWideChar(CP_UTF8, 0, cpszIn, -1, NULL, 0);
		pszResult = (TCHAR*)calloc(size, sizeof(TCHAR));
		MultiByteToWideChar(CP_UTF8, 0, cpszIn, -1, pszResult, size);
	}
	return pszResult;
}

char* utf16to8(const TCHAR* cpszIn)
{
	char* pszResult;
	if (!cpszIn || _tcslen(cpszIn) == 0) 
	{
		pszResult = (char*)calloc(1, sizeof(char));
	}
	else 
	{
		int len = WideCharToMultiByte(CP_UTF8, 0, cpszIn, -1, NULL, 0, 0, 0);
		pszResult = (char*)calloc(len, sizeof(char));
		WideCharToMultiByte(CP_UTF8, 0, cpszIn, -1, pszResult, len, 0, 0);
	}
	return pszResult;
}
