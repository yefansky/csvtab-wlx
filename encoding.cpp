#include "stdafx.h"

// https://stackoverflow.com/a/25023604/6121703
int detectCodePage(const char* data, int len) {
	int cp = 0;

	// BOM
	cp = strncmp(data, "\xEF\xBB\xBF", 3) == 0 ? CP_UTF8 :
		strncmp(data, "\xFE\xFF", 2) == 0 ? CP_UTF16BE :
		strncmp(data, "\xFF\xFE", 2) == 0 ? CP_UTF16LE :
		0;

	// By \n: UTF16LE - 0x0A 0x00, UTF16BE - 0x00 0x0A
	if (!cp) {
		int nCount = 0;
		int leadZeros = 0;
		int tailZeros = 0;
		for (int i = 0; i < len; i++) {
			if (data[i] != '\n')
				continue;

			nCount++;
			leadZeros += i > 0 && data[i - 1] == 0;
			tailZeros += (i < len - 1) && data[i + 1] == 0;
		}

		cp = nCount == 0 ? 0 :
			nCount == tailZeros ? CP_UTF16LE :
			nCount == leadZeros ? CP_UTF16BE :
			0;
	}

	if (!cp && isUtf8(data))
		cp = CP_UTF8;

	if (!cp)
		cp = CP_ACP;

	return cp;
}

// https://stackoverflow.com/a/1031773/6121703
BOOL isUtf8(const char* string) {
	if (!string)
		return FALSE;

	const unsigned char* bytes = (const unsigned char*)string;
	while (*bytes) {
		if ((bytes[0] == 0x09 || bytes[0] == 0x0A || bytes[0] == 0x0D || (0x20 <= bytes[0] && bytes[0] <= 0x7E))) {
			bytes += 1;
			continue;
		}

		if (((0xC2 <= bytes[0] && bytes[0] <= 0xDF) && (0x80 <= bytes[1] && bytes[1] <= 0xBF))) {
			bytes += 2;
			continue;
		}

		if ((bytes[0] == 0xE0 && (0xA0 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF)) ||
			(((0xE1 <= bytes[0] && bytes[0] <= 0xEC) || bytes[0] == 0xEE || bytes[0] == 0xEF) && (0x80 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF)) ||
			(bytes[0] == 0xED && (0x80 <= bytes[1] && bytes[1] <= 0x9F) && (0x80 <= bytes[2] && bytes[2] <= 0xBF))
			) {
			bytes += 3;
			continue;
		}

		if ((bytes[0] == 0xF0 && (0x90 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF)) ||
			((0xF1 <= bytes[0] && bytes[0] <= 0xF3) && (0x80 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF)) ||
			(bytes[0] == 0xF4 && (0x80 <= bytes[1] && bytes[1] <= 0x8F) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF))
			) {
			bytes += 4;
			continue;
		}

		return FALSE;
	}

	return TRUE;
}

TCHAR* utf8to16(const char* in) {
	TCHAR* out;
	if (!in || strlen(in) == 0) {
		out = (TCHAR*)calloc(1, sizeof(TCHAR));
	}
	else {
		DWORD size = MultiByteToWideChar(CP_UTF8, 0, in, -1, NULL, 0);
		out = (TCHAR*)calloc(size, sizeof(TCHAR));
		MultiByteToWideChar(CP_UTF8, 0, in, -1, out, size);
	}
	return out;
}

char* utf16to8(const TCHAR* in) {
	char* out;
	if (!in || _tcslen(in) == 0) {
		out = (char*)calloc(1, sizeof(char));
	}
	else {
		int len = WideCharToMultiByte(CP_UTF8, 0, in, -1, NULL, 0, 0, 0);
		out = (char*)calloc(len, sizeof(char));
		WideCharToMultiByte(CP_UTF8, 0, in, -1, out, len, 0, 0);
	}
	return out;
}
