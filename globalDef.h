#pragma once


#define LVS_EX_AUTOSIZECOLUMNS 0x10000000

#define WMU_UPDATE_CACHE       WM_USER + 1
#define WMU_UPDATE_GRID        WM_USER + 2
#define WMU_UPDATE_RESULTSET   WM_USER + 3
#define WMU_UPDATE_FILTER_SIZE WM_USER + 4
#define WMU_SET_HEADER_FILTERS WM_USER + 5
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 6
#define WMU_SET_CURRENT_CELL   WM_USER + 7
#define WMU_RESET_CACHE        WM_USER + 8
#define WMU_SET_FONT           WM_USER + 9
#define WMU_SET_THEME          WM_USER + 10
#define WMU_HIDE_COLUMN        WM_USER + 11
#define WMU_SHOW_COLUMNS       WM_USER + 12
#define WMU_SORT_COLUMN        WM_USER + 13
#define WMU_HOT_KEYS           WM_USER + 14  
#define WMU_HOT_CHARS          WM_USER + 15

#define IDC_MAIN               100
#define IDC_GRID               101
#define IDC_STATUSBAR          102
#define IDC_HEADER_EDIT        1000

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROWS          5001
#define IDM_COPY_COLUMN        5002
#define IDM_FILTER_ROW         5003
#define IDM_HEADER_ROW         5004
#define IDM_DARK_THEME         5005
#define IDM_HIDE_COLUMN        5008

#define IDM_ANSI               5010
#define IDM_UTF8               5011
#define IDM_UTF16LE            5012
#define IDM_UTF16BE            5013

#define IDM_COMMA              5020
#define IDM_SEMICOLON          5021
#define IDM_VBAR               5022
#define IDM_TAB                5023
#define IDM_COLON              5024

#define IDM_DEFAULT            5030
#define IDM_NO_PARSE           5031
#define IDM_NO_SHOW            5032

#define SB_VERSION             0
#define SB_CODEPAGE            1
#define SB_DELIMITER           2
#define SB_COMMENTS            3
#define SB_ROW_COUNT           4
#define SB_CURRENT_CELL        5
#define SB_AUXILIARY           6

#define MAX_COLUMN_COUNT       128
#define MAX_COLUMN_LENGTH      2000
#define MAX_FILTER_LENGTH      2000
//#define DELIMITERS             TEXT(",;|\t:")
const TCHAR DELIMITERS[] = TEXT(",;|\t:");
#define APP_NAME               TEXT("csvtab")
#define APP_VERSION            TEXT("1.0.6")

#define CP_UTF16LE             1200
#define CP_UTF16BE             1201

#define LCS_FINDFIRST          1
#define LCS_MATCHCASE          2
#define LCS_WHOLEWORDS         4
#define LCS_BACKWARDS          8

typedef struct {
	int size;
	DWORD PluginInterfaceVersionLow;
	DWORD PluginInterfaceVersionHi;
	char DefaultIniName[MAX_PATH];
} ListDefaultParamStruct;

static TCHAR iniPath[MAX_PATH] = { 0 };

void __stdcall ListCloseWindow(HWND hWnd);
LRESULT CALLBACK mainWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

HWND getMainWindow(HWND hWnd);
void setStoredValue(const TCHAR* name, int value);
int getStoredValue(const TCHAR* name, int defValue);
TCHAR* getStoredString(const TCHAR* name, const TCHAR* defValue);
int CALLBACK cbEnumTabStopChildren(HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
int detectCodePage(const char* data, int len);
TCHAR detectDelimiter(const TCHAR* data, BOOL skipComments);
void setClipboardText(const TCHAR* text);
BOOL isEOL(TCHAR c);
BOOL isNumber(TCHAR* val);
BOOL isUtf8(const char* string);
int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords);
BOOL hasString(const TCHAR* str, const TCHAR* sub, BOOL isCaseSensitive);
TCHAR* extractUrl(TCHAR* data);
void mergeSort(int indexes[], void* data, int l, int r, BOOL isBackward, BOOL isNums);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, const int cchTextMax);
void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState);

int detectCodePage(const char* data, int len);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);