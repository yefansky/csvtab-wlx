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
	int nSize;
	DWORD dwPluginInterfaceVersionLow;
	DWORD dwPluginInterfaceVersionHi;
	char szDefaultIniName[MAX_PATH];
} ListDefaultParamStruct;

void __stdcall ListCloseWindow(HWND hWnd);
LRESULT CALLBACK mainWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

HWND getMainWindow(HWND hWnd);
void setStoredValue(const TCHAR* cpszName, int nValue);
int getStoredValue(const TCHAR* cpszName, int nDefValue);
TCHAR* getStoredString(const TCHAR* cpszName, const TCHAR* cpszDefValue);
int CALLBACK cbEnumTabStopChildren(HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* cpszIn);
char* utf16to8(const TCHAR* cpszIn);
int detectCodePage(const char* cpszData, int nLen);
TCHAR detectDelimiter(const TCHAR* cpszData, BOOL bSkipComments);
void setClipboardText(const TCHAR* cpszText);
BOOL isEOL(TCHAR c);
BOOL isNumber(TCHAR* pszVal);
BOOL isUtf8(const char* cpszString);
int findString(TCHAR* pszText, TCHAR* pszWord, BOOL bIsMatchCase, BOOL bIsWholeWords);
BOOL hasString(const TCHAR* cpszStr, const TCHAR* cpszSub, BOOL bIsCaseSensitive);
TCHAR* extractUrl(TCHAR* pszData);
void mergeSort(int nIndexes[], void* pvData, int nL, int nR, BOOL bisBackward, BOOL bisNums);
int ListView_AddColumn(HWND hListWnd, TCHAR* pszColName, int nFmt);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, const int cnCchTextMax);
void Menu_SetItemState(HMENU hMenu, UINT wID, UINT uState);

int detectCodePage(const char* pszData, int nLen);
TCHAR* utf8to16(const char* cpszIn);
char* utf16to8(const TCHAR* cpszIn);