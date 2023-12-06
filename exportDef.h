#pragma once

#define USR_API extern "C" __declspec(dllexport)

USR_API HWND  ListLoad(HWND hListerWnd, const char* cpszFileToLoad, int nShowFlags);
USR_API HWND  ListLoadW(HWND hListerWnd, const TCHAR* cszFileToLoad, int nShowFlags);