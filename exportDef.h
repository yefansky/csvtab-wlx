#pragma once

#define USR_API extern "C" __declspec(dllexport)

USR_API HWND  ListLoad(HWND hListerWnd, const char* fileToLoad, int showFlags);
USR_API HWND  ListLoadW(HWND hListerWnd, const TCHAR* fileToLoad, int showFlags);