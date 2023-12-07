// stdafx.h: This is a precompiled header file.
// Files listed below are compiled only once, improving build performance for future builds.
// This also affects IntelliSense performance, including code completion and many code browsing features.
// However, files listed here are ALL re-compiled if any one of them is updated between builds.
// Do not add files here that you will be updating frequently as this negates the performance advantage.

#ifndef PCH_H
#define PCH_H

#define _CRT_SECURE_NO_WARNINGS

#ifndef UNICODE
#define UNICODE
#endif

#ifndef _UNICODE
#define _UNICODE
#endif

#include <assert.h>

#include <tchar.h>

// add headers that you want to pre-compile here
#include "framework.h"
#include <shellapi.h>
#include <CommCtrl.h>
#include <uxtheme.h>

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <uxtheme.h>
#include <locale.h>
#include <tchar.h>
#include <stdlib.h>
#include <stdio.h>
#include <ctype.h>
#include <math.h>

#include "globalDef.h"

#define KG_PROCESS_ERROR(C) do {if (!(C)) goto Exit0; } while(false);
#define KG_PROCESS_ERROR_RET_CODE(C, ret) do {if (!(C)) {nResult = ret; goto Exit0; }} while(false);


#define KG_PROCESS_SUCCESS(C) do {if ((C)) goto Exit1; } while(false);

#endif //PCH_H
