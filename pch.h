// pch.h: This is a precompiled header file.
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

#include <tchar.h>

// add headers that you want to pre-compile here
#include "framework.h"
#include <shellapi.h>
#include <CommCtrl.h>
#include <uxtheme.h>

#define USR_API extern "C" __declspec(dllexport)

#endif //PCH_H
