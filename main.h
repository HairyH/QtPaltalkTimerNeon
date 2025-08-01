/*********************************/
/* (c) 2020 Hairy Soft Solutions */
/*********************************/
#pragma once

// Including SDKDDKVer.h defines the highest available Windows platform.

// If you wish to build your application for a previous Windows platform, include WinSDKVer.h and
// set the _WIN32_WINNT macro to the platform you wish to support before including SDKDDKVer.h.

#define _WIN32_WINNT _WIN32_WINNT_VISTA
#include <SDKDDKVer.h>

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <CommCtrl.h>
#include <stdio.h>
#include <iostream>
// C RunTime Header Files
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>
#include <Richedit.h>
#include <string>
#include <thread>

// UIAutomation
#include <UIAutomation.h>
#include <ole2.h>
#include <atlbase.h> // For CComPtr
#include <cctype> // For std::iswlower and std::iswupper
#include <iterator> // Add this include for std::begin and std::end
#include <cwctype> // Add this include for std::iswlower and std::iswupper


// Resource Header
#include "resource1.h"

// New version Common Controls, for flat look

#pragma comment(lib,"comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "uiautomationcore.lib")
#pragma comment(linker,"\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")


// TODO: Change these to the App name 
char szAppName[] = "QtPaltalk Timer";
wchar_t wcAppName[] = L"QtPaltalk Timer";

#define msgba(h,x) MessageBoxA(h,x,szAppName,MB_OK)
#define msgbw(h,x) MessageBoxW(h,x,wcAppName,MB_OK)
//Quick Messagebox macro
#define msga(x) msgba(ghMain,x)
// Quick Messagebox WIDE String
#define msgw(x) msgbw(ghMain,x)


#define IDT_MICTIMER	5555
#define IDT_MONITORTIMER	6666
#define IDM_BEEP		5556
// Writes  strings (char*) to debug window
void OutputErrorDebugStringA(char* szError) {
	int iLen = strlen(szError);
	if (iLen > MAX_PATH) return;
	char szOut[MAX_PATH] = { 0 };
	sprintf_s(szOut, MAX_PATH, "Error: %s \n", szError);
	OutputDebugStringA(szOut);
}
// Writes wide Strings (wchar_t*) to debug window
void OutputErrorDebugStringW(wchar_t* wcError) {
	int iLen = wcslen(wcError);
	if (iLen > MAX_PATH) return;
	wchar_t wcOut[MAX_PATH] = { 0 };
	swprintf_s(wcOut, MAX_PATH, L"Error: %s \n", wcError);
}