
/*********************************/
/* (c) HairyH 2025               */
/*********************************/
#include "main.h"

using namespace std;
#undef DEBUG
//#define DEBUG

// Global Variables
HINSTANCE	hInst = 0;
HWND		ghMain = 0;
HWND        ghPtMain = 0;
HWND		ghEdit = 0;
HWND		ghInterval = 0;
HWND		ghLimit = 0;
HWND		ghClock = 0;
HWND		ghCbSend = 0;
// Debug related globals
char gszDebugMsg[MAX_PATH] = { '\0' };
wchar_t gwcDebugMsg[MAX_PATH] = { '\0' };
// Paltalk Handles
HWND ghPtRoom = NULL, ghPtLv = NULL;

// Font handles
HFONT ghFntClk = NULL;

// Timer related globals
BOOL gbMonitor = FALSE;
BOOL gbSendTxt = TRUE;
BOOL gbSendLimit = FALSE;
BOOL gbSendBold = TRUE;
BOOL gbBeep = FALSE;
BOOL gbAutoDot = FALSE;

int giMicTimerSeconds = 0;
int giInterval = 30;
int giLimit = 120;
// Nick related globals
char gszSavedNick[MAX_PATH] = { '0' };
char gszCurrentNick[MAX_PATH] = { '0' };
char gszDotMicUser[MAX_PATH] = { '0' };
wchar_t gwcRoomTitle[MAX_PATH] = { '0' };
int iMaxNicks = 0;
int iDrp = 0;
// UIAutomation related globals
IUIAutomationElement* emojiTextEditElement = nullptr;
IUIAutomationElement* automationElementRoom = nullptr;
CComPtr<IUIAutomation> automation;

//Quick Messagebox macro
#define msga(x) msgba(ghMain,x)
// Quick Messagebox WIDE String
#define msgw(x) msgbw(ghMain,x)

// Function prototypes
BOOL CALLBACK DlgMain(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
void GetPaltalkWindows(void);
BOOL CALLBACK EnumPaltalkWindows(HWND hWnd, LPARAM lParam);
BOOL InitClockDis(void);
BOOL InitIntervals(void);
BOOL InitMicLimits(void);
BOOL GetMicUser(void);
void CreateContextMenu(WPARAM wParam, LPARAM lparam);
void StartStopMonitoring(void);
void MicTimerStart(void);
void MicTimerReset(void);
void MicTimerTick(void);
void MonitorTimerTick(void);
wstring ConvertToBold(const wstring& inString);
void CopyPaste2Paltalk(char* szMsg);
// UIAutomation and Dot Mic User related functions
HRESULT __stdcall InitUIAutomation(void);
HRESULT __stdcall GetUIAutomationElementFromHWNDAndClassName(HWND hwnd, const wchar_t* className, IUIAutomationElement** foundElement);
HRESULT __stdcall FindWindowByTitle(const std::wstring& title, IUIAutomationElement** outElement);
void __stdcall DotMicUser(char* szMicUser);
static void SimulateRightClick(int x, int y);
static void SimulateLeftClick(int x, int y);
static void SendEnterTwice();
HRESULT __stdcall DotAndUnDotMicUser(char* szMicUser);

/// Main entry point of the app 
int APIENTRY WinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPSTR lpCmdLine, _In_ int nShowCmd)
{
	hInst = hInstance;
	InitCommonControls();
	LoadLibraryW(L"riched20.dll"); // comment if richedit is not used
	// TODO: Add any initiations as needed
	return DialogBox(hInst, MAKEINTRESOURCE(IDD_DLGMAIN), NULL, (DLGPROC)DlgMain);
}

/// Callback main message loop for the Application
BOOL CALLBACK DlgMain(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
	{
		ghMain = hwndDlg;
		ghInterval = GetDlgItem(hwndDlg, IDC_COMBO_INTERVAL);
		ghClock = GetDlgItem(hwndDlg, IDC_EDIT_CLOCK);
		ghCbSend = GetDlgItem(hwndDlg, IDC_CHECK1);
		ghLimit = GetDlgItem(hwndDlg, IDC_COMBO_MCLIMIT);
		InitClockDis();
		InitIntervals();
		InitMicLimits();
		SendDlgItemMessageW(hwndDlg, IDC_CHECK1, BM_SETCHECK, (WPARAM)BST_CHECKED, 0);
		SendDlgItemMessageW(hwndDlg, IDC_CHECK_BOLD, BM_SETCHECK, (WPARAM)BST_CHECKED, 0);

	}
	break; // return TRUE;
	case WM_CLOSE:
	{
		if (emojiTextEditElement)
			emojiTextEditElement->Release();
		CoUninitialize();
		if (ghFntClk)
			DeleteObject(ghFntClk);
		EndDialog(hwndDlg, 0);
	}
	return TRUE;
	case WM_CONTEXTMENU:
	{
		CreateContextMenu(wParam, lParam);
	}
	return TRUE;
	case WM_COMMAND:
	{
		switch (LOWORD(wParam))
		{
		case IDCANCEL:
		{
			EndDialog(hwndDlg, 0);
		}
		return TRUE;
		case IDOK:
		{
			GetPaltalkWindows();
		}
		return TRUE;
		case IDC_BUTTON_START:
		{
			StartStopMonitoring();
		}
		return TRUE;
		case IDC_CHECK1: // Send text to Paltalk
		{
			gbSendTxt = IsDlgButtonChecked(ghMain, IDC_CHECK1);
		}
		return TRUE;
		case IDC_CHECK2: // Send Mic time limit to Paltalk
		{
			gbSendLimit = IsDlgButtonChecked(ghMain, IDC_CHECK2);
		}
		return TRUE;
		case IDC_CHECK_BOLD: // Send Bold text to Paltalk
		{
			gbSendBold = IsDlgButtonChecked(ghMain, IDC_CHECK_BOLD);
		}
		return TRUE;
		case IDM_BEEP: // enable/disable beeps
		{
			if (gbBeep)
			{
				CheckMenuItem(GetMenu(ghMain), IDM_BEEP, MF_UNCHECKED);
				gbBeep = FALSE;
			}
			else
			{
				CheckMenuItem(GetMenu(ghMain), IDM_BEEP, MF_CHECKED);
				gbBeep = TRUE;
			}
		}
		return TRUE;
		case IDC_CHECKDOT: // Set Auto Dotting 
		{
			gbAutoDot = IsDlgButtonChecked(ghMain, IDC_CHECKDOT);
		}
		return TRUE;
		case IDC_COMBO_INTERVAL:
		{
			if (HIWORD(wParam) == CBN_SELCHANGE)
			{
				int iSel = SendMessage(ghInterval, CB_GETCURSEL, 0, 0);
				giInterval = SendMessage(ghInterval, CB_GETITEMDATA, (WPARAM)iSel, 0);
				char szTemp[100] = { 0 };
				sprintf_s(szTemp, "giInterval = %d \n", giInterval);
				OutputDebugStringA(szTemp);
			}
		}
		return TRUE;
		case IDC_COMBO_MCLIMIT:
		{
			if (HIWORD(wParam) == CBN_SELCHANGE)
			{
				int iSel = SendMessage(ghLimit, CB_GETCURSEL, 0, 0);
				giLimit = SendMessage(ghLimit, CB_GETITEMDATA, (WPARAM)iSel, 0);
				char szTemp[100] = { 0 };
				sprintf_s(szTemp, "giLimit = %d \n", giLimit);
				OutputDebugStringA(szTemp);
			}
		}
		return TRUE;
		}
	}
	case WM_TIMER:
	{
		if (wParam == IDT_MICTIMER)
		{
			MicTimerTick();
		}
		else if (wParam == IDT_MONITORTIMER)
		{
			MonitorTimerTick();
		}
	}

	default:
		return FALSE;
	}
	return FALSE;
}

/// Initialise Paltalk Control Handles
void GetPaltalkWindows(void)
{
	char szWinText[200] = { '0' };
	//wchar_t wcTitle[256] = { '0' };

	// Paltalk chat room window
	ghPtRoom = NULL; ghPtLv = NULL;

	// Getting the chat room window handle
	ghPtRoom = FindWindowW(L"DlgGroupChat Window Class", 0); // this is to find nick list
	if (GetWindowTextW(ghPtRoom, gwcRoomTitle, 254) < 1)
	{
		msga("No Paltalk Room Found!");
		return;
	}
	// Getting the main Paltalk window handle
	ghPtMain = FindWindowW(L"Qt5150QWindowIcon", gwcRoomTitle);  // this is to send text 
	if (ghPtMain == NULL)
	{
		msga("No Paltalk Main Window Found!");
		return;
	}
	// Initialise UIAutomation
	if (FAILED(InitUIAutomation())) {
		msga("Initializing UI Automation failed!");
	 }
	// Getting the Emoji Text Edit control UIAutomation element to send text to Paltalk
	HRESULT hr = GetUIAutomationElementFromHWNDAndClassName(ghPtMain, L"ui::controls::EmojiTextEdit", &emojiTextEditElement);
	if (FAILED(hr)) {
		std::cerr << "GetUIAutomationElementFromHWNDAndClassName failed: " << std::hex << hr << std::endl;
	}
	// Get the Chat Room windows handle
	if (ghPtRoom) 
	{   // we got it
		// Get the current thread ID and the target window's thread ID
		DWORD targetThreadId = GetWindowThreadProcessId(ghPtMain, 0);
		DWORD currentThreadId = GetCurrentThreadId();
		// Attach the input of the current thread to the target window's thread
		AttachThreadInput(currentThreadId, targetThreadId, true);
		// Bring the window to the foreground
		bool bRet = SetForegroundWindow(ghPtMain);
		// Give some time for the window to come into focus
		Sleep(500);
		// Detach input once done
		AttachThreadInput(currentThreadId, targetThreadId, false);
		// Make the Paltalk Room window the HWND_TOPMOST 
		SetWindowPos(ghPtMain, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);

		GetWindowTextA(ghPtRoom, szWinText, 200);
		// Set the title text to indicate which room we are timing
		wprintf_s(gwcRoomTitle, MAX_PATH, L"Timing: %s", szWinText);
		SendDlgItemMessageW(ghMain, IDC_STATIC_TITLE, WM_SETTEXT, (WPARAM)0, (LPARAM)gwcRoomTitle);
		// Finding the chat room window controls handles
		EnumChildWindows(ghPtRoom, EnumPaltalkWindows, 0);

	}
	else
	{
		msga("No Paltalk Window Found!");
	}

}

/// Enumeration Callback to Find the Control Windows
BOOL CALLBACK EnumPaltalkWindows(HWND hWnd, LPARAM lParam)
{
	char szListViewClass[] = "SysHeader32";
	char szMsg[MAX_PATH] = { 0 };
	char szClassNameBuffer[MAX_PATH] = { 0 };

	GetClassNameA(hWnd, szClassNameBuffer, MAX_PATH);

#ifdef DEBUG
	LONG lgIdCnt;
	lgIdCnt = GetWindowLongW(hWnd, GWL_ID);
	wsprintfA(szMsg, "windows class name: %s Control ID: %d \n", szClassNameBuffer, lgIdCnt);
	OutputDebugStringA(szMsg);
#endif // DEBUG

	if (strcmp(szListViewClass, szClassNameBuffer) == 0)
	{
		ghPtLv = hWnd;
#ifdef DEBUG
		sprintf_s(szMsg, "List View window handle: %p \n", ghPtLv);
		OutputDebugStringA(szMsg);
#endif // DEBUG
		return FALSE;
	}

	return TRUE;
}

/// Initialise Intervals Combo
BOOL InitIntervals(void)
{
	wchar_t wsTemp[256] = { '\0' };
	int iIndex = 0;
	int iSec = 0;
	int iMin = 0;

	for (int i = 30; i <= 180; i = i + 30)
	{
		iMin = i / 60;
		iSec = i % 60;
		swprintf_s(wsTemp, L"%d:%02d", iMin, iSec);
		iIndex = SendMessageW(ghInterval, CB_ADDSTRING, (WPARAM)0, (LPARAM)wsTemp);
		SendMessageW(ghInterval, CB_SETITEMDATA, (WPARAM)iIndex, (LPARAM)i);
	}

	SendMessageW(ghInterval, CB_SETCURSEL, (WPARAM)0, (LPARAM)0);
	giInterval = SendMessageW(ghInterval, CB_GETITEMDATA, (WPARAM)0, (LPARAM)0);
	return TRUE;
}

/// Initialise Mic time Limit selector Combo
BOOL InitMicLimits(void)
{
	wchar_t wsTemp[256] = { '\0' };
	int iIndex = 0;
	int iSec = 0;
	int iMin = 0;

	for (int i = 60; i <= 600; i = i + 30)
	{
		iMin = i / 60;
		iSec = i % 60;
		swprintf_s(wsTemp, L"%d:%02d", iMin, iSec);
		iIndex = SendMessageW(ghLimit, CB_ADDSTRING, (WPARAM)0, (LPARAM)wsTemp);
		giLimit = SendMessageW(ghLimit, CB_SETITEMDATA, (WPARAM)iIndex, (LPARAM)i);
	}

	SendMessageW(ghLimit, CB_SETCURSEL, (WPARAM)2, (LPARAM)2);
	giLimit = SendMessageW(ghLimit, CB_GETITEMDATA, (WPARAM)2, (LPARAM)0);
	return TRUE;

}

/// Initialise the Clock Display Window
BOOL InitClockDis(void)
{
	HDC hDC;
	int nHeight = 0;

	hDC = GetDC(ghMain);
	nHeight = -MulDiv(20, GetDeviceCaps(hDC, LOGPIXELSY), 72);
	ghFntClk = CreateFont(nHeight, 0, 0, 0, FW_BOLD, 0, 0, 0, 0, 0, 0, 0, 0, TEXT("Courier New"));
	SendMessageA(ghClock, WM_SETFONT, (WPARAM)ghFntClk, (LPARAM)TRUE);
	SendMessageA(ghClock, WM_SETTEXT, (WPARAM)0, (LPARAM)"00:00");

	return TRUE;
}

/// Starts the Mic timer
void MicTimerStart(void)
{
	SendMessageA(ghClock, WM_SETTEXT, 0, (LPARAM)"00:00");
	giMicTimerSeconds = 0;
	SetTimer(ghMain, IDT_MICTIMER, 1000, 0);
}

/// Reset the Timer
void MicTimerReset(void)
{
	KillTimer(ghMain, IDT_MICTIMER);
	giMicTimerSeconds = 0;
	SendMessageA(ghClock, WM_SETTEXT, 0, (LPARAM)"00:00");
}

/// Thi is called every 1sec. 
void MicTimerTick(void)
{
	char szClock[MAX_PATH] = { '\0' };
	char szMsg[MAX_PATH] = { '\0' };
	int iX = 60;
	int iMin;
	int iSec;

	giMicTimerSeconds++;

	iMin = giMicTimerSeconds / iX;
	iSec = giMicTimerSeconds % iX;
	// Refreshing the Clock display
	sprintf_s(szClock, MAX_PATH, "%02d:%02d", iMin, iSec);
	SendMessageA(ghClock, WM_SETTEXT, 0, (LPARAM)szClock);

	// Mic time limit
	if (giMicTimerSeconds == giLimit)
	{
		if (gbBeep)
		{
			//MessageBeep(MB_ICONEXCLAMATION);
			Beep(440, 1000);
		}
		// if mic time limit sending is enabled 
		if (gbSendLimit)
		{
			sprintf_s(szMsg, MAX_PATH, "%s on Mic for: %d:%02d min. TIME LIMIT!", gszCurrentNick, iMin, iSec);
			CopyPaste2Paltalk(szMsg);
		}
		else // just send the time on mic message 
		{
			sprintf_s(szMsg, MAX_PATH, "%s on Mic for: %d:%02d min.", gszCurrentNick, iMin, iSec);
			CopyPaste2Paltalk(szMsg);
		}
	}
	// if mic time = limit + 30 seconds
	else if (giMicTimerSeconds == giLimit + 30)
	{
		if (gbBeep)
		{
			//MessageBeep(MB_ICONEXCLAMATION);
			Beep(880, 1000);
		}
		if (gbSendLimit && gbAutoDot)
		{
			// Sending to the room auto dot is coming 
			sprintf_s(szMsg, MAX_PATH, "%s on Mic for: %d:%02d min. TIME LIMIT AUTO DOT!", gszCurrentNick, iMin, iSec);
			CopyPaste2Paltalk(szMsg);
			// Save the curret nick for dotting 
			sprintf_s(gszDotMicUser, MAX_PATH, "%s", gszCurrentNick);
			// Reset the timer
			MicTimerReset(); // Reset the mic timer
			// Reset saved nick
			sprintf_s(gszSavedNick, "a"); // Reset the saved nick
			//DotMicUser(gszDotMicUser); // Dot the mic user
			// Test New
			DotAndUnDotMicUser("MaxTimer");
			// Reset dotted mic user
			sprintf_s(gszDotMicUser, MAX_PATH, "a");
			
			return; // No need to send text every second
		}
		else
		{
			sprintf_s(szMsg, MAX_PATH, "%s on Mic for: %d:%02d min.", gszCurrentNick, iMin, iSec);
			CopyPaste2Paltalk(szMsg);
		}

	}

	// Work out when to send text to Paltalk
	else if (giMicTimerSeconds % giInterval == 0)
	{
		sprintf_s(szMsg, MAX_PATH, "%s on Mic for: %d:%02d min.", gszCurrentNick, iMin, iSec);
		CopyPaste2Paltalk(szMsg);
	}
}


void MonitorTimerTick(void)
{
	GetMicUser();
	// Failed to get current nick
	if (strlen(gszCurrentNick) < 2 && strlen(gszSavedNick) < 2) return;
	// no change keep going
	if (strcmp(gszCurrentNick, gszSavedNick) == 0)
	{
		iDrp = 0; // Reset dropout counter
	}
	// No nick but there was one before
	else if (strlen(gszCurrentNick) < 2 && strlen(gszSavedNick) > 2)
	{
		iDrp++; //to tolerate mic dropout

		char szTemp[MAX_PATH] = { '\0' };
		sprintf_s(szTemp, "Mic dropout count %d \n", iDrp);
		OutputDebugStringA(szTemp);

		if (iDrp > 4) // 5 second dropout
		{
			MicTimerReset(); //Stop the mic timing
			sprintf_s(gszSavedNick, "a"); // Reset the saved nick
			sprintf_s(szTemp, "5 dropouts: %d Reset Mic timer \n", iDrp);
			OutputDebugStringA(szTemp);
			iDrp = 0;
		}
	}
	// New nick on mic -> Start Mic Timer
	else if (strcmp(gszCurrentNick, gszSavedNick) != 0)
	{
		SYSTEMTIME sytUtc;
		char szMsg[MAX_PATH] = { '\0' };

		MicTimerStart();

		GetSystemTime(&sytUtc);
		int iLimitMin = giLimit / 60;
		int iLimitSec = giLimit % 60;
		sprintf_s(szMsg, "Start: %s on a %d:%02d min. Mic Session", gszCurrentNick, iLimitMin, iLimitSec);
		CopyPaste2Paltalk(szMsg);

		sprintf_s(szMsg, "Start: %s at: %02d:%02d:%02d UTC\n", gszCurrentNick, sytUtc.wHour, sytUtc.wMinute, sytUtc.wSecond);
		OutputDebugStringA(szMsg);

		strcpy_s(gszSavedNick, gszCurrentNick);
	}
}

/// Get the Mic user
BOOL GetMicUser(void)
{
	if (!ghPtLv) return FALSE;

	char szOut[MAX_PATH] = { '\0' };
	char szMsg[MAX_PATH] = { '\0' };
	wchar_t wsMsg[MAX_PATH] = { '\0' };
	BOOL bRet = FALSE;

	const int iSizeOfItemNameBuff = MAX_PATH * sizeof(char); //wchar_t
	LPSTR pXItemNameBuff = NULL;

	LVITEMA lviNick = { 0 };
	DWORD dwProcId;
	VOID* vpMemLvi;
	HANDLE hProc;
	int iNicks;
	int iImg = 0;

	LVITEMA lviRead = { 0 };

	GetWindowThreadProcessId(ghPtLv, &dwProcId);
	hProc = OpenProcess(
		PROCESS_VM_OPERATION |
		PROCESS_VM_READ |
		PROCESS_VM_WRITE, FALSE, dwProcId);

	if (hProc == NULL)return FALSE;

	vpMemLvi = VirtualAllocEx(hProc, NULL, sizeof(LVITEMA),
		MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
	if (vpMemLvi == NULL) return FALSE;

	pXItemNameBuff = (LPSTR)VirtualAllocEx(hProc, NULL, iSizeOfItemNameBuff,
		MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
	if (pXItemNameBuff == NULL) return FALSE;

	iNicks = ListView_GetItemCount(ghPtLv);
	if (iMaxNicks < iNicks)
	{
		iMaxNicks = iNicks;
		char szTemp[100] = { 0 };
		sprintf_s(szTemp, "%d", iMaxNicks);
		SendDlgItemMessageA(ghMain, IDC_EDIT2, WM_SETTEXT, (WPARAM)0, (LPARAM)szTemp);
	}
#ifdef DEBUG
	sprintf_s(szMsg, "Number of nicks: %d \n", iNicks);
	OutputDebugStringA(szMsg);
#endif // DEBUG
	for (int i = 0; i < iNicks; i++)
	{
		lviNick.mask = LVIF_IMAGE | LVIF_TEXT;
		lviNick.pszText = pXItemNameBuff;
		lviNick.cchTextMax = MAX_PATH - 1;
		lviNick.iItem = i;
		lviNick.iSubItem = 0;

		WriteProcessMemory(hProc, vpMemLvi, &lviNick, sizeof(lviNick), NULL);

		SendMessage(ghPtLv, 4171, (WPARAM)i, (LPARAM)vpMemLvi); // 4171

		ReadProcessMemory(hProc, vpMemLvi, &lviRead, sizeof(lviRead), NULL);

		iImg = lviRead.iImage;
		//sprintf_s(szMsg, "Item index: %d iImg: %d \n", i, iImg);
		//OutputDebugStringA(szMsg);

		if (iImg == 10)
		{
			SendMessage(ghPtLv, 4141, (WPARAM)i, (LPARAM)vpMemLvi);    // 4141

			ReadProcessMemory(hProc, lviRead.pszText, &gszCurrentNick, sizeof(gszCurrentNick), NULL);
			bRet = TRUE;
			break;
		}
		else
		{
			sprintf_s(gszCurrentNick, "a");
			bRet = FALSE;
		}

	}
#ifdef DEBUG
	sprintf_s(szMsg, "Image: %d Nickname: %s \n", iImg, gszCurrentNick);
	OutputDebugStringA(szMsg);
#endif // DEBUG
	// Cleanup 
	if (vpMemLvi != NULL) VirtualFreeEx(hProc, vpMemLvi, 0, MEM_RELEASE);
	if (pXItemNameBuff != NULL) VirtualFreeEx(hProc, pXItemNameBuff, 0, MEM_RELEASE);
	CloseHandle(hProc);

	return bRet;
}

void CreateContextMenu(WPARAM wParam, LPARAM lparam)
{
	HMENU hMenu = CreatePopupMenu();
	POINT pt;
	GetCursorPos(&pt);
	if (gbBeep)
		InsertMenuA(hMenu, 0, MF_BYCOMMAND | MF_STRING | MF_CHECKED, IDM_BEEP, "Beep");
	else
		InsertMenuA(hMenu, 0, MF_BYCOMMAND | MF_STRING | MF_UNCHECKED, IDM_BEEP, "Beep");

	TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_TOPALIGN, pt.x, pt.y, 0, ghMain, NULL);
}


/// Start Monitoring the mic que
void StartStopMonitoring(void)
{
	if (!ghPtRoom)
	{
		msga("Error: No Paltalk Room!\n [Get Pt] first!");
		return;
	}
	if (!gbMonitor)
	{
		SetTimer(ghMain, IDT_MONITORTIMER, 1000, NULL);
		gbMonitor = TRUE;
		SendDlgItemMessageW(ghMain, IDC_BUTTON_START, WM_SETTEXT, 0, (LPARAM)L"Stop Timing");
	}
	else
	{
		KillTimer(ghMain, IDT_MONITORTIMER);
		sprintf_s(gszSavedNick, "a");
		sprintf_s(gszCurrentNick, "a");
		MicTimerReset();
		gbMonitor = FALSE;
		SendDlgItemMessageW(ghMain, IDC_BUTTON_START, WM_SETTEXT, 0, (LPARAM)L"Start Timing");
	}
}

/// Sending Text to Paltalk 
void CopyPaste2Paltalk(char* szMsg)
{
	char szOut[MAX_PATH] = { '\0' };
	wchar_t wcOut[MAX_PATH] = { '\0' };
	HRESULT hr = S_OK;

	if (strlen(gszCurrentNick) < 2) return;
	else if (!gbSendTxt) return; // if text sending is disabled

	CComPtr<IUIAutomationLegacyIAccessiblePattern> pattern;
	hr = emojiTextEditElement->GetCurrentPatternAs(UIA_LegacyIAccessiblePatternId, IID_IUIAutomationLegacyIAccessiblePattern, (void**)&pattern);
	//GetCurrentPatternAs(UIA_ValuePatternId, IID_PPV_ARGS(&pattern));
	if (SUCCEEDED(hr)) {
		BSTR bstrOut = NULL;

		sprintf_s(szOut, MAX_PATH, "*** %s ***", szMsg);
		MultiByteToWideChar(CP_ACP, 0, szOut, -1, wcOut, MAX_PATH);
		if (gbSendBold) {
			wstring wstrOutBold = ConvertToBold(wcOut);
			bstrOut = SysAllocString(wstrOutBold.c_str());
		}
		else {
			bstrOut = SysAllocString(wcOut);
		}
		emojiTextEditElement->SetFocus();
		pattern->SetValue(bstrOut);
		SendMessageA(ghPtMain, WM_KEYDOWN, (WPARAM)VK_RETURN, 0);
		SendMessageA(ghPtMain, WM_KEYUP, (WPARAM)VK_RETURN, 0);
		SysFreeString(bstrOut);
	}
	else {
		std::cerr << "GetCurrentPatternAs failed: " << std::hex << hr << std::endl;
	}

}

/// Convert the string to bold
wstring ConvertToBold(const wstring& inString) {
	const int boldLowercaseOffset = 119737; // 'a' to bold 'a'
	const int boldUppercaseOffset = 119743; // 'A' to bold 'A'
	const int boldDigitOffset = 120734;      // '0' to bold '0'

	wstring result;

	for (wchar_t c : inString) {
		if (iswlower(c)) { // Lowercase letters
			DWORD codePoint = c + boldLowercaseOffset;
			if (codePoint <= 0xFFFF) {
				result += static_cast<wchar_t>(codePoint);
			}
			else {
				// Handle surrogate pairs for Unicode characters above U+FFFF
				wchar_t highSurrogate = (wchar_t)(0xD800 + (codePoint - 0x10000) / 0x400);
				wchar_t lowSurrogate = (wchar_t)(0xDC00 + (codePoint - 0x10000) % 0x400);
				result += highSurrogate;
				result += lowSurrogate;
			}
		}
		else if (iswupper(c)) { // Uppercase letters
			DWORD codePoint = c + boldUppercaseOffset;
			if (codePoint <= 0xFFFF) {
				result += static_cast<wchar_t>(codePoint);
			}
			else {
				wchar_t highSurrogate = (wchar_t)(0xD800 + (codePoint - 0x10000) / 0x400);
				wchar_t lowSurrogate = (wchar_t)(0xDC00 + (codePoint - 0x10000) % 0x400);
				result += highSurrogate;
				result += lowSurrogate;
			}
		}
		else if (iswdigit(c)) { // Digits
			DWORD codePoint = c + boldDigitOffset;
			if (codePoint <= 0xFFFF) {
				result += static_cast<wchar_t>(codePoint);
			}
			else {
				wchar_t highSurrogate = (wchar_t)(0xD800 + (codePoint - 0x10000) / 0x400);
				wchar_t lowSurrogate = (wchar_t)(0xDC00 + (codePoint - 0x10000) % 0x400);
				result += highSurrogate;
				result += lowSurrogate;
			}
		}
		else {
			result += c;
		}
	}

	return result;
}

/// Get the UIAutomation element from HWND and ClassName
HRESULT __stdcall GetUIAutomationElementFromHWNDAndClassName(HWND hwnd, const wchar_t* className, IUIAutomationElement** foundElement) {
	HRESULT hr = S_OK;

	CComPtr<IUIAutomationElement> elementRoom;
	hr = automation->ElementFromHandle(hwnd, &elementRoom);
	if (FAILED(hr)) {
		char szDebug[] = "ElementFromHandle failed: elementRoom\n";
		OutputDebugStringA(szDebug);
		return hr;
	}

	CComPtr<IUIAutomationCondition> classNameCondition;
	CComVariant classNameVariant(className);
	hr = automation->CreatePropertyCondition(UIA_ClassNamePropertyId, classNameVariant, &classNameCondition);
	if (FAILED(hr)) {
		char szDebug[] = "CreatePropertyCondition failed : className\n";
		OutputDebugStringA(szDebug);
		return hr;
	}

	hr = elementRoom->FindFirst(TreeScope_Subtree, classNameCondition, foundElement);
	if (FAILED(hr)) {
		char szDebug[] = "FindFirst failed: elementRoom\n";
		OutputDebugStringA(szDebug);
		return hr;
	}

	return S_OK;
}

// Find the Chat window by title
HRESULT __stdcall FindWindowByTitle(const std::wstring& title, IUIAutomationElement** outElement)
{
	HRESULT hr = S_OK;

	CComPtr<IUIAutomationElement> root = nullptr; // Desktop
	hr = automation->GetRootElement(&root);
	if (FAILED(hr)) {
		swprintf_s(gwcDebugMsg, MAX_PATH, L"Error getting root element");
		OutputDebugStringW(gwcDebugMsg);
		return hr;
	}

	CComPtr<IUIAutomationCondition> cond = nullptr;
	CComVariant NameVariant(title.c_str());
	hr = automation->CreatePropertyCondition(UIA_NamePropertyId, NameVariant, &cond);
		
	hr = root->FindFirst(TreeScope_Children, cond, outElement);
			
	return hr; 
}

// Auto Dotting after 30sec the time limit passed 
void __stdcall DotMicUser(char* szMicUser)
{
	char szMsg[MAX_PATH] = { '\0' };
	POINT ptCur = { 0 };
	POINT ptBaseButton = { 0 };

	IUIAutomationElement* roomEl = nullptr;
	std::wstring strRoomName(gwcRoomTitle);

	if (FAILED(FindWindowByTitle(strRoomName, &roomEl))) {
		swprintf_s(gwcDebugMsg, MAX_PATH, L"Room not found: %s \n", gwcRoomTitle);
		OutputDebugStringW(gwcDebugMsg);
		return;
	}

	//Get the Room window in focus
	roomEl->SetFocus();
	GetCursorPos(&ptCur);

	// Find the TalkingNowWidget element by class name
	VARIANT vClassName;
	vClassName.vt = VT_BSTR;
	vClassName.bstrVal = SysAllocString(L"ui::rooms::TalkingNowWidget"); // Replace with actual class name
	if (vClassName.bstrVal == nullptr) {
		//OutputErrorDebugStringW( L"Failed to allocate BSTR for class name." );
		roomEl->Release();
		return;
	}
	IUIAutomationElement* talkingNowEl = nullptr;
	IUIAutomationCondition* classCondition = nullptr;
	IUIAutomation* automation = nullptr;
	if (FAILED(CoCreateInstance(__uuidof(CUIAutomation), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&automation)))) {
		//OutputErrorDebugStringW(L"Failed to create UIAutomation instance.");
		SysFreeString(vClassName.bstrVal);
		roomEl->Release();
		return;
	}
	automation->CreatePropertyCondition(UIA_ClassNamePropertyId, vClassName, &classCondition);
	if (FAILED(roomEl->FindFirst(TreeScope_Descendants, classCondition, &talkingNowEl))) {
		//OutputErrorDebugStringW(L"TalkingNowWidget not found in room.") ;
		SysFreeString(vClassName.bstrVal);
		classCondition->Release();
		roomEl->Release();
		automation->Release();
		return;
	}
	SysFreeString(vClassName.bstrVal);
	classCondition->Release();

	// Right click the talkingNowEl element and send 2 ENTER to it 
	VARIANT vRect = { 0 };
	vRect.vt = VT_ARRAY | VT_R8; // The variant type is set for an array of doubles.

	talkingNowEl->GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, &vRect);
	SAFEARRAY* psa = nullptr;
	if (vRect.vt == (VT_ARRAY | VT_R8) && vRect.parray != nullptr) {
		psa = vRect.parray;
		double* pData = nullptr;
		SafeArrayAccessData(psa, (void**)&pData);

		if (pData != nullptr) {
			RECT r;
			r.left = static_cast<LONG>(pData[0]);
			r.top = static_cast<LONG>(pData[1]);
			r.right = static_cast<LONG>(pData[2]);
			r.bottom = static_cast<LONG>(pData[3]);

			int x = (r.left + (r.right / 2));
			int y = (r.top + (r.bottom / 2));

			//sprintf_s(szMsg, MAX_PATH, "X position = %d Y position = %d\n",x,y);
			//OutputDebugStringA(szMsg);

			roomEl->SetFocus();
			POINT p = { 0 };
			GetCursorPos(&p);

			SimulateRightClick(x, y);
			std::this_thread::sleep_for(std::chrono::milliseconds(500));
			SendEnterTwice();

		}

		SafeArrayUnaccessData(psa);
	}

	/***************************************************************************************************************/
	// Undoing the dot from the person we just dotted 
	std::this_thread::sleep_for(std::chrono::milliseconds(2000)); // Give time to dot 

	// Finding the mic queue title widget element
	IUIAutomationElement* micQueueTitleEl = nullptr;
	IUIAutomationCondition* classCondTitleItem = nullptr;
	VARIANT vTitleItemClass = { 0 };
	vTitleItemClass.vt = VT_BSTR;
	vTitleItemClass.bstrVal = SysAllocString(L"ui::rooms::member_list::MicQueueTitleItemWidget"); // Mic queue Title to search for nick
	if (vClassName.bstrVal == nullptr) {
		std::wcerr << L"Failed to allocate BSTR for class name. for TitleItemWidget" << std::endl;
	}
	// Setting up class name condition for TitleWidget
	if (FAILED(automation->CreatePropertyCondition(UIA_ClassNamePropertyId, vTitleItemClass, &classCondTitleItem))) {
		//OutputErrorDebugStringW(L"Error creating classCondTitleItem!");
		SysFreeString(vTitleItemClass.bstrVal);
		roomEl->Release();
		automation->Release();
		return;
	}
	// Getting the micQueueTitleEl element
	if (FAILED(roomEl->FindFirst(TreeScope_Descendants, classCondTitleItem, &micQueueTitleEl))) {
		//OutputErrorDebugStringW(L"Error finding Mic Queue Title Widget");
		SysFreeString(vTitleItemClass.bstrVal);
		classCondTitleItem->Release();
		roomEl->Release();
		automation->Release();
		return;
	}

	// if we got this far, we got the mic queue title element 
	if (classCondTitleItem) classCondTitleItem->Release();
	SysFreeString(vTitleItemClass.bstrVal);

	// Finding the Base Button element by class name
	IUIAutomationElement* baseButtonEl = nullptr;
	IUIAutomationCondition* condBaseButton = nullptr;

	VARIANT vbaseButtonClass;
	vbaseButtonClass.vt = VT_BSTR;
	vbaseButtonClass.bstrVal = SysAllocString(L"qtctrl::BaseButton");
	if (vbaseButtonClass.bstrVal == nullptr) {
		//OutputErrorDebugStringW(L"Failed to allocate BSTR for class name. BaseButton");
		if (micQueueTitleEl) micQueueTitleEl->Release();
		if (roomEl) roomEl->Release();
		if (automation) automation->Release();
		return;
	}

	// Setting up the condition for base button 
	if (FAILED(automation->CreatePropertyCondition(UIA_ClassNamePropertyId, vbaseButtonClass, &condBaseButton))) {
		//OutputErrorDebugStringW(L"Error failed to creat condition for BaseButton");
		SysFreeString(vbaseButtonClass.bstrVal);
		if (micQueueTitleEl) micQueueTitleEl->Release();
		if (roomEl) roomEl->Release();
		if (automation) automation->Release();
		return;
	}

	// Get the baseButton element
	if (FAILED(micQueueTitleEl->FindFirst(TreeScope_Descendants, condBaseButton, &baseButtonEl))) {
		//OutputErrorDebugStringW(L"Error finding BaseButton element");
		SysFreeString(vbaseButtonClass.bstrVal);
		condBaseButton->Release();
		roomEl->Release();
		automation->Release();
		return;
	}
	//Now we got the BaseButton element
	if (condBaseButton) condBaseButton->Release();
	SysFreeString(vbaseButtonClass.bstrVal);

	//Start finding the dotted person in the nick list
	baseButtonEl->GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, &vRect);
	SAFEARRAY* psa2 = nullptr;
	if (vRect.vt == (VT_ARRAY | VT_R8) && vRect.parray != nullptr) {
		psa2 = vRect.parray;
		double* pData = nullptr;
		SafeArrayAccessData(psa2, (void**)&pData);

		if (pData != nullptr) {
			RECT r;
			r.left = static_cast<LONG>(pData[0]);
			r.top = static_cast<LONG>(pData[1]);
			r.right = static_cast<LONG>(pData[2]);
			r.bottom = static_cast<LONG>(pData[3]);

			int x = (r.left + (r.right / 2));
			int y = (r.top + (r.bottom / 2));
			ptBaseButton.x = x;
			ptBaseButton.y = y;

			sprintf_s(szMsg, MAX_PATH, "X position = %d Y position = %d\n", x, y);
			OutputDebugStringA(szMsg);

			roomEl->SetFocus();

			// Activate the user list search box
			SimulateLeftClick(x, y);
			std::this_thread::sleep_for(std::chrono::milliseconds(300));
			// Send the dotted user name to the search box

			for (int i = 0; i < strlen(szMicUser); i++)
			{
				SendMessageA(ghPtMain, WM_CHAR, (WPARAM)szMicUser[i], 0);
			}

			// Wait to search 			
			std::this_thread::sleep_for(std::chrono::milliseconds(1000));

		}

		SafeArrayUnaccessData(psa2);
	}

	// Getting the ui::rooms::member_list::MemberItemWidget
	IUIAutomationElement* memberItemEl = nullptr;
	IUIAutomationCondition* condMemberItem = nullptr;
	VARIANT vcondMemberItem;
	vcondMemberItem.vt = VT_BSTR;
	vcondMemberItem.bstrVal = SysAllocString(L"ui::rooms::member_list::MemberItemWidget");

	if (vcondMemberItem.bstrVal == nullptr) {
		//OutputErrorDebugStringW(L"Failed to allocate BSTR for class name. vcondMemberItem");
		if (micQueueTitleEl) micQueueTitleEl->Release();
		if (baseButtonEl) baseButtonEl->Release();
		if (roomEl) roomEl->Release();
		if (automation) automation->Release();
		return;
	}

	// Setting up the condition for memberItemEl
	if (FAILED(automation->CreatePropertyCondition(UIA_ClassNamePropertyId, vcondMemberItem, &condMemberItem))) {
		//OutputErrorDebugStringW(L"Error failed to creat condition for FlatButton");
		SysFreeString(vcondMemberItem.bstrVal);
		if (micQueueTitleEl) micQueueTitleEl->Release();
		if (roomEl) roomEl->Release();
		if (automation) automation->Release();
		return;
	}

	// Get the memberItemEl element
	if (FAILED(roomEl->FindFirst(TreeScope_Descendants, condMemberItem, &memberItemEl))) {
		//OutputErrorDebugStringW(L"Error finding FlatButton element");
		SysFreeString(vcondMemberItem.bstrVal);
		if (baseButtonEl) baseButtonEl->Release();
		// if (condFlatButton) condFlatButton->Release();
		if (roomEl) roomEl->Release();
		if (automation) automation->Release();
		return;
	}
	//We got member Item, we can release the bstr 
	SysFreeString(vcondMemberItem.bstrVal);

	// Getting the bounding rect of member item.
	VARIANT vRect2 = { 0 };
	vRect.vt = VT_ARRAY | VT_R8; // The variant type is set for an array of doubles.
	memberItemEl->GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, &vRect2);
	SAFEARRAY* psa3 = nullptr;
	if (vRect2.vt == (VT_ARRAY | VT_R8) && vRect2.parray != nullptr) {
		psa3 = vRect2.parray;
		double* pData = nullptr;
		SafeArrayAccessData(psa3, (void**)&pData);

		if (pData != nullptr) {
			RECT r;
			r.left = static_cast<LONG>(pData[0]);
			r.top = static_cast<LONG>(pData[1]);
			r.right = static_cast<LONG>(pData[2]);
			r.bottom = static_cast<LONG>(pData[3]);

			int x = (r.left + (r.right / 2));
			int y = (r.top + (r.bottom / 2));

			sprintf_s(szMsg, MAX_PATH, "X: %d Y: %d   memberItemEl\n", x, y);
			OutputDebugStringA(szMsg);

			roomEl->SetFocus();
			POINT p = { 0 };
			GetCursorPos(&p);

			SimulateRightClick(x, y);
			std::this_thread::sleep_for(std::chrono::milliseconds(1000));
			SendEnterTwice();
			std::this_thread::sleep_for(std::chrono::milliseconds(1000));
			SimulateLeftClick(ptBaseButton.x, ptBaseButton.y);
			SetCursorPos(ptCur.x, ptCur.y);
		}

		SafeArrayUnaccessData(psa3);
	}

	// Cleanup the elements and automation
	if (memberItemEl) memberItemEl->Release();
	if (baseButtonEl) baseButtonEl->Release();
	if (micQueueTitleEl) micQueueTitleEl->Release();
	if (talkingNowEl) talkingNowEl->Release();
	if (automation) automation->Release();
	if (roomEl) roomEl->Release();

	// End of DotMicPerson 
}

// Initialize UI Automation
HRESULT __stdcall InitUIAutomation(void)
{
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr)) {
		sprintf_s(gszDebugMsg, MAX_PATH, "CoInitialize failed: 0x%008X", hr);
		OutputDebugStringA(gszDebugMsg);
		return hr;
	}

	hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&automation));
	if (FAILED(hr)) {
		sprintf_s(gszDebugMsg, MAX_PATH, "CoCreateInstance failed: 0x%008X", hr);
		OutputDebugStringA(gszDebugMsg);
		CoUninitialize();
		return hr;
	}
	return S_OK;
}

static void SimulateRightClick(int x, int y) {
	SetCursorPos(x, y);
	mouse_event(MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0);
	mouse_event(MOUSEEVENTF_RIGHTUP, x, y, 0, 0);
}

static void SimulateLeftClick(int x, int y) {
	SetCursorPos(x, y);
	mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
	mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
}

static void SendEnterTwice() {
	INPUT input[2] = {};
	for (int i = 0; i < 2; ++i) {
		input[0].type = INPUT_KEYBOARD;
		input[0].ki.wVk = VK_RETURN;
		SendInput(1, &input[0], sizeof(INPUT));

		input[0].ki.dwFlags = KEYEVENTF_KEYUP;
		SendInput(1, &input[0], sizeof(INPUT));
		input[0].ki.dwFlags = 0;
		std::this_thread::sleep_for(std::chrono::milliseconds(100));
	}
}

/// New doting mic user function
HRESULT __stdcall DotAndUnDotMicUser(char* szMicUser)
{
	HRESULT hr = S_OK;
	
	// Getting the chat room element
	CComPtr<IUIAutomationElement> elementRoom;
	std::wstring strRoomTitle(gwcRoomTitle);
	hr = FindWindowByTitle(strRoomTitle, &elementRoom);
	if (FAILED(hr)) {
	    char szDebug[] = "ElementFromHandle failed: elementRoom\n";
		OutputDebugStringA(szDebug);
		return hr;
	}

	// Finding the Mic Queue Title Widget element
	CComPtr<IUIAutomationElement> elementMicQueueTitle ;
	CComPtr<IUIAutomationCondition> classNameCondition;
	// Creating class condition
	CComVariant classNameVariant(L"automation->CreatePropertyCondition(UIA_ClassNamePropertyId, vTitleItemClass, &classCondTitleItem)");
	hr = automation->CreatePropertyCondition(UIA_ClassNamePropertyId, classNameVariant, &classNameCondition);
	if (FAILED(hr)) {
		swprintf_s(gwcDebugMsg, MAX_PATH, L"Class condition for MicQueueTitle failed!\n");
		OutputDebugStringW(gwcDebugMsg);
		return hr;
	}
	// Getting the elementMicQueueTitle
	hr = elementRoom->FindFirst(TreeScope_Descendants, classNameCondition, &elementMicQueueTitle);
	if (FAILED(hr)) {
		swprintf_s(gwcDebugMsg, MAX_PATH, L"Failed to find the elementMicQueueTitle!\n");
		OutputDebugStringW(gwcDebugMsg);
		return hr;
	}


	return hr;
}

// End of File