#include <cstdlib>
#include <iostream>
#include <windows.h>
#include <vector>
#include <string.h>
using namespace std;
BOOL CopyToClipboard(const char* pszData, const int nDataLen);
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam);
vector<HWND> hwnd_vector;
int main(int argc, char *argv[])
{
	EnumWindows(EnumWindowsProc, NULL);
	if (hwnd_vector.size() == 0)
	{
		cout << "there is no tencent window." << endl;
		system("PAUSE");
		return 0;
	}
	cout << "The windows which belong to Tencent are listed below." << endl;
	for (int i = 0; i < hwnd_vector.size(); i++)
	{
		char szWindowName[50] = { 0 };
		GetWindowTextA(hwnd_vector[i], szWindowName, 50);
		cout << i << "," << hwnd_vector[i] << "," << szWindowName << endl;
	}
	cout << endl << "please input the index of the right window:" << endl;
	unsigned int nIndex = 0;
	cin >> nIndex;
	cout << "please input the interval(ms):" << endl;
	unsigned int send_interval = 0;
	cin >> send_interval;
	cout << "please input the send times:" << endl;
	unsigned int send_times = 0;
	cin >> send_times;
	cout << "please input the message:" << endl;
	char msg[255];
	cin >> msg;
	int i = send_times;
	while (i > 0)
	{
		CopyToClipboard(msg, strlen(msg) + 1);
		SendMessage(hwnd_vector[nIndex], WM_PASTE, 0, 0);
		SendMessage(hwnd_vector[nIndex], WM_KEYDOWN, VK_RETURN, 0);
		i--;
		cout << "sent " << (send_times - i) << " time(s)," << "expect end after " << i*send_interval / 1000 << " seconds" << endl;
		Sleep(send_interval);
	}
	cout << "done.";
	system("PAUSE");
	return EXIT_SUCCESS;
}
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam)
{
	wchar_t lpClassName[50];
	GetClassName(hwnd, lpClassName, 50);
	if (_tcscmp(lpClassName, L"TXGuiFoundation") == 0)
	{
		hwnd_vector.push_back(hwnd);
		return true;
	}
}
BOOL CopyToClipboard(const char* pszData, const int nDataLen)
{
	if (OpenClipboard(NULL))
	{
		EmptyClipboard();
		HGLOBAL clipbuffer;
		char *buffer;
		clipbuffer = GlobalAlloc(GMEM_DDESHARE, nDataLen + 1);
		buffer = (char *)GlobalLock(clipbuffer);
		strcpy(buffer, pszData);
		GlobalUnlock(clipbuffer);
		SetClipboardData(CF_TEXT, clipbuffer);
		CloseClipboard();
		return TRUE;
	}
	return FALSE;
}






