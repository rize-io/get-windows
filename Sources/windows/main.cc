#include <cmath>
#include <cstdint>
#include <napi.h>
#include <iostream>
#include <shtypes.h>
#include <string>
#include <windows.h>
#include <psapi.h>
#include <version>
#include <dwmapi.h>
#include <algorithm>
#include <uiautomation.h>
#include <atlbase.h>
#include <objbase.h>

typedef int(__stdcall *lp_GetScaleFactorForMonitor)(HMONITOR, DEVICE_SCALE_FACTOR *);

struct OwnerWindowInfo {
	std::string path;
	std::string name;
};

template <typename T>
T getValueFromCallbackData(const Napi::CallbackInfo &info, unsigned handleIndex) {
	return reinterpret_cast<T>(info[handleIndex].As<Napi::Number>().Int64Value());
}

// Get wstring from string
std::wstring get_wstring(const std::string str) {
	return std::wstring(str.begin(), str.end());
}

std::string getFileName(const std::string &value) {
	char separator = '/';
#ifdef _WIN32
	separator = '\\';
#endif
	size_t index = value.rfind(separator, value.length());

	if (index != std::string::npos) {
		return (value.substr(index + 1, value.length() - index));
	}

	return ("");
}

// Convert wstring into utf8 string
std::string toUtf8(const std::wstring &str) {
	std::string ret;
	int len = WideCharToMultiByte(CP_UTF8, 0, str.c_str(), str.length(), NULL, 0, NULL, NULL);
	if (len > 0) {
		ret.resize(len);
		WideCharToMultiByte(CP_UTF8, 0, str.c_str(), str.length(), &ret[0], len, NULL, NULL);
	}

	return ret;
}

// Return window title in utf8 string
std::string getWindowTitle(const HWND hwnd) {
	int bufsize = GetWindowTextLengthW(hwnd) + 1;
	LPWSTR t = new WCHAR[bufsize];
	GetWindowTextW(hwnd, t, bufsize);

	std::wstring ws(t);
	std::string title = toUtf8(ws);

	return title;
}

// Return description from file version info
std::string getDescriptionFromFileVersionInfo(const BYTE *pBlock) {
	UINT bufLen = 0;
	struct LANGANDCODEPAGE {
		WORD wLanguage;
		WORD wCodePage;
	} * lpTranslate;

	LANGANDCODEPAGE codePage{0x040904E4};
	// Get language struct
	if (VerQueryValueW((LPVOID *)pBlock, (LPCWSTR)L"\\VarFileInfo\\Translation", (LPVOID *)&lpTranslate, &bufLen)) {
		codePage = lpTranslate[0];
	}

	wchar_t fileDescriptionKey[256];
	wsprintfW(fileDescriptionKey, L"\\StringFileInfo\\%04x%04x\\FileDescription", codePage.wLanguage, codePage.wCodePage);
	wchar_t *fileDescription = NULL;
	UINT fileDescriptionSize;
	// Get description file
	if (VerQueryValueW((LPVOID *)pBlock, fileDescriptionKey, (LPVOID *)&fileDescription, &fileDescriptionSize)) {
		return toUtf8(fileDescription);
	}

	return "";
}

// Return process path and name
OwnerWindowInfo getProcessPathAndName(const HANDLE &phlde) {
	DWORD dwSize{MAX_PATH};
	wchar_t exeName[MAX_PATH]{};
	QueryFullProcessImageNameW(phlde, 0, exeName, &dwSize);
	std::string path = toUtf8(exeName);
	std::string name = getFileName(path);

	DWORD dwHandle = 0;
	wchar_t *wspath(exeName);
	DWORD infoSize = GetFileVersionInfoSizeW(wspath, &dwHandle);

	if (infoSize != 0) {
		BYTE *pVersionInfo = new BYTE[infoSize];
		std::unique_ptr<BYTE[]> skey_automatic_cleanup(pVersionInfo);
		if (GetFileVersionInfoW(wspath, NULL, infoSize, pVersionInfo) != 0) {
			std::string nname = getDescriptionFromFileVersionInfo(pVersionInfo);
			if (nname != "") {
				name = nname;
			}
		}
	}

	return {path, name};
}

OwnerWindowInfo newOwner;

BOOL CALLBACK EnumChildWindowsProc(HWND hwnd, LPARAM lParam) {
	// Get process ID
	DWORD processId{0};
	GetWindowThreadProcessId(hwnd, &processId);
	OwnerWindowInfo *ownerInfo = (OwnerWindowInfo *)lParam;
	// Get process Handler
	HANDLE phlde{OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, processId)};

	if (phlde != NULL) {
		newOwner = getProcessPathAndName(phlde);
		CloseHandle(hwnd);
		if (ownerInfo->path != newOwner.path) {
			return FALSE;
		}
	}

	return TRUE;
}

// Find the address bar in Google Chrome
HRESULT getChromeAddressBarUIAElement(HWND hwnd, IUIAutomationElement** ppAddressBar)
{
	CComPtr<IUIAutomation> pAutomation;
	HRESULT hr = CoCreateInstance(__uuidof(CUIAutomation), nullptr, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&pAutomation);
	if (FAILED(hr)) return hr;

	CComPtr<IUIAutomationElement> pRootElement;
	hr = pAutomation->ElementFromHandle(hwnd, &pRootElement);
	if (FAILED(hr)) return hr;

	CComPtr<IUIAutomationCondition> pAccessKeyCondition;
	hr = pAutomation->CreatePropertyCondition(UIA_AccessKeyPropertyId, CComVariant("Ctrl+L"), &pAccessKeyCondition);
	if (FAILED(hr)) return hr;

	CComPtr<IUIAutomationCondition> pNameCondition;
	hr = pAutomation->CreatePropertyCondition(UIA_NamePropertyId, CComVariant("Address and search bar"), &pNameCondition);
	if (FAILED(hr)) return hr;

	CComPtr<IUIAutomationCondition> pOrCondition;
	hr = pAutomation->CreateAndCondition(pAccessKeyCondition, pNameCondition, &pOrCondition);
	if (FAILED(hr)) return hr;

	hr = pRootElement->FindFirst(TreeScope_Subtree, pOrCondition, ppAddressBar);

	return hr;
}

// Print information about an UI Automation Element
void printElementInfo(IUIAutomationElement* element) {
	CComBSTR bstrName;
	CONTROLTYPEID controlId;
	CComBSTR bstrLocalizedControlType;

	if (SUCCEEDED(element->get_CurrentName(&bstrName)) && bstrName) {
		std::wstring wstrName(bstrName, SysStringLen(bstrName));
		std::string strName(wstrName.begin(), wstrName.end());
		std::cout << "Element Name: " << strName << std::endl;
	}

	if (SUCCEEDED(element->get_CurrentControlType(&controlId))) {
		std::cout << "Element ControlType: " << controlId << std::endl;
	}

	if (SUCCEEDED(element->get_CurrentLocalizedControlType(&bstrLocalizedControlType)) && bstrLocalizedControlType) {
		std::wstring wstrLocalizedControlType(bstrLocalizedControlType, SysStringLen(bstrLocalizedControlType));
		std::string strLocalizedControlType(wstrLocalizedControlType.begin(), wstrLocalizedControlType.end());
		std::cout << "Element LocalizedControlType: " << strLocalizedControlType << std::endl;
	}
}

// Get the URL from Google Chrome
std::string getChromeUrl(HWND hwnd) {
	std::string url;

	CComPtr<IUIAutomationElement> pAddressBar;
	HRESULT hr = getChromeAddressBarUIAElement(hwnd, &pAddressBar);

	if (SUCCEEDED(hr) && pAddressBar)
	{
		printElementInfo(pAddressBar);
		CComPtr<IUIAutomationValuePattern> pValuePattern;
		hr = pAddressBar->GetCurrentPattern(UIA_ValuePatternId, (IUnknown**)&pValuePattern);

		if (SUCCEEDED(hr) && pValuePattern)
		{
			CComBSTR bstrValue;
			hr = pValuePattern->get_CurrentValue(&bstrValue);

			if (SUCCEEDED(hr) && bstrValue)
			{
				url = CW2A(bstrValue, CP_UTF8);
			}
		}
	}

	return url;
}

// Return window information
Napi::Value getWindowInformation(const HWND &hwnd, const Napi::CallbackInfo &info) {
	Napi::Env env{info.Env()};

	if (hwnd == NULL) {
		return env.Null();
	}

	// Get process ID
	DWORD processId{0};
	GetWindowThreadProcessId(hwnd, &processId);
	// Get process Handler
	HANDLE phlde{OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, processId)};

	if (phlde == NULL) {
		return env.Null();
	}

	OwnerWindowInfo ownerInfo = getProcessPathAndName(phlde);

	// ApplicationFrameHost & Universal Windows Platform Support
	if (getFileName(ownerInfo.path) == "ApplicationFrameHost.exe") {
		newOwner = (OwnerWindowInfo) * new OwnerWindowInfo();
		BOOL result = EnumChildWindows(hwnd, (WNDENUMPROC)EnumChildWindowsProc, (LPARAM)&ownerInfo);
		if (result == FALSE && newOwner.name.size())
		{
			ownerInfo = newOwner;
		}
	}

	if (ownerInfo.name == "Widgets.exe") {
		return env.Null();
	}

	PROCESS_MEMORY_COUNTERS memoryCounter;
	BOOL memoryResult = GetProcessMemoryInfo(phlde, &memoryCounter, sizeof(memoryCounter));

	CloseHandle(phlde);

	if (memoryResult == 0) {
		return env.Null();
	}

	RECT lpRect;
	BOOL rectResult = GetWindowRect(hwnd, &lpRect);

	if (rectResult == 0) {
		return env.Null();
	}

	Napi::Object owner = Napi::Object::New(env);

	owner.Set(Napi::String::New(env, "processId"), processId);
	owner.Set(Napi::String::New(env, "path"), ownerInfo.path);
	owner.Set(Napi::String::New(env, "name"), ownerInfo.name);

	Napi::Object bounds = Napi::Object::New(env);

	bounds.Set(Napi::String::New(env, "x"), lpRect.left);
	bounds.Set(Napi::String::New(env, "y"), lpRect.top);
	bounds.Set(Napi::String::New(env, "width"), lpRect.right - lpRect.left);
	bounds.Set(Napi::String::New(env, "height"), lpRect.bottom - lpRect.top);

	Napi::Object activeWinObj = Napi::Object::New(env);

	activeWinObj.Set(Napi::String::New(env, "platform"), Napi::String::New(env, "windows"));
	activeWinObj.Set(Napi::String::New(env, "id"), (LONG)hwnd);
	activeWinObj.Set(Napi::String::New(env, "title"), getWindowTitle(hwnd));
	activeWinObj.Set(Napi::String::New(env, "owner"), owner);
	activeWinObj.Set(Napi::String::New(env, "bounds"), bounds);
	activeWinObj.Set(Napi::String::New(env, "memoryUsage"), memoryCounter.WorkingSetSize);

	// Check if the window is Google Chrome
	if (ownerInfo.name == "Google Chrome") {
		HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

		if (SUCCEEDED(hr)) {
			std::string chromeUrl = getChromeUrl(hwnd);
			printf("chromeUrl: %s", chromeUrl.c_str());
			activeWinObj.Set(Napi::String::New(env, "url"), Napi::String::New(env, chromeUrl));
		}

		CoUninitialize();
	}

	return activeWinObj;
}

// List of HWND used for EnumDesktopWindows callback
std::vector<HWND> _windows;

// EnumDesktopWindows callback
BOOL CALLBACK EnumDekstopWindowsProc(HWND hwnd, LPARAM lParam) {
	if (IsWindow(hwnd) && IsWindowEnabled(hwnd) && IsWindowVisible(hwnd)) {
		WINDOWINFO winInfo{};
		GetWindowInfo(hwnd, &winInfo);

		if (
			(winInfo.dwExStyle & WS_EX_TOOLWINDOW) == 0
			&& (winInfo.dwStyle & WS_CAPTION) == WS_CAPTION
			&& (winInfo.dwStyle & WS_CHILD) == 0
		) {
			int ClockedVal;
			DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, (PVOID)&ClockedVal, sizeof(ClockedVal));
			if (ClockedVal == 0) {
				_windows.push_back(hwnd);
			}
		}
	}

	return TRUE;
};

// Method to get active window
Napi::Value getActiveWindow(const Napi::CallbackInfo &info) {
	HWND hwnd = GetForegroundWindow();
	Napi::Value value = getWindowInformation(hwnd, info);
	return value;
}

// Method to get an array of open windows
Napi::Array getOpenWindows(const Napi::CallbackInfo &info) {
	Napi::Env env{info.Env()};

	Napi::Array values = Napi::Array::New(env);

	_windows.clear();

	if (EnumDesktopWindows(NULL, (WNDENUMPROC)EnumDekstopWindowsProc, NULL)) {
		uint32_t i = 0;
		for (HWND _win : _windows) {
			Napi::Value value = getWindowInformation(_win, info);
			if (value != env.Null()) {
				values.Set(i++, value);
			}
		}
	}

	return values;
}

Napi::Object Init(Napi::Env env, Napi::Object exports) {
	exports.Set(Napi::String::New(env, "getActiveWindow"), Napi::Function::New(env, getActiveWindow));
	exports.Set(Napi::String::New(env, "getOpenWindows"), Napi::Function::New(env, getOpenWindows));
	return exports;
}

NODE_API_MODULE(addon, Init)
