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
#include <chrono>
#include <functional>

typedef int(__stdcall *lp_GetScaleFactorForMonitor)(HMONITOR, DEVICE_SCALE_FACTOR *);
typedef std::function<bool(IUIAutomationElement*)> ElementMatcher;

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

	LANGANDCODEPAGE codePage{0x0409, 0x04E4};
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

bool ownerHasName(const OwnerWindowInfo& ownerInfo, const std::string exeName, const std::string appName) {
	if (!ownerInfo.path.empty()) {
		std::string path = ownerInfo.path;
		size_t lastBackslash = path.find_last_of('\\');
		if (lastBackslash != std::string::npos) {
			std::string lastSection = path.substr(lastBackslash + 1);
			if (lastSection.find(exeName) != std::string::npos) {
				return true;
			}
		}
	}

	if (!ownerInfo.name.empty()) {
		std::string name = ownerInfo.name;
		if (name.find(appName) != std::string::npos) {
			return true;
		}
	}

	return false;
}

bool isGoogleChrome(const OwnerWindowInfo& ownerInfo) {
	return ownerHasName(ownerInfo, "chrome", "Google Chrome");
}

bool isBraveBrowser(const OwnerWindowInfo& ownerInfo) {
	return ownerHasName(ownerInfo, "brave", "Brave Browser");
}

bool isMicrosoftEdge(const OwnerWindowInfo& ownerInfo) {
	return ownerHasName(ownerInfo, "msedge", "Microsoft Edge");
}

bool isFirefox(const OwnerWindowInfo& ownerInfo) {
	return ownerHasName(ownerInfo, "firefox", "Firefox");
}

bool isOperaBrowser(const OwnerWindowInfo& ownerInfo) {
	return ownerHasName(ownerInfo, "opera", "Opera Internet Browser");
}

bool isSupportedBrowser(const OwnerWindowInfo& ownerInfo) {
	return isGoogleChrome(ownerInfo) || isBraveBrowser(ownerInfo) || isMicrosoftEdge(ownerInfo) || isFirefox(ownerInfo) || isOperaBrowser(ownerInfo);
}

IUIAutomationElement* findUIAElementRecursively(IUIAutomationElement* element, int depth, int& iteration, ElementMatcher matcher, bool skipChildren = false) {
	if (element == nullptr) {
		return nullptr;
	}

	if (depth == 0 && iteration != 0) {
		return nullptr;
	}

	iteration++;

	CONTROLTYPEID controlId;
	HRESULT hr = element->get_CurrentControlType(&controlId);
	if (FAILED(hr)) {
		return nullptr;
	}

	if (controlId == UIA_DocumentControlTypeId || controlId == UIA_MenuBarControlTypeId || controlId == UIA_MenuControlTypeId || controlId == UIA_TabControlTypeId || controlId == UIA_CustomControlTypeId) {
		skipChildren = true;
	}

   if (matcher(element)) {
		element->AddRef();
		return element;
	}

	CComPtr<IUIAutomationTreeWalker> pTreeWalker;
	CComPtr<IUIAutomation> pAutomation;

	hr = CoCreateInstance(__uuidof(CUIAutomation), nullptr, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&pAutomation);
	if (FAILED(hr)) {
		return nullptr;
	}

	hr = pAutomation->get_RawViewWalker(&pTreeWalker);
	if (FAILED(hr)) {
		return nullptr;
	}

	if (!skipChildren) {
		CComPtr<IUIAutomationElement> pFirstChild;
		hr = pTreeWalker->GetFirstChildElement(element, &pFirstChild);
		if (SUCCEEDED(hr)) {
			IUIAutomationElement* result = findUIAElementRecursively(pFirstChild, depth + 1, iteration, matcher);
			if (result) {
				return result;
			}
		}
	}

	CComPtr<IUIAutomationElement> pNextSibling;
	hr = pTreeWalker->GetNextSiblingElement(element, &pNextSibling);
	if (SUCCEEDED(hr)) {
		IUIAutomationElement* result = findUIAElementRecursively(pNextSibling, depth, iteration, matcher, controlId == UIA_DocumentControlTypeId);
		if (result) {
			return result;
		}
	}

	return nullptr;
}

HRESULT findUIAElement(HWND hwnd, IUIAutomationElement** ppAddressBar, ElementMatcher matcher) {
	HRESULT hr = S_OK;
	CComPtr<IUIAutomation> pAutomation;

	hr = CoCreateInstance(__uuidof(CUIAutomation), nullptr, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&pAutomation);
	if (FAILED(hr)) {
		return hr;
	}

	CComPtr<IUIAutomationElement> pRootElement;
	hr = pAutomation->ElementFromHandle(hwnd, &pRootElement);
	if (FAILED(hr)) {
		return hr;
	}

	int iteration = 0;
	IUIAutomationElement* result = findUIAElementRecursively(pRootElement, 0, iteration, matcher);
	if (result) {
		*ppAddressBar = result;
		return S_OK;
	} else {
		return E_FAIL;
	}
}

bool isButtonControlType(IUIAutomationElement* element) {
	CONTROLTYPEID controlId;
	HRESULT hr = element->get_CurrentControlType(&controlId);
	if (FAILED(hr)) {
		return false;
	}

	return controlId != UIA_ButtonControlTypeId;
}

bool isEditControlType(IUIAutomationElement* element) {
	CONTROLTYPEID controlId;
	HRESULT hr = element->get_CurrentControlType(&controlId);
	if (FAILED(hr)) {
		return false;
	}

	return controlId != UIA_EditControlTypeId;
}

bool matchElementName(IUIAutomationElement* element, const std::string& targetName) {
	CComBSTR bstrName;
	if (SUCCEEDED(element->get_CurrentName(&bstrName)) && bstrName) {
		std::wstring wstrName(bstrName, SysStringLen(bstrName));
		std::string strName(wstrName.begin(), wstrName.end());

		if (strName.find(targetName) != std::string::npos) {
			return true;
		}
	}

	return false;
}

bool matchElementLegacyDescription(IUIAutomationElement* element, const std::string& targetDescription) {
	CComPtr<IUIAutomationLegacyIAccessiblePattern> pLegacyIAccessiblePattern;
	HRESULT hr = element->GetCurrentPatternAs(UIA_LegacyIAccessiblePatternId, __uuidof(IUIAutomationLegacyIAccessiblePattern), (void**)&pLegacyIAccessiblePattern);
	if (SUCCEEDED(hr) && pLegacyIAccessiblePattern) {
		CComBSTR bstrDescription;
		hr = pLegacyIAccessiblePattern->get_CurrentDescription(&bstrDescription);
		if (SUCCEEDED(hr) && bstrDescription) {
			std::wstring wstrDescription(bstrDescription, SysStringLen(bstrDescription));
			std::string strDescription(wstrDescription.begin(), wstrDescription.end());

			if (strDescription.find(targetDescription) != std::string::npos) {
				return true;
			}
		}
	}

	return false;
}

ElementMatcher googleChromeAddressBarMatcher = [](IUIAutomationElement* element) -> bool {
	return isEditControlType(element) && (matchElementName(element, "Ctrl+L") || matchElementName(element, "Address and search bar"));
};

ElementMatcher googleChromeIncognitoMatcher = [](IUIAutomationElement* element) -> bool {
	return isButtonControlType(element) && matchElementName(element, "Incognito");
};

ElementMatcher braveBrowserIncognitoMatcher = [](IUIAutomationElement* element) -> bool {
	return isButtonControlType(element) && (matchElementName(element, "Private") || matchElementLegacyDescription(element, "This is a private window with Tor"));
};

ElementMatcher microsoftEdgeIncognitoMatcher = [](IUIAutomationElement* element) -> bool {
	return isButtonControlType(element) && matchElementName(element, "InPrivate");
};

ElementMatcher firefoxAddressBarMatcher = [](IUIAutomationElement* element) -> bool {
	return isEditControlType(element) && matchElementName(element, "Search with");
};

ElementMatcher firefoxIncognitoMatcher = [](IUIAutomationElement* element) -> bool {
	return isButtonControlType(element) && matchElementName(element, "Mozilla Firefox Private Browsing");
};

ElementMatcher operaBrowserAddressBarMatcher = [](IUIAutomationElement* element) -> bool {
	return isButtonControlType(element) && matchElementName(element, "Address field");
};

ElementMatcher operaBrowserIncognitoMatcher = [](IUIAutomationElement* element) -> bool {
	return isButtonControlType(element) && matchElementName(element, "Opera (Private)");
};

std::string getUrl(HWND hwnd, const OwnerWindowInfo& ownerInfo) {
	std::string url;

	ElementMatcher matcher;

	if (isGoogleChrome(ownerInfo)) {
		matcher = googleChromeAddressBarMatcher;
	}

	if (isBraveBrowser(ownerInfo)) {
		matcher = googleChromeAddressBarMatcher;
	}

	if (isMicrosoftEdge(ownerInfo)) {
		matcher = googleChromeAddressBarMatcher;
	}

	if (isFirefox(ownerInfo)) {
		matcher = firefoxAddressBarMatcher;
	}

	if (isOperaBrowser(ownerInfo)) {
		matcher = operaBrowserAddressBarMatcher;
	}

	CComPtr<IUIAutomationElement> pAddressBar;
	HRESULT hr = findUIAElement(hwnd, &pAddressBar, matcher);

	if (SUCCEEDED(hr) && pAddressBar)
	{
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

std::string getMode(HWND hwnd, const OwnerWindowInfo& ownerInfo) {
	std::string mode;

	ElementMatcher matcher;

	if (isGoogleChrome(ownerInfo)) {
		matcher = googleChromeIncognitoMatcher;
	}

	if (isBraveBrowser(ownerInfo)) {
		matcher = braveBrowserIncognitoMatcher;
	}

	if (isMicrosoftEdge(ownerInfo)) {
		matcher = microsoftEdgeIncognitoMatcher;
	}

	if (isFirefox(ownerInfo)) {
		matcher = firefoxIncognitoMatcher;
	}

	if (isOperaBrowser(ownerInfo)) {
		matcher = operaBrowserIncognitoMatcher;
	}

	CComPtr<IUIAutomationElement> pIncognito;
	HRESULT hr = findUIAElement(hwnd, &pIncognito, matcher);

	if (SUCCEEDED(hr) && pIncognito)
	{
		mode = "incognito";
	} else {
		mode = "normal";
	}

	return mode;
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

	if (isSupportedBrowser(ownerInfo)) {
		HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

		if (SUCCEEDED(hr)) {
			std::string url = getUrl(hwnd, ownerInfo);
			std::string mode = getMode(hwnd, ownerInfo);
			activeWinObj.Set(Napi::String::New(env, "url"), Napi::String::New(env, url));
			activeWinObj.Set(Napi::String::New(env, "mode"), Napi::String::New(env, mode));
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
