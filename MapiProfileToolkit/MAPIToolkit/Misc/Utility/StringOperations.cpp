#include "StringOperations.h"
#include <Windows.h>
#include <oleauto.h>
#include <algorithm>
#include <sstream>
namespace MAPIToolkit
{
	std::string ConvertMultiByteToStdString(LPSTR lpStr)
	{
		return std::string(lpStr);
	}

	std::wstring ConvertWideCharToStdWstring(LPWSTR lpwStr)
	{
		return std::wstring(lpwStr);
	}

	std::string ConvertWideCharToStdString(LPWSTR lpwStr)
	{
		LPSTR lpszMultiByte = new CHAR[lstrlenW(lpwStr) + 1];
		WideCharToMultiByte(CP_ACP, 0,
			lpwStr,
			-1,
			lpszMultiByte,
			lstrlenW(lpwStr) + 1,
			0, 0);
		return std::string(lpszMultiByte);
	}

	LPWSTR ConvertMultiByteToWideChar(LPSTR lpStr)
	{
		int a = lstrlenA(lpStr);
		BSTR unicodestr = SysAllocStringLen(NULL, a);
		MultiByteToWideChar(CP_ACP, 0, lpStr, a, unicodestr, a);
		return unicodestr;
	}

	LPSTR ConvertWideCharToMultiByte(LPWSTR lpwStr)
	{
		LPSTR lpszMultiByte = new CHAR[lstrlenW(lpwStr) + 1];
		WideCharToMultiByte(CP_ACP, 0,
			lpwStr,
			-1,
			lpszMultiByte,
			lstrlenW(lpwStr) + 1,
			0, 0);
		return lpszMultiByte;
	}

	LPSTR ConvertWideCharToMultiByte(const wchar_t* wcharVal)
	{
		std::wstring tempString = wcharVal;
		LPWSTR lpwStr = (LPWSTR)tempString.c_str();
		LPSTR lpszMultiByte = new CHAR[lstrlenW(lpwStr) + 1];
		WideCharToMultiByte(CP_ACP, 0,
			lpwStr,
			-1,
			lpszMultiByte,
			lstrlenW(lpwStr) + 1,
			0, 0);
		return lpszMultiByte;
	}

	bool WStringReplace(std::wstring* wstr, const std::wstring original, const std::wstring replacement) {
		size_t start_pos = wstr->find(original);
		if (start_pos == std::wstring::npos)
			return false;
		wstr->replace(start_pos, original.length(), replacement);
		return true;
	}

	std::wstring SubstringToEnd(std::wstring wszStringToFind, std::wstring wszStringToTrim)
	{
		std::transform(wszStringToTrim.begin(), wszStringToTrim.end(), wszStringToTrim.begin(), ::tolower);
		std::transform(wszStringToFind.begin(), wszStringToFind.end(), wszStringToFind.begin(), ::tolower);
		size_t pos = wszStringToTrim.find(wszStringToFind);
		if (pos != std::wstring::npos)
		{
			return wszStringToTrim.substr(pos + wszStringToFind.length(), std::wstring::npos);
		}
		else
		{
			return wszStringToTrim;
		}
	}


	std::wstring SubstringFromStart(std::wstring wszStringToFind, std::wstring wszStringToTrim)
	{
		std::transform(wszStringToTrim.begin(), wszStringToTrim.end(), wszStringToTrim.begin(), ::tolower);
		std::transform(wszStringToFind.begin(), wszStringToFind.end(), wszStringToFind.begin(), ::tolower);
		size_t pos = wszStringToTrim.find(wszStringToFind);
		if (pos != std::wstring::npos)
		{
			return wszStringToTrim.substr(0, pos - 1);
		}
		else
		{
			return wszStringToTrim;
		}
	}

	std::wstring ConvertStringToWstring(std::wstring& szString)
	{
		std::wstring wsTmp(szString.begin(), szString.end());
		return wsTmp;
	}

	LPWSTR ConvertStdStringToWideChar(std::wstring szValue)
	{
		// Set up a SPropValue array for the properties you need to configure.
		LPSTR lpStr = (LPSTR)szValue.c_str();
		int a = lstrlenA(lpStr);
		BSTR unicodestr = SysAllocStringLen(NULL, a);
		MultiByteToWideChar(CP_ACP, 0, lpStr, a, unicodestr, a);
		return unicodestr;
	}

	LPWSTR ConvertStdStringToWideChar(const wchar_t* szValue)
	{
		// Set up a SPropValue array for the properties you need to configure.
		std::wstring tempString = szValue;
		LPSTR lpStr = (LPSTR)tempString.c_str();
		int a = lstrlenA(lpStr);
		BSTR unicodestr = SysAllocStringLen(NULL, a);
		MultiByteToWideChar(CP_ACP, 0, lpStr, a, unicodestr, a);
		return unicodestr;
	}

	BSTR ConvertStdStringToBstr(const wchar_t* szValue)
	{
		return SysAllocString(szValue);
	}

	std::wstring ConvertIntToString(int t)
	{
		std::wstringstream wss;
		wss << t;
		return wss.str();
	}

}