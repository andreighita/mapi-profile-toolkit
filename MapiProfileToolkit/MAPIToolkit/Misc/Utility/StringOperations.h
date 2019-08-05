#pragma once
#include <string>
#include <Windows.h>

namespace MAPIToolkit
{
	std::string ConvertMultiByteToStdString(LPSTR lpStr);

	std::wstring ConvertWideCharToStdWstring(LPWSTR lpwStr);

	std::string ConvertWideCharToStdString(LPWSTR lpwStr);

	LPWSTR ConvertMultiByteToWideChar(LPSTR lpStr);

	LPSTR ConvertWideCharToMultiByte(LPWSTR lpwStr);

	LPSTR ConvertWideCharToMultiByte(const wchar_t* wcharVal);

	bool WStringReplace(std::wstring* wstr, const std::wstring original, const std::wstring replacement);

	std::wstring SubstringToEnd(std::wstring wszStringToFind, std::wstring wszStringToTrim);

	std::wstring SubstringFromStart(std::wstring wszStringToFind, std::wstring wszStringToTrim);

	std::wstring ConvertStringToWstring(std::wstring& szString);

	LPWSTR ConvertStdStringToWideChar(std::wstring szValue);

	LPWSTR ConvertStdStringToWideChar(const wchar_t* szValue);

	BSTR ConvertStdStringToBstr(const wchar_t* szValue);

	std::wstring ConvertIntToString(int t);
}