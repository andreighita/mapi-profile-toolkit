#pragma once
#include "stdafx.h"

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

bool WStringReplace(std::wstring * wstr, const std::wstring original, const std::wstring replacement) {
	size_t start_pos = wstr->find(original);
	if (start_pos == std::wstring::npos)
		return false;
	wstr->replace(start_pos, original.length(), replacement);
	return true;
}

