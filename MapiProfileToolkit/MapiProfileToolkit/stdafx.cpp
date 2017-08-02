#include "stdafx.h"

// TODO: reference any additional headers you need in STDAFX.H
// and not in this file

std::wstring ConvertStringToWstring(std::string & szString)
{
	std::wstring wsTmp(szString.begin(), szString.end());
	return wsTmp;
}

LPWSTR ConvertStdStringToWideChar(std::string szValue)
{
	// Set up a SPropValue array for the properties you need to configure.
	LPSTR lpStr = (LPSTR)szValue.c_str();
	int a = lstrlenA(lpStr);
	BSTR unicodestr = SysAllocStringLen(NULL, a);
	MultiByteToWideChar(CP_ACP, 0, lpStr, a, unicodestr, a);
	return unicodestr;
}