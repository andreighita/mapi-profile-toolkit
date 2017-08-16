/*
* © 2015 Microsoft Corporation
*
* written by Andrei Ghita
*
* Microsoft provides programming examples for illustration only, without warranty either expressed or implied.
* This includes, but is not limited to, the implied warranties of merchantability or fitness for a particular purpose.
* This article assumes that you are familiar with the programming language that is being demonstrated and with
* the tools that are used to create and to debug procedures. Microsoft support engineers can help explain the
* functionality of a particular procedure, but they will not modify these examples to provide added functionality
* or construct procedures to meet your specific requirements.
*/

// stdafx.cpp : source file that includes just the standard includes
// ProfileToolkit.pch will be the pre-compiled header
// stdafx.obj will contain the pre-compiled type information

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