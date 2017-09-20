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

// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#include <stdio.h>
#include <tchar.h>
#include <iostream>
#include <strsafe.h>
#include <atlstr.h>
#include <iomanip>
#include <sstream>
#include "Logger.h"

#pragma comment (lib, "mapi32.lib")
#pragma warning(disable:4996) // _CRT_SECURE_NO_WARNINGS

#define EC_HRES(_hRes) \
	do { \
		hRes = _hRes; \
		if (FAILED(hRes)) \
																{ \
			std::cout << "FAILED! hr = " << std::hex << hRes << ".  LINE = " << std::dec << __LINE__ << "\n"; \
			std::cout << " >>> " << (wchar_t*) L#_hRes <<  "\n"; \
			goto Error; \
																} \
								} while (0)

#define EC_HRES_LOG(_hRes, loggerMode) \
	do { \
		hRes = _hRes; \
		if (FAILED(hRes)) \
					{ \
			std::wostringstream oss; \
			oss << L"Error " << std::hex << _hRes << L" in file " << __FILE__ << L" at line " << std::dec << __LINE__ ; \
			Logger::Write(logLevelError, oss.str()); \
			goto Error; \
																} \
								} while (0)

#define EC_HRES_MSG(_hRes, wszMessage) \
	do { \
		hRes = _hRes; \
		if (FAILED(hRes)) \
																		{ \
			std::wostringstream oss; \
			oss << L"Method: " << __FUNCTIONW__ << L"\n Message: "<< wszMessage << L"\nFile: " << __FILE__ << L"\nLine:  " << std::dec << __LINE__ << L"\nError: " << std::hex << _hRes ; \
			Logger::Write(logLevelError, oss.str()); \
			goto Error; \
																		} \
									} while (0)

std::wstring ConvertStringToWstring(std::string & szString);
LPWSTR ConvertStdStringToWideChar(std::string szValue);



// TODO: reference additional headers your program requires here
