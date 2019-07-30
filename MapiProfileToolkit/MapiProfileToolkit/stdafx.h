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
#include <algorithm>
#include <MAPIX.h>
#include <MAPIUtil.h>
#include <MAPIAux.h>
#include "ExtraMAPIDefs.h"
#include <mapidefs.h>
#include <guiddef.h>
#include <iostream>
#include <string>
#include <utility>
#include <vector>
#include "ToolkitObjects.h"
#include "Misc/Utility/StringOperations.h"

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
		Logger::Write(logLevelInfo, wszMessage); \
		if (FAILED(hRes)) \
																		{ \
			std::wostringstream oss; \
			oss << L"Method: " << __FUNCTIONW__ << L"\nFile: " << __FILE__ << L"\nLine:  " << std::dec << __LINE__ << L"\nError: " << std::hex << _hRes ; \
			Logger::Write(logLevelError, oss.str()); \
			goto Error; \
																		} \
									} while (0)

#define FLAGCHECK(variable, flag) (flag == (variable & flag))

#define VALUECHECK(variable, value) (value == variable)

std::wstring ConvertStringToWstring(std::string & szString);
LPWSTR ConvertStdStringToWideChar(std::string szValue);



// TODO: reference additional headers your program requires here
