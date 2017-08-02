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
			Logger::Write(logLevelError, oss.str() , loggerMode); \
			goto Error; \
																} \
								} while (0)

#define EC_HRES_MSG(_hRes, uidErrorMsg) \
	do { \
		hRes = _hRes; \
		if (FAILED(hRes)) \
																		{ \
			std::cout << "FAILED! hr = " << uidErrorMsg << std::hex << hRes << ".  LINE = " << std::dec << __LINE__ << "\n"; \
			std::cout << " >>> " << (wchar_t*) L#_hRes <<  "\n"; \
			goto Error; \
																		} \
									} while (0)


std::wstring ConvertStringToWstring(std::string & szString);
LPWSTR ConvertStdStringToWideChar(std::string szValue);



// TODO: reference additional headers your program requires here
