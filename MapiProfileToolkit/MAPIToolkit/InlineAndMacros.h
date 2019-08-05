#pragma once
#include <string>
#include <Windows.h>
#include "Logger.h"
#include <sstream>
namespace MAPIToolkit
{
#define HCK(_hRes) \
	do { \
		hRes = _hRes; \
		if (FAILED(hRes)) \
																{ \
			std::cout << "FAILED! hr = " << std::hex << hRes << ".  LINE = " << std::dec << __LINE__ << "\n"; \
			std::cout << " >>> " << std::hex << _hRes <<  "\n"; \
			goto Error; \
																} \
								} while (0)

#define HCKL(_hRes, loggerMode) \
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

#define HCKM(_hRes, wszMessage) \
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

#define FCHK(variable, flag) (flag == (variable & flag))

#define VCHK(variable, value) (value == variable)
}