#pragma once

#include "targetver.h"


#include <MAPIX.h>
#include <MAPIUtil.h>
#include <MAPIAux.h>
#include "ExtraMAPIDefs.h"
#include "EdkMdb.h"
#include <MAPIGuid.h>
#include <MSPST.h>
#include <mapidefs.h>
#include <guiddef.h>
#include <initguid.h>
#define USES_IID_IMAPIProp 
#define USES_IID_IMsgServiceAdmin2
#define USES_IID_IMAPISession
#include <MAPIAux.h>	
#include <stdio.h>
#include <tchar.h>
#include <strsafe.h>
#include <atlstr.h>
#include <iomanip>
#include <sstream>
#include "Logger.h"
#include <algorithm>
#include <iostream>
#include <string>
#include <utility>
#include <vector>
#include "ToolkitObjects.h"
#include "MAPIObjects.h"
#include "Misc/Utility/StringOperations.h"

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Activeds.lib")
#pragma comment(lib, "adsiid.lib")
#pragma comment(lib, "mapi32.lib")
#pragma comment(lib, "Crypt32.lib")

#pragma warning(disable:4996) // _CRT_SECURE_NO_WARNINGS

#define EC_HRES(_hRes) \
	do { \
		hRes = _hRes; \
		if (FAILED(hRes)) \
																{ \
			std::cout << "FAILED! hr = " << std::hex << hRes << ".  LINE = " << std::dec << __LINE__ << "\n"; \
			std::cout << " >>> " << std::hex << _hRes <<  "\n"; \
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

#define FCHK(variable, flag) (flag == (variable & flag))

#define VCHK(variable, value) (value == variable)

std::wstring ConvertStringToWstring(std::wstring & szString);
LPWSTR ConvertStdStringToWideChar(std::wstring szValue);



// TODO: reference additional headers your program requires here
