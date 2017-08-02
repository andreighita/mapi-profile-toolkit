#pragma once
#include "stdafx.h"
#include <objbase.h>
#include <wchar.h>
#include <activeds.h>
#include <Iads.h>
#include <sddl.h>
#include <wchar.h>
#include <initguid.h>
#define USES_IID_IADsADSystemInfo
#define USES_IID_IDirectorySearch
#define USES_IID_IADs
typedef IADs FAR * LPADS;
typedef IDirectorySearch FAR * LPDIRECTORYSEARCH;
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Activeds.lib")
#pragma comment (lib, "adsiid.lib")

std::wstring GetUserDn();
std::wstring FindPrimarySMTPAddress(std::wstring wszUserDn);