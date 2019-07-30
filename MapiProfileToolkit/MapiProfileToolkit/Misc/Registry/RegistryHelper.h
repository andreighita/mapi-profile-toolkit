#pragma once
#include "../../stdafx.h"

std::wstring __cdecl GetStringValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName);