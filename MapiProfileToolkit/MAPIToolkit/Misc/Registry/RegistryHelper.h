#pragma once
#include "pch.h"

std::wstring __cdecl GetStringValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName);