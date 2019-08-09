#pragma once
#include <string>
#include <windows.h>
#include "Logger.h"

namespace MAPIToolkit
{
	std::wstring __cdecl GetRegStringValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName);
	BOOL __cdecl WriteRegStringValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName, LPCTSTR lpszValueData);
	BOOL __cdecl WriteRegDwordValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName, DWORD dwValueData);
	BOOL __cdecl WriteRegBinaryValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName, BYTE* pbValueData);
	BOOL ReadAllValues(HKEY hRegistryHive, LPCTSTR lpszKeyName);
}