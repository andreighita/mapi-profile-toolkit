#pragma once
#include <string>
#include <windows.h>

namespace MAPIToolkit
{
	std::wstring __cdecl GetRegStringValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName);
}