#pragma once
#include "pch.h"
namespace MAPIToolkit
{
	std::wstring __cdecl GetRegStringValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName);
}