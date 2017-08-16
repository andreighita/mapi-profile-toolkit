// QueryKey - Enumerates the subkeys of key and its associated values.
//     hKey - Key whose subkeys and values are to be enumerated.

#include "stdafx.h"
#include "RegistryHelper.h"
#include <windows.h>
#include <stdio.h>
#include <tchar.h>

#define MAX_KEY_LENGTH 255
#define MAX_VALUE_NAME 16383

void QueryKey(HKEY hKey)
{
	TCHAR    achKey[MAX_KEY_LENGTH];   // buffer for subkey name
	DWORD    cbName;                   // size of name string 
	TCHAR    achClass[MAX_PATH] = TEXT("");  // buffer for class name 
	DWORD    cchClassName = MAX_PATH;  // size of class string 
	DWORD    cSubKeys = 0;               // number of subkeys 
	DWORD    cbMaxSubKey;              // longest subkey size 
	DWORD    cchMaxClass;              // longest class string 
	DWORD    cValues;              // number of values for key 
	DWORD    cchMaxValue;          // longest value name 
	DWORD    cbMaxValueData;       // longest value data 
	DWORD    cbSecurityDescriptor; // size of security descriptor 
	FILETIME ftLastWriteTime;      // last write time 

	DWORD i, retCode;

	TCHAR  achValue[MAX_VALUE_NAME];
	DWORD cchValue = MAX_VALUE_NAME;

	// Get the class name and the value count. 
	retCode = RegQueryInfoKey(
		hKey,                    // key handle 
		achClass,                // buffer for class name 
		&cchClassName,           // size of class string 
		NULL,                    // reserved 
		&cSubKeys,               // number of subkeys 
		&cbMaxSubKey,            // longest subkey size 
		&cchMaxClass,            // longest class string 
		&cValues,                // number of values for this key 
		&cchMaxValue,            // longest value name 
		&cbMaxValueData,         // longest value data 
		&cbSecurityDescriptor,   // security descriptor 
		&ftLastWriteTime);       // last write time 

								 // Enumerate the subkeys, until RegEnumKeyEx fails.

	if (cSubKeys)
	{
		printf("\nNumber of subkeys: %d\n", cSubKeys);

		for (i = 0; i<cSubKeys; i++)
		{
			cbName = MAX_KEY_LENGTH;
			retCode = RegEnumKeyEx(hKey, i,
				achKey,
				&cbName,
				NULL,
				NULL,
				NULL,
				&ftLastWriteTime);
			if (retCode == ERROR_SUCCESS)
			{
				_tprintf(TEXT("(%d) %s\n"), i + 1, achKey);
			}
		}
	}

	// Enumerate the key values. 

	if (cValues)
	{
		printf("\nNumber of values: %d\n", cValues);

		for (i = 0, retCode = ERROR_SUCCESS; i<cValues; i++)
		{
			cchValue = MAX_VALUE_NAME;
			achValue[0] = '\0';
			retCode = RegEnumValue(hKey, i,
				achValue,
				&cchValue,
				NULL,
				NULL,
				NULL,
				NULL);

			if (retCode == ERROR_SUCCESS)
			{
				_tprintf(TEXT("(%d) %s\n"), i + 1, achValue);
			}
		}
	}
}

bool __cdecl GetValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName, REGSAM regSam, DWORD dwSearchedType, LPBYTE * lpszValueData, DWORD * dwValueDataSize)
{
	HKEY hKey;
	bool fFound = false;
	LSTATUS lStatus = ERROR_SUCCESS;
	lStatus = RegOpenKeyEx(hRegistryHive,
		lpszKeyName,
		0,
		regSam,
		&hKey);
	if ( lStatus == ERROR_SUCCESS
		)
	{
		DWORD dwType;
		DWORD dwSize;
		LPBYTE lpbLookupDataValue = NULL;

		if (ERROR_SUCCESS == RegQueryValueExW(hKey, lpszValueName, NULL, &dwType, NULL, &dwSize))
		{
			if (dwSearchedType == dwType)
			{
				lpbLookupDataValue = (LPBYTE)malloc(dwSize);
				ZeroMemory(lpbLookupDataValue, dwSize);
				if (ERROR_SUCCESS == RegQueryValueExW(hKey, lpszValueName, NULL, &dwType, lpbLookupDataValue, &dwSize))
				{
					memcpy(dwValueDataSize, &dwSize, sizeof(DWORD));
					memcpy(lpszValueData, &lpbLookupDataValue, sizeof(&lpbLookupDataValue));
					fFound = true;
				}
			}
		}
	}

	RegCloseKey(hKey);
	return fFound;
}

std::wstring __cdecl GetStringValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName)
{
	LPBYTE lpbTempValueData = NULL;
	DWORD dwValueDataSize;

	if (GetValue(hRegistryHive, lpszKeyName, lpszValueName, KEY_READ, REG_SZ, &lpbTempValueData, &dwValueDataSize))
	{
		if (lpbTempValueData)
		{
			return std::wstring((LPWSTR)lpbTempValueData);
		}
		else return L"";
	}
	else if (GetValue(hRegistryHive, lpszKeyName, lpszValueName, KEY_READ | KEY_WOW64_64KEY, REG_SZ, &lpbTempValueData, &dwValueDataSize))
	{
		if (lpbTempValueData)
		{
			return std::wstring((LPWSTR)lpbTempValueData);
		}
		else return L"";
	}
	if (GetValue(hRegistryHive, lpszKeyName, lpszValueName, KEY_READ | KEY_WOW64_32KEY, REG_SZ, &lpbTempValueData, &dwValueDataSize))
	{
		if (lpbTempValueData)
		{
			return std::wstring((LPWSTR)lpbTempValueData);
		}
		else return L"";
	}
	else
		return L"";
}


