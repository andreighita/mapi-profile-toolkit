#pragma once
#include <string>
#include <Windows.h>
#include <algorithm>

std::string ConvertMultiByteToStdString(LPSTR lpStr);

std::wstring ConvertWideCharToStdWstring(LPWSTR lpwStr);

std::string ConvertWideCharToStdString(LPWSTR lpwStr);

LPWSTR ConvertMultiByteToWideChar(LPSTR lpStr);

LPSTR ConvertWideCharToMultiByte(LPWSTR lpwStr);

bool WStringReplace(std::wstring* wstr, const std::wstring original, const std::wstring replacement);

std::wstring SubstringToEnd(std::wstring wszStringToFind, std::wstring wszStringToTrim);

std::wstring SubstringFromStart(std::wstring wszStringToFind, std::wstring wszStringToTrim);
