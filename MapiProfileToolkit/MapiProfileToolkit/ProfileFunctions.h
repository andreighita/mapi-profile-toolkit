#include "stdafx.h"
#include "Logger.h"
#include "MAPIObjects.h"
#include "MAPIX.h"
#include "MAPIUtil.h"

std::wstring GetDefaultProfileName(LoggingMode loggingMode);
ULONG GetProfileCount(LoggingMode loggingMode);
HRESULT GetProfiles(ULONG ulProfileCount, ProfileInfo * profileInfo, LoggingMode loggingMode);
HRESULT GetProfile(LPWSTR lpszProfileName, ProfileInfo * profileInfo, LoggingMode loggingMode);
HRESULT UpdateCachedModeConfig(LPSTR lpszProfileName, ULONG ulSectionIndex, ULONG ulCachedModeOwner, ULONG ulCachedModeShared, ULONG ulCachedModePublicFolders, int iCachedModeMonths, LoggingMode loggingMode);
HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszOldPath, LPWSTR lpszNewPath, bool bMoveFiles, LoggingMode loggingMode);
HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszNewPath, bool bMoveFiles, LoggingMode loggingMode); 
