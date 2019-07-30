#pragma once

#include "../stdafx.h"
#include "MAPIObjects.h"

#include "Profile.h"

HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszOldPath, LPWSTR lpszNewPath, bool bMoveFiles);

HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszNewPath, bool bMoveFiles);

HRESULT HrCreatePstService(LPSERVICEADMIN2 lpServiceAdmin2, LPMAPIUID* lppServiceUid, LPWSTR lpszServiceName, ULONG ulResourceFlags, ULONG ulPstConfigFlag, LPWSTR lpszPstPathW, LPWSTR lpszDisplayName);

