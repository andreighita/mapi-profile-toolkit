#pragma once

#include "stdafx.h"

#include "StringOperations.h"
#include "ToolkitObjects.h"
#include "MAPIObjects.h"
#include "Profile.h"

HRESULT HrSetCachedMode(LPWSTR lpwszProfileName, BOOL bDefaultProfile, BOOL bAllProfiles, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths, int iOutlookVersion);

HRESULT HrSetCachedModeOneProfile(LPWSTR lpwszProfileName, ProfileInfo* pProfileInfo, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths, int iOutlookVersion);

HRESULT HrSetCachedModeOneService(LPSTR lpszProfileName, LPMAPIUID lpServiceUid, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths, int iOutlookVersion);

