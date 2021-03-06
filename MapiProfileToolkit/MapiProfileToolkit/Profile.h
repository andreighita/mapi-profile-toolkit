/*
* � 2015 Microsoft Corporation
*
* written by Andrei Ghita
*
* Microsoft provides programming examples for illustration only, without warranty either expressed or implied.
* This includes, but is not limited to, the implied warranties of merchantability or fitness for a particular purpose.
* This article assumes that you are familiar with the programming language that is being demonstrated and with
* the tools that are used to create and to debug procedures. Microsoft support engineers can help explain the
* functionality of a particular procedure, but they will not modify these examples to provide added functionality
* or construct procedures to meet your specific requirements.
*/

#include "stdafx.h"
#include "Logger.h"
#include <initguid.h>
#define USES_IID_IMAPIProp 
#define USES_IID_IMsgServiceAdmin2
#include "MAPIObjects.h"
#include <MAPIX.h>
#include <MAPIUtil.h>
#include <MAPIAux.h>
#include "MAPIObjects.h"
#include "EdkMdb.h"
#include <MAPIGuid.h>
#include <MAPIAux.h>	
#include <MSPST.h>
#include <WinBase.h>
#include <Shlwapi.h>
#include <string>
#include <utility>
#include <iostream>
#include <algorithm> 
#include <vector>
#include "ToolkitObjects.h"
#include "StringOperations.h"
#include "ExchangeAccount.h"
#include "AdditionalMailbox.h"
#include "PST.h"

LPWSTR GetDefaultProfileNameLP();

// GetDefaultProfileName
// returns a std::wstring value with the name of the default Outlook profile
std::wstring GetDefaultProfileName();

// GetProfileCount
// returns the number of mapi profiles for the current user
ULONG GetProfileCount();

HRESULT HrGetProfiles(ULONG ulProfileCount, ProfileInfo* profileInfo);

HRESULT HrDeleteProfile(LPWSTR lpszProfileName);

HRESULT HrCreateProfile(LPWSTR lpszProfileName);

HRESULT HrCreateProfile(LPWSTR lpszProfileName, LPSERVICEADMIN2* lppSvcAdmin2);

HRESULT HrSetDefaultProfile(LPWSTR lpszProfileName);

// Outlook 2016
HRESULT HrCloneProfile(ProfileInfo* profileInfo);

// Outlook 2013
HRESULT HrSimpleCloneProfile(ProfileInfo* profileInfo, bool bSetDefaultProfile);

VOID PrintProfile(ProfileInfo* profileInfo);

HRESULT HrGetProfile(LPWSTR lpszProfileName, ProfileInfo* profileInfo);

// HrDeleteProvider
// Deletes the provider with the specified UID from the service with the specified UID in a given profile
HRESULT HrDeleteProvider(LPWSTR lpwszProfileName, LPMAPIUID lpServiceUid, LPMAPIUID lpProviderUid);

// HrGetSections
// Returns the EMSMDB and StoreProvider sections of a service
HRESULT HrGetSections(LPSERVICEADMIN2 lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT* lppEmsMdbSection, LPPROFSECT* lppStoreProviderSection);

// HrGetSections
// Returns the EMSMDB and StoreProvider sections of a service
HRESULT HrGetSections(LPSERVICEADMIN lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT* lppEmsMdbSection, LPPROFSECT* lppStoreProviderSection);