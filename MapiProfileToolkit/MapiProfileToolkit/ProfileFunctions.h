/*
* © 2015 Microsoft Corporation
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

LPWSTR GetDefaultProfileNameLP();
std::wstring GetDefaultProfileName();
ULONG GetProfileCount();
HRESULT HrGetProfiles(ULONG ulProfileCount, ProfileInfo * profileInfo);
//HRESULT GetProfile(LPWSTR lpszProfileName, ProfileInfo * profileInfo);
HRESULT HrSetCachedModeOneService(LPSTR lpszProfileName, LPMAPIUID lpServiceUid, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths);
HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszOldPath, LPWSTR lpszNewPath, bool bMoveFiles);
HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszNewPath, bool bMoveFiles);
HRESULT HrCreateProfile(LPWSTR lpszProfileName);
HRESULT HrCreateProfile(LPWSTR lpszProfileName, LPSERVICEADMIN2 *lppSvcAdmin2);
HRESULT HrSetDefaultProfile(LPWSTR lpszProfileName);
HRESULT HrCloneProfile(ProfileInfo * profileInfo);
VOID PrintProfile(ProfileInfo * profileInfo);
HRESULT HrGetProfile(LPWSTR lpszProfileName, ProfileInfo * profileInfo);
HRESULT HrPromoteDelegates(LPWSTR lpwszProfileName, BOOL bDefaultProfile, BOOL bAllProfiles, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ULONG ulConnectMode);
HRESULT HrPromoteDelegatesOneProfile(LPWSTR profileName, ProfileInfo * pProfileInfo, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ULONG ulConnectMode);
HRESULT HrDeleteProvider(LPWSTR lpwszProfileName, LPMAPIUID lpServiceUid, LPMAPIUID lpProviderUid);
HRESULT HrCreatePstService(LPSERVICEADMIN2 lpServiceAdmin2, LPMAPIUID * lppServiceUid, LPWSTR lpszServiceName, ULONG ulResourceFlags, ULONG ulPstConfigFlag, LPWSTR lpszPstPathW, LPWSTR lpszDisplayName);
HRESULT HrGetDefaultMsemsServiceAdminProviderPtr(LPWSTR lpwszProfileName, LPPROVIDERADMIN * lppProvAdmin, LPMAPIUID * lppServiceUid);

HRESULT HrGetSections(LPSERVICEADMIN lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT * lppEmsMdbSection, LPPROFSECT * lppStoreProviderSection);
HRESULT HrGetSections(LPSERVICEADMIN2 lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT * lppEmsMdbSection, LPPROFSECT * lppStoreProviderSection);

HRESULT HrCreateMsemsServiceModernExt(BOOL bDefaultProfile, LPWSTR lpwszProfileName, ULONG ulResourceFlags,	ULONG ulProfileConfigFlags,	ULONG ulCachedModeMonths, LPWSTR lpszSmtpAddress, LPWSTR lpszDisplayName);
HRESULT HrCreateMsemsServiceModern(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszDisplayName);
HRESULT HrCreateMsemsServiceLegacyUnresolved(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer);
HRESULT HrCreateMsemsServiceROH(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszMailboxLegacyDn,
	LPWSTR lpszUnresolvedServer,
	LPWSTR lpszRohProxyServer,
	LPWSTR lpszProfileServerDn,
	LPWSTR lpszAutodiscoverUrl);
HRESULT HrCreateMsemsServiceMOH(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszMailboxDn,
	LPWSTR lpszServerDn,
	LPWSTR lpszMailStoreInternalUrl,
	LPWSTR lpszMailStoreExternalUrl,
	LPWSTR lpszAddressBookInternalUrl,
	LPWSTR lpszAddressBookExternalUrl);

HRESULT HrAddDelegateMailboxModern(
	BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	BOOL bDefaultService,
	int iServiceIndex,
	LPWSTR lpszwDisplayName,
	LPWSTR lpszwSMTPAddress);

HRESULT HrAddDelegateMailbox(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	BOOL bDefaultService,
	int iServiceIndex,
	LPWSTR lpszwMailboxDisplay,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer,
	LPWSTR lpszwServerDN,
	LPWSTR lpszwSMTPAddress,
	LPWSTR lpRohProxyserver,
	ULONG ulRohProxyServerFlags,
	ULONG ulRohProxyServerAuthPackage,
	LPWSTR lpwszMapiHttpMailStoreInternalUrl);
HRESULT HrAddDelegateMailboxLegacy(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	BOOL bDefaultService,
	int iServiceIndex,
	LPWSTR lpszwMailboxDisplay,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer,
	LPWSTR lpszwServerDN);
HRESULT HrPromoteOneDelegate(LPWSTR lpwszProfileName, int iOutlookVersion, ULONG ulConnectMode, MailboxInfo mailboxInfo);
HRESULT HrSetCachedModeOneProfile(LPWSTR lpwszProfileName, ProfileInfo * pProfileInfo, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths);

HRESULT HrCreateMsemsService(ULONG ulProifileMode, LPWSTR lpwszProfileName, int iOutlookVersion, ServiceOptions * pServiceOptions);
HRESULT HrCreateMsemsServiceOneProfile(LPWSTR lpwszProfileName, int iOutlookVersion, ServiceOptions * pServiceOptions);