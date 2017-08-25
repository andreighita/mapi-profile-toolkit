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
#include <EdkMdb.h>
#include <MAPIGuid.h>
#include <MAPIAux.h>	
#include <MSPST.h>
#include <WinBase.h>
#include <Shlwapi.h>


std::wstring GetDefaultProfileName(LoggingMode loggingMode);
ULONG GetProfileCount(LoggingMode loggingMode);
HRESULT GetProfiles(ULONG ulProfileCount, ProfileInfo * profileInfo, LoggingMode loggingMode);
HRESULT GetProfile(LPWSTR lpszProfileName, ProfileInfo * profileInfo, LoggingMode loggingMode);
HRESULT UpdateCachedModeConfig(LPSTR lpszProfileName, ULONG ulSectionIndex, ULONG ulCachedModeOwner, ULONG ulCachedModeShared, ULONG ulCachedModePublicFolders, int iCachedModeMonths, LoggingMode loggingMode);
HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszOldPath, LPWSTR lpszNewPath, bool bMoveFiles, LoggingMode loggingMode);
HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszNewPath, bool bMoveFiles, LoggingMode loggingMode);
HRESULT HrCreateProfile(LPWSTR lpszProfileName, LPSERVICEADMIN2 *lppSvcAdmin2);
HRESULT HrSetDefaultProfile(LPWSTR lpszProfileName);
HRESULT HrCloneProfile(ProfileInfo * profileInfo, LoggingMode loggingMode);
VOID PrintProfile(ProfileInfo * profileInfo);
HRESULT HrGetProfile(LPWSTR lpszProfileName, ProfileInfo * profileInfo, LoggingMode loggingMode);

HRESULT HrCreatePstService(LPSERVICEADMIN2 lpServiceAdmin2, LPMAPIUID * lppServiceUid, LPWSTR lpszServiceName, ULONG ulResourceFlags, ULONG ulPstConfigFlag, LPWSTR lpszPstPathW, LPWSTR lpszDisplayName);

HRESULT HrGetSections(LPSERVICEADMIN2 lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT * lppEmsMdbSection, LPPROFSECT * lppStoreProviderSection);
HRESULT HrCreateMsemsServiceModernExt(LPSERVICEADMIN2 lpServiceAdmin2,
	LPMAPIUID * lppServiceUid,
	ULONG ulResourceFlags,
	ULONG ulProfileConfigFlags,
	ULONG ulCachedModeMonths,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszDisplayName);
HRESULT HrCreateMsemsServiceModern(LPSERVICEADMIN2 lpServiceAdmin2,
	LPMAPIUID * lppServiceUid,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszDisplayName);
HRESULT HrCreateMsemsServiceLegacyUnresolved(LPSERVICEADMIN2 lpServiceAdmin2,
	LPMAPIUID * lppServiceUid,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer);
HRESULT HrCreateMsemsServiceROH(LPSERVICEADMIN2 lpServiceAdmin2,
	LPMAPIUID * lppServiceUid,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszMailboxLegacyDn,
	LPWSTR lpszUnresolvedServer,
	LPWSTR lpszRohProxyServer,
	LPWSTR lpszProfileServerDn,
	LPWSTR lpszAutodiscoverUrl);
HRESULT HrCreateMsemsServiceMOH(LPSERVICEADMIN2 lpServiceAdmin2,
	LPMAPIUID * lppServiceUid,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszMailboxDn,
	LPWSTR lpszMailStoreInternalUrl,
	LPWSTR lpszMailStoreExternalUrl,
	LPWSTR lpszAddressBookInternalUrl,
	LPWSTR lpszAddressBookExternalUrl);


HRESULT HrAddDelegateMailboxModern(MAPIUID uidService,
	LPSERVICEADMIN2 lpSvcAdmin,
	LPWSTR lpszwDisplayName,
	LPWSTR lpszwSMTPAddress);
HRESULT HrAddDelegateMailbox(MAPIUID uidService,
	LPSERVICEADMIN2 lpSvcAdmin,
	LPWSTR lpszwMailboxDisplay,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer,
	LPWSTR lpszwServerDN,
	LPWSTR lpszwSMTPAddress);
HRESULT HrAddDelegateMailboxLegacy(MAPIUID uidService,
	LPSERVICEADMIN2 lpSvcAdmin,
	LPWSTR lpszwMailboxDisplay,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer,
	LPWSTR lpszwServerDN);