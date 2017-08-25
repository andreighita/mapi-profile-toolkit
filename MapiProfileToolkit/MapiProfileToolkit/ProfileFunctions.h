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