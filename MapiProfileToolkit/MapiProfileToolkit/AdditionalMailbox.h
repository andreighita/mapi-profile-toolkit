#pragma once

#include "stdafx.h"
#include "Profile.h"
#include "StringOperations.h"
#include "ToolkitObjects.h"
#include "MAPIObjects.h"
#include <MAPIUtil.h>
#include "ExtraMAPIDefs.h"

// HrAddDelegateMailboxModern
// Adds a delegate mailbox to a given service. The property set is one for Outlook 2016 where all is needed is:
// - the SMTP address of the mailbox
// - the Display Name for the mailbox
HRESULT HrAddDelegateMailboxModern(
	BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	BOOL bDefaultService,
	ULONG iServiceIndex,
	LPWSTR lpszwDisplayName,
	LPWSTR lpszwSMTPAddress);

HRESULT HrAddDelegateMailbox(ULONG ulProifileMode, LPWSTR lpwszProfileName, ULONG ulServiceMode, int iServiceIndex, int iOutlookVersion, MailboxOptions* pMailboxOptions);

HRESULT HrAddDelegateMailboxOneProfile(LPWSTR lpwszProfileName, int iOutlookVersion, ULONG ulServiceMode, int iServiceIndex, MailboxOptions* pMailboxOptions);

HRESULT HrAddDelegateMailbox(ULONG ulProifileMode, LPWSTR lpwszProfileName, ULONG ulServiceMode, int iServiceIndex, int iOutlookVersion, MailboxOptions* pMailboxOptions);

HRESULT HrAddDelegateMailboxOneProfile(LPWSTR lpwszProfileName, int iOutlookVersion, ULONG ulServiceMode, int iServiceIndex, MailboxOptions* pMailboxOptions);

HRESULT HrAddDelegateMailbox(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	BOOL bDefaultService,
	ULONG ulServiceIndex,
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
	ULONG ulServiceIndex,
	LPWSTR lpszwMailboxDisplay,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer,
	LPWSTR lpszwServerDN);

HRESULT HrPromoteDelegates(LPWSTR lpwszProfileName, BOOL bDefaultProfile, BOOL bAllProfiles, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ULONG ulConnectMode);

HRESULT HrPromoteDelegatesOneProfile(LPWSTR lpwszProfileName, ProfileInfo* pProfileInfo, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ULONG ulConnectMode);

HRESULT HrPromoteOneDelegate(LPWSTR lpwszProfileName, int iOutlookVersion, ULONG ulConnectMode, MailboxInfo mailboxInfo);


