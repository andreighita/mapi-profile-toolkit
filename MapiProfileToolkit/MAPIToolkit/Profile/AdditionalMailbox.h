#pragma once

#include "pch.h"
#include "Profile.h"
#include "../Misc/Utility/StringOperations.h"

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

HRESULT HrAddDelegateMailbox(ProfileMode profileMode, LPWSTR lpwszProfileName, ServiceMode serviceMode, int iServiceIndex, int iOutlookVersion, ProviderOptions * pMailboxOptions);

HRESULT HrAddDelegateMailboxOneProfile(LPWSTR lpwszProfileName, int iOutlookVersion, ServiceMode serviceMode, int iServiceIndex, ProviderOptions * pMailboxOptions);

HRESULT HrAddDelegateMailbox(ProfileMode profileMode, LPWSTR lpwszProfileName, ServiceMode serviceMode, int iServiceIndex, int iOutlookVersion, ProviderOptions * pMailboxOptions);

HRESULT HrAddDelegateMailboxOneProfile(LPWSTR lpwszProfileName, int iOutlookVersion, ServiceMode serviceMode, int iServiceIndex, ProviderOptions * pMailboxOptions);

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

HRESULT HrPromoteDelegates(LPWSTR lpwszProfileName, BOOL bDefaultProfile, BOOL bAllProfiles, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ConnectMode connectMode);

HRESULT HrPromoteDelegatesOneProfile(LPWSTR lpwszProfileName, ProfileInfo * pProfileInfo, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ConnectMode connectMode);

HRESULT HrPromoteOneDelegate(LPWSTR lpwszProfileName, int iOutlookVersion, ConnectMode connectMode, MailboxInfo mailboxInfo);


