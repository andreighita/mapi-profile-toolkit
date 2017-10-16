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
#include <MAPIDefS.h>

#define PR_PST_CONFIG_FLAGS PROP_TAG(PT_LONG, 0x6770)
#define PR_PST_PATH_W PROP_TAG(PT_UNICODE, 0x6700)

enum { CACHEDMODE_DISABLED = 1, CACHEDMODE_ENABLED };
enum { SERVICE_PRIMARY = 1, SERVICE_SECONDARY };
enum { PROFILETYPE_UNKNOWN, PROFILETYPE_PRIMARY, PROFILETYPE_DELEGATE, PROFILETYPE_PUBLICFOLDERS };

enum { PSTTYPE_ANSI = 0, PSTTYPE_UNICODE = 0x80000000};

struct MailboxInfo
{
	std::wstring wszDisplayName; // PR_DISPLAY_NAME
	std::wstring wszSmtpAddress; // PR_PROFILE_USER_SMTP_EMAIL_ADDRESS
	std::wstring wszProfileMailbox; // PR_PROFILE_MAILBOX
	std::wstring wszProfileServerDN; // PR_PROFILE_SERVER_DN
	std::wstring wszRohProxyServer; // PR_ROH_PROXY_SERVER
	std::wstring wszProfileServer; // PR_PROFILE_SERVER
	std::wstring wszProfileServerFqdnW; // PR_PROFILE_SERVER_FQDN_W
	std::wstring wszAutodiscoverUrl; // PR_PROFILE_LKG_AUTODISCOVER_URL
	std::wstring wszMailStoreInternalUrl; // PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL
	std::wstring wszMailStoreExternalUrl; // PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL
	std::wstring wszAddressBookInternalUrl; // PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL
	std::wstring wszAddressBookExternalUrl; // PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL
	BOOL bPrimaryMailbox; 
	ULONG ulResourceFlags; // PR_RESOURCE_FLAGS
	ULONG ulRohProxyAuthScheme; // PR_PROFILE_RPC_PROXY_SERVER_AUTH_PACKAGE
	ULONG ulRohFlags; // PR_ROH_FLAGS
	ULONG ulProfileType; // PR_PROFILE_TYPE
	MAPIUID muidProviderUid;
	MAPIUID muidServiceUid;
	BOOL bIsOnlineArchive;
};

struct PstInfo
{
	std::wstring wszDisplayName;
	std::wstring wszPstPath; // PR_PST_PATH_W
	ULONG ulPstType; 
	ULONG ulPstConfigFlags; // PR_PST_CONFIG_FLAGS
};

struct MapiProperty
{
	std::wstring wszNamedPropertyName;
	std::wstring wszPropertyTag;
	ULONG ulNamedPropertyValue;
};

struct ProviderInfo
{
	MapiProperty * mapiProperties;
};

struct ExchangeAccountInfo
{
	std::wstring wszDisplayName;
	std::wstring wszDatafilePath;
	std::wstring wszEmailAddress;
	BOOL bCachedModeEnabledOwner;
	BOOL bCachedModeEnabledShared;
	BOOL bCachedModeEnabledPublicFolders;
	int iCachedModeMonths;
	std::wstring szUserName;
	std::wstring szUserEmailSmtpAddress;
	ULONG ulMailboxCount;
	std::wstring wszRohProxyServer;
	std::wstring wszUnresolvedServer;
	std::wstring wszHomeServerName;
	std::wstring wszHomeServerDN;
	MailboxInfo * accountMailboxes;
	ULONG ulProfileConfigFlags;

};

struct EMSMdbSection
{
	std::wstring wszDisplayName;
	std::wstring wszDatafilePath;
	std::wstring wszSmtpAddress;
	BOOL bCachedModeEnabledOwner;
	BOOL bCachedModeEnabledShared;
	BOOL bCachedModeEnabledPublicFolders;
	int iCachedModeMonths;
	std::wstring szUserName;
	std::wstring szUserEmailSmtpAddress;
	ULONG ulMailboxCount;
	std::wstring wszRohProxyServer;
	std::wstring wszUnresolvedServer;
	std::wstring wszHomeServerName;
	std::wstring wszHomeServerDN;
	MailboxInfo * accountMailboxes;
	ULONG ulProfileConfigFlags;
};

struct AddressBookProviderInfo
{
	std::wstring szDisplayName;
	std::wstring szAbServerName;
	std::wstring szAbUsername;
	ULONG ulMaxEntries;
	ULONG ulTimeout;
	ULONG ulSlowTimeout;
	BOOL bUseSSL;
	BOOL bUsePSA;
	ULONG ulAbServerPort;
};

struct EasAccountInfo
{
	std::wstring szDisplayName;
	std::wstring szDataFilePath; // PR_PROFILE_OFFLINE_STORE_PATH_W
};

struct OtherServiceInfo
{
	std::wstring szDisplayName;
	std::wstring szServiceName;
};

struct ServiceInfo
{
	std::wstring wszServiceName;
	std::wstring wszDisplayName;
	ULONG ulServiceType; // MSEMS = SERVICETYPE_EXCHANGEACCOUNT; EMABLT = SERVICETYPE_ADDRESSBOOKPROVIDER; MSPST_MS/MSUPST_MS = SERVICETYPE_PST; EAS = SERVICETYPE_EASACCOUNT;
	BOOL bDefaultStore; // PR_RESOURCE_FLAGS = 
	ULONG ulResourceFlags; // PR_RESOURCE_FLAGS = SERVICE_DEFAULT_STORE
	EasAccountInfo * easAccountInfo;
	ExchangeAccountInfo * exchangeAccountInfo;
	AddressBookProviderInfo * addressBookProviderInfo;
	PstInfo * pstInfo;
	MAPIUID muidServiceUid;
};

struct ProfileInfo
{
	std::wstring wszProfileName;
	BOOL bDefaultProfile;
	ULONG ulServiceCount;
	ServiceInfo * profileServices;
};


#pragma once
