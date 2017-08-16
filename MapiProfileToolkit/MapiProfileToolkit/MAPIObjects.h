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

#define PR_PST_CONFIG_FLAGS PROP_TAG(PT_LONG, 0x6770)
#define PR_PST_PATH_W PROP_TAG(PT_UNICODE, 0x6700)

enum { CACHEDMODE_DISABLED = 1, CACHEDMODE_ENABLED };
enum { SERVICE_PRIMARY = 1, SERVICE_SECONDARY };
enum { ENTRYTYPE_UNKNOWN, ENTRYTYPE_PRIMARY, ENTRYTYPE_DELEGATE, ENTRYTYPE_PUBLIC_FOLDERS };
enum { SERVICETYPE_OTHER, SERVICETYPE_EXCHANGEACCOUNT, SERVICETYPE_ADDRESSBOOKPROVIDER, SERVICETYPE_PST, SERVICETYPE_EASACCOUNT };
enum { PSTTYPE_ANSI = 0, PSTTYPE_UNICODE = 0x80000000};

struct MailboxInfo
{
	std::wstring szDisplayName;
	BOOL bDefaultMailbox;
	ULONG ulEntryType;
};

struct PstInfo
{
	std::wstring szDisplayName;
	std::wstring szPstPath; // PR_PST_PATH_W
	ULONG ulPstType; // PR_PST_CONFIG_FLAGS
};

struct MapiProperty
{
	std::wstring szPropertyName;
	std::wstring szPropertyTag;
	std::wstring szPropertyValue;
};

struct ProviderInfo
{
	MapiProperty * mapiProperties;
};

struct ExchangeAccountInfo
{
	std::wstring szServiceDisplayName;
	std::wstring szDatafilePath;
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
	std::wstring szServiceName;
	ULONG ulServiceType; // MSEMS = SERVICETYPE_EXCHANGEACCOUNT; EMABLT = SERVICETYPE_ADDRESSBOOKPROVIDER; MSPST_MS/MSUPST_MS = SERVICETYPE_PST; EAS = SERVICETYPE_EASACCOUNT;
	BOOL bDefaultStore; // PR_RESOURCE_FLAGS = SERVICE_DEFAULT_STORE
	EasAccountInfo * easAccountInfo;
	ExchangeAccountInfo * exchangeAccountInfo;
	AddressBookProviderInfo * addressBookProviderInfo;
	PstInfo * pstInfo;
};

struct ProfileInfo
{
	std::wstring szProfileName;
	BOOL bDefaultProfile;
	ULONG ulServiceCount;
	ServiceInfo * profileServices;
};


#pragma once
