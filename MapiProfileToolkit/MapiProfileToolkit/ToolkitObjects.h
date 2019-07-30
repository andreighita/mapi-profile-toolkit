#pragma once

#include "stdafx.h"



#define Unspecified					= 0x00000000
#define ProfileAdd					= 0x00000001
#define ProfileClone				= 0x00000002
#define ProfileEdit					= 0x00000004
#define ProfileList					= 0x00000008
#define ProfileListAll				= 0x00000010
#define ProfileRemove				= 0x00000020
#define ProfileRemoveAll			= 0x00000040
#define ProfileUpdate				= 0x00000080
#define ProviderAdd					= 0x00000100
#define ProviderEdit				= 0x00000200
#define ProviderList				= 0x00000400
#define ProviderListAll				= 0x00000800
#define ProviderRemove				= 0x00001000
#define ProviderRemoveAll			= 0x00002000
#define ProviderUpdate				= 0x00004000
#define ServiceAdd					= 0x00008000
#define ServiceEdit					= 0x00010000
#define ServiceDisableCachedMode	= 0x00020000
#define ServiceEnableCachedMode		= 0x00040000
#define ServiceList					= 0x00080000
#define ServiceListAll				= 0x00100000
#define ServiceRemove				= 0x00200000
#define ServiceRemoveAll			= 0x00400000
#define ServiceUpdate				= 0x00800000
#define PromoteDelegate				= 0x01000000

enum ProfileMode
{
	Unknown = 0,
	Default,
	Specific,
	All
};

enum ConnectMode
{
	Unknown = 0,
	RoH,
	MoH
};

enum AddressBookAction
{
	Unknown = 0,
	ListSpecific,
	ListAll,
	Update,
	Create
};



enum ServiceMode
{
	Unspecified,
	Default,
	Specific,
	All
};

enum ServiceType
{
	Unknown,
	Mailbox,
	Pst,
	AddressBook
};

enum ProviderType 
{ 
	Unknown,
	PrimaryMailbox,
	Delegate,
	PublicFolder
};

enum ExportMode
{ 
	Unknown,
	Export,
	NoExport
};

enum 
{ 
	INPUTMODE_USERINPUT, 
	INPUTMODE_ACTIVEDIRECTORY 
};

// Make sure any changes in here are reflected in Logger.h as well
enum LoggingMode
{ 
	Unknown,
	None,
	Console,
	File,
	ConsoleAndFile
};

struct ProfileOptions
{
	ProfileMode profileMode;					// pm
	std::wstring wszProfileName;			// pn
	bool bSetDefaultProfile;				// pd
};

struct ServiceOptions
{
	bool bSetDefaultservice;				// ssds
	int iCachedModeMonths;					// scmm		| 0 = all; 1, 3, 6, 12 or 24 for the same number of months;
	int iServiceIndex;						// si		| service index from the list of services
	std::wstring wszAddressBookExternalUrl; // sabe 
	std::wstring wszAddressBookInternalUrl; // sabi
	std::wstring wszAutodiscoverUrl;		// sau
	std::wstring wszMailboxDisplayName;		// smdn
	std::wstring wszMailboxLegacyDN;		// smldn
	std::wstring wszMailStoreExternalUrl;	// smse
	std::wstring wszMailStoreInternalUrl;	// smsi
	std::wstring wszProfileName;			// spn
	std::wstring wszRohProxyServer;			// srps
	std::wstring wszServerDisplayName;		// ssdn
	std::wstring wszServerLegacyDN;			// ssldn
	std::wstring wszSmtpAddress;			// ssa
	std::wstring wszUnresolvedServer;		// sus
	std::wstring wszUnresolvedUser;			// suu
	ULONG ulCachedModeOwner;				// scmo		| 1 = disabled; 2 = enabled; 
	ULONG ulCachedModePublicFolder;			// scmpf	| 1 = disabled; 2 = enabled; 
	ULONG ulCachedModeShared;				// scms		| 1 = disabled; 2 = enabled; 
	ULONG ulConfigFlags;					// scfgf		| PR_PROFILE_CONFIG_FLAGS
	ConnectMode connectMode;					// scnctm		| ROH or MOH
	ULONG ulProfileMode;					// spm		| PROFILEMODE_DEFAULT = 1, PROFILEMODE_ONE = 2, PROFILEMODE_ALL = 3
	ULONG ulResourceFlags;					// srf		| PR_RESOURCES_FLAGS
	ServiceType serviceType;
	ServiceMode serviceMode;
};

struct MailboxOptions
{
	std::wstring wszProfileName;			// mpn		| Profile Name
	ULONG ulProfileMode;					// mpm		| PROFILEMODE_DEFAULT = 1, PROFILEMODE_ONE = 2, PROFILEMODE_ALL = 3
	ULONG ulServiceIndex;					// msi		| Service Index from 
	bool bDefaultService;					// mds		| Default service in profile
	std::wstring wszSmtpAddress;			// msa		| 
	std::wstring wszMailboxLegacyDN;		// mmldn
	std::wstring wszMailboxDisplayName;		// mmdn
	std::wstring wszServerLegacyDN;			// msldn
	std::wstring wszServerDisplayName;		// msdn
	std::wstring wszRohProxyServer;			// mrps
	std::wstring wszMailStoreExternalUrl;	// mmse
	std::wstring wszMailStoreInternalUrl;	// mmsi
	ULONG ulMailboxType;					
	ULONG ulMailboxAction;					
	ULONG ulRohProxyServerFlags;			// mrpsf
	ULONG ulRohProxyServerAuthPackage;		// mrpsap
};

struct AddressBookOptions
{
	ULONG ulRunningMode; // 1 = List one; 2 = List all; 2 = Update; 3 = Create 
	ULONG ulProfileMode; // 1 = default; 2 = specific;
	std::wstring szProfileName;
	std::wstring szABDisplayName;
	std::wstring	szConfigFilePath;
	std::wstring	szABServerName;
};

struct DataFileOptions
{
	int iPstIndex;
	ULONG ulPstType;
	std::wstring wszPstPath;
	std::wstring wszDisplayName;
	bool bMovePst;
	std::wstring wszPstOldPath;
	std::wstring wszPstNewPath;
};

struct UpdateSmtpAddress
{
	ULONG ulAdTimeout;
	ULONG ulInputMode;
	std::wstring szADsPAth;
	std::wstring szOldDomainName;
	std::wstring szNewDomainName;
};
struct RuntimeOptions
{
	ToolkitAction toolkitAction;
	LoggingMode loggingMode;
	std::wstring wszExportPath;
	ExportMode exportMode; // 0 = no export; 1 = export;
	std::wstring wszLogFilePath;
	int iOutlookVersion;
	ProfileOptions * profileOptions;
	ServiceOptions * serviceOptions;
	MailboxOptions * mailboxOptions;
	DataFileOptions * dataFileOptions;
	AddressBookOptions * addressBookOptions;
};

struct ScenarioAddMAilbox
{
	bool bLegacy;
};




