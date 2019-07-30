#pragma once

#include "stdafx.h"

#define WIN32_LEAN_AND_MEAN

#define ACTION_UNSPECIFIED					0x00000000

#define ACTION_PROFILE_ADD					0x00000001
#define ACTION_PROFILE_CLONE				0x00000002
#define ACTION_PROFILE_UPDATE				0x00000004
#define ACTION_PROFILE_LIST					0x00000008
#define ACTION_PROFILE_LISTALL				0x00000010
#define ACTION_PROFILE_REMOVE				0x00000020
#define ACTION_PROFILE_REMOVEALL			0x00000040
#define ACTION_PROFILE_SETDEFAULT			0x00000080
#define ACTION_PROFILE_PROMOTEDELEGATES		0x00000100
#define ACTION_PROFILE_PROMOTEONEDELEGATE	0x00000200

#define ACTION_PROVIDER_ADD					0x00000400
#define ACTION_PROVIDER_UPDATE				0x00000800
#define ACTION_PROVIDER_LIST				0x00001000
#define ACTION_PROVIDER_LISTALL				0x00002000
#define ACTION_PROVIDER_REMOVE				0x00004000
#define ACTION_PROVIDER_REMOVEALL			0x00008000

#define ACTION_SERVICE_ADD					0x00010000
#define ACTION_SERVICE_UPDATE				0x00020000
#define ACTION_SERVICE_SETCACHEDMODE		0x00040000
#define ACTION_SERVICE_LIST					0x00080000 
#define ACTION_SERVICE_LISTALL				0x00100000 
#define ACTION_SERVICE_REMOVE				0x00200000 
#define ACTION_SERVICE_REMOVEALL			0x00400000 

//#define ACTION_1								0x00800000
//#define ACTION_2								0x01000000
//#define ACTION_3								0x02000000 
//#define ACTION_4								0x04000000
//#define ACTION_5								0x08000000
//#define ACTION_6								0x10000000
//#define ACTION_7								0x20000000
//#define ACTION_8								0x40000000
//#define ACTION_9								0x80000000

typedef enum
{
		Mode_Unknown,
		Mode_Default,
		Mode_Specific,
		Mode_All
	}  ProfileMode;

	typedef ProfileMode ServiceMode;

	typedef enum 
	{
		ConnectMode_Unknown, 
		ConnectMode_RpcOverHttp = 1,
		ConnectMode_MapiOverHttp
	} ConnectMode;

	typedef enum 
	{
		ServiceType_Unknown,
		ServiceType_Mailbox ,
		ServiceType_Pst,
		ServiceType_AddressBook
	} ServiceType;

	typedef enum 
	{
		PrimaryMailbox = 1,
		Delegate,
		PublicFolder
	} ProviderType;

	typedef enum 
	{
		Export = 1,
		NoExport
	} ExportMode;

	typedef enum 
	{
		User = 1,
		Directory,
		File
	} InputMode;

	typedef enum 
	{
		LoggingModeNone = 1,
		LoggingModeConsole,
		LoggingModeFile,
		LoggingModeConsoleAndFile
	} LoggingMode;

	typedef enum
	{
		Enabled = 1,
		Disabled
	} CachedMode;

	typedef enum { logLevelInfo, logLevelWarning, logLevelError, logLevelSuccess, logLevelFailed, logLevelDebug } LogLevel;
	typedef enum  { logCallStatusSuccess, logCallStatusError, logCallStatusNoFile, logCallStatusLoggingDisabled } LogCallStatus;

	struct ServiceOptions;
	struct AddressBookOptions;
	struct ProviderOptions;
	struct DataFileOptions;

struct ProfileOptions
{
	ProfileMode profileMode;					// pm
	std::wstring wszProfileName;			// pn
	bool bSetDefaultProfile;				// pd
	ServiceOptions * serviceOptions;
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
	CachedMode cachedModeOwner;				// scmo		| 1 = disabled; 2 = enabled; 
	CachedMode cachedModePublicFolders;			// scmpf	| 1 = disabled; 2 = enabled; 
	CachedMode cachedModeShared;				// scms		| 1 = disabled; 2 = enabled; 
	ULONG ulConfigFlags;					// scfgf		| PR_PROFILE_CONFIG_FLAGS
	ConnectMode connectMode;					// scnctm		| ROH or MOH
	ULONG ulProfileMode;					// spm		| PROFILEMODE_DEFAULT = 1, PROFILEMODE_ONE = 2, PROFILEMODE_ALL = 3
	ULONG ulResourceFlags;					// srf		| PR_RESOURCES_FLAGS
	ServiceType serviceType;
	ServiceMode serviceMode;
	AddressBookOptions * addressBookOptions;
	ProviderOptions * providerOptions;
	DataFileOptions * dataFileOptions;
};

struct ProviderOptions
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
	ProviderType providerType;
	ULONG ulRohProxyServerFlags;			// mrpsf
	ULONG ulRohProxyServerAuthPackage;		// mrpsap
};

struct AddressBookOptions
{
	std::wstring wszProfileName;
	std::wstring wszABDisplayName;
	std::wstring wszConfigFilePath;
	std::wstring wszABServerName;
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
	InputMode inputMode;
	std::wstring szADsPAth;
	std::wstring szOldDomainName;
	std::wstring szNewDomainName;
};
struct RuntimeOptions
{
	ULONG action;
	LoggingMode loggingMode;
	std::wstring wszExportPath;
	ExportMode exportMode; // 0 = no export; 1 = export;
	std::wstring wszLogFilePath;
	int iOutlookVersion;
	ProfileOptions * profileOptions;



};

struct ScenarioAddMAilbox
{
	bool bLegacy;
};




