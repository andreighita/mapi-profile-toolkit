#pragma once

#include "stdafx.h"

#define PROFILEMODE_DEFAULT (ULONG)0 
#define PROFILEMODE_ONE (ULONG)1
#define PROFILEMODE_ALL (ULONG)2

#define CONNECT_ROH (ULONG)1
#define CONNECT_MOH (ULONG)2

// 1 = List one; 2 = List all; 3 = Update; 4 = Create 
#define ADDRESSBOOK_LIST_ONE (ULONG)1 
#define ADDRESSBOOK_LIST_ALL (ULONG)2
#define ADDRESSBOOK_UPDATE (ULONG)3
#define ADDRESSBOOK_CREATE (ULONG)4

enum 
{ 
	SCENARIO_PROFILE = 1,
	SCENARIO_SERVICE,
	SCENARIO_MAILBOX,
	SCENARIO_DATAFILE,
	SCENARIO_LDAP,
	SCENARIO_CUSTOM
};

enum { ACTIONTYPE_STANDARD = 1, ACTIONTYPE_CUSTOM };

enum { STANDARDACTION_ADD = 1, STANDARDACTION_REMOVE, STANDARDACTION_EDIT, STANDARDACTION_LIST };

enum { ACTION_UNKNOWN = 0, ACTION_ADD = 1, ACTION_REMOVE, ACTION_EDIT, ACTION_LIST, ACTION_UPDATE, ACTION_PROMOTEDELEGATE, ACTION_ADDDELEGATE, ACTION_CLONE, ACTION_SIMPLECLONE, ACTION_ENABLECACHEDMODE};

enum 
{ 
	CUSTOMACTION_PROMOTEMAILBOXTOSERVICE = 1, 
	CUSTOMACTION_EDITCACHEDMODECONFIGURATION,
	CUSTOMACTION_UPDATESMTPADDRESS,
	CUSTOMACTION_CHANGEPSTLOCATION,
	CUSTOMACTION_REMOVEORPHANEDDATAFILES
};

enum 
{ 
	SERVICEMODE_DEFAULT = 1, 
	SERVICEMODE_ONE, 
	SERVICEMODE_ALL 
};

enum
{
	SERVICETYPE_OTHER,
	SERVICETYPE_MAILBOX,
	SERVICETYPE_PST,
	SERVICETYPE_ADDRESSBOOK
};

enum
{
	MAILBOXTYPE_PRIMARY = 1,
	MAILBOXTYPE_DELEGATE,
	MAILBOXTYPE_PUBLICFOLDER
};

enum 
{ 
	MAILBOXMODE_DEFAULT = 1, 
	MAILBOXMODE_ONE, 
	MAILBOXMODE_ALL 
};

enum 
{ 
	EXPORTMODE_NOEXPORT = 0, 
	EXPORTMODE_EXPORT 
};

enum 
{ 
	INPUTMODE_USERINPUT, 
	INPUTMODE_ACTIVEDIRECTORY 
};

// Make sure any changes in here are reflected in Logger.h as well
enum 
{ 
	LOGGINGMODE_NONE,
	LOGGINGMODE_CONSOLE, 
	LOGGINGMODE_FILE, 
	LOGGINGODE_CONSOLE_AND_FILE
};

struct ProfileOptions
{
	ULONG ulProfileMode;					// pm
	std::wstring wszProfileName;			// pn
	bool bSetDefaultProfile;				// pd
	ULONG ulProfileAction;
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
	ULONG ulConnectMode;					// scnctm		| ROH or MOH
	ULONG ulProfileMode;					// spm		| PROFILEMODE_DEFAULT = 1, PROFILEMODE_ONE = 2, PROFILEMODE_ALL = 3
	ULONG ulResourceFlags;					// srf		| PR_RESOURCES_FLAGS
	ULONG ulServiceType;
	ULONG ulServiceAction;
	ULONG ulServiceMode;
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
	ULONG ulScenario;
	ULONG ulActionType;
	ULONG ulAction;
	ULONG ulLoggingMode;
	std::wstring wszExportPath;
	BOOL bExportMode; // 0 = no export; 1 = export;
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




