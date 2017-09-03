#pragma once

#include "stdafx.h"

enum 
{ 
	SCENARIO_PROFILE = 1,
	SCENARIO_SERVICE,
	SCENARIO_MAILBOX,
	SCENARIO_DATAFILE,
	SCENARIO_LDAP,
	SCENARIO_CUSTOM
};

enum { ACTIONTYPE_STANDARD, ACTIONTYPE_CUSTOM };

enum { STANDARDACTION_ADD, STANDARDACTION_REMOVE, STANDARDACTION_EDIT, STANDARDACTION_LIST };

enum 
{ 
	CUSTOMACTION_PROMOTEMAILBOXTOSERVICE, 
	CUSTOMACTION_EDITCACHEDMODECONFIGURATION,
	CUSTOMACTION_UPDATESMTPADDRESS,
	CUSTOMACTION_CHANGEPSTLOCATION,
	CUSTOMACTION_REMOVEORPHANEDDATAFILES
};

enum 
{ 
	PROFILEMODE_DEFAULT = 1, 
	PROFILEMODE_ONE, 
	PROFILEMODE_ALL 
};

enum 
{ 
	SERVICEMODE_DEFAULT = 1, 
	SERVICEMODE_ONE, 
	SERVICEMODE_ALL 
};
enum 
{ 
	MAILBOXMODE_DEFAULT = 1, 
	MAILBOXMODE_ONE, 
	MAILBOXMODE_ALL 
};
enum 
{ 
	CONNECT_ROH = 1, 
	CONNECT_MOH 
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
enum 
{ 
	LOGGINGODE_CONSOLE_AND_FILE, 
	LOGGINGMODE_CONSOLE, 
	LOGGINGMODE_FILE, 
	LOGGINGMODE_NONE 
};

struct ProfileOptions
{
	ULONG ulProfileMode;					// pm
	std::wstring wszProfileName;			// pn
	bool bSetDefaultProfile;				// pd
};

struct ServiceOptions
{
	bool bDefaultservice;					// sds
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
	ULONG iOutlookVersion;					// sov		| 2007, 2010, 2013 or 2016
	ULONG ulProfileMode;					// spm		| PROFILEMODE_DEFAULT = 1, PROFILEMODE_ONE = 2, PROFILEMODE_ALL = 3
	ULONG ulResourceFlags;					// srf		| PR_RESOURCES_FLAGS

};

struct MailboxOptions
{
	std::wstring wszProfileName;			// mpn		| Profile Name
	ULONG ulProfileMode;					// mpm		| PROFILEMODE_DEFAULT = 1, PROFILEMODE_ONE = 2, PROFILEMODE_ALL = 3
	ULONG ulServiceIndex;					// msi		| Service Index from 
	bool bDefaultService;					// mds		| Default service in profile
	int iOutlookVersion;					// mov		| 2007, 2010, 2013 or 2016
	std::wstring wszSmtpAddress;			// msa		| 
	std::wstring wszMailboxLegacyDN;		// mmldn
	std::wstring wszMailboxDisplayName;		// mmdn
	std::wstring wszServerLegacyDN;			// msldn
	std::wstring wszServerDisplayName;		// msdn
};

struct LdapOptions
{

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
	ProfileOptions * profileOptions;
	ServiceOptions * serviceOptions;
	MailboxOptions * mailboxOptions;
	DataFileOptions * dataFileOptions;
	LdapOptions * ldapOptions;
};

struct ScenarioAddMAilbox
{
	bool bLegacy;
};




