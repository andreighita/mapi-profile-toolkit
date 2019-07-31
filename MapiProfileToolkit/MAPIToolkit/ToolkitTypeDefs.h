#pragma once
#include <string>
#include <Windows.h>
#include <MAPIX.h>
#include <MAPIAux.h>

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
#define ACTION_SERVICE_CHANGEDATAFILEPATH	0x00800000

//#define ACTION_1								0x01000000
//#define ACTION_2								0x02000000 
//#define ACTION_3								0x04000000
//#define ACTION_4								0x08000000
//#define ACTION_5								0x10000000
//#define ACTION_6								0x20000000
//#define ACTION_7								0x40000000
//#define ACTION_8								0x80000000

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
	ServiceType_Mailbox,
	ServiceType_Pst,
	ServiceType_AddressBook,
	ServiceType_All
} ServiceType;

typedef enum
{
	ProviderType_PrimaryMailbox = 1,
	ProviderType_Delegate,
	ProviderType_PublicFolder,
	ProviderType_All
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
	CachedMode_Enabled = 1,
	CachedMode_Disabled
} CachedMode;

typedef enum { logLevelInfo, logLevelWarning, logLevelError, logLevelSuccess, logLevelFailed, logLevelDebug } LogLevel;
typedef enum { logCallStatusSuccess, logCallStatusError, logCallStatusNoFile, logCallStatusLoggingDisabled } LogCallStatus;


struct UpdateSmtpAddress
{
	ULONG ulAdTimeout;
	InputMode inputMode;
	std::wstring szADsPAth;
	std::wstring szOldDomainName;
	std::wstring szNewDomainName;
};





