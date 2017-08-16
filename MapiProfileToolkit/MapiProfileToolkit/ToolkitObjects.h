#pragma once

#include "stdafx.h"

enum {
	SCENARIO_CREATEPROFILE = 1,
	SCENARIO_CLONEPROFILE,
	SCENARIO_ADDSERVICE,
	SCENARIO_ADDMAILBOX,
	SCENARIO_ADDPST,
	SCENARIO_ADDLDAPADDRESSLIST,
	SCENARIO_SETCACHEDMODECONFIGURATION,
	SCENARIO_LISTPROFILECONFIGURATION,
	SCENARIO_LISTPSTS,
	SCENARIO_CONVERTADDITIONALMAILBOXESTOACCOUNTS
};

enum { PROFILEMODE_DEFAULT = 1, PROFILEMODE_SPECIFIC, PROFILEMODE_ALL };
enum { READWRITEMODE_READ = 1, READWRITEMODE_WRITE };
enum { EXPORTMODE_NOEXPORT = 0, EXPORTMODE_EXPORT };
enum { RUNNINGMODE_PROFILE = 0, RUNNINGMODE_PST};
enum { INPUTMODE_USERINPUT, INPUTMODE_ACTIVEDIRECTORY };
enum { LOGGINGODE_CONSOLE_AND_FILE, LOGGINGMODE_CONSOLE, LOGGINGMODE_FILE, LOGGINGMODE_NONE};

struct RuntimeOptions
{
	ULONG ulRunningMode; // 1 = read; 2 = write; 
	ULONG ulReadWriteMode;
	ULONG ulProfileMode; // 1 = default; 2 = specific; 3 = all;
	ULONG ulLoggingMode;
	std::wstring szProfileName;
	ULONG ulServiceIndex;
	ULONG ulCachedModeOwner; // 1 = disabled; 2 = enabled; 
	ULONG ulCachedModeShared; // 1 = disabled; 2 = enabled; 
	ULONG ulCachedModePublicFolder; // 1 = disabled; 2 = enabled; 
	int iCachedModeMonths; // 0 = all; 1, 3, 6, 12 or 24 for the same number of months; 
	std::wstring szExportPath;
	BOOL bExportMode; // 0 = no export; 1 = export;
	BOOL bNoHeader;
	ULONG ulAdTimeout;
	std::wstring szADsPAth;
	std::wstring szLogFilePath;
	std::wstring szOldDomainName;
	std::wstring szNewDomainName;
	std::wstring szPstOldPath;
	std::wstring szPstNewPath;
	bool bPstMoveFiles;
};

