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

// ProfileToolkit.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

#include <MAPIX.h>
#include <MAPIUtil.h>
#include "ProfileFunctions.h"
#include "MapiProfileToolkit.h"
#include "ToolkitObjects.h"
#include <iostream>
#include <string>
#include <utility>
#include <algorithm>  
#include "XMLHelper.h"
#include "RegistryHelper.h"
#include "Logger.h"

BOOL Is64BitProcess(void)
{
#if defined(_WIN64)
	return TRUE;   // 64-bit program
#else
	return FALSE;
#endif
}

BOOL _cdecl IsCorrectBitness()
{

	std::wstring szOLVer = L"";
	std::wstring szOLBitness = L"";
	szOLVer = GetStringValue(HKEY_CLASSES_ROOT, TEXT("Outlook.Application\\CurVer"), NULL);
	if (szOLVer != L"")
	{
		if (szOLVer == L"Outlook.Application.16")
		{
			szOLBitness = GetStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\16.0\\Outlook"), TEXT("Bitness"));
			if (szOLBitness != L"")
			{
				if (szOLBitness == L"x64")
				{
					if (Is64BitProcess())
						return TRUE;
				}
				else if (szOLBitness == L"x86")
				{
					if (Is64BitProcess())
						return FALSE;
					else
						return TRUE;
				}
				else return FALSE;
			}
		}
		else if (szOLVer == L"Outlook.Application.15")
		{
			szOLBitness = GetStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\15.0\\Outlook"), TEXT("Bitness"));
			if (szOLBitness != L"")
			{
				if (szOLBitness == L"x64")
				{
					if (Is64BitProcess())
						return TRUE;
				}
				else if (szOLBitness == L"x86")
				{
					if (Is64BitProcess())
						return FALSE;
					else
						return TRUE;
				}
				else return FALSE;
			}
		}
		else if (szOLVer == L"Outlook.Application.14")
		{
			szOLBitness = GetStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\14.0\\Outlook"), TEXT("Bitness"));
			if (szOLBitness != L"")
			{
				if (szOLBitness == L"x64")
				{
					if (Is64BitProcess())
						return TRUE;
				}
				else if (szOLBitness == L"x86")
				{
					if (Is64BitProcess())
						return FALSE;
					else
						return TRUE;
				}
				else return FALSE;
			}
		}
		else return FALSE;
		return FALSE;
	}
	else return FALSE;
}

BOOL RunScenario(int argc, _TCHAR* argv[], RuntimeOptions * pRunOpts)
{
	if (!pRunOpts) return FALSE;
	ZeroMemory(pRunOpts, sizeof(RuntimeOptions));

	for (int i = 1; i < argc; i++)
	{
		switch (argv[i][0])
		{
		case '#':
			if (0 == argv[i][1])
			{
				// Bad argument - get out of here
				return false;
			}
			switch (tolower(argv[i][1]))
			{
			case '4':
				if (tolower(argv[i][2]) == 'c')
				{
					if (i + 1 < argc)
					{
						std::wstring runningMode = argv[i + 1];
						std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
						if (runningMode == L"all")
						{
							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_ALL;
							i++;
						}
						else if (runningMode == L"one")
						{
							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_SPECIFIC;
							i++;
						}
						else if (runningMode == L"default")
						{
							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_DEFAULT;
							i++;
						}
						else
						{
							return false;
						}
					}
				}
				else if (tolower(argv[i][2]) == 'n')
				{
					if (i + 1 < argc)
					{
						pRunOpts->szProfileName = argv[i + 1];
						i++;
					}
					else return false;
				}
				else return false;
				break;
			case 'c':
				if (tolower(argv[i][2]) == 'm')
				{
					if ((tolower(argv[i][3]) == 'o'))
					{
						if (i + 1 < argc)
						{
							std::wstring runningMode = argv[i + 1];
							std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
							if (runningMode == L"enable")
							{
								pRunOpts->ulCachedModeOwner = CACHEDMODE_ENABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else if (runningMode == L"disable")
							{
								pRunOpts->ulCachedModeOwner = CACHEDMODE_DISABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else
							{
								return false;
							}
						}
						else return false;
					}
					else if ((tolower(argv[i][3]) == 'm'))
					{
						if (i + 1 < argc)
						{
							pRunOpts->iCachedModeMonths = _wtoi(argv[i + 1]);
							i++;
						}
						else return false;
					}
					else if ((tolower(argv[i][3]) == 's'))
					{
						if (i + 1 < argc)
						{

							std::wstring runningMode = argv[i + 1];
							std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
							if (runningMode == L"enable")
							{
								pRunOpts->ulCachedModeShared = CACHEDMODE_ENABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else if (runningMode == L"disable")
							{
								pRunOpts->ulCachedModeShared = CACHEDMODE_DISABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else
							{
								return false;
							}
						}
						else return false;
					}
					else if ((tolower(argv[i][3]) == 'p'))
					{
						if (i + 1 < argc)
						{

							std::wstring runningMode = argv[i + 1];
							std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
							if (runningMode == L"enable")
							{
								pRunOpts->ulCachedModePublicFolder = CACHEDMODE_ENABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else if (runningMode == L"disable")
							{
								pRunOpts->ulCachedModePublicFolder = CACHEDMODE_DISABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else
							{
								return false;
							}
						}
						else return false;
					}
					else return false;
					break;
				}
				else return false;
				break;
			case 's':
				if (tolower(argv[i][2]) == 'i')
				{
					if (i + 1 < argc)
					{
						pRunOpts->ulServiceIndex = _wtoi(argv[i + 1]);
						i++;
					}
					else return false;
				}
				else return false;
				break;
			case 'r':
				if (tolower(argv[i][2]) == 'm')
				{
					if (i + 1 < argc)
					{
						std::wstring runningMode = argv[i + 1];
						std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
						if (runningMode == L"profile")
						{
							pRunOpts->ulRunningMode = (ULONG)RUNNINGMODE_PROFILE;
							i++;
						}
						else if (runningMode == L"pst")
						{
							pRunOpts->ulRunningMode = (ULONG)RUNNINGMODE_PST;
							i++;
						}
						else if (runningMode == L"addressbook")
						{
							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_DEFAULT;
							i++;
						}
						else
						{
							return false;
						}
					}
				}
				else return false;
				break;
			case 'n':
				if (tolower(argv[i][2]) == 'h')
				{
					pRunOpts->bNoHeader = true;
					i++;
				}
				else if (tolower(argv[i][2]) == 'p')
				{
					pRunOpts->szPstNewPath = argv[i + 1];
					i++;
				}
				else return false;
				break;
			case 'm':
				if (tolower(argv[i][2]) == 'f')
				{
					pRunOpts->bPstMoveFiles = true;
					i++;
				}
				else return false;
				break;
			case 'e':
				if (tolower(argv[i][2]) == 'p')
				{
					std::wstring exportPath = argv[i + 1];
					std::transform(exportPath.begin(), exportPath.end(), exportPath.begin(), ::tolower);
					pRunOpts->szExportPath = exportPath;
					pRunOpts->bExportMode = EXPORTMODE_EXPORT;
					i++;
				}
				else return false;
				break;
			case 'o':
				if (tolower(argv[i][2]) == 'p')
				{
					pRunOpts->szPstOldPath = argv[i + 1];
					i++;
				}
				else return false;
				break;
			case 'l':
				if (tolower(argv[i][2]) == 'm')
				{
					pRunOpts->ulLoggingMode = _wtoi(argv[i + 1]);
					i++;
				}
				else if (tolower(argv[i][2]) == 'p')
				{
					pRunOpts->szLogFilePath = argv[i + 1];
					Logger::Initialise(pRunOpts->szLogFilePath);
					i++;
				}
				else return false;
				break;
			case '?':
			default:
				// display help
				pRunOpts->ulLoggingMode = LOGGINGMODE_CONSOLE;
				return false;
				break;
			}
		}
	}

	if (pRunOpts->ulLoggingMode == loggingModeConsoleandFile && pRunOpts->szLogFilePath.empty())
	{
		pRunOpts->ulLoggingMode = loggingModeConsole;
	}

	if (pRunOpts->ulLoggingMode == loggingModeFile && pRunOpts->szLogFilePath.empty())
	{
		pRunOpts->ulLoggingMode = loggingModeNone;
	}

	// If no profile mode or index or name specified then use default
	if (!pRunOpts->szProfileName.empty())
	{
		if (pRunOpts->ulProfileMode == 0)
		{
			pRunOpts->ulProfileMode = PROFILEMODE_DEFAULT;
		}

	}

	// If running mode is RUNNINGMODE_WRITE then expect a profile section name or a service index or a service type
	if (pRunOpts->ulReadWriteMode == READWRITEMODE_WRITE)
	{
		if (pRunOpts->ulServiceIndex >= 1)
		{
			return true;
		}
		else return false;
	}
	return true;
}

BOOL ParseArgs(int argc, _TCHAR* argv[], RuntimeOptions * pRunOpts)
{
	if (!pRunOpts) return FALSE;

	

	// Setting running mode to Read as a default
	pRunOpts->ulReadWriteMode = READWRITEMODE_READ;
	pRunOpts->ulRunningMode = RUNNINGMODE_PROFILE;
	pRunOpts->bExportMode = EXPORTMODE_NOEXPORT;
	pRunOpts->bPstMoveFiles = false;
	pRunOpts->ulLoggingMode = loggingModeNone;
	for (int i = 1; i < argc; i++)
	{
		switch (argv[i][0])
		{
		case '-':
		case '/':
		case '\\':
			if (0 == argv[i][1])
			{
				// Bad argument - get out of here
				return false;
			}
			switch (tolower(argv[i][1]))
			{
			case 'p':
				if (tolower(argv[i][2]) == 'm')
				{
					if (i + 1 < argc)
					{
						std::wstring runningMode = argv[i + 1];
						std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
						if (runningMode == L"all")
						{
							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_ALL;
							i++;
						}
						else if (runningMode == L"one")
						{
							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_SPECIFIC;
							i++;
						}
						else if (runningMode == L"default")
						{
							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_DEFAULT;
							i++;
						}
						else
						{
							return false;
						}
					}
				}
				else if (tolower(argv[i][2]) == 'n')
				{
					if (i + 1 < argc)
					{
						pRunOpts->szProfileName = argv[i + 1];
						i++;
					}
					else return false;
				}
				else return false;
				break;
			case 'c':
				if (tolower(argv[i][2]) == 'm')
				{
					if ((tolower(argv[i][3]) == 'o'))
					{
						if (i + 1 < argc)
						{
							std::wstring runningMode = argv[i + 1];
							std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
							if (runningMode == L"enable")
							{
								pRunOpts->ulCachedModeOwner = CACHEDMODE_ENABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else if (runningMode == L"disable")
							{
								pRunOpts->ulCachedModeOwner = CACHEDMODE_DISABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else
							{
								return false;
							}
						}
						else return false;
					}
					else if ((tolower(argv[i][3]) == 'm'))
					{
						if (i + 1 < argc)
						{
							pRunOpts->iCachedModeMonths = _wtoi(argv[i + 1]);
							i++;
						}
						else return false;
					}
					else if ((tolower(argv[i][3]) == 's'))
					{
						if (i + 1 < argc)
						{

							std::wstring runningMode = argv[i + 1];
							std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
							if (runningMode == L"enable")
							{
								pRunOpts->ulCachedModeShared = CACHEDMODE_ENABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else if (runningMode == L"disable")
							{
								pRunOpts->ulCachedModeShared = CACHEDMODE_DISABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else
							{
								return false;
							}
						}
						else return false;
					}
					else if ((tolower(argv[i][3]) == 'p'))
					{
						if (i + 1 < argc)
						{

							std::wstring runningMode = argv[i + 1];
							std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
							if (runningMode == L"enable")
							{
								pRunOpts->ulCachedModePublicFolder = CACHEDMODE_ENABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else if (runningMode == L"disable")
							{
								pRunOpts->ulCachedModePublicFolder = CACHEDMODE_DISABLED;
								pRunOpts->ulReadWriteMode = READWRITEMODE_WRITE;
								i++;
							}
							else
							{
								return false;
							}
						}
						else return false;
					}
					else return false;
					break;
				}
				else return false;
				break;
			case 's':
				if (tolower(argv[i][2]) == 'i')
				{
					if (i + 1 < argc)
					{
						pRunOpts->ulServiceIndex = _wtoi(argv[i + 1]);
						i++;
					}
					else return false;
				}
				else return false;
				break;
			case 'r':
				if (tolower(argv[i][2]) == 'm')
				{
					if (i + 1 < argc)
					{
						std::wstring runningMode = argv[i + 1];
						std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
						if (runningMode == L"profile")
						{
							pRunOpts->ulRunningMode = (ULONG)RUNNINGMODE_PROFILE;
							i++;
						}
						else if (runningMode == L"pst")
						{
							pRunOpts->ulRunningMode = (ULONG)RUNNINGMODE_PST;
							i++;
						}
						else if (runningMode == L"addressbook")
						{
							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_DEFAULT;
							i++;
						}
						else
						{
							return false;
						}
					}
				}
				else return false;
				break;
			case 'n':
				if (tolower(argv[i][2]) == 'h')
				{
					pRunOpts->bNoHeader = true;
					i++;
				}
				else if (tolower(argv[i][2]) == 'p')
				{
					pRunOpts->szPstNewPath = argv[i + 1];
					i++;
				}
				else return false;
				break;
			case 'm':
				if (tolower(argv[i][2]) == 'f')
				{
					pRunOpts->bPstMoveFiles = true;
					i++;
				}
				else return false;
				break;
			case 'e':
				if (tolower(argv[i][2]) == 'p')
				{
					std::wstring exportPath = argv[i + 1];
					std::transform(exportPath.begin(), exportPath.end(), exportPath.begin(), ::tolower);
					pRunOpts->szExportPath = exportPath;
					pRunOpts->bExportMode = EXPORTMODE_EXPORT;
					i++;
				}
				else return false;
				break;
			case 'o':
				if (tolower(argv[i][2]) == 'p')
				{
					pRunOpts->szPstOldPath = argv[i + 1];
					i++;
				}
				else return false;
				break;
			case 'l':
				if (tolower(argv[i][2]) == 'm')
				{
					pRunOpts->ulLoggingMode = _wtoi(argv[i + 1]);
					i++;
				}
				else if (tolower(argv[i][2]) == 'p')
				{
					pRunOpts->szLogFilePath = argv[i + 1];
					Logger::Initialise(pRunOpts->szLogFilePath);
					i++;
				}
				else return false;
				break;
			case '?':
			default:
				// display help
				pRunOpts->ulLoggingMode = LOGGINGMODE_CONSOLE;
				return false;
				break;
			}
		}
	}

	if (pRunOpts->ulLoggingMode == loggingModeConsoleandFile && pRunOpts->szLogFilePath.empty())
	{
		pRunOpts->ulLoggingMode = loggingModeConsole;
	}

	if (pRunOpts->ulLoggingMode == loggingModeFile && pRunOpts->szLogFilePath.empty())
	{
		pRunOpts->ulLoggingMode = loggingModeNone;
	}

	// If no profile mode or index or name specified then use default
	if (!pRunOpts->szProfileName.empty())
	{
		if (pRunOpts->ulProfileMode == 0)
		{
			pRunOpts->ulProfileMode = PROFILEMODE_DEFAULT;
		}

	}

	// If running mode is RUNNINGMODE_WRITE then expect a profile section name or a service index or a service type
	if (pRunOpts->ulReadWriteMode == READWRITEMODE_WRITE)
	{
		if (pRunOpts->ulServiceIndex >= 1)
		{
			return true;
		}
		else return false;
	}
	return true;
}

void DisplayUsage()
{
	printf("ProfileToolkit - Profile Examination Tool\n");
	printf("    Lists profile settings and optionally enables or disables cached exchange \n");
	printf("    mode.\n");
	printf("\n");
	printf("Usage: ProfileToolkit [-?] [-pm <all, one, default>] [-pn profilename] \n");
	printf("       [-si serviceIndex] [-cmo <enable, disable>] [-cms <enable, disable>] \n");
	printf("       [-cmp <enable, disable>]	[-cmm <0, 1, 3, 6, 12, 24>] [-ep exportpath]\n");
	printf("\n");
	printf("Options:\n");
	printf("       -pm:    \"all\" to process all profiles.\n");
	printf("               \"default\" to process the default profile.\n");
	printf("               \"one\" to process a specific profile. Prifile Name needs to be \n");
	printf("               specified using -pn.\n");
	printf("               Default profile will be used if -pm is not used.\n");
	printf("       -pn:    Name of the profile to process.\n");
	printf("               Default profile will be used if -pn is not used.\n");
	printf("\n");
	printf("       -si:    Index of the account to process from previous export.\n");
	printf("       	       Must be used in conjunction with -pm one -pn profile or -pm default.\n");
	printf("\n");
	printf("       -cmo:   \"enable\" or \"disable\" for enabling or disabling cached Exchange \n");
	printf("               mode on the owner's mailbox.\n");
	printf("       	       Must be used in conjunction with -pm one -pn profile and -si index.\n");
	printf("       -cms:   \"enable\" or \"disable\" for enabling or disabling cached Exchange \n");
	printf("               mode on shared folders (delegate).\n");
	printf("       	       Must be used in conjunction with -pm one -pn profile and -si index.\n");
	printf("       -cmp:   \"enable\" or \"disable\" for enabling or disabling cached Exchange \n");
	printf("               mode on public folders favorites.\n");
	printf("       	       Must be used in conjunction with -pm one -pn profile and -si index.\n");
	printf("       -cmm:   0 for all or 1, 3, 6, 12 or 24 for the same number of months to sync\n");
	printf("       	       Must be used in conjunction with -pm one -pn profile, -si index and.\n");
	printf("       	       -cmo enable.\n");
	printf("\n");
	printf("       -ep:    exportPath for exporting settings to disk.\n");
	printf("\n");
	printf("       -?      Displays this usage information.\n");
}


LoggingMode loggingMode;

static std::wofstream ofsLogFile;
static std::wstring szLogFilePath;
static bool bIsLogFileOpen;


void _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hRes = S_OK;

	// Using the toolkip options to manage the runtime options
	RuntimeOptions tkOptions = { 0 };
	loggingMode = loggingModeNone;
	// Parse the command line arguments
	if (!ParseArgs(argc, argv, &tkOptions))
	{
		if (tkOptions.ulLoggingMode != loggingModeNone)
		{
			DisplayUsage();
		}
		return;
	}

	loggingMode = LoggingMode(tkOptions.ulLoggingMode);
	// Check the curren't process' bitness vs Outlook's bitness and only run it if matched to avoid MAPI dialog boxes.
	if (!IsCorrectBitness())
	{
		Logger::Write(logLevelFailed, L"Unable to resolve bitness or bitness not matched.", loggingMode);
		return;
	}
	Logger::Write(logLevelSuccess, L"Bitness matched.", loggingMode);

	try
	{
		MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
		if (SUCCEEDED(MAPIInitialize(&MAPIINIT)))
		{
			Logger::Write(logLevelSuccess, L"MAPI Initialised", loggingMode);
			switch (tkOptions.ulRunningMode)
			{
			case RUNNINGMODE_PROFILE:
				switch (tkOptions.ulProfileMode)
				{
				case PROFILEMODE_ALL:
					if (tkOptions.ulReadWriteMode == READWRITEMODE_READ)
					{
						ULONG ulProfileCount = GetProfileCount(loggingMode);
						ProfileInfo * profileInfo = new ProfileInfo[ulProfileCount];
						ZeroMemory(profileInfo, sizeof(ProfileInfo) * ulProfileCount);
						Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for all profiles", loggingMode);
						EC_HRES_LOG(GetProfiles(ulProfileCount, profileInfo, loggingMode), loggingMode);
						if (tkOptions.szExportPath != L"")
						{
							Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles", loggingMode);
							ExportXML(ulProfileCount, profileInfo, tkOptions.szExportPath, loggingMode);
						}
						else
						{
							Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles", loggingMode);
							ExportXML(ulProfileCount, profileInfo, L"", loggingMode);
						}
					}
					break;
				case PROFILEMODE_SPECIFIC:
					if (tkOptions.ulReadWriteMode == READWRITEMODE_READ)
					{
						ProfileInfo profileInfo;
						Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for profile: " + tkOptions.szProfileName, loggingMode);
						EC_HRES_LOG(GetProfile((LPWSTR)tkOptions.szProfileName.c_str(), &profileInfo, loggingMode), loggingMode);
						if (tkOptions.szExportPath != L"")
						{
							Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile", loggingMode);
							ExportXML(1, &profileInfo, tkOptions.szExportPath, loggingMode);
						}
						else
						{
							Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile", loggingMode);
							ExportXML(1, &profileInfo, L"", loggingMode);
						}
					}
					else if (tkOptions.ulReadWriteMode == READWRITEMODE_WRITE)
					{
						Logger::Write(logLevelInfo, L"Updating cached mode configuration on profile: " + tkOptions.szProfileName, loggingMode);
						EC_HRES_LOG(UpdateCachedModeConfig((LPSTR)tkOptions.szProfileName.c_str(), tkOptions.ulServiceIndex, tkOptions.ulCachedModeOwner, tkOptions.ulCachedModeShared, tkOptions.ulCachedModePublicFolder, tkOptions.iCachedModeMonths, loggingMode), loggingMode);
					}

					break;
				case PROFILEMODE_DEFAULT:
					std::wstring szDefaultProfileName = GetDefaultProfileName(loggingMode);
					if (!szDefaultProfileName.empty())
					{
						tkOptions.szProfileName = szDefaultProfileName;
					}
					if (tkOptions.ulReadWriteMode == READWRITEMODE_READ)
					{
						ProfileInfo profileInfo;
						Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for default profile: " + tkOptions.szProfileName, loggingMode);
						EC_HRES_LOG(GetProfile((LPWSTR)tkOptions.szProfileName.c_str(), &profileInfo, loggingMode), loggingMode);
						if (tkOptions.szExportPath != L"")
						{
							Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile", loggingMode);
							ExportXML(1, &profileInfo, tkOptions.szExportPath, loggingMode);
						}
						else
						{
							Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile", loggingMode);
							ExportXML(1, &profileInfo, L"", loggingMode);
						}
					}
					else if (tkOptions.ulReadWriteMode == READWRITEMODE_WRITE)
					{
						Logger::Write(logLevelInfo, L"Updating cached mode configuration on default profile: " + tkOptions.szProfileName, loggingMode);
						EC_HRES_LOG(UpdateCachedModeConfig((LPSTR)tkOptions.szProfileName.c_str(), tkOptions.ulServiceIndex, tkOptions.ulCachedModeOwner, tkOptions.ulCachedModeShared, tkOptions.ulCachedModePublicFolder, tkOptions.iCachedModeMonths, loggingMode), loggingMode);
					}
					break;
				}
				break;
			case RUNNINGMODE_PST:
				if (tkOptions.szPstOldPath.empty())
				{
					EC_HRES_LOG(UpdatePstPath((LPWSTR)tkOptions.szProfileName.c_str(), (LPWSTR)tkOptions.szPstNewPath.c_str(), tkOptions.bPstMoveFiles, loggingMode), loggingMode);
				}
				else
				{
					EC_HRES_LOG(UpdatePstPath((LPWSTR)tkOptions.szProfileName.c_str(), (LPWSTR)tkOptions.szPstOldPath.c_str(), (LPWSTR)tkOptions.szPstNewPath.c_str(), tkOptions.bPstMoveFiles, loggingMode), loggingMode);
				}
				break;
			};
			MAPIUninitialize();
		}
	}
	catch (int exception)
	{
		std::wostringstream oss; \
			oss << L"Error " << std::dec << exception << L" encountered";
		Logger::Write(logLevelError, oss.str(), loggingMode);
	}

Error:
	goto Cleanup;
Cleanup:
	// Free up memory

	return;
}

