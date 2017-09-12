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

// Is64BitProcess
// Returns true if 64 bit process or false if 32 bit.
BOOL Is64BitProcess(void)
{
#if defined(_WIN64)
	return TRUE;   // 64-bit program
#else
	return FALSE;
#endif
}

// IsCorrectBitness
// Matches the App bitness against Outlook's bitness to avoid MAPI version errors at startup
// The execution will only continue if the bitness is matched.
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
		else if (szOLVer == L"Outlook.Application.12")
		{
			if (Is64BitProcess())
				return FALSE;
		}
		else return FALSE;
		return FALSE;
	}
	else return FALSE;
}

BOOL ValidateScenario(int argc, _TCHAR* argv[], RuntimeOptions * pRunOpts)
{
	if (!pRunOpts) return FALSE;
	ZeroMemory(pRunOpts, sizeof(RuntimeOptions));
	int iThreeParam = 0;

	for (int i = 1; i < argc; i++)
	{
		if (iThreeParam < 3)
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
				case 'p':
					pRunOpts->ulScenario = SCENARIO_PROFILE;
					iThreeParam++;
					break;
				case 's':
					pRunOpts->ulScenario = SCENARIO_SERVICE;
					iThreeParam++;
					break;
				case 'm':
					pRunOpts->ulScenario = SCENARIO_MAILBOX;
					iThreeParam++;
					break;
				case 'd':
					pRunOpts->ulScenario = SCENARIO_DATAFILE;
					iThreeParam++;
					break;
				case 'l':
					pRunOpts->ulScenario = SCENARIO_LDAP;
					iThreeParam++;
					break;
				case 'c':
					pRunOpts->ulScenario = SCENARIO_CUSTOM;
					iThreeParam++;
					break;
				default:
					return false;
				}
				break;
			case '-':
			case '/':
			case '\\':
				if (0 == argv[i][1])
				{
					return false;
				}
				switch (tolower(argv[i][1]))
				{
				case 't':
					if (i + 1 < argc)
					{
						std::wstring wszActionType = argv[i + 1];
						std::transform(wszActionType.begin(), wszActionType.end(), wszActionType.begin(), ::tolower);
						if (wszActionType == L"standard")
						{
							pRunOpts->ulActionType = ACTIONTYPE_STANDARD;
							iThreeParam++;
							i++;
							break;
						}
						else if (wszActionType == L"custom")
						{
							pRunOpts->ulActionType = ACTIONTYPE_CUSTOM;
							iThreeParam++;
							i++;
							break;
						}
						else
						{
							return false;
						}
					}
					break;
				case 'a':
					if (i + 1 < argc)
					{
						std::wstring wszAction = argv[i + 1];
						std::transform(wszAction.begin(), wszAction.end(), wszAction.begin(), ::tolower);
						if (pRunOpts->ulActionType == ACTIONTYPE_STANDARD)
						{
							if (wszAction == L"add")
							{
								pRunOpts->ulAction = STANDARDACTION_ADD;
								iThreeParam++;
								i++;
							}
							else if (wszAction == L"remove")
							{
								pRunOpts->ulAction = STANDARDACTION_REMOVE;
								iThreeParam++;
								i++;
							}
							else if (wszAction == L"edit")
							{
								pRunOpts->ulAction = STANDARDACTION_EDIT;
								iThreeParam++;
								i++;
							}
							else if (wszAction == L"list")
							{
								pRunOpts->ulAction = STANDARDACTION_LIST;
								iThreeParam++;
								i++;
							}
							else
							{
								return false;
							}
						}
						else if (pRunOpts->ulActionType == ACTIONTYPE_CUSTOM)
						{
							if (wszAction == L"promotemailboxtoservice")
							{
								pRunOpts->ulAction = CUSTOMACTION_PROMOTEMAILBOXTOSERVICE;
								iThreeParam++;
								i++;
							}
							else if (wszAction == L"editcachedmodeconfiguration")
							{
								pRunOpts->ulAction = CUSTOMACTION_EDITCACHEDMODECONFIGURATION;
								iThreeParam++;
								i++;
							}
							else if (wszAction == L"updatesmtpaddress")
							{
								pRunOpts->ulAction = CUSTOMACTION_UPDATESMTPADDRESS;
								iThreeParam++;
								i++;
							}
							else if (wszAction == L"changepstlocation")
							{
								pRunOpts->ulAction = CUSTOMACTION_CHANGEPSTLOCATION;
								iThreeParam++;
								i++;
							}
							else if (wszAction == L"removeorphaneddatafiles")
							{
								pRunOpts->ulAction = CUSTOMACTION_REMOVEORPHANEDDATAFILES;
								iThreeParam++;
								i++;
							}
							else
							{
								return false;
							}
						}
					}
					break;
				case 'l':
					if (tolower(argv[i][2]) == 'm')
					{
						pRunOpts->ulLoggingMode = _wtoi(argv[i + 1]);
						i++;
					}
					else if (tolower(argv[i][2]) == 'p')
					{
						pRunOpts->wszLogFilePath = argv[i + 1];
						Logger::Initialise(pRunOpts->wszLogFilePath);
						i++;
					}
					else return false;
					break;
				case 'e':
					if (tolower(argv[i][2]) == 'p')
					{
						std::wstring wszExportPath = argv[i + 1];
						std::transform(wszExportPath.begin(), wszExportPath.end(), wszExportPath.begin(), ::tolower);
						pRunOpts->wszExportPath = wszExportPath;
						pRunOpts->bExportMode = EXPORTMODE_EXPORT;
						i++;
					}
					else return false;
					break;
				case '?':
					return false;
				default:
					// display help
					pRunOpts->ulLoggingMode = LOGGINGMODE_CONSOLE;
				}
			}
		}
		else break;
	}
	switch (pRunOpts->ulScenario)
	{
	case SCENARIO_PROFILE:
		pRunOpts->profileOptions = new ProfileOptions();
		return ParseArgsProfile(argc, argv, pRunOpts->profileOptions);
		break;
	case SCENARIO_SERVICE:
		pRunOpts->serviceOptions = new ServiceOptions();
		return ParseArgsService(argc, argv, pRunOpts->serviceOptions);
		break;
	case SCENARIO_MAILBOX:
		pRunOpts->mailboxOptions = new MailboxOptions();
		return ParseArgsMailbox(argc, argv, pRunOpts->mailboxOptions);
		break;
	case SCENARIO_DATAFILE:
		pRunOpts->dataFileOptions = new DataFileOptions();
		return FALSE;
		break;
	case SCENARIO_LDAP:
		pRunOpts->ldapOptions = new LdapOptions();
		return FALSE;
		break;
	case SCENARIO_CUSTOM:
		return FALSE;
		break;
	default:
		return FALSE;
	}
}

BOOL ParseArgsProfile(int argc, _TCHAR* argv[], ProfileOptions * profileOptions)
{
	if (!profileOptions) return FALSE;
	profileOptions->bSetDefaultProfile = false;
	profileOptions->ulProfileMode = PROFILEMODE_DEFAULT;

	for (int i = 1; i < argc; i++)
	{
		switch (argv[i][0])
		{
		case '-':
		case '/':
		case '\\':
			if (0 == argv[i][1])
			{
				return false;
			}
			switch (tolower(argv[i][1]))
			{
			case 'p':
				if (tolower(argv[i][2]) == 'm')
				{
					if (i + 1 < argc)
					{
						std::wstring profileMode = argv[i + 1];
						std::transform(profileMode.begin(), profileMode.end(), profileMode.begin(), ::tolower);
						if (profileMode == L"all")
						{
							profileOptions->ulProfileMode = (ULONG)PROFILEMODE_ALL;
							i++;
						}
						else if (profileMode == L"one")
						{
							profileOptions->ulProfileMode = (ULONG)PROFILEMODE_ONE;
							i++;
						}
						else if (profileMode == L"default")
						{
							profileOptions->ulProfileMode = (ULONG)PROFILEMODE_DEFAULT;
							i++;
						}
						else return false;
					}
				}
				else if (tolower(argv[i][2]) == 'n')
				{
					if (i + 1 < argc)
					{
						profileOptions->wszProfileName = argv[i + 1];
						i++;
					}
					else return false;
				}
				else if (tolower(argv[i][2]) == 'd')
				{
					profileOptions->bSetDefaultProfile = true;
				}
				else return false;
				break;
			}
			break;
		}
	}
	return true;
}

BOOL ParseArgsService(int argc, _TCHAR* argv[], ServiceOptions * serviceOptions)
{
	if (!serviceOptions) return FALSE;
	serviceOptions->bDefaultservice = false;
	serviceOptions->bSetDefaultservice = false;
	serviceOptions->ulProfileMode = PROFILEMODE_DEFAULT;

	for (int i = 1; i < argc; i++)
	{
		switch (argv[i][0])
		{
		case '-':
		case '/':
		case '\\':
			if (0 == argv[i][1])
			{
				return false;
			}
			switch (tolower(argv[i][1]))
			{
			case 's':
				if (tolower(argv[i][2]) == 'a')
				{
					if (tolower(argv[i][3]) == 'b')
					{
						if (tolower(argv[i][4]) == 'e')
						{
							if (i + 1 < argc)
							{
								// sabe		| wszAddressBookExternalUrl
								std::wstring wszAddressBookExternalUrl = argv[i + 1];
								std::transform(wszAddressBookExternalUrl.begin(), wszAddressBookExternalUrl.end(), wszAddressBookExternalUrl.begin(), ::tolower);
								serviceOptions->wszAddressBookExternalUrl = wszAddressBookExternalUrl;
								i++;
							}
							else return false;
						}
						else if (tolower(argv[i][4]) == 'i')
						{
							if (i + 1 < argc)
							{
								// sabi		| wszAddressBookExternalUrl
								std::wstring wszAddressBookInternalUrl = argv[i + 1];
								std::transform(wszAddressBookInternalUrl.begin(), wszAddressBookInternalUrl.end(), wszAddressBookInternalUrl.begin(), ::tolower);
								serviceOptions->wszAddressBookInternalUrl = wszAddressBookInternalUrl;
								i++;
							}
							else return false;
						}
						else return false;
					}
					else if (tolower(argv[i][3]) == 'u')
					{
						if (i + 1 < argc)
						{
							// sau		| wszAutodiscoverUrl
							std::wstring wszAutodiscoverUrl = argv[i + 1];
							std::transform(wszAutodiscoverUrl.begin(), wszAutodiscoverUrl.end(), wszAutodiscoverUrl.begin(), ::tolower);
							serviceOptions->wszAutodiscoverUrl = wszAutodiscoverUrl;
							i++;
						}
						else return false;
					}
					else return false;
				}
				else if (tolower(argv[i][2]) == 'c')
				{
					if (tolower(argv[i][3]) == 'f')
					{
						if (tolower(argv[i][4]) == 'g')
						{
							if (tolower(argv[i][5]) == 'f')
							{
								if (i + 1 < argc)
								{
									// scfgf	| ulConfigFlags
									serviceOptions->ulConfigFlags = _wtoi(argv[i + 1]);
									i++;
								}
								else return false;
							}
							else return false;
						}
						else return false;
					}
					else if (tolower(argv[i][3]) == 'm')
					{
						if (tolower(argv[i][4]) == 'm')
						{
							if (i + 1 < argc)
							{
								// scmm	| iCachedModeMonths
								serviceOptions->iCachedModeMonths = _wtoi(argv[i + 1]);
								i++;
							}
							else return false;
						}
						else if (tolower(argv[i][4]) == 'o')
						{
							if (i + 1 < argc)
							{
								// scmo	| ulCachedModeOwner
								serviceOptions->ulCachedModeOwner = _wtoi(argv[i + 1]);
								i++;
							}
							else return false;
						}
						else if (tolower(argv[i][4]) == 'p')
						{
							if (tolower(argv[i][5]) == 'f')
							{
								if (i + 1 < argc)
								{
									// scmpf	| ulCachedModePublicFolder
									serviceOptions->ulCachedModePublicFolder = _wtoi(argv[i + 1]);
									i++;
								}
								else return false;
							}
							else return false;
						}
						else if (tolower(argv[i][4]) == 's')
						{
							if (i + 1 < argc)
							{
								// scms	| ulCachedModeShared
								serviceOptions->ulCachedModeShared = _wtoi(argv[i + 1]);
								i++;
							}
							else return false;
						}
						else return false;
					}
					else if (tolower(argv[i][3]) == 'n')
					{
						if (tolower(argv[i][4]) == 'c')
						{
							if (tolower(argv[i][5]) == 't')
							{
								if (tolower(argv[i][6]) == 'm')
								{
									if (i + 1 < argc)
									{
										// scnctm	| ulConnectMode
										serviceOptions->ulConnectMode = _wtoi(argv[i + 1]);
										i++;
									}
									else return false;
								}
								else return false;
							}
							else return false;
						}
						else return false;
					}
					else return false;
				}
				else if (tolower(argv[i][2]) == 'd')
				{
					if (tolower(argv[i][3]) == 's')
					{
						// sds	| bDefaultservice;
						serviceOptions->bDefaultservice = true;
					}
				}
				else if (tolower(argv[i][2]) == 'i')
				{
					if (i + 1 < argc)
					{
						// si	| iServiceIndex
						serviceOptions->iServiceIndex = _wtoi(argv[i + 1]);
						i++;
					}
					else return false;
				}
				else if (tolower(argv[i][2]) == 'm')
				{
					if (tolower(argv[i][3]) == 'd')
					{
						if (tolower(argv[i][4]) == 'n')
						{
							if (i + 1 < argc)
							{
								// smdn		| wszMailboxDisplayName
								std::wstring wszMailboxDisplayName = argv[i + 1];
								std::transform(wszMailboxDisplayName.begin(), wszMailboxDisplayName.end(), wszMailboxDisplayName.begin(), ::tolower);
								serviceOptions->wszMailboxDisplayName = wszMailboxDisplayName;
								i++;
							}
							else return false;
						}
					}
					else if (tolower(argv[i][3]) == 'l')
					{
						if (tolower(argv[i][4]) == 'd')
						{
							if (tolower(argv[i][5]) == 'n')
							{
								if (i + 1 < argc)
								{
									// smldn		| wszMailboxLegacyDN
									std::wstring wszMailboxLegacyDN = argv[i + 1];
									std::transform(wszMailboxLegacyDN.begin(), wszMailboxLegacyDN.end(), wszMailboxLegacyDN.begin(), ::tolower);
									serviceOptions->wszMailboxLegacyDN = wszMailboxLegacyDN;
									i++;
								}
								else return false;
							}
						}
					}
					else if (tolower(argv[i][3]) == 's')
					{
						if (tolower(argv[i][4]) == 'e')
						{
							if (i + 1 < argc)
							{
								// smse	| wszMailStoreExternalUrl
								std::wstring wszMailStoreExternalUrl = argv[i + 1];
								std::transform(wszMailStoreExternalUrl.begin(), wszMailStoreExternalUrl.end(), wszMailStoreExternalUrl.begin(), ::tolower);
								serviceOptions->wszMailStoreExternalUrl = wszMailStoreExternalUrl;
								i++;
							}
							else return false;
						}
						else if (tolower(argv[i][4]) == 'i')
						{
							if (i + 1 < argc)
							{
								// smsi	| wszMailStoreExternalUrl
								std::wstring wszMailStoreInternalUrl = argv[i + 1];
								std::transform(wszMailStoreInternalUrl.begin(), wszMailStoreInternalUrl.end(), wszMailStoreInternalUrl.begin(), ::tolower);
								serviceOptions->wszMailStoreInternalUrl = wszMailStoreInternalUrl;
								i++;
							}
							else return false;
						}
					}
				}
				else if (tolower(argv[i][2]) == 'o')
				{
					if (tolower(argv[i][2]) == 'v')
					{
						if (i + 1 < argc)
						{
							// mov		| iOutlookVersion
							std::wstring wszOutlookVersion = argv[i + 1];
							std::transform(wszOutlookVersion.begin(), wszOutlookVersion.end(), wszOutlookVersion.begin(), ::tolower);
							if (wszOutlookVersion == L"2007")
							{
								serviceOptions->iOutlookVersion = 2007;
								i++;
							}
							else if (wszOutlookVersion == L"2010")
							{
								serviceOptions->iOutlookVersion = 2010;
								i++;
							}
							else if (wszOutlookVersion == L"2013")
							{
								serviceOptions->iOutlookVersion = 2013;
								i++;
							}
							else if (wszOutlookVersion == L"2016")
							{
								serviceOptions->iOutlookVersion = 2016;
								i++;
							}
							else return false;
						}
						else return false;
					}
					else return false;
				}
				else if (tolower(argv[i][2]) == 'p')
				{
					if (tolower(argv[i][3]) == 'm')
					{
						// spm		| ulProfileMode
						if (i + 1 < argc)
						{
							std::wstring profileMode = argv[i + 1];
							std::transform(profileMode.begin(), profileMode.end(), profileMode.begin(), ::tolower);
							if (profileMode == L"all")
							{
								serviceOptions->ulProfileMode = (ULONG)PROFILEMODE_ALL;
								i++;
							}
							else if (profileMode == L"one")
							{
								serviceOptions->ulProfileMode = (ULONG)PROFILEMODE_ONE;
								i++;
							}
							else if (profileMode == L"default")
							{
								serviceOptions->ulProfileMode = (ULONG)PROFILEMODE_DEFAULT;
								i++;
							}
							else return false;
						}
						else return false;
					}
					else if (tolower(argv[i][3]) == 'n')
					{
						if (i + 1 < argc)
						{
							// spn		| wszProfileName
							std::wstring wszProfileName = argv[i + 1];
							std::transform(wszProfileName.begin(), wszProfileName.end(), wszProfileName.begin(), ::tolower);
							serviceOptions->wszProfileName = wszProfileName;
							i++;
						}
						else return false;
					}
				}
				else if (tolower(argv[i][2]) == 'r')
				{
					if (tolower(argv[i][3]) == 'f')
					{
						// srf		| ulResourceFlags
						if (i + 1 < argc)
						{
							// si	| iServiceIndex
							serviceOptions->iServiceIndex = _wtoi(argv[i + 1]);
							i++;
						}
						else return false;
					}
					else if (tolower(argv[i][3]) == 'p')
					{
						if (tolower(argv[i][4]) == 's')
						{

							if (i + 1 < argc)
							{
								// srps		| wszRohProxyServer
								std::wstring wszRohProxyServer = argv[i + 1];
								std::transform(wszRohProxyServer.begin(), wszRohProxyServer.end(), wszRohProxyServer.begin(), ::tolower);
								serviceOptions->wszRohProxyServer = wszRohProxyServer;
								i++;
							}
							else return false;
						}
						else return false;
					}
					else return false;
				}
				else if (tolower(argv[i][2]) == 's')
				{
					if (tolower(argv[i][3]) == 'a')
					{
						if (i + 1 < argc)
						{
							// ssa	| wszSmtpAddress
							std::wstring wszSmtpAddress = argv[i + 1];
							std::transform(wszSmtpAddress.begin(), wszSmtpAddress.end(), wszSmtpAddress.begin(), ::tolower);
							serviceOptions->wszSmtpAddress = wszSmtpAddress;
							i++;
						}
						else return false;
					}
					else if (tolower(argv[i][3]) == 'd')
					{
						if (tolower(argv[i][4]) == 'n')
						{
							if (i + 1 < argc)
							{
								// ssdn	| wszServerDisplayName
								std::wstring wszServerDisplayName = argv[i + 1];
								std::transform(wszServerDisplayName.begin(), wszServerDisplayName.end(), wszServerDisplayName.begin(), ::tolower);
								serviceOptions->wszServerDisplayName = wszServerDisplayName;
								i++;
							}
							else return false;
						}
						else if (tolower(argv[i][4]) == 's')
						{
							// ssds	| bSetDefaultservice
							serviceOptions->bSetDefaultservice = true;
						}
						else return false;
					}
					else if (tolower(argv[i][3]) == 'l')
					{
						if (tolower(argv[i][4]) == 'd')
						{
							if (tolower(argv[i][5]) == 'n')
							{
								if (i + 1 < argc)
								{
									// ssldn	| wszServerLegacyDN
									std::wstring wszServerLegacyDN = argv[i + 1];
									std::transform(wszServerLegacyDN.begin(), wszServerLegacyDN.end(), wszServerLegacyDN.begin(), ::tolower);
									serviceOptions->wszServerLegacyDN = wszServerLegacyDN;
									i++;
								}
								else return false;
							}
							else return false;
						}
						else return false;
					}
				}
				else if (tolower(argv[i][2]) == 'u')
				{
					if (tolower(argv[i][3]) == 's')
					{

						if (i + 1 < argc)
						{
							// sus	| wszUnresolvedServer
							std::wstring wszUnresolvedServer = argv[i + 1];
							std::transform(wszUnresolvedServer.begin(), wszUnresolvedServer.end(), wszUnresolvedServer.begin(), ::tolower);
							serviceOptions->wszUnresolvedServer = wszUnresolvedServer;
							i++;
						}
						else return false;
					}
					if (tolower(argv[i][3]) == 'u')
					{
						if (i + 1 < argc)
						{
							// suu	| wszUnresolvedUser
							std::wstring wszUnresolvedUser = argv[i + 1];
							std::transform(wszUnresolvedUser.begin(), wszUnresolvedUser.end(), wszUnresolvedUser.begin(), ::tolower);
							serviceOptions->wszUnresolvedUser = wszUnresolvedUser;
							i++;
						}
						else return false;
					}
					else return false;
				}
				else return false;
				break;
			}
			break;
		}
	}
	return true;
}

BOOL ParseArgsMailbox(int argc, _TCHAR* argv[], MailboxOptions * mailboxOptions)
{
	if (!mailboxOptions) return FALSE;

	mailboxOptions->bDefaultService = false;

	for (int i = 1; i < argc; i++)
	{
		switch (argv[i][0])
		{
		case '-':
		case '/':
		case '\\':
			if (0 == argv[i][1])
			{
				return false;
			}
			switch (tolower(argv[i][1]))
			{
			case 'm':
				if (tolower(argv[i][2]) == 'd')
				{
					if (tolower(argv[i][3]) == 's')
					{
						// mds		| bDefaultService
						mailboxOptions->bDefaultService = true;
					}
				}
				else if (tolower(argv[i][2]) == 'm')
				{
					if (tolower(argv[i][3]) == 'd')
					{
						if (tolower(argv[i][4]) == 'n')
						{
							if (i + 1 < argc)
							{
								// mmdn		| wszMailboxDisplayName
								std::wstring wszMailboxDisplayName = argv[i + 1];
								std::transform(wszMailboxDisplayName.begin(), wszMailboxDisplayName.end(), wszMailboxDisplayName.begin(), ::tolower);
								mailboxOptions->wszMailboxDisplayName = wszMailboxDisplayName;
								i++;
							}
						}
					}
					if (tolower(argv[i][3]) == 'l')
					{
						if (tolower(argv[i][4]) == 'd')
						{
							if (tolower(argv[i][5]) == 'n')
							{
								if (i + 1 < argc)
								{
									// mmldn		| wszMailboxLegacyDN
									std::wstring wszMailboxLegacyDN = argv[i + 1];
									std::transform(wszMailboxLegacyDN.begin(), wszMailboxLegacyDN.end(), wszMailboxLegacyDN.begin(), ::tolower);
									mailboxOptions->wszMailboxLegacyDN = wszMailboxLegacyDN;
									i++;
								}
							}
						}
					}
				}
				else if (tolower(argv[i][2]) == 'o')
				{
					if (tolower(argv[i][3]) == 'v')
					{
						if (i + 1 < argc)
						{
							// mov		| iOutlookVersion
							std::wstring wszOutlookVersion = argv[i + 1];
							std::transform(wszOutlookVersion.begin(), wszOutlookVersion.end(), wszOutlookVersion.begin(), ::tolower);
							if (wszOutlookVersion == L"2007")
							{
								mailboxOptions->iOutlookVersion = 2007;
								i++;
							}
							else if (wszOutlookVersion == L"2010")
							{
								mailboxOptions->iOutlookVersion = 2010;
								i++;
							}
							else if (wszOutlookVersion == L"2013")
							{
								mailboxOptions->iOutlookVersion = 2013;
								i++;
							}
							else if (wszOutlookVersion == L"2016")
							{
								mailboxOptions->iOutlookVersion = 2016;
								i++;
							}
							else return false;
						}
					}
				}
				else if (tolower(argv[i][2]) == 'p')
				{
					if (tolower(argv[i][3]) == 'n')
					{
						if (i + 1 < argc)
						{
							// mpn		| wszProfileName
							std::wstring wszProfileName = argv[i + 1];
							std::transform(wszProfileName.begin(), wszProfileName.end(), wszProfileName.begin(), ::tolower);
							mailboxOptions->wszProfileName = wszProfileName;
							i++;
						}
					}
					else if (tolower(argv[i][3]) == 'm')
					{
						if (i + 1 < argc)
						{
							// mpm		| ulProfileMode
							std::wstring profileMode = argv[i + 1];
							std::transform(profileMode.begin(), profileMode.end(), profileMode.begin(), ::tolower);
							if (profileMode == L"all")
							{
								mailboxOptions->ulProfileMode = (ULONG)PROFILEMODE_ALL;
								i++;
							}
							else if (profileMode == L"one")
							{
								mailboxOptions->ulProfileMode = (ULONG)PROFILEMODE_ONE;
								i++;
							}
							else if (profileMode == L"default")
							{
								mailboxOptions->ulProfileMode = (ULONG)PROFILEMODE_DEFAULT;
								i++;
							}
							else return false;
						}
					}
					else return false;
				}
				else if (tolower(argv[i][2]) == 's')
				{
					if (tolower(argv[i][3]) == 'a')
					{
						if (i + 1 < argc)
						{
							// msa		| wszSmtpAddress
							std::wstring wszSmtpAddress = argv[i + 1];
							std::transform(wszSmtpAddress.begin(), wszSmtpAddress.end(), wszSmtpAddress.begin(), ::tolower);
							mailboxOptions->wszSmtpAddress = wszSmtpAddress;
							i++;
						}
					}
					else if (tolower(argv[i][3]) == 'd')
					{
						if (tolower(argv[i][4]) == 'n')
						{
							if (i + 1 < argc)
							{
								// msdn		| wszServerDisplayName
								std::wstring wszServerDisplayName = argv[i + 1];
								std::transform(wszServerDisplayName.begin(), wszServerDisplayName.end(), wszServerDisplayName.begin(), ::tolower);
								mailboxOptions->wszServerDisplayName = wszServerDisplayName;
								i++;
							}
						}
					}
					if (tolower(argv[i][3]) == 'i')
					{
						if (i + 1 < argc)
						{
							// msi	| ulServiceIndex
							mailboxOptions->ulServiceIndex = _wtoi(argv[i + 1]);;
							i++;
						}
					}
					if (tolower(argv[i][3]) == 'l')
					{
						if (tolower(argv[i][4]) == 'd')
						{
							if (tolower(argv[i][5]) == 'n')
							{
								if (i + 1 < argc)
								{
									// msldn	| wszServerLegacyDN
									std::wstring wszServerLegacyDN = argv[i + 1];
									std::transform(wszServerLegacyDN.begin(), wszServerLegacyDN.end(), wszServerLegacyDN.begin(), ::tolower);
									mailboxOptions->wszServerLegacyDN = wszServerLegacyDN;
									i++;
								}
							}
						}
					}
				}
				else return false;
				break;
			}
			break;
		}
		
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
	RuntimeOptions * tkOptions = new RuntimeOptions();
	loggingMode = loggingModeNone;
	// Parse the command line arguments
	if (!ValidateScenario(argc, argv, tkOptions))
	{
		if (tkOptions->ulLoggingMode != loggingModeNone)
		{
			DisplayUsage();
		}
		return;
	}

	MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
	if (SUCCEEDED(MAPIInitialize(&MAPIINIT)))
	{

		switch (tkOptions->ulScenario)
		{
		case SCENARIO_PROFILE:
			if (tkOptions->ulActionType == ACTIONTYPE_STANDARD)
			{
				if (tkOptions->ulAction == STANDARDACTION_ADD)
				{
					// this only works with the default profile for now
					HrCreateProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str());
					break;
				}
			}
		case SCENARIO_SERVICE:
			if (tkOptions->ulActionType == ACTIONTYPE_STANDARD)
			{
				if (tkOptions->ulAction == STANDARDACTION_ADD)
				{
					if (tkOptions->mailboxOptions->iOutlookVersion == 2007)
					{
						HrCreateMsemsServiceLegacyUnresolved((tkOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
							(LPWSTR)tkOptions->serviceOptions->wszProfileName.c_str(),
							(LPWSTR)tkOptions->serviceOptions->wszMailboxLegacyDN.c_str(),
							(LPWSTR)tkOptions->serviceOptions->wszServerDisplayName.c_str(),
							loggingMode);
					}
					else if ((tkOptions->mailboxOptions->iOutlookVersion == 2010) || (tkOptions->mailboxOptions->iOutlookVersion == 2013))
					{
						// this only works with the default profile for now
						if (tkOptions->serviceOptions->ulConnectMode == CONNECT_ROH)
						{
							HrCreateMsemsServiceROH((tkOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
								(LPWSTR)tkOptions->serviceOptions->wszProfileName.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszSmtpAddress.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszMailboxLegacyDN.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszUnresolvedServer.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszRohProxyServer.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszServerLegacyDN.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszAutodiscoverUrl.c_str(),
								loggingMode);
						}
						else if (tkOptions->serviceOptions->ulConnectMode == CONNECT_MOH)
						{
							HrCreateMsemsServiceMOH((tkOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
								(LPWSTR)tkOptions->serviceOptions->wszProfileName.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszSmtpAddress.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszMailboxLegacyDN.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszMailStoreInternalUrl.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszMailStoreExternalUrl.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszAddressBookInternalUrl.c_str(),
								(LPWSTR)tkOptions->serviceOptions->wszAddressBookExternalUrl.c_str(),
								loggingMode);
						}
						else
						{

						}
					}
					else // default to the 2016 logic
					{
						// this only works with the default profile for now
						HrCreateMsemsServiceModern((tkOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
							(LPWSTR)tkOptions->serviceOptions->wszProfileName.c_str(),
							(LPWSTR)tkOptions->serviceOptions->wszSmtpAddress.c_str(),
							(LPWSTR)tkOptions->serviceOptions->wszMailboxDisplayName.c_str(),
							loggingMode);
					}
				}
			}
			break;
		case SCENARIO_MAILBOX:
			if (tkOptions->ulActionType == ACTIONTYPE_STANDARD)
			{
				if (tkOptions->ulAction == STANDARDACTION_ADD)
				{
					if (tkOptions->mailboxOptions->iOutlookVersion == 2007)
					{
						// this only works with the default profile for now
						HrAddDelegateMailboxLegacy((tkOptions->mailboxOptions->ulProfileMode == PROFILEMODE_DEFAULT),
							(LPWSTR)tkOptions->mailboxOptions->wszProfileName.c_str(),
							tkOptions->mailboxOptions->bDefaultService,
							tkOptions->mailboxOptions->ulServiceIndex,
							(LPWSTR)tkOptions->mailboxOptions->wszMailboxDisplayName.c_str(),
							(LPWSTR)tkOptions->mailboxOptions->wszMailboxLegacyDN.c_str(),
							(LPWSTR)tkOptions->mailboxOptions->wszServerDisplayName.c_str(),
							(LPWSTR)tkOptions->mailboxOptions->wszServerLegacyDN.c_str(),
							loggingMode);
					}
					else if ((tkOptions->mailboxOptions->iOutlookVersion == 2010) || (tkOptions->mailboxOptions->iOutlookVersion == 2013))
					{
						// this only works with the default profile for now
						HrAddDelegateMailbox((tkOptions->mailboxOptions->ulProfileMode == PROFILEMODE_DEFAULT),
							(LPWSTR)tkOptions->mailboxOptions->wszProfileName.c_str(),
							tkOptions->mailboxOptions->bDefaultService,
							tkOptions->mailboxOptions->ulServiceIndex,
							(LPWSTR)tkOptions->mailboxOptions->wszMailboxDisplayName.c_str(),
							(LPWSTR)tkOptions->mailboxOptions->wszMailboxLegacyDN.c_str(),
							(LPWSTR)tkOptions->mailboxOptions->wszServerDisplayName.c_str(),
							(LPWSTR)tkOptions->mailboxOptions->wszServerLegacyDN.c_str(),
							(LPWSTR)tkOptions->mailboxOptions->wszSmtpAddress.c_str(),
							NULL,
							0,
							0,
							NULL,
							loggingMode);
					}
					else // default to the 2016 logic
					{
						// this only works with the default profile for now
						HrAddDelegateMailboxModern((tkOptions->mailboxOptions->ulProfileMode == PROFILEMODE_DEFAULT),
							(LPWSTR)tkOptions->mailboxOptions->wszProfileName.c_str(),
							tkOptions->mailboxOptions->bDefaultService,
							tkOptions->mailboxOptions->ulServiceIndex,
							(LPWSTR)tkOptions->mailboxOptions->wszMailboxDisplayName.c_str(),
							(LPWSTR)tkOptions->mailboxOptions->wszSmtpAddress.c_str(),
							loggingMode);
					}
				}
			}
			break;
		}
		MAPIUninitialize();
	}
	//loggingMode = LoggingMode(tkOptions.ulLoggingMode);
	//// Check the curren't process' bitness vs Outlook's bitness and only run it if matched to avoid MAPI dialog boxes.
	//if (!IsCorrectBitness())
	//{
	//	Logger::Write(logLevelFailed, L"Unable to resolve bitness or bitness not matched.", loggingMode);
	//	return;
	//}
	//Logger::Write(logLevelSuccess, L"Bitness matched.", loggingMode);

	try
	{
		//MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
		//if (SUCCEEDED(MAPIInitialize(&MAPIINIT)))
		//{
		//	Logger::Write(logLevelSuccess, L"MAPI Initialised", loggingMode);
		//	//switch (tkOptions.ulScenario)
		//	//{
		//	//case SCENARIO_PROFILE:
		//	//	switch (tkOptions.profileOptions->ulProfileMode)
		//	//	{
		//	//	case PROFILEMODE_ALL:
		//	//			ULONG ulProfileCount = GetProfileCount(loggingMode);
		//	//			ProfileInfo * profileInfo = new ProfileInfo[ulProfileCount];
		//	//			ZeroMemory(profileInfo, sizeof(ProfileInfo) * ulProfileCount);
		//	//			Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for all profiles", loggingMode);
		//	//			EC_HRES_LOG(GetProfiles(ulProfileCount, profileInfo, loggingMode), loggingMode);
		//	//			if (tkOptions.wszExportPath != L"")
		//	//			{
		//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles", loggingMode);
		//	//				ExportXML(ulProfileCount, profileInfo, tkOptions.wszExportPath, loggingMode);
		//	//			}
		//	//			else
		//	//			{
		//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles", loggingMode);
		//	//				ExportXML(ulProfileCount, profileInfo, L"", loggingMode);
		//	//			}
		//	//		
		//	//		break;
		//	//	case PROFILEMODE_ONE:
		//	//		
		//	//			ProfileInfo profileInfo;
		//	//			Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for profile: " + tkOptions.profileOptions->wszProfileName, loggingMode);
		//	//			EC_HRES_LOG(GetProfile((LPWSTR)tkOptions.profileOptions->wszProfileName.c_str(), &profileInfo, loggingMode), loggingMode);
		//	//			if (tkOptions.wszExportPath != L"")
		//	//			{
		//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile", loggingMode);
		//	//				ExportXML(1, &profileInfo, tkOptions.szExportPath, loggingMode);
		//	//			}
		//	//			else
		//	//			{
		//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile", loggingMode);
		//	//				ExportXML(1, &profileInfo, L"", loggingMode);
		//	//			}
		//	//		break;
		//	//	case PROFILEMODE_DEFAULT:
		//	//		std::wstring szDefaultProfileName = GetDefaultProfileName(loggingMode);
		//	//		if (!szDefaultProfileName.empty())
		//	//		{
		//	//			tkOptions.szProfileName = szDefaultProfileName;
		//	//		}
		//	//		if (tkOptions.ulReadWriteMode == READWRITEMODE_READ)
		//	//		{
		//	//			ProfileInfo profileInfo;
		//	//			Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for default profile: " + tkOptions.szProfileName, loggingMode);
		//	//			EC_HRES_LOG(GetProfile((LPWSTR)tkOptions.szProfileName.c_str(), &profileInfo, loggingMode), loggingMode);
		//	//			if (tkOptions.szExportPath != L"")
		//	//			{
		//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile", loggingMode);
		//	//				ExportXML(1, &profileInfo, tkOptions.szExportPath, loggingMode);
		//	//			}
		//	//			else
		//	//			{
		//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile", loggingMode);
		//	//				ExportXML(1, &profileInfo, L"", loggingMode);
		//	//			}
		//	//		}
		//	//		else if (tkOptions.ulReadWriteMode == READWRITEMODE_WRITE)
		//	//		{
		//	//			Logger::Write(logLevelInfo, L"Updating cached mode configuration on default profile: " + tkOptions.szProfileName, loggingMode);
		//	//			EC_HRES_LOG(UpdateCachedModeConfig((LPSTR)tkOptions.szProfileName.c_str(), tkOptions.ulServiceIndex, tkOptions.ulCachedModeOwner, tkOptions.ulCachedModeShared, tkOptions.ulCachedModePublicFolder, tkOptions.iCachedModeMonths, loggingMode), loggingMode);
		//	//		}
		//	//		break;
		//	//	}
		//	//	break;
		//	//case RUNNINGMODE_PST:
		//	//	if (tkOptions.szPstOldPath.empty())
		//	//	{
		//	//		EC_HRES_LOG(UpdatePstPath((LPWSTR)tkOptions.szProfileName.c_str(), (LPWSTR)tkOptions.szPstNewPath.c_str(), tkOptions.bPstMoveFiles, loggingMode), loggingMode);
		//	//	}
		//	//	else
		//	//	{
		//	//		EC_HRES_LOG(UpdatePstPath((LPWSTR)tkOptions.szProfileName.c_str(), (LPWSTR)tkOptions.szPstOldPath.c_str(), (LPWSTR)tkOptions.szPstNewPath.c_str(), tkOptions.bPstMoveFiles, loggingMode), loggingMode);
		//	//	}
		//	//	break;
		//	//};
		//	//MAPIUninitialize();
		//}
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