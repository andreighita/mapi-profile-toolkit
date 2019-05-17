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
#include "Profile.h"
#include "MapiProfileToolkit.h"
#include "ToolkitObjects.h"
#include <iostream>
#include <string>
#include <utility>
#include <algorithm>  
#include "XMLHelper.h"
#include "RegistryHelper.h"
#include "Logger.h"
#include <vector>
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

// GetOutlookVersion
int GetOutlookVersion()
{
	std::wstring szOLVer = L"";
	szOLVer = GetStringValue(HKEY_CLASSES_ROOT, TEXT("Outlook.Application\\CurVer"), NULL);
	if (szOLVer != L"")
	{
		if (szOLVer == L"Outlook.Application.16")
		{
			return 2016;
		}
		else if (szOLVer == L"Outlook.Application.15")
		{
			return 2013;
		}
		else if (szOLVer == L"Outlook.Application.14")
		{
			return 2010;
		}
		else if (szOLVer == L"Outlook.Application.12")
		{
			return 2007;
		}
		return 0;
	}
	else return 0;
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
	std::vector<std::string> wszDiscardedArgs;
	if (!pRunOpts) return FALSE;
	ZeroMemory(pRunOpts, sizeof(RuntimeOptions));
	int iThreeParam = 0;
	pRunOpts->iOutlookVersion = GetOutlookVersion();
	pRunOpts->ulActionType = ACTIONTYPE_STANDARD;
	pRunOpts->ulLoggingMode = loggingModeConsole;

	pRunOpts->profileOptions = new ProfileOptions();
	pRunOpts->profileOptions->ulProfileMode = PROFILEMODE_DEFAULT;
	pRunOpts->serviceOptions = new ServiceOptions();
	pRunOpts->serviceOptions->ulServiceMode = SERVICEMODE_DEFAULT;
	pRunOpts->serviceOptions->ulConnectMode = CONNECT_MOH;
	pRunOpts->mailboxOptions = new MailboxOptions();

	for (int i = 1; i < argc; i++)
	{
		std::wstring wsArg = argv[i];
		std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

		if ((wsArg == L"-exportpath") || (wsArg == L"-ep"))
		{
			std::wstring wszExportPath = argv[i + 1];
			std::transform(wszExportPath.begin(), wszExportPath.end(), wszExportPath.begin(), ::tolower);
			pRunOpts->wszExportPath = wszExportPath;
			pRunOpts->bExportMode = EXPORTMODE_EXPORT;
			i++;
		}
		else if ((wsArg == L"-profile") || (wsArg == L"-p"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"add")
				{
					pRunOpts->profileOptions->ulProfileAction = ACTION_ADD;
					i++;
				}
				else if (wszValue == L"edit")
				{
					pRunOpts->profileOptions->ulProfileAction = ACTION_EDIT;
					i++;
				}
				else if (wszValue == L"remove")
				{
					pRunOpts->profileOptions->ulProfileAction = ACTION_REMOVE;
					i++;
				}
				else if (wszValue == L"list")
				{
					pRunOpts->profileOptions->ulProfileAction = ACTION_LIST;
					i++;
				}
				else if (wszValue == L"clone")
				{
					pRunOpts->profileOptions->ulProfileAction = ACTION_CLONE;
					i++;
				}
				else if (wszValue == L"simpleclone")
				{
					pRunOpts->profileOptions->ulProfileAction = ACTION_SIMPLECLONE;
					i++;
				}
				else
				{
					return false;
				}
			}
		}
		else if ((wsArg == L"-profilemode") || (wsArg == L"-pm"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"default")
				{
					pRunOpts->profileOptions->ulProfileMode = PROFILEMODE_DEFAULT;
					i++;

				}
				else if (wszValue == L"one")
				{
					pRunOpts->profileOptions->ulProfileMode = PROFILEMODE_ONE;
					i++;

				}
				else if (wszValue == L"all")
				{
					pRunOpts->profileOptions->ulProfileMode = PROFILEMODE_ALL;
					i++;

				}
				else return false;
			}
		}
		else if ((wsArg == L"-profilename") || (wsArg == L"-pn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->wszProfileName = argv[i + 1];
				pRunOpts->profileOptions->ulProfileMode = PROFILEMODE_ONE;
				i++;

			}
		}
		else if ((wsArg == L"-setdefaultprofile") || (wsArg == L"-sdp"))
		{
			pRunOpts->profileOptions->bSetDefaultProfile = true;

		}
		else if ((wsArg == L"-service") || (wsArg == L"-s"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"add")
				{
					pRunOpts->serviceOptions->ulServiceAction = ACTION_ADD;
					i++;
				}
				else if (wszValue == L"edit")
				{
					pRunOpts->serviceOptions->ulServiceAction = ACTION_EDIT;
					i++;
				}
				else if (wszValue == L"remove")
				{
					pRunOpts->serviceOptions->ulServiceAction = ACTION_REMOVE;
					i++;
				}
				else if (wszValue == L"list")
				{
					pRunOpts->serviceOptions->ulServiceAction = ACTION_LIST;
					i++;
				}
				else if (wszValue == L"update")
				{
					pRunOpts->serviceOptions->ulServiceAction = ACTION_UPDATE;
					i++;
				}
				else if (wszValue == L"enablecachedmode")
				{
					pRunOpts->serviceOptions->ulServiceAction = ACTION_ENABLECACHEDMODE;
					i++;
				}

				else
				{
					return false;
				}
			}
		}
		else if ((wsArg == L"-servicetype") || (wsArg == L"-st"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"mailbox")
				{
					pRunOpts->serviceOptions->ulServiceType = SERVICETYPE_MAILBOX;
					i++;
				}
				else if (wszValue == L"pst")
				{
					pRunOpts->serviceOptions->ulServiceType = SERVICETYPE_PST;
					i++;
				}
				else if (wszValue == L"addressbook")
				{
					pRunOpts->serviceOptions->ulServiceType = SERVICETYPE_ADDRESSBOOK;
					i++;
				}
				else
				{
					return false;
				}
			}
		}
		else if ((wsArg == L"-servicemode") || (wsArg == L"-sm"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"default")
				{
					pRunOpts->serviceOptions->ulServiceMode = SERVICEMODE_DEFAULT;
					i++;

				}
				else if (wszValue == L"one")
				{
					pRunOpts->serviceOptions->ulServiceMode = SERVICEMODE_ONE;
					i++;

				}
				else if (wszValue == L"all")
				{
					pRunOpts->serviceOptions->ulServiceMode = SERVICEMODE_ALL;
					i++;

				}
				else return false;
			}
		}
		else if ((wsArg == L"-mailbox") || (wsArg == L"-m"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"add")
				{
					pRunOpts->mailboxOptions->ulMailboxAction = ACTION_ADD;
					i++;
				}
				else if (wszValue == L"edit")
				{
					pRunOpts->mailboxOptions->ulMailboxAction = ACTION_EDIT;
					i++;
				}
				else if (wszValue == L"remove")
				{
					pRunOpts->mailboxOptions->ulMailboxAction = ACTION_REMOVE;
					i++;
				}
				else if (wszValue == L"list")
				{
					pRunOpts->mailboxOptions->ulMailboxAction = ACTION_LIST;
					i++;
				}
				else if (wszValue == L"promotedelegates")
				{
					pRunOpts->mailboxOptions->ulMailboxAction = ACTION_PROMOTEDELEGATE;
					i++;
				}
				else
				{
					return false;
				}
			}
		}
		else if ((wsArg == L"-mailboxtype") || (wsArg == L"-mt"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"primary")
				{
					pRunOpts->mailboxOptions->ulMailboxType = MAILBOXTYPE_PRIMARY;
					i++;
				}
				else if (wszValue == L"delegate")
				{
					pRunOpts->mailboxOptions->ulMailboxType = MAILBOXTYPE_DELEGATE;
					i++;
				}
				else if (wszValue == L"publicfolder")
				{
					pRunOpts->mailboxOptions->ulMailboxType = MAILBOXTYPE_PUBLICFOLDER;
					i++;
				}
				else
				{
					return false;
				}
			}
		}
		else if ((wsArg == L"-setdefaultservice") || (wsArg == L"-sds"))
		{
			pRunOpts->serviceOptions->bSetDefaultservice = true;

		}
		else if ((wsArg == L"-cachedmodemonths") || (wsArg == L"-cmm"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->iCachedModeMonths = _wtoi(argv[i + 1]);
				i++;

			}
		}
		else if ((wsArg == L"-serviceindex") || (wsArg == L"-si"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->iServiceIndex = _wtoi(argv[i + 1]);
				i++;

			}
		}
		else if ((wsArg == L"-abexternalurl") || (wsArg == L"-abeu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszAddressBookExternalUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-abinternalurl") || (wsArg == L"-abiu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszAddressBookInternalUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-autodiscoverurl") || (wsArg == L"-au"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszAutodiscoverUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-mailboxdisplayname") || (wsArg == L"-mdn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszMailboxDisplayName = argv[i + 1];
				pRunOpts->mailboxOptions->wszMailboxDisplayName = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-mailboxlegacydn") || (wsArg == L"-mldn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszMailboxLegacyDN = argv[i + 1];
				pRunOpts->mailboxOptions->wszMailboxLegacyDN = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-mailstoreexternalurl") || (wsArg == L"-mseu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszMailStoreExternalUrl = argv[i + 1];
				pRunOpts->mailboxOptions->wszMailStoreExternalUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-mailstoreinternalurl") || (wsArg == L"-msiu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszMailStoreInternalUrl = argv[i + 1];
				pRunOpts->mailboxOptions->wszMailStoreExternalUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-rohproxyserver") || (wsArg == L"-rps"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszRohProxyServer = argv[i + 1];
				pRunOpts->mailboxOptions->wszRohProxyServer = argv[i + 1];
				i++;
			}
		}
		else if ((wsArg == L"-rohproxyserverflags") || (wsArg == L"-rpsf"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->mailboxOptions->ulRohProxyServerFlags = _wtoi(argv[i + 1]);
				i++;
			}
		}
		else if ((wsArg == L"-rohproxyserverauthpackage") || (wsArg == L"-mrpsap"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->mailboxOptions->ulRohProxyServerAuthPackage = _wtoi(argv[i + 1]);
				i++;
			}
		}
		else if ((wsArg == L"-serverdisplayname") || (wsArg == L"-sdn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszServerDisplayName = argv[i + 1];
				pRunOpts->mailboxOptions->wszServerDisplayName = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-serverlegacydn") || (wsArg == L"-sldn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszServerLegacyDN = argv[i + 1];
				pRunOpts->mailboxOptions->wszServerLegacyDN = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-smtpaddress") || (wsArg == L"-sa"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszSmtpAddress = argv[i + 1];
				pRunOpts->mailboxOptions->wszSmtpAddress = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-unresolvedserver") || (wsArg == L"-us"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszUnresolvedServer = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-unresolveduser") || (wsArg == L"-uu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->wszUnresolvedUser = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-cachedmodeowner") || (wsArg == L"-cmo"))
		{
			pRunOpts->serviceOptions->ulCachedModeOwner = true;

		}
		else if ((wsArg == L"-cachedmodepublicfolder") || (wsArg == L"-cmpf"))
		{
			pRunOpts->serviceOptions->ulCachedModePublicFolder = true;

		}
		else if ((wsArg == L"-cachedmodeshared") || (wsArg == L"-cms"))
		{
			pRunOpts->serviceOptions->ulCachedModeShared = true;

		}
		else if ((wsArg == L"-configflags") || (wsArg == L"-cf"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->ulConfigFlags = _wtol(argv[i + 1]);
				i++;

			}
		}
		else if ((wsArg == L"-connectmode") || (wsArg == L"-cm"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"roh")
				{
					pRunOpts->serviceOptions->ulConnectMode = CONNECT_ROH;
					i++;

				}
				if (wszValue == L"moh")
				{
					pRunOpts->serviceOptions->ulConnectMode = CONNECT_MOH;
					i++;

				}
			}
		}
		else if ((wsArg == L"-resourceflags") || (wsArg == L"-rf"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->serviceOptions->ulResourceFlags = _wtol(argv[i + 1]);
				i++;

			}
		}
		else if ((wsArg == L"-cachedmodeowner") || (wsArg == L"-cmo"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"enable")
				{
					pRunOpts->serviceOptions->ulCachedModeOwner = CACHEDMODE_ENABLED;
					i++;

				}
				if (wszValue == L"disable")
				{
					pRunOpts->serviceOptions->ulCachedModeOwner = CACHEDMODE_DISABLED;
					i++;

				}
			}
		}
		else if ((wsArg == L"-cachedmodeshared") || (wsArg == L"-cms"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"enable")
				{
					pRunOpts->serviceOptions->ulCachedModeShared = CACHEDMODE_ENABLED;
					i++;

				}
				if (wszValue == L"disable")
				{
					pRunOpts->serviceOptions->ulCachedModeShared = CACHEDMODE_DISABLED;
					i++;

				}
			}
		}
		else if ((wsArg == L"-cachedmodepublicfolders") || (wsArg == L"-cmpf"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"enable")
				{
					pRunOpts->serviceOptions->ulCachedModePublicFolder = CACHEDMODE_ENABLED;
					i++;

				}
				if (wszValue == L"disable")
				{
					pRunOpts->serviceOptions->ulCachedModePublicFolder = CACHEDMODE_DISABLED;
					i++;

				}
			}
		}
		else return false;
	}
	return true;
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
						profileOptions->ulProfileMode = PROFILEMODE_ONE;
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
	serviceOptions->ulServiceMode = SERVICEMODE_DEFAULT;
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
						serviceOptions->ulServiceMode = SERVICEMODE_DEFAULT;
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
							serviceOptions->ulProfileMode = PROFILEMODE_ONE;
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
					else if (tolower(argv[i][3]) == 'u')
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

	// Parse the command line arguments
	if (!ValidateScenario(argc, argv, tkOptions))
	{
		if (tkOptions->ulLoggingMode != loggingModeNone)
		{
			DisplayUsage();
		}
		return;
	}
	Logger::SetLoggingMode((LoggingMode)tkOptions->ulLoggingMode);

	loggingMode = (LoggingMode)tkOptions->ulLoggingMode;
	ProfileInfo profInfo;
	ProfileInfo * lpProfInfo = &profInfo; 

	MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
	if (SUCCEEDED(MAPIInitialize(&MAPIINIT)))
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo * profileInfo = new ProfileInfo[ulProfileCount];
		//HrGetProfiles(ulProfileCount, profileInfo);

		switch (tkOptions->profileOptions->ulProfileAction)
		{
		case ACTION_ADD:
			EC_HRES_MSG(HrCreateProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str()), L"Calling HrCreateProfile");
			if (tkOptions->serviceOptions->ulServiceAction == ACTION_ADD)
			EC_HRES_LOG(HrCreateMsemsService(PROFILEMODE_ONE,
				(LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(),
				tkOptions->iOutlookVersion,
				tkOptions->serviceOptions), L"Calling HrCreateMsemsService");
			break;
		case ACTION_EDIT:
			if (tkOptions->profileOptions->bSetDefaultProfile)
			{
				HrSetDefaultProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str());
			}
			switch (tkOptions->serviceOptions->ulServiceAction)
			{
			case ACTION_ADD:
				EC_HRES_LOG(HrCreateMsemsService(tkOptions->profileOptions->ulProfileMode == PROFILEMODE_DEFAULT, 
					(LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(), 
					tkOptions->iOutlookVersion, 
					tkOptions->serviceOptions), L"Calling HrCreateMsemsService");
			case ACTION_EDIT:

				switch (tkOptions->mailboxOptions->ulMailboxAction)
				{
				case ACTION_ADD:
					EC_HRES_LOG(HrAddDelegateMailbox(tkOptions->profileOptions->ulProfileMode,
						(LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(),
						tkOptions->serviceOptions->ulServiceMode == SERVICEMODE_DEFAULT,
						tkOptions->serviceOptions->iServiceIndex,
						tkOptions->iOutlookVersion,
						tkOptions->mailboxOptions), L"Calling HrAddDelegateMailbox");
					break;
				case ACTION_EDIT:
				case ACTION_REMOVE:
					break;
				case ACTION_PROMOTEDELEGATE:
					EC_HRES_LOG(HrPromoteDelegates((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(),
						tkOptions->profileOptions->ulProfileMode == PROFILEMODE_DEFAULT,
						tkOptions->profileOptions->ulProfileMode == PROFILEMODE_ALL,
						tkOptions->serviceOptions->iServiceIndex,
						tkOptions->serviceOptions->ulServiceMode == SERVICEMODE_DEFAULT,
						tkOptions->serviceOptions->ulServiceMode == SERVICEMODE_ALL,
						tkOptions->iOutlookVersion,
						tkOptions->serviceOptions->ulConnectMode), L"Calling HrPromoteDelegates");
					// If Caching options were specified then update the cached mode configuration accordingly
					if ((tkOptions->serviceOptions->ulCachedModeOwner > 0) || (tkOptions->serviceOptions->ulCachedModeShared > 0) || (tkOptions->serviceOptions->ulCachedModePublicFolder > 0))
					{
						// This is not yet implemented
					}
					break;
				case ACTION_ADDDELEGATE:

				case ACTION_LIST:

					break;
				};

			case ACTION_UPDATE:
				break;
			case ACTION_LIST:
				break;
			case ACTION_ENABLECACHEDMODE:
				if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_DEFAULT)
				{
					EC_HRES_LOG(HrSetCachedMode((LPWSTR)GetDefaultProfileName().c_str(), true, false, -1, tkOptions->serviceOptions->ulServiceMode == SERVICEMODE_DEFAULT, tkOptions->serviceOptions->ulServiceMode == SERVICEMODE_ALL, tkOptions->serviceOptions->ulCachedModeOwner == 1, tkOptions->serviceOptions->ulCachedModeShared == 1, tkOptions->serviceOptions->ulCachedModePublicFolder == 1, tkOptions->serviceOptions->iCachedModeMonths, tkOptions->iOutlookVersion), L"HrSetCachedMode");
				}
				else if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_ONE)
				{
					EC_HRES_LOG(HrSetCachedMode((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(), false, false, -1, tkOptions->serviceOptions->ulServiceMode == SERVICEMODE_DEFAULT, tkOptions->serviceOptions->ulServiceMode == SERVICEMODE_ALL,  tkOptions->serviceOptions->ulCachedModeOwner == 1, tkOptions->serviceOptions->ulCachedModeShared == 1, tkOptions->serviceOptions->ulCachedModePublicFolder == 1, tkOptions->serviceOptions->iCachedModeMonths, tkOptions->iOutlookVersion), L"HrSetCachedMode");
				}
				
				break;
			};
			break;
		case ACTION_LIST:
			EC_HRES_LOG(HrListProfiles(tkOptions->profileOptions, tkOptions->wszExportPath), L"Calling HrListProfiles");
			break;
		case ACTION_CLONE:

			MAPIAllocateBuffer(sizeof(ProfileInfo), (LPVOID*)lpProfInfo);
			ZeroMemory(lpProfInfo, sizeof(ProfileInfo));

			if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_DEFAULT)
			{
				EC_HRES_LOG(HrGetProfile((LPWSTR)GetDefaultProfileName().c_str(), &profInfo), L"Calling HrGetProfile");
			}
			else if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_ONE)
			{
				EC_HRES_LOG(HrGetProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(), &profInfo), L"Calling HrGetProfile");

			}
			EC_HRES_LOG(HrCloneProfile(&profInfo), L"Calling HrCloneProfile");
			break;
		case ACTION_SIMPLECLONE:

			MAPIAllocateBuffer(sizeof(ProfileInfo), (LPVOID*)lpProfInfo);
			ZeroMemory(lpProfInfo, sizeof(ProfileInfo));

			if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_DEFAULT)
			{
				EC_HRES_LOG(HrGetProfile((LPWSTR)GetDefaultProfileName().c_str(), &profInfo), L"Calling HrGetProfile");
			}
			else if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_ONE)
			{
				EC_HRES_LOG(HrGetProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(), &profInfo), L"Calling HrGetProfile");

			}
			EC_HRES_LOG(HrSimpleCloneProfile(&profInfo, tkOptions->profileOptions->bSetDefaultProfile), L"Calling HrCloneProfile");
			break;
		case ACTION_REMOVE:
			if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_DEFAULT)
			{
				EC_HRES_LOG(HrDeleteProfile((LPWSTR)GetDefaultProfileName().c_str()), L"HrDeleteProfile");
			}
			else if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_ONE)
			{
				EC_HRES_LOG(HrDeleteProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str()), L"HrDeleteProfile");
			}


			break;
		};

		//switch (tkOptions->ulScenario)
		//{
		//case SCENARIO_PROFILE:
		//	if (tkOptions->ulActionType == ACTIONTYPE_STANDARD)
		//	{
		//		if (tkOptions->ulAction == STANDARDACTION_ADD)
		//		{
		//			// this only works with the default profile for now
		//			HrCreateProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str());

		//		}

		//		else if (tkOptions->ulAction == STANDARDACTION_LIST)
		//		{

		//			if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_ALL)
		//			{
		//				ULONG ulProfileCount = GetProfileCount();
		//				ProfileInfo * profileInfo = new ProfileInfo[ulProfileCount];
		//				ZeroMemory(profileInfo, sizeof(ProfileInfo) * ulProfileCount);
		//				Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for all profiles");
		//				EC_HRES_MSG(HrGetProfiles(ulProfileCount, profileInfo), L"Calling HrGetProfiles");
		//				if (tkOptions->wszExportPath != L"")
		//				{
		//					Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles");
		//					ExportXML(ulProfileCount, profileInfo, tkOptions->wszExportPath);
		//				}
		//				else
		//				{
		//					Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles");
		//					ExportXML(ulProfileCount, profileInfo, L"");
		//				}

		//			}
		//			if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_ONE)
		//			{
		//				ProfileInfo * pProfileInfo = new ProfileInfo();
		//				Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for profile: " + tkOptions->profileOptions->wszProfileName);
		//				EC_HRES_MSG(HrGetProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(), pProfileInfo), L"Calling HrGetProfile");
		//				if (tkOptions->wszExportPath != L"")
		//				{
		//					Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile");
		//					ExportXML(1, pProfileInfo, tkOptions->wszExportPath);
		//				}
		//				else
		//				{
		//					Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile");
		//					ExportXML(1, pProfileInfo, L"");
		//				}

		//			}
		//			if (tkOptions->profileOptions->ulProfileMode == PROFILEMODE_DEFAULT)
		//			{
		//				std::wstring szDefaultProfileName = GetDefaultProfileName();
		//				if (!szDefaultProfileName.empty())
		//				{
		//					tkOptions->profileOptions->wszProfileName = szDefaultProfileName;
		//				}

		//				ProfileInfo * pProfileInfo = new ProfileInfo();
		//				Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for default profile: " + tkOptions->profileOptions->wszProfileName);
		//				EC_HRES_MSG(HrGetProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(), pProfileInfo), L"Calling HrGetProfile");
		//				if (tkOptions->wszExportPath != L"")
		//				{
		//					Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile");
		//					ExportXML(1, pProfileInfo, tkOptions->wszExportPath);
		//				}
		//				else
		//				{
		//					Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile");
		//					ExportXML(1, pProfileInfo, L"");
		//				}
		//			}
		//		}
		//	}
		//	break;
		//case SCENARIO_SERVICE:
		//	if (tkOptions->ulActionType == ACTIONTYPE_STANDARD)
		//	{
		//		if (tkOptions->ulAction == STANDARDACTION_ADD)
		//		{
		//			if (tkOptions->iOutlookVersion == 2007)
		//			{
		//				HrCreateMsemsServiceLegacyUnresolved((tkOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
		//					(LPWSTR)tkOptions->serviceOptions->wszProfileName.c_str(),
		//					(LPWSTR)tkOptions->serviceOptions->wszMailboxLegacyDN.c_str(),
		//					(LPWSTR)tkOptions->serviceOptions->wszServerDisplayName.c_str());
		//			}
		//			else if ((tkOptions->iOutlookVersion == 2010) || (tkOptions->iOutlookVersion == 2013))
		//			{
		//				// this only works with the default profile for now
		//				if (tkOptions->serviceOptions->ulConnectMode == CONNECT_ROH)
		//				{
		//					HrCreateMsemsServiceROH((tkOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
		//						(LPWSTR)tkOptions->serviceOptions->wszProfileName.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszSmtpAddress.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszMailboxLegacyDN.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszUnresolvedServer.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszRohProxyServer.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszServerLegacyDN.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszAutodiscoverUrl.c_str());
		//				}
		//				else if (tkOptions->serviceOptions->ulConnectMode == CONNECT_MOH)
		//				{
		//					HrCreateMsemsServiceMOH((tkOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
		//						(LPWSTR)tkOptions->serviceOptions->wszProfileName.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszSmtpAddress.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszMailboxLegacyDN.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszServerLegacyDN.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszMailStoreInternalUrl.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszMailStoreExternalUrl.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszAddressBookInternalUrl.c_str(),
		//						(LPWSTR)tkOptions->serviceOptions->wszAddressBookExternalUrl.c_str());
		//				}
		//				else
		//				{

		//				}
		//			}
		//			else // default to the 2016 logic
		//			{
		//				// this only works with the default profile for now
		//				HrCreateMsemsServiceModern((tkOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
		//					(LPWSTR)tkOptions->serviceOptions->wszProfileName.c_str(),
		//					(LPWSTR)tkOptions->serviceOptions->wszSmtpAddress.c_str(),
		//					(LPWSTR)tkOptions->serviceOptions->wszMailboxDisplayName.c_str());
		//			}
		//		}
		//	}
		//	break;
		//case SCENARIO_MAILBOX:
		//	if (tkOptions->ulActionType == ACTIONTYPE_STANDARD)
		//	{
		//		if (tkOptions->ulAction == STANDARDACTION_ADD)
		//		{
		//			if (tkOptions->iOutlookVersion == 2007)
		//			{
		//				// this only works with the default profile for now
		//				HrAddDelegateMailboxLegacy((tkOptions->mailboxOptions->ulProfileMode == PROFILEMODE_DEFAULT),
		//					(LPWSTR)tkOptions->mailboxOptions->wszProfileName.c_str(),
		//					tkOptions->mailboxOptions->bDefaultService,
		//					tkOptions->mailboxOptions->ulServiceIndex,
		//					(LPWSTR)tkOptions->mailboxOptions->wszMailboxDisplayName.c_str(),
		//					(LPWSTR)tkOptions->mailboxOptions->wszMailboxLegacyDN.c_str(),
		//					(LPWSTR)tkOptions->mailboxOptions->wszServerDisplayName.c_str(),
		//					(LPWSTR)tkOptions->mailboxOptions->wszServerLegacyDN.c_str());
		//			}
		//			else if ((tkOptions->iOutlookVersion == 2010) || (tkOptions->iOutlookVersion == 2013))
		//			{
		//				// this only works with the default profile for now
		//				HrAddDelegateMailbox((tkOptions->mailboxOptions->ulProfileMode == PROFILEMODE_DEFAULT),
		//					(LPWSTR)tkOptions->mailboxOptions->wszProfileName.c_str(),
		//					tkOptions->mailboxOptions->bDefaultService,
		//					tkOptions->mailboxOptions->ulServiceIndex,
		//					(LPWSTR)tkOptions->mailboxOptions->wszMailboxDisplayName.c_str(),
		//					(LPWSTR)tkOptions->mailboxOptions->wszMailboxLegacyDN.c_str(),
		//					(LPWSTR)tkOptions->mailboxOptions->wszServerDisplayName.c_str(),
		//					(LPWSTR)tkOptions->mailboxOptions->wszServerLegacyDN.c_str(),
		//					(LPWSTR)tkOptions->mailboxOptions->wszSmtpAddress.c_str(),
		//					NULL,
		//					0,
		//					0,
		//					NULL);
		//			}
		//			else // default to the 2016 logic
		//			{
		//				// this only works with the default profile for now
		//				HrAddDelegateMailboxModern((tkOptions->mailboxOptions->ulProfileMode == PROFILEMODE_DEFAULT),
		//					(LPWSTR)tkOptions->mailboxOptions->wszProfileName.c_str(),
		//					tkOptions->mailboxOptions->bDefaultService,
		//					tkOptions->mailboxOptions->ulServiceIndex,
		//					(LPWSTR)tkOptions->mailboxOptions->wszMailboxDisplayName.c_str(),
		//					(LPWSTR)tkOptions->mailboxOptions->wszSmtpAddress.c_str());
		//			}
		//		}
		//	}
		//	break;
		//}
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
		//	//			EC_HRES_MSG(GetProfiles(ulProfileCount, profileInfo, loggingMode), loggingMode);
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
		//	//			EC_HRES_MSG(GetProfile((LPWSTR)tkOptions.profileOptions->wszProfileName.c_str(), &profileInfo, loggingMode), loggingMode);
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
		//	//			EC_HRES_MSG(GetProfile((LPWSTR)tkOptions.szProfileName.c_str(), &profileInfo, loggingMode), loggingMode);
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
		//	//			EC_HRES_MSG(UpdateCachedModeConfig((LPSTR)tkOptions.szProfileName.c_str(), tkOptions.ulServiceIndex, tkOptions.ulCachedModeOwner, tkOptions.ulCachedModeShared, tkOptions.ulCachedModePublicFolder, tkOptions.iCachedModeMonths, loggingMode), loggingMode);
		//	//		}
		//	//		break;
		//	//	}
		//	//	break;
		//	//case RUNNINGMODE_PST:
		//	//	if (tkOptions.szPstOldPath.empty())
		//	//	{
		//	//		EC_HRES_MSG(UpdatePstPath((LPWSTR)tkOptions.szProfileName.c_str(), (LPWSTR)tkOptions.szPstNewPath.c_str(), tkOptions.bPstMoveFiles, loggingMode), loggingMode);
		//	//	}
		//	//	else
		//	//	{
		//	//		EC_HRES_MSG(UpdatePstPath((LPWSTR)tkOptions.szProfileName.c_str(), (LPWSTR)tkOptions.szPstOldPath.c_str(), (LPWSTR)tkOptions.szPstNewPath.c_str(), tkOptions.bPstMoveFiles, loggingMode), loggingMode);
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
		Logger::Write(logLevelError, oss.str());
	}

Error:
	goto Cleanup;
Cleanup:
	// Free up memory

	return;
}


HRESULT HrListProfiles(ProfileOptions * pProfileOptions, std::wstring wszExportPath)
{
	HRESULT hRes = S_OK;
	if (pProfileOptions->ulProfileMode == PROFILEMODE_ALL)
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo * profileInfo = new ProfileInfo[ulProfileCount];
		ZeroMemory(profileInfo, sizeof(ProfileInfo) * ulProfileCount);
		Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for all profiles");
		EC_HRES_MSG(HrGetProfiles(ulProfileCount, profileInfo), L"Calling HrGetProfiles");
		if (wszExportPath != L"")
		{
			Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles");
			ExportXML(ulProfileCount, profileInfo, wszExportPath);
		}
		else
		{
			Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles");
			ExportXML(ulProfileCount, profileInfo, L"");
		}

	}
	if (pProfileOptions->ulProfileMode == PROFILEMODE_ONE)
	{
		ProfileInfo * pProfileInfo = new ProfileInfo();
		Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for profile: " + pProfileOptions->wszProfileName);
		EC_HRES_MSG(HrGetProfile((LPWSTR)pProfileOptions->wszProfileName.c_str(), pProfileInfo), L"Calling HrGetProfile");
		if (wszExportPath != L"")
		{
			Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile");
			ExportXML(1, pProfileInfo, wszExportPath);
		}
		else
		{
			Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile");
			ExportXML(1, pProfileInfo, L"");
		}

	}
	if (pProfileOptions->ulProfileMode == PROFILEMODE_DEFAULT)
	{
		std::wstring szDefaultProfileName = GetDefaultProfileName();
		if (!szDefaultProfileName.empty())
		{
			pProfileOptions->wszProfileName = szDefaultProfileName;
		}

		ProfileInfo * pProfileInfo = new ProfileInfo();
		Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for default profile: " + pProfileOptions->wszProfileName);
		EC_HRES_MSG(HrGetProfile((LPWSTR)pProfileOptions->wszProfileName.c_str(), pProfileInfo), L"Calling HrGetProfile");
		if (wszExportPath != L"")
		{
			Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile");
			ExportXML(1, pProfileInfo, wszExportPath);
		}
		else
		{
			Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile");
			ExportXML(1, pProfileInfo, L"");
		}
	}

Error:
	return hRes;
}