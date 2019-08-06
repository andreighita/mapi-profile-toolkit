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
#pragma once
#include "stdafx.h"
#include "MapiProfileToolkit.h"

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
		if (szOLVer == L"Outlook.Application.19")
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
		else if (szOLVer == L"Outlook.Application.16")
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

BOOL ValidateScenario(int argc, _TCHAR* argv[], RuntimeOptions* pRunOpts)
{
	std::vector<std::string> wszDiscardedArgs;
	if (!pRunOpts) return FALSE;
	ZeroMemory(pRunOpts, sizeof(RuntimeOptions));
	pRunOpts->action = ACTION_UNSPECIFIED;
	int iThreeParam = 0;
	pRunOpts->iOutlookVersion = GetOutlookVersion();
	pRunOpts->loggingMode = LoggingMode::LoggingModeConsole;

	pRunOpts->profileOptions = new ProfileOptions();
	pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Default;
	pRunOpts->profileOptions->serviceOptions = new ServiceOptions();
	pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_Default;
	pRunOpts->profileOptions->serviceOptions->connectMode = ConnectMode::ConnectMode_RpcOverHttp;
	pRunOpts->profileOptions->serviceOptions->providerOptions = new ProviderOptions();
	pRunOpts->profileOptions->serviceOptions->addressBookOptions = new AddressBookOptions();

	for (int i = 1; i < argc; i++)
	{
		std::wstring wsArg = argv[i];
		std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

		if ((wsArg == L"-exportpath") || (wsArg == L"-ep"))
		{
			std::wstring wszExportPath = argv[i + 1];
			std::transform(wszExportPath.begin(), wszExportPath.end(), wszExportPath.begin(), ::tolower);
			pRunOpts->wszExportPath = wszExportPath;
			pRunOpts->exportMode = ExportMode::Export;
			i++;
		}
		else if ((wsArg == L"-addressbookdisplayname") || (wsArg == L"-abdn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszABDisplayName = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-addressbookservername") || (wsArg == L"-absn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszABServerName = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-addressbookconfigfilepath") || (wsArg == L"-abcfp"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-profile") || (wsArg == L"-p"))
		{
			if (i + 1 < argc)
			{
				std::wstring wszValue = argv[i + 1];
				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
				if (wszValue == L"add")
				{
					pRunOpts->action |= ACTION_PROFILE_ADD;
					i++;
				}
				else if (wszValue == L"update")
				{
					pRunOpts->action |= ACTION_PROFILE_RENAME;
					i++;
				}
				else if (wszValue == L"remove")
				{
					pRunOpts->action |= ACTION_PROFILE_REMOVE;
					i++;
				}
				else if (wszValue == L"removeall")
				{
					pRunOpts->action |= ACTION_PROFILE_REMOVEALL;
					i++;
				}
				else if (wszValue == L"list")
				{
					pRunOpts->action |= ACTION_PROFILE_LIST;
					i++;
				}
				else if (wszValue == L"listall")
				{
					pRunOpts->action |= ACTION_PROFILE_LISTALL;
					i++;
				}
				else if (wszValue == L"clone")
				{
					pRunOpts->action |= ACTION_PROFILE_CLONE;
					i++;
				}
				else if (wszValue == L"promotedelegates")
				{
					pRunOpts->action |= ACTION_PROFILE_PROMOTEDELEGATES;
					i++;
				}
				else if (wszValue == L"setdefault")
				{
					pRunOpts->action |= ACTION_PROFILE_SETDEFAULT;
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
					pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Default;
					i++;

				}
				else if (wszValue == L"specific")
				{
					pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Specific;
					i++;

				}
				else if (wszValue == L"all")
				{
					pRunOpts->profileOptions->profileMode = ProfileMode::Mode_All;
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
				pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Specific;
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
					pRunOpts->action |= ACTION_SERVICE_ADD;
					i++;
				}
				else if (wszValue == L"update")
				{
					pRunOpts->action |= ACTION_SERVICE_UPDATE;
					i++;
				}
				else if (wszValue == L"remove")
				{
					pRunOpts->action |= ACTION_SERVICE_REMOVE;
					i++;
				}
				else if (wszValue == L"removeall")
				{
					pRunOpts->action |= ACTION_SERVICE_REMOVEALL;
					i++;
				}
				else if (wszValue == L"list")
				{
					pRunOpts->action |= ACTION_SERVICE_LIST;
					i++;
				}
				else if (wszValue == L"listall")
				{
					pRunOpts->action |= ACTION_SERVICE_LISTALL;
					i++;
				}
				else if (wszValue == L"setcachedmode")
				{
					pRunOpts->action |= ACTION_SERVICE_SETCACHEDMODE;
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
					pRunOpts->profileOptions->serviceOptions->serviceType = ServiceType::ServiceType_ExchangeAccount;
					i++;
				}
				else if (wszValue == L"pst")
				{
					pRunOpts->profileOptions->serviceOptions->serviceType = ServiceType::ServiceType_DataFile;
					i++;
				}
				else if (wszValue == L"addressbook")
				{
					pRunOpts->profileOptions->serviceOptions->serviceType = ServiceType::ServiceType_AddressBook;
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
					pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_Default;
					i++;

				}
				else if (wszValue == L"specific")
				{
					pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_Specific;
					i++;

				}
				else if (wszValue == L"all")
				{
					pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_All;
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
					pRunOpts->action |= ACTION_PROVIDER_ADD;
					i++;
				}
				else if (wszValue == L"update")
				{
					pRunOpts->action |= ACTION_PROVIDER_UPDATE;
					i++;
				}
				else if (wszValue == L"remove")
				{
					pRunOpts->action |= ACTION_PROVIDER_REMOVE;
					i++;
				}
				else if (wszValue == L"removeall")
				{
					pRunOpts->action |= ACTION_PROVIDER_REMOVEALL;
					i++;
				}
				else if (wszValue == L"list")
				{
					pRunOpts->action |= ACTION_PROVIDER_LIST;
					i++;
				}
				else if (wszValue == L"listall")
				{
					pRunOpts->action |= ACTION_PROVIDER_LIST;
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
					pRunOpts->profileOptions->serviceOptions->providerOptions->providerType = ProviderType::PrimaryMailbox;
					i++;
				}
				else if (wszValue == L"delegate")
				{
					pRunOpts->profileOptions->serviceOptions->providerOptions->providerType = ProviderType::Delegate;
					i++;
				}
				else if (wszValue == L"publicfolder")
				{
					pRunOpts->profileOptions->serviceOptions->providerOptions->providerType = ProviderType::PublicFolder;
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
			pRunOpts->profileOptions->serviceOptions->bSetDefaultservice = true;

		}
		else if ((wsArg == L"-cachedmodemonths") || (wsArg == L"-cmm"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->iCachedModeMonths = _wtoi(argv[i + 1]);
				i++;

			}
		}
		else if ((wsArg == L"-serviceindex") || (wsArg == L"-si"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->iServiceIndex = _wtoi(argv[i + 1]);
				i++;

			}
		}
		else if ((wsArg == L"-abexternalurl") || (wsArg == L"-abeu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszAddressBookExternalUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-abinternalurl") || (wsArg == L"-abiu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszAddressBookInternalUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-autodiscoverurl") || (wsArg == L"-au"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszAutodiscoverUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-mailboxdisplayname") || (wsArg == L"-mdn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszMailboxDisplayName = argv[i + 1];
				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailboxDisplayName = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-mailboxlegacydn") || (wsArg == L"-mldn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszMailboxLegacyDN = argv[i + 1];
				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailboxLegacyDN = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-mailstoreexternalurl") || (wsArg == L"-mseu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszMailStoreExternalUrl = argv[i + 1];
				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailStoreExternalUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-mailstoreinternalurl") || (wsArg == L"-msiu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszMailStoreInternalUrl = argv[i + 1];
				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailStoreExternalUrl = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-rohproxyserver") || (wsArg == L"-rps"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszRohProxyServer = argv[i + 1];
				pRunOpts->profileOptions->serviceOptions->providerOptions->wszRohProxyServer = argv[i + 1];
				i++;
			}
		}
		else if ((wsArg == L"-rohproxyserverflags") || (wsArg == L"-rpsf"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->providerOptions->ulRohProxyServerFlags = _wtoi(argv[i + 1]);
				i++;
			}
		}
		else if ((wsArg == L"-rohproxyserverauthpackage") || (wsArg == L"-mrpsap"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->providerOptions->ulRohProxyServerAuthPackage = _wtoi(argv[i + 1]);
				i++;
			}
		}
		else if ((wsArg == L"-serverdisplayname") || (wsArg == L"-sdn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszServerDisplayName = argv[i + 1];
				pRunOpts->profileOptions->serviceOptions->providerOptions->wszServerDisplayName = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-serverlegacydn") || (wsArg == L"-sldn"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszServerLegacyDN = argv[i + 1];
				pRunOpts->profileOptions->serviceOptions->providerOptions->wszServerLegacyDN = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-smtpaddress") || (wsArg == L"-sa"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszSmtpAddress = argv[i + 1];
				pRunOpts->profileOptions->serviceOptions->providerOptions->wszSmtpAddress = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-unresolvedserver") || (wsArg == L"-us"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszUnresolvedServer = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-unresolveduser") || (wsArg == L"-uu"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->wszUnresolvedUser = argv[i + 1];
				i++;

			}
		}
		else if ((wsArg == L"-cachedmodeowner") || (wsArg == L"-cmo"))
		{
			pRunOpts->profileOptions->serviceOptions->cachedModeOwner = CachedMode::Enabled;

		}
		else if ((wsArg == L"-cachedmodepublicfolder") || (wsArg == L"-cmpf"))
		{
			pRunOpts->profileOptions->serviceOptions->cachedModePublicFolders = CachedMode::Enabled;

		}
		else if ((wsArg == L"-cachedmodeshared") || (wsArg == L"-cms"))
		{
			pRunOpts->profileOptions->serviceOptions->cachedModeShared = CachedMode::Enabled;

		}
		else if ((wsArg == L"-configflags") || (wsArg == L"-cf"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->ulConfigFlags = _wtol(argv[i + 1]);
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
					pRunOpts->profileOptions->serviceOptions->connectMode = ConnectMode::ConnectMode_RpcOverHttp;
					i++;

				}
				if (wszValue == L"moh")
				{
					pRunOpts->profileOptions->serviceOptions->connectMode = ConnectMode::ConnectMode_MapiOverHttp;
					i++;

				}
			}
		}
		else if ((wsArg == L"-resourceflags") || (wsArg == L"-rf"))
		{
			if (i + 1 < argc)
			{
				pRunOpts->profileOptions->serviceOptions->ulResourceFlags = _wtol(argv[i + 1]);
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
					pRunOpts->profileOptions->serviceOptions->cachedModeOwner = CachedMode::Enabled;
					i++;

				}
				if (wszValue == L"disable")
				{
					pRunOpts->profileOptions->serviceOptions->cachedModeOwner = CachedMode::Disabled;
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
					pRunOpts->profileOptions->serviceOptions->cachedModeShared = CachedMode::Enabled;
					i++;

				}
				if (wszValue == L"disable")
				{
					pRunOpts->profileOptions->serviceOptions->cachedModeShared = CachedMode::Disabled;
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
					pRunOpts->profileOptions->serviceOptions->cachedModePublicFolders = CachedMode::Enabled;
					i++;

				}
				if (wszValue == L"disable")
				{
					pRunOpts->profileOptions->serviceOptions->cachedModePublicFolders = CachedMode::Disabled;
					i++;

				}
			}
		}
		else return false;
	}

	// Address Book specific validation
	if VCHK(pRunOpts->profileOptions->serviceOptions->serviceType, ServiceType::ServiceType_AddressBook)
	{
		if FCHK(pRunOpts->action, ACTION_SERVICE_ADD)
		{
			if (pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath.empty())
			{
				return false;
			}
			else if (FCHK(pRunOpts->action, ACTION_SERVICE_UPDATE) ||
				FCHK(pRunOpts->action, ACTION_SERVICE_LIST) ||
				FCHK(pRunOpts->action, ACTION_SERVICE_REMOVE))
			{
				if (pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszABDisplayName.empty())
				{
					return false;
				}

			}
		}
	}
	return true;
}

void DisplayUsage()
{
	std::wprintf(L"ProfileToolkit - Profile Examination Tool\n");
	std::wprintf(L"    Lists profile settings and optionally enables or disables cached exchange \n");
	std::wprintf(L"    mode.\n");
	std::wprintf(L"\n");
	std::wprintf(L"Usage: ProfileToolkit [-?] [-pm <all, one, default>] [-pn profilename] \n");
	std::wprintf(L"       [-si serviceIndex] [-cmo <enable, disable>] [-cms <enable, disable>] \n");
	std::wprintf(L"       [-cmp <enable, disable>]	[-cmm <0, 1, 3, 6, 12, 24>] [-ep exportpath]\n");
	std::wprintf(L"\n");
	std::wprintf(L"Usage: ProfileToolkit [-?] [-profile <all, one, default>] [-profilename profilename] \n");
	std::wprintf(L"       [-si serviceIndex] [-cmo <enable, disable>] [-cms <enable, disable>] \n");
	std::wprintf(L"       [-cmp <enable, disable>]	[-cmm <0, 1, 3, 6, 12, 24>] [-ep exportpath]\n");
	std::wprintf(L"       [-addressbook <create, update, listall, listone>] [-addressbookdisplayname <displayname>] \n");
	std::wprintf(L"       [-addressbookservername <servername>] [-addressbookconfigfilepath <configfilepath>] \n");
	std::wprintf(L"\n");
	std::wprintf(L"Options:\n");
	std::wprintf(L"    -profile:                  \"all\" to process all profiles.\n");
	std::wprintf(L"							   \"default\" to process the default profile.\n");
	std::wprintf(L"                               \"one\" to process a specific profile. Prifile Name needs to be \n");
	std::wprintf(L"                               specified using -pn.\n");
	std::wprintf(L"                               Default profile will be used if -pm is not used.\n");
	std::wprintf(L"   -profilename:               Name of the profile to process.\n");
	std::wprintf(L"                               Default profile will be used if -pn is not used.\n");
	std::wprintf(L"\n");
	std::wprintf(L"   -addressbook:               \"create\" to create a new address book service.\n");
	std::wprintf(L"                               \"update\" to update an existing address book service. The display name.\n");
	std::wprintf(L"                               of the address book to update needs to be specified using -addressbookdisplayname.\n");
	std::wprintf(L"                               \"listall\" Tto list all address book services in the profile. \n");
	std::wprintf(L"                               \"listone\" to list an existing address book service. The display name.\n");
	std::wprintf(L"                               of the address book to list needs to be specified using -addressbookdisplayname.\n");
	std::wprintf(L"   -addressbookdisplayname:    The display name of the address book to create, update or list.\n");
	std::wprintf(L"   -addressbookservername:     The display name of the LDAP server configure in the address book.\n");
	std::wprintf(L"   -addressbookconfigfilepath: The display name of the LDAP server configure in the address book.\n");
	std::wprintf(L"   -profilename:               Name of the profile to process.\n");
	std::wprintf(L"                               Default profile will be used if -pn is not used.\n");
	std::wprintf(L"\n");
	//wprintf(L"       -si:    Index of the account to process from previous export.\n");
	//wprintf(L"       	       Must be used in conjunction with -pm one -pn profile or -pm default.\n");
	//wprintf(L"\n");
	//wprintf(L"       -cmo:   \"enable\" or \"disable\" for enabling or disabling cached Exchange \n");
	//wprintf(L"               mode on the owner's mailbox.\n");
	//wprintf(L"       	       Must be used in conjunction with -pm one -pn profile and -si index.\n");
	//wprintf(L"       -cms:   \"enable\" or \"disable\" for enabling or disabling cached Exchange \n");
	//wprintf(L"               mode on shared folders (delegate).\n");
	//wprintf(L"       	       Must be used in conjunction with -pm one -pn profile and -si index.\n");
	//wprintf(L"       -cmp:   \"enable\" or \"disable\" for enabling or disabling cached Exchange \n");
	//wprintf(L"               mode on public folders favorites.\n");
	//wprintf(L"       	       Must be used in conjunction with -pm one -pn profile and -si index.\n");
	//wprintf(L"       -cmm:   0 for all or 1, 3, 6, 12 or 24 for the same number of months to sync\n");
	//wprintf(L"       	       Must be used in conjunction with -pm one -pn profile, -si index and.\n");
	//wprintf(L"       	       -cmo enable.\n");
	//wprintf(L"\n");
	//wprintf(L"       -ep:    exportPath for exporting settings to disk.\n");
	//wprintf(L"\n");
	std::wprintf(L"       -?      Displays this usage information.\n");
}

LoggingMode loggingMode;

static std::wofstream ofsLogFile;
static std::wstring szLogFilePath;
static bool bIsLogFileOpen;

void _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hRes = S_OK;

	// Using the toolkip options to manage the runtime options
	RuntimeOptions* tkOptions = new RuntimeOptions();

	// Parse the command line arguments
	if (!ValidateScenario(argc, argv, tkOptions))
	{
		if (tkOptions->loggingMode != LoggingMode::LoggingModeNone)
		{
			DisplayUsage();
		}
		return;
	}
	else
	{
		if (!tkOptions->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath.empty())
		{
			// If a path was specified 
			if (!PathFileExists(LPCWSTR(tkOptions->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath.c_str())))
			{
				std::wprintf(L"WARNING: The specified file \"%s\" does not exsits.\n", LPTSTR(tkOptions->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath.c_str()));
				return;
			}
		}
	}
	Logger::SetLoggingMode((LoggingMode)tkOptions->loggingMode);

	loggingMode = (LoggingMode)tkOptions->loggingMode;
	ProfileInfo profInfo;
	ProfileInfo* lpProfInfo = &profInfo;

	MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
	if (SUCCEEDED(MAPIInitialize(&MAPIINIT)))
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo* profileInfo = new ProfileInfo[ulProfileCount];
		//HrGetProfiles(ulProfileCount, profileInfo);

		if (!FCHK(tkOptions->action, ACTION_UNSPECIFIED))
		{

			// Do we want to ADD a new profile?
			if FCHK(tkOptions->action, ACTION_PROFILE_ADD)
			{
				EC_HRES_MSG(HrCreateProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str()), L"Calling HrCreateProfile");

				// Do we also want to add a service?
				if FCHK(tkOptions->action, ACTION_SERVICE_ADD)
					EC_HRES_LOG(HrCreateMsemsService(tkOptions->profileOptions->profileMode,
					(LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(),
						tkOptions->iOutlookVersion,
						tkOptions->profileOptions->serviceOptions), L"Calling HrCreateMsemsService");
			}

			if FCHK(tkOptions->action, ACTION_PROFILE_SETDEFAULT)
			{
				HrSetDefaultProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str());
			}

			if FCHK(tkOptions->action, ACTION_PROFILE_RENAME)
			{
				// Are we adding a service?
				if FCHK(tkOptions->action, ACTION_SERVICE_ADD)
				{
					EC_HRES_LOG(HrCreateMsemsService(tkOptions->profileOptions->profileMode,
						(LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(),
						tkOptions->iOutlookVersion,
						tkOptions->profileOptions->serviceOptions), L"Calling HrCreateMsemsService");
				}

				if FCHK(tkOptions->action, ACTION_SERVICE_UPDATE)
				{
					if FCHK(tkOptions->action, ACTION_PROVIDER_ADD)
					{
						if VCHK(tkOptions->profileOptions->serviceOptions->providerOptions->providerType, ProviderType::Delegate)
						{
							EC_HRES_LOG(HrAddDelegateMailbox(tkOptions->profileOptions->profileMode,
								(LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(),
								tkOptions->profileOptions->serviceOptions->serviceMode,
								tkOptions->profileOptions->serviceOptions->iServiceIndex,
								tkOptions->iOutlookVersion,
								tkOptions->profileOptions->serviceOptions->providerOptions), L"Calling HrAddDelegateMailbox");
						}
					}
				}

				if FCHK(tkOptions->action, ACTION_PROFILE_PROMOTEDELEGATES)
				{
					EC_HRES_LOG(HrPromoteDelegates((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(),
						VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_Default),
						VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_All),
						tkOptions->profileOptions->serviceOptions->iServiceIndex,
						VCHK(tkOptions->profileOptions->serviceOptions->serviceMode, ServiceMode::Mode_Default),
						VCHK(tkOptions->profileOptions->serviceOptions->serviceMode, ServiceMode::Mode_All),
						tkOptions->iOutlookVersion,
						tkOptions->profileOptions->serviceOptions->connectMode), L"Calling HrPromoteDelegates");
					// If Caching options were specified then update the cached mode configuration accordingly
					// To do: add cached mode support
				}

				if FCHK(tkOptions->action, ACTION_SERVICE_SETCACHEDMODE)
					if VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_Default)
					{
						EC_HRES_LOG(HrSetCachedMode((LPWSTR)GetDefaultProfileName().c_str(), true, false, -1,
							VCHK(tkOptions->profileOptions->serviceOptions->serviceMode, ServiceMode::Mode_Default),
							VCHK(tkOptions->profileOptions->serviceOptions->serviceMode, ServiceMode::Mode_All),
							VCHK(tkOptions->profileOptions->serviceOptions->cachedModeOwner, CachedMode::Enabled),
							VCHK(tkOptions->profileOptions->serviceOptions->cachedModeShared, CachedMode::Enabled),
							VCHK(tkOptions->profileOptions->serviceOptions->cachedModePublicFolders, CachedMode::Enabled),
							tkOptions->profileOptions->serviceOptions->iCachedModeMonths, tkOptions->iOutlookVersion), L"HrSetCachedMode");
					}
					else if VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_Specific)
					{
						EC_HRES_LOG(HrSetCachedMode((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(), false, false, -1,
							VCHK(tkOptions->profileOptions->serviceOptions->serviceMode, ServiceMode::Mode_Default),
							VCHK(tkOptions->profileOptions->serviceOptions->serviceMode, ServiceMode::Mode_All),
							VCHK(tkOptions->profileOptions->serviceOptions->cachedModeOwner, CachedMode::Enabled),
							VCHK(tkOptions->profileOptions->serviceOptions->cachedModeShared, CachedMode::Enabled),
							VCHK(tkOptions->profileOptions->serviceOptions->cachedModePublicFolders, CachedMode::Enabled),
							tkOptions->profileOptions->serviceOptions->iCachedModeMonths, tkOptions->iOutlookVersion), L"HrSetCachedMode");
					}
					else if VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_All)
					{
						Logger::Write(LogLevel::logLevelFailed, L"Functionality not yet implemented");
					}
			}


			if FCHK(tkOptions->action, ACTION_PROFILE_LISTALL)
			{
				EC_HRES_LOG(HrListProfiles(tkOptions->profileOptions, tkOptions->wszExportPath), L"Calling HrListProfiles");
			}

			if FCHK(tkOptions->action, ACTION_PROFILE_CLONE)
			{
				MAPIAllocateBuffer(sizeof(ProfileInfo), (LPVOID*)lpProfInfo);
				ZeroMemory(lpProfInfo, sizeof(ProfileInfo));

				if VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_Default)
				{
					EC_HRES_LOG(HrGetProfile((LPWSTR)GetDefaultProfileName().c_str(), &profInfo), L"Calling HrGetProfile");
				}
				else if VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_Specific)
					EC_HRES_LOG(HrGetProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str(), &profInfo), L"Calling HrGetProfile");

			}
			EC_HRES_LOG(HrCloneProfile(&profInfo), L"Calling HrCloneProfile");


			if FCHK(tkOptions->action, ACTION_PROFILE_REMOVE)
			{
				if VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_Default)
				{
					EC_HRES_LOG(HrDeleteProfile((LPWSTR)GetDefaultProfileName().c_str()), L"HrDeleteProfile");
				}
				else if VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_Specific)
				{
					EC_HRES_LOG(HrDeleteProfile((LPWSTR)tkOptions->profileOptions->wszProfileName.c_str()), L"HrDeleteProfile");
				}
			};

			if VCHK(tkOptions->profileOptions->serviceOptions->serviceType, ServiceType::ServiceType_AddressBook)
			{

				// SOME LDAP AB LOGIC

				LPPROFADMIN lpProfAdmin = NULL;		// profile administration object pointer
				LPSERVICEADMIN lpSvcAdmin = NULL;	// service administration object pointer
				MAPIUID mapiUid = { 0 };			// MAPIUID structure
				LPMAPIUID lpMapiUid = &mapiUid;		// pointer to a MAPIUID structure
				BOOL fValidPath = false;
				BOOL fServiceExists = false;
				// Create a new ABProvider instance and set the service name to EMABLT (Address Book service)
				ABProvider pABProvider = { 0 };
				pABProvider.lpszServiceName = L"EMABLT";

				// Make sure the file path is valid and parse the XML to populate the ABProvider parameters
				if (!tkOptions->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath.empty())
				{
					fValidPath = true;
					EC_HRES_MSG(ParseConfigXml(LPTSTR(tkOptions->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath.c_str()), &pABProvider), L"Parsing AB config file");
				}

				// If we're processing the default profile then fetch the name of it and populate that in the runtime options.
				if VCHK(tkOptions->profileOptions->profileMode, ProfileMode::Mode_Default)
				{
					tkOptions->profileOptions->wszProfileName = GetDefaultProfileName();
					if (tkOptions->profileOptions->wszProfileName.empty())
					{
						wprintf(L"ERROR: No default profile found, please specify a valid profile name.");
						return;
					}
				}

				// Create a profile administration object.
				EC_HRES_MSG(MAPIAdminProfiles(0,		// Bitmask of flags indicating options for the service entry function. 
					&lpProfAdmin), L"Getting a profile admin interface pointer");					// Pointer to a pointer to the new profile administration object.
				wprintf(L"Retrieved IProfAdmin interface pointer.\n");

				// Get access to a message service administration object for making changes to the message services in a profile. 
				EC_HRES_MSG(lpProfAdmin->AdminServices(LPTSTR(tkOptions->profileOptions->wszProfileName.c_str()),	// A pointer to the name of the profile to be modified. The lpszProfileName parameter must not be NULL.
					NULL,																			// Always NULL. 
					NULL,																			// A handle of the parent window for any dialog boxes or windows that this method displays.
					0,																				// A bitmask of flags that controls the retrieval of the message service administration object. The following flags can be set:
					&lpSvcAdmin), L"Getting a service admin interface pointer");																	// A pointer to a pointer to a message service administration object.
				wprintf(L"Retrieved IMsgServiceAdmin interface pointer.\n");

				if FCHK(tkOptions->action, ACTION_SERVICE_ADD)
				{
					wprintf(L"Running in Create mode.\n");
					if (fValidPath)
					{
						// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
						// display name (and optionally, the ldap server name) supplied
						EC_HRES(CheckABServiceExists(lpSvcAdmin, pABProvider.lpszDisplayName, pABProvider.lpszServerName, &mapiUid, &fServiceExists));
						if (!fServiceExists)
							// If no existing service is found then call CreateAbService to create the new service
							EC_HRES(CreateABService(lpSvcAdmin, &pABProvider));
						else
							wprintf(L"The specified AB already exists.\n");
					}
					else
						wprintf(L"ERROR: Invalid input file or invalid file path.");

				}

				if FCHK(tkOptions->action, ACTION_SERVICE_UPDATE)
				{
									wprintf(L"Running in Update mode.\n");
				if (fValidPath)
				{
					if (!tkOptions->profileOptions->serviceOptions->addressBookOptions->wszABServerName.empty())
					{
						// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
						// display name (and optionally, the ldap server name) supplied
						EC_HRES(CheckABServiceExists(lpSvcAdmin, LPTSTR(tkOptions->profileOptions->serviceOptions->addressBookOptions->wszABDisplayName.c_str()), LPTSTR(tkOptions->profileOptions->serviceOptions->addressBookOptions->wszABServerName.c_str()), &mapiUid, &fServiceExists));
					}
					else
						// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
						// display name (and optionally, the ldap server name) supplied
						EC_HRES(CheckABServiceExists(lpSvcAdmin, LPTSTR(tkOptions->profileOptions->serviceOptions->addressBookOptions->wszABDisplayName.c_str()), &mapiUid, &fServiceExists));
					if (fServiceExists)
						// If the searched for service is found then call UpdateABService to update the service properties
						EC_HRES(UpdateABService(lpSvcAdmin, &pABProvider, lpMapiUid));
					else
						wprintf(L"The specified AB doesn't exist.\n");
				}
				else
					wprintf(L"ERROR: Invalid input file or invalid file path.");
				}

				if FCHK(tkOptions->action, ACTION_SERVICE_LISTALL)
				{
					wprintf(L"Running in List mode.\n");
					// Calling ListAllABServices to list all the existing Ldap AB Servies in the selected profile
					EC_HRES(ListAllABServices(lpSvcAdmin));
				}

				if FCHK(tkOptions->action, ACTION_SERVICE_REMOVE)
				{
									wprintf(L"Running in Remove mode.\n");
				if (!tkOptions->profileOptions->serviceOptions->addressBookOptions->wszABServerName.empty())
				{
					// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
					// display name (and optionally, the ldap server name) supplied
					EC_HRES(CheckABServiceExists(lpSvcAdmin, LPTSTR(tkOptions->profileOptions->serviceOptions->addressBookOptions->wszABDisplayName.c_str()), LPTSTR(tkOptions->profileOptions->serviceOptions->addressBookOptions->wszABServerName.c_str()), &mapiUid, &fServiceExists));
				}
				else
					// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
					// display name (and optionally, the ldap server name) supplied
					EC_HRES(CheckABServiceExists(lpSvcAdmin, LPTSTR(tkOptions->profileOptions->serviceOptions->addressBookOptions->wszABDisplayName.c_str()), &mapiUid, &fServiceExists));
				if (fServiceExists)
					// If the searched for service is found then call RemoveABService to update the service properties
					EC_HRES(RemoveABService(lpSvcAdmin, lpMapiUid));
				else
					wprintf(L"The specified AB doesn't exist.\n");
				}

			}
		}
	}



#pragma region SomeAncientCodeFromYesterYear
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
	//				HrCreateMsemsServiceLegacyUnresolved((tkOptions->profileOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
	//					(LPWSTR)tkOptions->profileOptions->serviceOptions->wszProfileName.c_str(),
	//					(LPWSTR)tkOptions->profileOptions->serviceOptions->wszMailboxLegacyDN.c_str(),
	//					(LPWSTR)tkOptions->profileOptions->serviceOptions->wszServerDisplayName.c_str());
	//			}
	//			else if ((tkOptions->iOutlookVersion == 2010) || (tkOptions->iOutlookVersion == 2013))
	//			{
	//				// this only works with the default profile for now
	//				if (tkOptions->profileOptions->serviceOptions->ulConnectMode == CONNECT_ROH)
	//				{
	//					HrCreateMsemsServiceROH((tkOptions->profileOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszProfileName.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszSmtpAddress.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszMailboxLegacyDN.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszUnresolvedServer.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszRohProxyServer.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszServerLegacyDN.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszAutodiscoverUrl.c_str());
	//				}
	//				else if (tkOptions->profileOptions->serviceOptions->ulConnectMode == CONNECT_MOH)
	//				{
	//					HrCreateMsemsServiceMOH((tkOptions->profileOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszProfileName.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszSmtpAddress.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszMailboxLegacyDN.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszServerLegacyDN.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszMailStoreInternalUrl.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszMailStoreExternalUrl.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszAddressBookInternalUrl.c_str(),
	//						(LPWSTR)tkOptions->profileOptions->serviceOptions->wszAddressBookExternalUrl.c_str());
	//				}
	//				else
	//				{

	//				}
	//			}
	//			else // default to the 2016 logic
	//			{
	//				// this only works with the default profile for now
	//				HrCreateMsemsServiceModern((tkOptions->profileOptions->serviceOptions->ulProfileMode == PROFILEMODE_DEFAULT),
	//					(LPWSTR)tkOptions->profileOptions->serviceOptions->wszProfileName.c_str(),
	//					(LPWSTR)tkOptions->profileOptions->serviceOptions->wszSmtpAddress.c_str(),
	//					(LPWSTR)tkOptions->profileOptions->serviceOptions->wszMailboxDisplayName.c_str());
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
#pragma endregion
	MAPIUninitialize();



#pragma region SomeMoreStuff
	//loggingMode = LoggingMode(tkOptions.ulLoggingMode);
	//// Check the curren't process' bitness vs Outlook's bitness and only run it if matched to avoid MAPI dialog boxes.
	//if (!IsCorrectBitness())
	//{
	//	Logger::Write(logLevelFailed, L"Unable to resolve bitness or bitness not matched.", loggingMode);
	//	return;
	//}
	//Logger::Write(logLevelSuccess, L"Bitness matched.", loggingMode);

	//try
	//{
	//	//MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
	//	//if (SUCCEEDED(MAPIInitialize(&MAPIINIT)))
	//	//{
	//	//	Logger::Write(logLevelSuccess, L"MAPI Initialised", loggingMode);
	//	//	//switch (tkOptions.ulScenario)
	//	//	//{
	//	//	//case SCENARIO_PROFILE:
	//	//	//	switch (tkOptions.profileOptions->ulProfileMode)
	//	//	//	{
	//	//	//	case PROFILEMODE_ALL:
	//	//	//			ULONG ulProfileCount = GetProfileCount(loggingMode);
	//	//	//			ProfileInfo * profileInfo = new ProfileInfo[ulProfileCount];
	//	//	//			ZeroMemory(profileInfo, sizeof(ProfileInfo) * ulProfileCount);
	//	//	//			Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for all profiles", loggingMode);
	//	//	//			EC_HRES_MSG(GetProfiles(ulProfileCount, profileInfo, loggingMode), loggingMode);
	//	//	//			if (tkOptions.wszExportPath != L"")
	//	//	//			{
	//	//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles", loggingMode);
	//	//	//				ExportXML(ulProfileCount, profileInfo, tkOptions.wszExportPath, loggingMode);
	//	//	//			}
	//	//	//			else
	//	//	//			{
	//	//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for all profiles", loggingMode);
	//	//	//				ExportXML(ulProfileCount, profileInfo, L"", loggingMode);
	//	//	//			}
	//	//	//		
	//	//	//		break;
	//	//	//	case PROFILEMODE_ONE:
	//	//	//		
	//	//	//			ProfileInfo profileInfo;
	//	//	//			Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for profile: " + tkOptions.profileOptions->wszProfileName, loggingMode);
	//	//	//			EC_HRES_MSG(GetProfile((LPWSTR)tkOptions.profileOptions->wszProfileName.c_str(), &profileInfo, loggingMode), loggingMode);
	//	//	//			if (tkOptions.wszExportPath != L"")
	//	//	//			{
	//	//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile", loggingMode);
	//	//	//				ExportXML(1, &profileInfo, tkOptions.szExportPath, loggingMode);
	//	//	//			}
	//	//	//			else
	//	//	//			{
	//	//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for profile", loggingMode);
	//	//	//				ExportXML(1, &profileInfo, L"", loggingMode);
	//	//	//			}
	//	//	//		break;
	//	//	//	case PROFILEMODE_DEFAULT:
	//	//	//		std::wstring szDefaultProfileName = GetDefaultProfileName(loggingMode);
	//	//	//		if (!szDefaultProfileName.empty())
	//	//	//		{
	//	//	//			tkOptions.szProfileName = szDefaultProfileName;
	//	//	//		}
	//	//	//		if (tkOptions.ulReadWriteMode == READWRITEMODE_READ)
	//	//	//		{
	//	//	//			ProfileInfo profileInfo;
	//	//	//			Logger::Write(logLevelInfo, L"Retrieving MAPI Profile information for default profile: " + tkOptions.szProfileName, loggingMode);
	//	//	//			EC_HRES_MSG(GetProfile((LPWSTR)tkOptions.szProfileName.c_str(), &profileInfo, loggingMode), loggingMode);
	//	//	//			if (tkOptions.szExportPath != L"")
	//	//	//			{
	//	//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile", loggingMode);
	//	//	//				ExportXML(1, &profileInfo, tkOptions.szExportPath, loggingMode);
	//	//	//			}
	//	//	//			else
	//	//	//			{
	//	//	//				Logger::Write(logLevelInfo, L"Exporting MAPI Profile information for default profile", loggingMode);
	//	//	//				ExportXML(1, &profileInfo, L"", loggingMode);
	//	//	//			}
	//	//	//		}
	//	//	//		else if (tkOptions.ulReadWriteMode == READWRITEMODE_WRITE)
	//	//	//		{
	//	//	//			Logger::Write(logLevelInfo, L"Updating cached mode configuration on default profile: " + tkOptions.szProfileName, loggingMode);
	//	//	//			EC_HRES_MSG(UpdateCachedModeConfig((LPSTR)tkOptions.szProfileName.c_str(), tkOptions.ulServiceIndex, tkOptions.ulCachedModeOwner, tkOptions.ulCachedModeShared, tkOptions.ulCachedModePublicFolder, tkOptions.iCachedModeMonths, loggingMode), loggingMode);
	//	//	//		}
	//	//	//		break;
	//	//	//	}
	//	//	//	break;
	//	//	//case RUNNINGMODE_PST:
	//	//	//	if (tkOptions.szPstOldPath.empty())
	//	//	//	{
	//	//	//		EC_HRES_MSG(UpdatePstPath((LPWSTR)tkOptions.szProfileName.c_str(), (LPWSTR)tkOptions.szPstNewPath.c_str(), tkOptions.bPstMoveFiles, loggingMode), loggingMode);
	//	//	//	}
	//	//	//	else
	//	//	//	{
	//	//	//		EC_HRES_MSG(UpdatePstPath((LPWSTR)tkOptions.szProfileName.c_str(), (LPWSTR)tkOptions.szPstOldPath.c_str(), (LPWSTR)tkOptions.szPstNewPath.c_str(), tkOptions.bPstMoveFiles, loggingMode), loggingMode);
	//	//	//	}
	//	//	//	break;
	//	//	//};
	//	//	//MAPIUninitialize();
	//	//}
	//}
	//catch (int exception)
	//{
	//	std::wostringstream oss; \
	//		oss << L"Error " << std::dec << exception << L" encountered";
	//	Logger::Write(logLevelError, oss.str());
	//}
#pragma endregion

	Error:
		 goto Cleanup;
	 Cleanup:
		 // Free up memory

		 return;
}


