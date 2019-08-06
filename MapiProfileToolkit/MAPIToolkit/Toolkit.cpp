#pragma once
// MAPIToolkit.cpp : Defines the functions for the static library.
//

#include "framework.h"
#include <mapidefs.h>
#include <guiddef.h>
#include <initguid.h>
#define USES_IID_IMAPIProp
#define USES_IID_IMsgServiceAdmin2
#include <atlchecked.h>
#include "ToolkitTypeDefs.h"
#include "RegistryHelper.h"
#include <iterator> 
#include <algorithm>
#include "ServiceWorker.h"
#include "ExchangeAccountWorker.h"
#include "AddressBookWorker.h"
#include "DataFileWorker.h"
#include "Logger.h"
#include "Toolkit.h"
#include "InlineAndMacros.h"
#include "Profile//Profile.h"
#include "PrimaryMailboxWorker.h"
#include "DelegateWorker.h"
#include "PublicFoldersWorker.h"
#include "ProfileWorker.h"
#include "ProviderWorker.h"


#pragma comment(lib, "Advapi32.lib")
#pragma comment(lib, "Mapi32.lib")
#pragma comment(lib, "Crypt32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shlwapi.lib")

#pragma warning(disable:4996) // _CRT_SECURE_NO_WARNINGS

namespace MAPIToolkit
{
	std::map<std::wstring, ULONG> Toolkit::g_actionsMap =
	{
		{ L"addprofile",			ACTION_PROFILE_ADD},
		{ L"addprovider",			ACTION_PROVIDER_ADD },
		{ L"addservice",			ACTION_SERVICE_ADD },
		{ L"changedatafilepath",	ACTION_SERVICE_CHANGEDATAFILEPATH },
		{ L"cloneprofile",			ACTION_PROFILE_CLONE },
		{ L"promotedelegates",		ACTION_PROFILE_PROMOTEDELEGATES },
		{ L"listallprofiles",		ACTION_PROFILE_LISTALL },
		{ L"listallproviders",		ACTION_PROVIDER_LISTALL },
		{ L"listallservices",		ACTION_SERVICE_LISTALL },
		{ L"listprofile",			ACTION_PROFILE_LIST },
		{ L"listprovider",			ACTION_PROVIDER_LIST },
		{ L"listservice",			ACTION_SERVICE_LIST },
		{ L"promoteonedelegate",	ACTION_PROFILE_PROMOTEONEDELEGATE },
		{ L"removeallprofiles",		ACTION_PROFILE_REMOVEALL },
		{ L"removeallproviders",	ACTION_PROVIDER_REMOVEALL },
		{ L"removeallservices",		ACTION_SERVICE_REMOVEALL },
		{ L"removeprofile",			ACTION_PROFILE_REMOVE },
		{ L"removeprovider",		ACTION_PROVIDER_REMOVE },
		{ L"removeservice",			ACTION_SERVICE_REMOVE },
		{ L"setcachedmode",			ACTION_SERVICE_SETCACHEDMODE },
		{ L"setdefaultprofile",		ACTION_PROFILE_SETDEFAULT },
		{ L"setdefaultservice",		ACTION_SERVICE_SETDEFAULT },
		{ L"renameprofile",			ACTION_PROFILE_RENAME },
		{ L"updateprovider",		ACTION_PROVIDER_UPDATE },
		{ L"updateservice",			ACTION_SERVICE_UPDATE }
	};

	std::map<std::wstring, ULONG> Toolkit::g_profileModeMap =
	{
		{ L"default",	PROFILEMODE_DEFAULT},
		{ L"specific",	PROFILEMODE_SPECIFIC },
		{ L"all",		PROFILEMODE_ALL }
	};

	std::map<std::wstring, ULONG> Toolkit::g_serviceModeMap =
	{
		{ L"default",	SERVICEMODE_DEFAULT},
		{ L"specific",	SERVICEMODE_SPECIFIC },
		{ L"all",		SERVICEMODE_ALL }
	};

	std::map<std::wstring, ULONG> Toolkit::g_serviceTypeMap =
	{
		{ L"addressbook",		SERVICETYPE_ADDRESSBOOK},
		{ L"datafile",			SERVICETYPE_DATAFILE },
		{ L"exchangeaccount",	SERVICETYPE_EXCHANGEACCOUNT },
		{ L"all",				SERVICETYPE_ALL }
	};

	ULONG Toolkit::m_action;
	int Toolkit::m_OutlookVersion;
	ULONG Toolkit::m_loggingMode;
	ServiceWorker* Toolkit::m_serviceWorker;
	ProviderWorker* Toolkit::m_providerWorker;
	ProfileWorker* Toolkit::m_profileWorker;
	ULONG Toolkit::m_profileCount;
	std::wstring Toolkit::m_wszExportPath;
	ULONG Toolkit::m_exportMode; // 0 = no export; 1 = export;
	std::wstring Toolkit::m_wszLogFilePath;
	ULONG Toolkit::m_profileMode; // pm
	LPPROFADMIN Toolkit::m_pProfAdmin;

	// Is64BitProcess
// Returns true if 64 bit process or false if 32 bit.
	BOOL Toolkit::Is64BitProcess(void)
	{
#if defined(_WIN64)
		return TRUE;   // 64-bit program
#else
		return FALSE;
#endif
	}



	void Toolkit::DisplayUsage()
	{
		std::wprintf(L"MAPIToolkit - MAPI profile utility\n");
		std::wprintf(L"    Allows management of Outlook / MAPI profiles at the command line.\n");
		std::wprintf(L"\n");
		std::wprintf(L"Usage: [-?] \n");
		std::wprintf(L"       [-action <addservice, listallservices, listservice, removeallservices, removeservice, updateservice>]\n");
		std::wprintf(L"       [-profilemode <default, specific, all>]\n");
		std::wprintf(L"       [-profilename name]\n");
		std::wprintf(L"       [-servicetype <addressbook>]\n");
		std::wprintf(L"       [-addressbookdisplayname displayname]\n");
		std::wprintf(L"       [-addressbookservername servername]\n");
		std::wprintf(L"       [-addressbookconfigfilepath configfilepath]\n");
		std::wprintf(L"\n");
		std::wprintf(L"Options:\n");
		std::wprintf(L" -?:			                Displays this information.\n");
		std::wprintf(L" -action:                    \"addservice\" adds a service with the type specified by \"servicetype\".\n");
		std::wprintf(L"								\"listallservices\" lists all services with the type specified by \"servicetype\".\n");
		std::wprintf(L"                             \"listservice\" lists a specific service with the type specified by \"servicetype\".\n");
		std::wprintf(L"                             \"removeallservices\" removes all services with the type specified by \"servicetype\".\n");
		std::wprintf(L"                             \"removeservice\" removes a specific service with the type specified by \"servicetype\".\n");
		std::wprintf(L"                             \"updateservice\" updates a specific service with the type specified by \"servicetype\".\n");
		std::wprintf(L" -profilemode:               \"default\" to run the selected action on the default profile.\n");
		std::wprintf(L"                             \"specific\" to run the selected action on the profile specified by the \"profilename\" value.\n");
		std::wprintf(L"                             \"all\" to run the selected action against all profiles.\n");
		std::wprintf(L"                             The default profile will be used if a profile mode is not specified.\n");
		std::wprintf(L"\n");
		std::wprintf(L" -profilename:               Name of the profile to run the specified actiona against. The profile name is mandatory\n");
		std::wprintf(L"                             if \"profilename\" is set to \"specific\" \n");
		std::wprintf(L" -servicetype:               \"addressbook\" to run an addressbook specific operation.\n");
		std::wprintf(L" -servicetype:               This is the only operation currently supported.\n");
		std::wprintf(L" -addressbookdisplayname:    The display name of the address book service to create, update, list or remove.\n");
		std::wprintf(L" -addressbookservername:     The display name of the LDAP server configured in the address book.\n");
		std::wprintf(L" -addressbookconfigfilepath: The path towards the address book configuration XML to use.\n");
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

	// GetOutlookVersion
	int Toolkit::GetOutlookVersion()
	{
		std::wstring szOLVer = L"";
		szOLVer = GetRegStringValue(HKEY_CLASSES_ROOT, TEXT("Outlook.Application\\CurVer"), NULL);
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
	BOOL _cdecl Toolkit::IsCorrectBitness()
	{

		std::wstring szOLVer = L"";
		std::wstring szOLBitness = L"";
		szOLVer = GetRegStringValue(HKEY_CLASSES_ROOT, TEXT("Outlook.Application\\CurVer"), NULL);
		if (szOLVer != L"")
		{
			if (szOLVer == L"Outlook.Application.19")
			{
				szOLBitness = GetRegStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\16.0\\Outlook"), TEXT("Bitness"));
				if (szOLBitness != L"")
				{
					if (szOLBitness == L"x64")
					{
						if (Toolkit::Is64BitProcess())
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
				szOLBitness = GetRegStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\16.0\\Outlook"), TEXT("Bitness"));
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
				szOLBitness = GetRegStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\15.0\\Outlook"), TEXT("Bitness"));
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
				szOLBitness = GetRegStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\14.0\\Outlook"), TEXT("Bitness"));
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

	BOOL Toolkit::Initialise()
	{
		m_action = ACTION_UNSPECIFIED;
		m_OutlookVersion = GetOutlookVersion();
		m_loggingMode = LOGGINGMODE_CONSOLE;
		m_serviceWorker = NULL;
		m_providerWorker = NULL;
		m_profileWorker = NULL;
		m_profileCount = 0;
		m_wszExportPath = L"";
		m_exportMode = EXPORTMODE_EXPORT; // 0 = no export; 1 = export;
		m_wszLogFilePath = L"";
		m_profileMode = PROFILEMODE_DEFAULT;
		m_pProfAdmin = NULL;

		MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };

		HCKM(MAPIInitialize(&MAPIINIT), L"Initialising MAPI");
		HCKM(MAPIAdminProfiles(0, &m_pProfAdmin), L"Getting profile administration interface pointer.");

		return TRUE;
	}

	VOID Toolkit::Uninitialise()
	{
		if (m_pProfAdmin) m_pProfAdmin->Release();
		MAPIUninitialize();
	}

	void Toolkit::Run(int argc, wchar_t* argv[])
	{
		m_action = ACTION_UNSPECIFIED;
		m_OutlookVersion = GetOutlookVersion();
		m_loggingMode = LOGGINGMODE_CONSOLE;
		m_serviceWorker = NULL;
		m_providerWorker = NULL;

		if (ParseParams(argc, argv))
		{

		}
		else
			DisplayUsage();
	}

	void Toolkit::RunAction()
	{
		switch (m_action)
		{

		case ACTION_UNSPECIFIED:
		{
			Logger::Write(LOGLEVEL_FAILED, L"You must specify an action");
			break;
		}
		case ACTION_PROFILE_ADD:
		{
			if (m_profileWorker)
			{
				m_profileWorker->AddProfile();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROFILE_CLONE:
		{
			if (m_profileWorker)
			{
				m_profileWorker->CloneProfile();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROFILE_RENAME:
		{
			if (m_profileWorker)
			{
				m_profileWorker->RenameProfile();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROFILE_LIST:
		{
			if (m_profileWorker)
			{
				m_profileWorker->ListProfile();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROFILE_LISTALL:
		{
			if (m_profileWorker)
			{
				m_profileWorker->ListAllProfiles();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROFILE_REMOVE:
		{
			if (m_profileWorker)
			{
				m_profileWorker->RemoveProfile();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROFILE_REMOVEALL:
		{
			if (m_profileWorker)
			{
				m_profileWorker->RemoveAllProfiles();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROFILE_SETDEFAULT:
		{
			if (m_profileWorker)
			{
				m_profileWorker->SetDefaultProfile();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROFILE_PROMOTEDELEGATES:
		{
			if (m_profileWorker)
			{
				m_profileWorker->PromoteDelegates();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROFILE_PROMOTEONEDELEGATE:
		{
			if (m_profileWorker)
			{
				m_profileWorker->PromoteOneDelegate();
			}
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROVIDER_ADD:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROVIDER_UPDATE:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROVIDER_LIST:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROVIDER_LISTALL:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROVIDER_REMOVE:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_PROVIDER_REMOVEALL:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_SERVICE_ADD:
		{
			if (m_serviceWorker)
			{
				if (SERVICETYPE_ADDRESSBOOK == m_serviceWorker->m_serviceType)
				{
					((AddressBookWorker*)m_serviceWorker)->AddAddressBookService();
					Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				}
				else if (SERVICETYPE_EXCHANGEACCOUNT == m_serviceWorker->m_serviceType)
				{
					((ExchangeAccountWorker*)m_serviceWorker)->AddExchangeAccount();
					Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				}
				else if (SERVICETYPE_DATAFILE == m_serviceWorker->m_serviceType)
				{
					((DataFileWorker*)m_serviceWorker)->AddDataFile();
					Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				}
			}
			break;
		}
		case ACTION_SERVICE_UPDATE:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_SERVICE_SETCACHEDMODE:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_SERVICE_LIST:
		{
			if (m_serviceWorker)
			{
				if (SERVICETYPE_ADDRESSBOOK == m_serviceWorker->m_serviceType)
				{
					((AddressBookWorker*)m_serviceWorker)->ListAddressBookService();
				}
			}
			else
				Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_SERVICE_LISTALL:
		{
			if (m_serviceWorker)
			{
				if (SERVICETYPE_ADDRESSBOOK == m_serviceWorker->m_serviceType)
				{
					((AddressBookWorker*)m_serviceWorker)->ListAllAddressBookServices();
				}
			}
			else
				Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_SERVICE_REMOVE:
		{
			if (m_serviceWorker)
			{
				if (SERVICETYPE_ADDRESSBOOK == m_serviceWorker->m_serviceType)
				{
					((AddressBookWorker*)m_serviceWorker)->RemoveAddressBookService();
				}
			}
			else
				Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_SERVICE_REMOVEALL:
		{
			if (m_serviceWorker)
			{
				if (SERVICETYPE_ADDRESSBOOK == m_serviceWorker->m_serviceType)
				{
					((AddressBookWorker*)m_serviceWorker)->RemoveAllAddressBookServices();
				}
			}
			else
				Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_SERVICE_CHANGEDATAFILEPATH:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case ACTION_SERVICE_SETDEFAULT:
		{
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		}
	}

	BOOL Toolkit::ParseParams(int argc, wchar_t* argv[])
	{
		// general toolkit
		for (int i = 1; i < argc; i++)
		{
			std::wstring wsArg = argv[i];
			std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

			if ((wsArg == L"-action") || (wsArg == L"-a"))
			{
				if (i + 1 < argc)
				{
					std::wstring wszValue = argv[i + 1];
					std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
					try
					{
						m_action = g_actionsMap.at(wszValue);
					}
					catch (const std::exception& e)
					{
						Logger::Write(LOGLEVEL_FAILED, L"The specified action is not valid. Valid options are 'addprofile', 'addprovider', 'addservice', 'changedatafilepath', 'cloneprofile', 'promotedelegates', 'listallprofiles', 'listallproviders', 'listallservices', 'listprofile', 'listprovider', 'listservice', 'promoteonedelegate', 'removeallprofiles', 'removeallproviders', 'removeallservices', 'removeprofile', 'removeprovider', 'removeservice', 'setcachedmode', 'setdefaultprofile', 'setdefaultservice', 'renameprofile', 'updateprovider', and 'updateservice'");
						return false;
					}
				}
				else
				{
					Logger::Write(LOGLEVEL_FAILED, L"You must specify a valid action. Valid options are 'addprofile', 'addprovider', 'addservice', 'changedatafilepath', 'cloneprofile', 'promotedelegates', 'listallprofiles', 'listallproviders', 'listallservices', 'listprofile', 'listprovider', 'listservice', 'promoteonedelegate', 'removeallprofiles', 'removeallproviders', 'removeallservices', 'removeprofile', 'removeprovider', 'removeservice', 'setcachedmode', 'setdefaultprofile', 'setdefaultservice', 'renameprofile', 'updateprovider', and 'updateservice'");
					return false;
				}
			}
			else if ((wsArg == L"-exportpath") || (wsArg == L"-ep"))
			{
				std::wstring wszExportPath = argv[i + 1];
				std::transform(wszExportPath.begin(), wszExportPath.end(), wszExportPath.begin(), ::tolower);
				m_wszExportPath = wszExportPath;
				i++;
			}
			else if ((wsArg == L"-exportmode") || (wsArg == L"-em"))
			{
				if (i + 1 < argc)
				{
					std::wstring wszValue = argv[i + 1];
					std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
					if (wszValue == L"export")
					{
						m_exportMode = EXPORTMODE_EXPORT;
						i++;

					}
					else if (wszValue == L"noexport")
					{
						m_exportMode = EXPORTMODE_NOEXPORT;
						i++;

					}
					else
					{
						Logger::Write(LOGLEVEL_FAILED, L"The specified value is not a valid export mode. Valid options are 'export' and 'noexport'");
						return false;
					}
				}
				else
				{
					Logger::Write(LOGLEVEL_FAILED, L"You must specify a valid export mode. Valid options are 'export' and 'noexport'");
					return false;
				}
			}
			else if ((wsArg == L"-profilemode") || (wsArg == L"-pm"))
			{
				if (i + 1 < argc)
				{
					std::wstring wszValue = argv[i + 1];
					std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
					try
					{
						m_action = g_profileModeMap.at(wszValue);
					}
					catch (const std::exception& e)
					{
						Logger::Write(LOGLEVEL_FAILED, L"The specified profile mode is not valid. Valid options are 'default', 'specific', and 'all'.\n");
						return false;
					}
				}
				else
				{
					Logger::Write(LOGLEVEL_FAILED, L"You must specify a valid profile mode. Valid options are 'default', 'specific', and 'all'.\n");
					return false;
				}
			}
			else if ((wsArg == L"-loggingMode") || (wsArg == L"-lm"))
			{
				if (i + 1 < argc)
				{
					std::wstring wszValue = argv[i + 1];
					std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
					if (wszValue == L"none")
					{
						m_loggingMode = LOGGINGMODE_NONE;
						i++;

					}
					else if (wszValue == L"console")
					{
						m_loggingMode = LOGGINGMODE_CONSOLE;
						i++;

					}
					else if (wszValue == L"file")
					{
						m_loggingMode = LOGGINGMODE_FILE;
						i++;
					}
					else if (wszValue == L"consoleandfile")
					{
						m_loggingMode = LOGGINGMODE_ALL;
						i++;
					}
					else
					{
						Logger::Write(LOGLEVEL_FAILED, L"The specified value is not a logging mode. Valid options are 'none', 'console', 'file', and 'consoleandfile'");
						return false;
					}
				}
				else
				{
					Logger::Write(LOGLEVEL_FAILED, L"You must specify a logging mode. Valid options are 'none', 'console', 'file', and 'consoleandfile'");
					return false;
				}
			}
		}

		// profile worker
		m_profileWorker = new ProfileWorker();

		for (int i = 1; i < argc; i++)
		{
			std::wstring wsArg = argv[i];
			std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

			if ((wsArg == L"-profilename") || (wsArg == L"-pn"))
			{
				if (i + 1 < argc)
				{

					m_profileWorker->profileName = argv[i + 1];
					m_profileMode = PROFILEMODE_SPECIFIC;
					i++;
				}
			}
		}

		// If a specific profile is needed then make sure a profile name was specified
		if (VCHK(m_profileMode, PROFILEMODE_SPECIFIC) && m_profileWorker->profileName.empty())
		{
			Logger::Write(LOGLEVEL_FAILED, L"You must either specify a profile name or pass 'default' for the value of thethe 'profilemode' parameter.");
			return false;
		}

		// create service worker
		for (int i = 1; i < argc; i++)
		{
			std::wstring wsArg = argv[i];
			std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

			if ((wsArg == L"-servicetype") || (wsArg == L"-st"))
			{
				if (i + 1 < argc)
				{
					std::wstring wszValue = argv[i + 1];
					std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
					try
					{
						m_action = g_serviceTypeMap.at(wszValue);
					}
					catch (const std::exception& e)
					{
						Logger::Write(LOGLEVEL_FAILED, L"The specified service type is not valid. Valid options are 'addressbook', 'datafile', 'exchangeaccount', and 'all'.\n");
						return false;
					}
				}
				else
				{
					Logger::Write(LOGLEVEL_FAILED, L"You must specify a valid service type. Valid options are 'addressbook', 'datafile', 'exchangeaccount', and 'all'.\n");
					return false;
				}
			}
		}



		if (m_serviceWorker)
		{
			// configure address book worker
			if (m_serviceWorker->m_serviceType == SERVICETYPE_ADDRESSBOOK)
			{
				for (int i = 1; i < argc; i++)
				{
					std::wstring wsArg = argv[i];
					std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

					if ((wsArg == L"-addressbookdisplayname") || (wsArg == L"-abdn"))
					{
						if (i + 1 < argc)
						{
							((AddressBookWorker*)m_serviceWorker)->wszABDisplayName = argv[i + 1];
							i++;
						}
					}
					else if ((wsArg == L"-addressbookservername") || (wsArg == L"-absn"))
					{
						if (i + 1 < argc)
						{
							((AddressBookWorker*)m_serviceWorker)->wszABServerName = argv[i + 1];
							i++;
						}
					}
					else if ((wsArg == L"-addressbookconfigfilepath") || (wsArg == L"-abcfp"))
					{
						if (i + 1 < argc)
						{
							((AddressBookWorker*)m_serviceWorker)->wszConfigFilePath = argv[i + 1];
							i++;
						}
					}
				}
			}
		}

		// Provider 
		for (int i = 1; i < argc; i++)
		{
			std::wstring wsArg = argv[i];
			std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

			if ((wsArg == L"-providertype") || (wsArg == L"-mt"))
			{
				if (i + 1 < argc)
				{
					std::wstring wszValue = argv[i + 1];
					std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
					if (wszValue == L"primarymailbox")
					{
						m_providerWorker = new PrimaryMailboxWorker();
						m_providerWorker->providerType = PROVIDERTYPE_PRIMARYMAILBOX;
						i++;
					}
					else if (wszValue == L"delegate")
					{
						m_providerWorker = new DelegateWorker();
						m_providerWorker->providerType = PROVIDERTYPE_DELEGATE;
						i++;
					}
					else if (wszValue == L"publicfolders")
					{
						m_providerWorker = new PublicFoldersWorker();
						m_providerWorker->providerType = PROVIDERTYPE_PUBLICFOLDERS;
						i++;
					}
					else
					{
						Logger::Write(LOGLEVEL_FAILED, L"The provider type specified is not valid. Valid entries are 'primarymailbox', 'delegate', and 'publicfolders'.");
						return false;
					}
				}
				else
				{
					Logger::Write(LOGLEVEL_FAILED, L"You must specify a provider type. Valid entries are 'primarymailbox', 'delegate', and 'publicfolders'.");
					return false;
				}
			}
		}

		return true;

	}
}