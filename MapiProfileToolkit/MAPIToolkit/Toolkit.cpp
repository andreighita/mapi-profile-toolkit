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
#include "Misc/Utility/StringOperations.h"
#include "Misc/XML/AddressBookXmlParser.h"
#include <vector>


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

	std::map<std::wstring, std::wstring> Toolkit::g_addressBookMap =
	{
		{ L"servicename",		L"EMABLT"},
		{ L"displayname",		L""},
		{ L"servername",		L"" },
		{ L"serverport",		L"389" },
		{ L"usessl",			L"false" },
		{ L"username",			L"" },
		{ L"password",			L"" },
		{ L"requirespa",		L"false" },
		{ L"searchtimeout",		L"60" },
		{ L"maxentries",		L"100" },
		{ L"defaultsearchbase",	L"true" },
		{ L"customsearchbase",	L"" },
		{ L"enablebrowsing",	L"false" },
		{ L"configfilepath",	L"" }
	};

	std::map<std::wstring, std::wstring> Toolkit::g_toolkitMap =
	{
		{ L"action",			L""},
		{ L"outlookversion",	L"" },
		{ L"loggingMode",		L"" },
		{ L"profileCount",		L"" },
		{ L"exportpath",		L"" },
		{ L"exportmode",		L"" },
		{ L"logfilepath",		L"" },
		{ L"profilemode",		L"default" },
		{ L"profilename",		L""},
		{ L"servicemode",		L"default" },
		{ L"servicetype",		L"" },
		{ L"providermode",		L"default" },
		{ L"providertype",		L"" },
		{ L"configfilepath",	L"" }
	};

	std::map<std::wstring, std::wstring> Toolkit::g_regKeyMap =
	{
		{ L"toolkit",				L"SOFTWARE\\Microsoft Ltd\\MAPIToolkit"},
		{ L"addressbook",			L"SOFTWARE\\Microsoft Ltd\\MAPIToolkit\\AddressBook" }
	};

	std::map <int, std::wstring > Toolkit::g_hexMap =
	{
		{ 0, L"0"},
		{ 1, L"1"},
		{ 2, L"2"},
		{ 3, L"3"},
		{ 4, L"4"},
		{ 5, L"5"},
		{ 6, L"6"},
		{ 7, L"7"},
		{ 8, L"8"},
		{ 9, L"9"},
		{ 10, L"A"},
		{ 11, L"B"},
		{ 12, L"C"},
		{ 13, L"D"},
		{ 14, L"E"},
		{ 15, L"F"}
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
	ULONG Toolkit::m_serviceType;
	BOOL Toolkit::m_registry = FALSE;

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
		std::wprintf(L"       Allows the management of Outlook / MAPI profiles at the command line.\n");
		std::wprintf(L"\n");
		std::wprintf(L"Usage: [-?] \n");
		std::wprintf(L"       [-action <addservice, listallservices, listservice, removeallservices, removeservice, updateservice>]\n");
		std::wprintf(L"       [-profilemode <default, specific, all>]\n");
		std::wprintf(L"       [-profilename name]\n");
		std::wprintf(L"       [-servicetype <addressbook>]\n");
		std::wprintf(L"       [-displayname name]\n");
		std::wprintf(L"       [-servername name]\n");
		std::wprintf(L"       [-serverport port]\n");
		std::wprintf(L"       [-usessl <true, false>]\n");
		std::wprintf(L"       [-username username]\n");
		std::wprintf(L"       [-password password]\n");
		std::wprintf(L"       [-requirespa <true, false>]\n");
		std::wprintf(L"       [-searchtimeout timeout]\n");
		std::wprintf(L"       [-maxentries maxentries]\n");
		std::wprintf(L"       [-defaultsearchbase <true, false>]\n");
		std::wprintf(L"       [-customsearchbase searchbase]\n");
		std::wprintf(L"       [-enablebrowsing <true, false>]\n");
		std::wprintf(L"       [-configfilepath path]\n");
		std::wprintf(L"       \n");
		std::wprintf(L"Options:\n");
		std::wprintf(L" -?:			                Displays this information.\n");
		std::wprintf(L" -action:                    \"addservice\" adds a service with the type specified by \"servicetype\".\n");
		std::wprintf(L"								\"listallservices\" lists all services with the type specified by \"servicetype\".\n");
		std::wprintf(L"                             \"listservice\" lists a specific service with the type specified by \"servicetype\".\n");
		std::wprintf(L"                             \"removeallservices\" removes all services with the type specified by \"servicetype\".\n");
		std::wprintf(L"                             \"removeservice\" removes a specific service with the type specified by \"servicetype\".\n");
		std::wprintf(L"                             \"updateservice\" updates a specific service with the type specified by \"servicetype\".\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -profilemode:               \"default\" to run the selected action on the default profile.\n");
		std::wprintf(L"                             \"specific\" to run the selected action on the profile specified by the \"profilename\" value.\n");
		std::wprintf(L"                             \"all\" to run the selected action against all profiles.\n");
		std::wprintf(L"                             The default profile will be used if a profile mode is not specified.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -profilename:               Name of the profile to run the specified actiona against. The profile name is mandatory\n");
		std::wprintf(L"                             if \"profilename\" is set to \"specific\" \n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -servicetype:               \"addressbook\" to run an addressbook specific operation.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -servicetype:               This is the only operation currently supported.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -displayname:               The display name of the address book service to create, update, list or remove.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -servername:                The display name of the LDAP server configured in the address book.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -configfilepath:            The path towards the address book configuration XML to use.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -serverport:                The LDAP port to connect to. The standard port for Active Directory is 389.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -usessl:                    \"true\" if a SSL connection is required. The default value is \"false\".\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -username:                  The Username to use for authenticating in the form of domain\\username, UPN or just the username \n");
		std::wprintf(L"                             if domain name not applicable or not required. Leave blank if a username and password are not required.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -password:                  The Password to use for authenticating. This must be a clear text passord. It will be encrypted via \n");
		std::wprintf(L"                             CryptoAPIand stored in the AB settings. \n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -requirespa:                \"true\" if Secure Password Authentication is required is required. The default value is \"false\".\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -searchtimeout:             The number of seconds before the search request times out. The default value is 60 seconds.]\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -maxentries:                The maximum number of results returned by a search in this AB. The default value is 100.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -defaultsearchbase:         \"true\" the default search base is to be used. The default value is \"true\". \n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -customsearchbase:          Custom search base in case DefaultSearchBase is set to False. \n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -enablebrowsing:            Indicates whether browsing the AB contens is supported.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -configfilepath:            The path towards the address book configuration XML to use.\n");
		std::wprintf(L"                             \n");
		std::wprintf(L" -?                          Displays this usage information.\n");
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
		HRESULT hRes = S_OK;
		m_action = ACTION_UNSPECIFIED;
		m_OutlookVersion = GetOutlookVersion();
		m_loggingMode = LOGGINGMODE_CONSOLE;
		m_serviceWorker = NULL;
		m_providerWorker = NULL;
		m_profileCount = 0;
		m_wszExportPath = L"";
		m_exportMode = EXPORTMODE_EXPORT; // 0 = no export; 1 = export;
		m_wszLogFilePath = L"";
		m_profileMode = PROFILEMODE_DEFAULT;
		m_pProfAdmin = NULL;

		MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
		CHK_HR_MSG(CoInitialize(NULL), L"Initialising the COM library on the current thread");
		CHK_HR_MSG(MAPIInitialize(&MAPIINIT), L"Initialising MAPI");
		CHK_HR_MSG(MAPIAdminProfiles(0, &m_pProfAdmin), L"Getting profile administration interface pointer.");
	Error:
		goto CleanUp;
	CleanUp:
		return TRUE;
	}

	VOID Toolkit::Uninitialise()
	{
		if (m_pProfAdmin) m_pProfAdmin->Release();
		MAPIUninitialize();
		CoUninitialize();
	}

	BOOL Toolkit::SaveConfig()
	{
		for (auto const& keyValuePair : g_toolkitMap)
		{
			if (!keyValuePair.second.empty())
				if (!WriteRegStringValue(HKEY_CURRENT_USER, (LPCTSTR)g_regKeyMap.at(L"toolkit").c_str(), (LPCTSTR)keyValuePair.first.c_str(), (LPCTSTR)keyValuePair.second.c_str())) return FALSE;
		}

		if (0 == wcscmp(g_toolkitMap.at(L"servicetype").c_str(), L"addressbook"))
		{
			for (auto const& keyValuePair : g_addressBookMap)
			{
				if (!keyValuePair.second.empty())
					if (!WriteRegStringValue(HKEY_CURRENT_USER, (LPCTSTR)g_regKeyMap.at(L"addressbook").c_str(), (LPCTSTR)keyValuePair.first.c_str(), (LPCTSTR)keyValuePair.second.c_str())) return FALSE;
			}
		}
		return TRUE;
	}

	BOOL Toolkit::ReadConfig()
	{
		ReadRegConfig(HKEY_CURRENT_USER, (LPCTSTR)g_regKeyMap.at(L"toolkit").c_str(), &g_toolkitMap);
		ReadRegConfig(HKEY_CURRENT_USER, (LPCTSTR)g_regKeyMap.at(L"addressbook").c_str(), &g_addressBookMap);
		return TRUE;
	}

	void Toolkit::Run(int argc, wchar_t* argv[])
	{
		Initialise();

		if (ParseParams(argc, argv))
		{
			RunAction();
		}
		else
			DisplayUsage();

		Uninitialise();
	}

	BOOL Toolkit::AddService(LPSERVICEADMIN2 pServiceAdmin)
	{
		switch (g_serviceTypeMap.at(g_toolkitMap.at(L"servicetype")))
		{
		case SERVICETYPE_ADDRESSBOOK:
		{
			if (AddAddressBookService(pServiceAdmin))
				Logger::Write(LOGLEVEL_SUCCESS, L"Address book service succesfully added");
			break;
		}
		case SERVICETYPE_EXCHANGEACCOUNT:
		{
			((ExchangeAccountWorker*)m_serviceWorker)->AddExchangeAccount();
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		case SERVICETYPE_DATAFILE:
		{
			((DataFileWorker*)m_serviceWorker)->AddDataFile();
			Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		}
		return 0;
	}

	BOOL Toolkit::AddAddressBookService(LPSERVICEADMIN2 pServiceAdmin)
	{
		HRESULT hRes = S_OK;
		if (!g_addressBookMap.at(L"servername").empty())
		{
			// if no display name specified, use the server name instead
			if (g_addressBookMap.at(L"displayname").empty())
			{
				g_addressBookMap.at(L"displayname") = g_addressBookMap.at(L"servername");
			}
			BOOL bServiceExists = false;
			CHK_HR_MSG(CheckABServiceExists(pServiceAdmin, (LPTSTR)g_addressBookMap.at(L"displayname").c_str(), (LPTSTR)g_addressBookMap.at(L"servicename").c_str(), 0, &bServiceExists), L"Cheking if service already exists");
			if (!bServiceExists)
				return (SUCCEEDED(CreateABService(pServiceAdmin)));
		}

	Error:
		goto CleanUp;
	CleanUp:
		return (SUCCEEDED(hRes));

	}

	void Toolkit::UpdateAddressBookService()
	{
	}

	void Toolkit::ListAddressBookService()
	{
	}

	void Toolkit::ListAllAddressBookServices()
	{
	}

	void Toolkit::RemoveAddressBookService()
	{
	}

	void Toolkit::RemoveAllAddressBookServices()
	{
	}

	void Toolkit::RunAction()
	{
		HRESULT hRes = S_OK;
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
		default:
		{
			std::vector<std::wstring> vProfileNames;
			switch (g_profileModeMap.at(g_toolkitMap.at(L"profilemode")))
			{
			case PROFILEMODE_ALL:
				CHK_BOOL_MSG(GetProfileNames(m_pProfAdmin, &vProfileNames), L"Retrieving profile names");
				for (auto const& profileName : vProfileNames) {
					if (RunActionOneProfile(profileName))
					{
						Logger::Write(LOGLEVEL_SUCCESS, L"Action succesfully run on profile: " + profileName);
					}
					else
					{
						Logger::Write(LOGLEVEL_FAILED, L"Unable to run action on profile: " + profileName);
					}
				}
				break;
			case PROFILEMODE_SPECIFIC:
				if (RunActionOneProfile(g_toolkitMap.at(L"profilename")))
				{
					Logger::Write(LOGLEVEL_SUCCESS, L"Action succesfully run on profile: " + g_toolkitMap.at(L"profilename"));
				}
				else
				{
					Logger::Write(LOGLEVEL_FAILED, L"Unable to run action on profile: " + g_toolkitMap.at(L"profilename"));
				}
				break;
			}
			break;
		}

		}

	Error:
		goto CleanUp;
	CleanUp:
		return;
	}

	BOOL Toolkit::RunActionOneProfile(std::wstring wszProfileName)
	{
		HRESULT hRes = NULL;
		LPSERVICEADMIN pServiceAdmin = NULL;
		LPSERVICEADMIN2 pServiceAdmin2 = NULL;
		// Retrieves pointers to the supported interfaces on an object.

		CHK_HR_MSG(m_pProfAdmin->AdminServices((LPTSTR)wszProfileName.c_str(), NULL, NULL, MAPI_UNICODE, (LPSERVICEADMIN*)&pServiceAdmin), L"Getting admin services interface pointer for profile " + wszProfileName);
		CHK_HR_MSG(pServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)& pServiceAdmin2), L"Getting LPSERVICEADMIN2 interface pointer for profile ");
		
		switch (m_action)
		{

		case ACTION_SERVICE_ADD:
		{
			if (AddService(pServiceAdmin2))
			{

			}
			break;
		}
		case ACTION_SERVICE_LISTALL:
		{
			if (m_serviceWorker)
			{
				if (SERVICETYPE_ADDRESSBOOK == m_serviceType)
				{
					((AddressBookWorker*)m_serviceWorker)->ListAllAddressBookServices();
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
				if (SERVICETYPE_ADDRESSBOOK == m_serviceType)
				{
					((AddressBookWorker*)m_serviceWorker)->RemoveAllAddressBookServices();
				}
			}
			else
				Logger::Write(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}

		}

	Error:
		goto CleanUp;
	CleanUp:
		return SUCCEEDED(hRes);
	}

	BOOL Toolkit::RunActionOneService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID pMapiUid)
	{
		HRESULT hRes = NULL;
		LPPROVIDERADMIN pProviderAdmin = NULL;
		std::wstring wstr(reinterpret_cast<wchar_t*>(pMapiUid), 16 / sizeof(wchar_t));
		CHK_HR_MSG(pServiceAdmin->AdminProviders(pMapiUid, NULL, &pProviderAdmin), L"Getting profider admin interface pointer for service with UID: " + MapiUidToString(pMapiUid));

		switch (m_action)
		{

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
				if (SERVICETYPE_ADDRESSBOOK == m_serviceType)
				{
					((AddressBookWorker*)m_serviceWorker)->ListAddressBookService();
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
				if (SERVICETYPE_ADDRESSBOOK == m_serviceType)
				{
					((AddressBookWorker*)m_serviceWorker)->RemoveAddressBookService();
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

	Error:
		goto CleanUp;
	CleanUp:
		return SUCCEEDED(hRes);
	}

	BOOL Toolkit::ParseParams(int argc, wchar_t* argv[])
	{
		HRESULT hRes = S_OK;
		// parse the actions
		std::wstring wszActionItem;
		std::wstringstream wss(g_toolkitMap.at(L"action"));
		// check if we're supposed to list the help menu
		for (int i = 1; i < argc; i++)
		{
			std::wstring wsArg = argv[i];
			std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

			if (wsArg == L"-?")
			{
				return false;
			}
		}

		// check if we're supposed to read the configuration from the registry
		for (int i = 1; i < argc; i++)
		{
			std::wstring wsArg = argv[i];
			std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

			if (wsArg == L"-registry")
			{
				ReadConfig();
			}
		}

		// general toolkit
		for (int i = 1; i < argc; i++)
		{
			switch (argv[i][0])
			{
			case '-':
				std::wstring wsArg = SubstringFromStart(1, argv[i]);
				std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

				try
				{
					if (i + 1 < argc)
					{
						g_toolkitMap.at(wsArg) = argv[i + 1];
						i++;
					};
				}
				catch (const std::exception& e)
				{

				}
				break;
			}
		}

		// profile worker
		m_profileWorker = new ProfileWorker();
		m_profileWorker->profileName = GetDefaultProfileName();
		m_profileCount = GetProfileCount();

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
		if (VCHK(g_profileModeMap.at(g_toolkitMap.at(L"profilemode")), PROFILEMODE_SPECIFIC) && g_toolkitMap.at(L"profilename").c_str())
		{
			Logger::Write(LOGLEVEL_FAILED, L"You must either specify a profile name or pass 'default' for the value of thethe 'profilemode' parameter.");
			return false;
		}
		else if (VCHK(g_profileModeMap.at(g_toolkitMap.at(L"profilemode")), PROFILEMODE_DEFAULT))
		{
			g_toolkitMap.at(L"profilename") == GetDefaultProfileName();
			g_toolkitMap.at(L"profilemode") == L"specific";
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
						m_serviceType = g_serviceTypeMap.at(wszValue);

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



		// configure address book worker
		if (m_serviceType == SERVICETYPE_ADDRESSBOOK)
		{
			m_serviceWorker = new AddressBookWorker();

			for (int i = 1; i < argc; i++)
			{
				switch (argv[i][0])
				{
				case '-':
				case '/':
				case '\\':
					std::wstring wsArg = SubstringFromStart(1, argv[i]);
					std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

					try
					{
						if (i + 1 < argc)
						{
							g_addressBookMap.at(wsArg) = argv[i + 1];
							i++;
						};
					}
					catch (const std::exception& e)
					{

					}
					break;
				}
			}

			CHK_HR_MSG(ParseAddressBookXml((LPTSTR)g_addressBookMap.at(L"configfilepath").c_str()), L"Parsing configuration XML file");
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


		while (std::getline(wss, wszActionItem, L'|'))
			m_action |= g_actionsMap.at(wszActionItem);

		SaveConfig();
	

	Error:
		goto CleanUp;
	CleanUp:
		return SUCCEEDED(hRes);
	}
}