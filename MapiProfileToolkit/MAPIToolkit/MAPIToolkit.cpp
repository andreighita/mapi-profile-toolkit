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
#include "Toolkit.h"
#include "RegistryHelper.h"
#include <iterator> 
#include <algorithm>
#include "ServiceWorker.h"
#include "ExchangeAccountWorker.h"
#include "AddressBookWorker.h"
#include "DataFileWorker.h"
#include "Logger.h"
#include "MAPIToolkit.h"
#include "InlineAndMacros.h"
#include "Profile//Profile.h"
#include "PrimaryMailboxWorker.h"
#include "DelegateWorker.h"
#include "PublicFoldersWorker.h"


#pragma comment(lib, "Advapi32.lib")
#pragma comment(lib, "Mapi32.lib")
#pragma comment(lib, "Crypt32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shlwapi.lib")

#pragma warning(disable:4996) // _CRT_SECURE_NO_WARNINGS

namespace MAPIToolkit
{
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
	BOOL _cdecl IsCorrectBitness()
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

	void Run(int argc, wchar_t* argv[])
	{
		m_toolkit = new Toolkit();
		m_toolkit->action = ACTION_UNSPECIFIED;
		m_toolkit->iOutlookVersion = GetOutlookVersion();
		m_toolkit->loggingMode = LoggingMode::LoggingModeConsole;
		m_toolkit->m_serviceWorker = NULL;
		m_toolkit->m_providerWorker = NULL;

		if (ParseParams(argc, argv))
		{
			
		}
	}

	BOOL ParseParams(int argc, wchar_t* argv[])
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
					if (wszValue == L"addprofile")
					{
						m_toolkit->action = ACTION_PROFILE_ADD;
						i++;
					}
					else if (wszValue == L"cloneprofile")
					{
						m_toolkit->action = ACTION_PROFILE_CLONE;
						i++;
					}
					else if (wszValue == L"updateprofile")
					{
						m_toolkit->action = ACTION_PROFILE_UPDATE;
						i++;
					}
					else if (wszValue == L"listprofile")
					{
						m_toolkit->action = ACTION_PROFILE_LIST;
						i++;
					}
					else if (wszValue == L"listallprofiles")
					{
						m_toolkit->action = ACTION_PROFILE_LISTALL;
						i++;
					}
					else if (wszValue == L"removeprofile")
					{
						m_toolkit->action = ACTION_PROFILE_REMOVE;
						i++;
					}
					else if (wszValue == L"removeallprofiles")
					{
						m_toolkit->action = ACTION_PROFILE_REMOVEALL;
						i++;
					}
					else if (wszValue == L"setdefaultprofile")
					{
						m_toolkit->action = ACTION_PROFILE_SETDEFAULT;
						i++;
					}
					else if (wszValue == L"promotedelegates")
					{
						m_toolkit->action = ACTION_PROFILE_PROMOTEDELEGATES;
						i++;
					}
					else if (wszValue == L"promoteonedelegate")
					{
						m_toolkit->action = ACTION_PROFILE_PROMOTEONEDELEGATE;
						i++;
					}
					else if (wszValue == L"addprovider")
					{
						m_toolkit->action = ACTION_PROVIDER_ADD;
						i++;
					}
					else if (wszValue == L"updateprovider")
					{
						m_toolkit->action = ACTION_PROVIDER_UPDATE;
						i++;
					}
					else if (wszValue == L"listprovider")
					{
						m_toolkit->action = ACTION_PROVIDER_LIST;
						i++;
					}
					else if (wszValue == L"listallproviders")
					{
						m_toolkit->action = ACTION_PROVIDER_LISTALL;
						i++;
					}
					else if (wszValue == L"removeprovider")
					{
						m_toolkit->action = ACTION_PROVIDER_REMOVE;
						i++;
					}
					else if (wszValue == L"removeallproviders")
					{
						m_toolkit->action = ACTION_PROVIDER_REMOVEALL;
						i++;
					}
					else if (wszValue == L"addservice")
					{
						m_toolkit->action = ACTION_SERVICE_ADD;
						i++;
					}
					else if (wszValue == L"updateservice")
					{
						m_toolkit->action = ACTION_SERVICE_UPDATE;
						i++;
					}
					else if (wszValue == L"setcachedmode")
					{
						m_toolkit->action = ACTION_SERVICE_SETCACHEDMODE;
						i++;
					}
					else if (wszValue == L"listservice")
					{
						m_toolkit->action = ACTION_SERVICE_LIST;
						i++;
					}
					else if (wszValue == L"listallservices")
					{
						m_toolkit->action = ACTION_SERVICE_LISTALL;
						i++;
					}
					else if (wszValue == L"removeservice")
					{
						m_toolkit->action = ACTION_SERVICE_REMOVE;
						i++;
					}
					else if (wszValue == L"removeallservices")
					{
						m_toolkit->action = ACTION_SERVICE_REMOVEALL;
						i++;
					}
					else if (wszValue == L"changedatafilepath")
					{
						m_toolkit->action = ACTION_SERVICE_CHANGEDATAFILEPATH;
						i++;
					}
					else if (wszValue == L"setdefaultservice")
					{
						m_toolkit->action = ACTION_SERVICE_SETDEFAULT;
						i++;
					}
					else
					{
						Logger::Write(LogLevel::logLevelFailed, L"The specified action is not valid. Valid options are 'addprofile', 'addprovider', 'addservice', 'changedatafilepath', 'cloneprofile', 'promotedelegates', 'listallprofiles', 'listallproviders', 'listallservices', 'listprofile', 'listprovider', 'listservice', 'promoteonedelegate', 'removeallprofiles', 'removeallproviders', 'removeallservices', 'removeprofile', 'removeprovider', 'removeservice', 'setcachedmode', 'setdefaultprofile', 'setdefaultservice', 'updateprofile', 'updateprovider', and 'updateservice'");
						return false;
					}
				}
				else
				{
					Logger::Write(LogLevel::logLevelFailed, L"You must specify a valid export mode. Valid options are 'export' and 'noexport'");
					return false;
				}
			}
			else if ((wsArg == L"-exportpath") || (wsArg == L"-ep"))
			{
				std::wstring wszExportPath = argv[i + 1];
				std::transform(wszExportPath.begin(), wszExportPath.end(), wszExportPath.begin(), ::tolower);
				m_toolkit->wszExportPath = wszExportPath;
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
						m_toolkit->exportMode = ExportMode::Export;
						i++;

					}
					else if (wszValue == L"noexport")
					{
						m_toolkit->exportMode = ExportMode::NoExport;
						i++;

					}
					else
					{
						Logger::Write(LogLevel::logLevelFailed, L"The specified value is not a valid export mode. Valid options are 'export' and 'noexport'");
						return false;
					}
				}
				else
				{
					Logger::Write(LogLevel::logLevelFailed, L"You must specify a valid export mode. Valid options are 'export' and 'noexport'");
					return false;
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
						m_toolkit->m_profileWorker = new ProfileWorker();
						m_toolkit->profileMode = ProfileMode::Mode_Default;
						i++;

					}
					else if (wszValue == L"specific")
					{
						m_toolkit->m_profileWorker = new ProfileWorker();
						m_toolkit->profileMode = ProfileMode::Mode_Specific;
						i++;

					}
					else if (wszValue == L"all")
					{
						m_toolkit->profileCount = GetProfileCount();
						m_toolkit->m_profileWorker = new ProfileWorker[m_toolkit->profileCount];
						m_toolkit->profileMode = ProfileMode::Mode_All;
						i++;
					}
					else
					{
						Logger::Write(LogLevel::logLevelFailed, L"The specified value is not a valid profile mode. Valid options are 'default', 'specific', and 'all'");
						return false;
					}
				}
				else
				{
					Logger::Write(LogLevel::logLevelFailed, L"You must specify a valid profile mode. Valid options are 'default', 'specific', and 'all'");
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
						m_toolkit->loggingMode = LoggingMode::LoggingModeNone;
						i++;

					}
					else if (wszValue == L"console")
					{
						m_toolkit->loggingMode = LoggingMode::LoggingModeConsole;
						i++;

					}
					else if (wszValue == L"file")
					{
						m_toolkit->loggingMode = LoggingMode::LoggingModeFile;
						i++;
					}
					else if (wszValue == L"consoleandfile")
					{
						m_toolkit->loggingMode = LoggingMode::LoggingModeConsoleAndFile;
						i++;
					}
					else
					{
						Logger::Write(LogLevel::logLevelFailed, L"The specified value is not a logging mode. Valid options are 'none', 'console', 'file', and 'consoleandfile'");
						return false;
					}
				}
				else
				{
					Logger::Write(LogLevel::logLevelFailed, L"You must specify a logging mode. Valid options are 'none', 'console', 'file', and 'consoleandfile'");
					return false;
				}
			}
		}

		// profile worker
		m_toolkit->m_profileWorker = new ProfileWorker();

		for (int i = 1; i < argc; i++)
		{
			std::wstring wsArg = argv[i];
			std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

			if ((wsArg == L"-profilename") || (wsArg == L"-pn"))
			{
				if (i + 1 < argc)
				{

					m_toolkit->m_profileWorker->profileName = argv[i + 1];
					m_toolkit->profileMode = ProfileMode::Mode_Specific;
					i++;
				}
			}
		}

		// If a specific profile is needed then make sure a profile name was specified
		if (VCHK(m_toolkit->profileMode, ProfileMode::Mode_Specific) && m_toolkit->m_profileWorker->profileName.empty())
		{
			Logger::Write(LogLevel::logLevelFailed, L"You must either specify a profile name or pass 'default' for the value of thethe 'profilemode' parameter.");
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
					if (wszValue == L"mailbox")
					{
						m_toolkit->m_serviceWorker = new ExchangeAccountWorker();
						m_toolkit->m_serviceWorker->m_serviceType = ServiceType::ServiceType_Mailbox;
						i++;
					}
					else if (wszValue == L"pst")
					{
						m_toolkit->m_serviceWorker = new DataFileWorker();
						m_toolkit->m_serviceWorker->m_serviceType = ServiceType::ServiceType_Pst;
						i++;
					}
					else if (wszValue == L"addressbook")
					{
						m_toolkit->m_serviceWorker = new AddressBookWorker();
						m_toolkit->m_serviceWorker->m_serviceType = ServiceType::ServiceType_AddressBook;
						i++;
					}
					else
					{
						Logger::Write(LogLevel::logLevelFailed, L"The specified value is not a service type. Valid options are 'mailbox', 'pst', and 'addressbook'");
						return false;
					}
				}
				else
				{
					Logger::Write(LogLevel::logLevelFailed, L"You must specify a valid service type. Valid options are 'mailbox', 'pst', and 'addressbook'");
					return false;
				}
			}
		}

		// configure address book worker
		if (m_toolkit->m_serviceWorker->m_serviceType == ServiceType::ServiceType_AddressBook)
		{
			for (int i = 1; i < argc; i++)
			{
				std::wstring wsArg = argv[i];
				std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

				if ((wsArg == L"-addressbookdisplayname") || (wsArg == L"-abdn"))
				{
					if (i + 1 < argc)
					{
						((AddressBookWorker*)m_toolkit->m_serviceWorker)->wszABDisplayName = argv[i + 1];
						i++;
					}
				}
				else if ((wsArg == L"-addressbookservername") || (wsArg == L"-absn"))
				{
					if (i + 1 < argc)
					{
						((AddressBookWorker*)m_toolkit->m_serviceWorker)->wszABServerName = argv[i + 1];
						i++;
					}
				}
				else if ((wsArg == L"-addressbookconfigfilepath") || (wsArg == L"-abcfp"))
				{
					if (i + 1 < argc)
					{
						((AddressBookWorker*)m_toolkit->m_serviceWorker)->wszConfigFilePath = argv[i + 1];
						i++;
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
						m_toolkit->m_providerWorker = new PrimaryMailboxWorker();
						m_toolkit->m_providerWorker->providerType = ProviderType::ProviderType_PrimaryMailbox;
						i++;
					}
					else if (wszValue == L"delegate")
					{
						m_toolkit->m_providerWorker = new DelegateWorker();
						m_toolkit->m_providerWorker->providerType = ProviderType::ProviderType_Delegate;
						i++;
					}
					else if (wszValue == L"publicfolders")
					{
						m_toolkit->m_providerWorker = new PublicFoldersWorker();
						m_toolkit->m_providerWorker->providerType = ProviderType::ProviderType_PublicFolder;
						i++;
					}
					else
					{
						Logger::Write(LogLevel::logLevelFailed, L"The provider type specified is not valid. Valid entries are 'primarymailbox', 'delegate', and 'publicfolders'.");
						return false;
					}
				}
				else
				{
					Logger::Write(LogLevel::logLevelFailed, L"You must specify a provider type. Valid entries are 'primarymailbox', 'delegate', and 'publicfolders'.");
					return false;
				}
			}
		}

		return true;

	}




}