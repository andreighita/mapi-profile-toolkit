// MAPIToolkit.cpp : Defines the functions for the static library.
//

#include "pch.h"
#include "framework.h"
#include <atlchecked.h>
#include "Toolkit.h"
#include "RegistryHelper.h"
#include <iterator> 
#include <map> 
#include <algorithm>
#include "ServiceWorker.h"
#include "ExchangeAccountWorker.h"
#include "AddressBookWorker.h"
#include "DataFileWorker.h"
#include "Logger.h"

namespace MAPIToolkit
{
	BOOL ParseParams(_TCHAR* argv[], int argc, Toolkit* pToolkit);
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

	static Toolkit * m_toolkit;

	void Initialise(int argc, _TCHAR* argv[])
	{
		m_toolkit = new Toolkit();
		m_toolkit->action = ACTION_UNSPECIFIED;
		m_toolkit->iOutlookVersion = GetOutlookVersion();
		m_toolkit->loggingMode = LoggingMode::LoggingModeConsole;
		if (ParseParams(argv, argc, m_toolkit))
		{

		}
	}

	BOOL ParseParams(_TCHAR* argv[], int argc, Toolkit* pToolkit)
	{
		// general toolkit
		for (int i = 1; i < argc; i++)
		{
			std::wstring wsArg = argv[i];
			std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

			if ((wsArg == L"-exportpath") || (wsArg == L"-ep"))
			{
				std::wstring wszExportPath = argv[i + 1];
				std::transform(wszExportPath.begin(), wszExportPath.end(), wszExportPath.begin(), ::tolower);
				pToolkit->wszExportPath = wszExportPath;
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
						pToolkit->exportMode = ExportMode::Export;
						i++;

					}
					else if (wszValue == L"noexport")
					{
						pToolkit->exportMode = ExportMode::NoExport;
						i++;

					}
					else
					{
						Logger::Write(LogLevel::logLevelFailed, L"The specified value is not a valid export mode. Valid options are 'export' and 'noexport'");
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
						pToolkit->profileMode = ProfileMode::Mode_Default;
						i++;

					}
					else if (wszValue == L"specific")
					{
						pToolkit->profileMode = ProfileMode::Mode_Specific;
						i++;

					}
					else if (wszValue == L"all")
					{
						pToolkit->profileMode = ProfileMode::Mode_All;
						i++;

					}
					else
					{
						Logger::Write(LogLevel::logLevelFailed, L"The specified value is not a valid profile mode. Valid options are 'default', 'specific', and 'all'");
						return false;
					}
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
						pToolkit->loggingMode = LoggingMode::LoggingModeNone;
						i++;

					}
					else if (wszValue == L"console")
					{
						pToolkit->loggingMode = LoggingMode::LoggingModeConsole;
						i++;

					}
					else if (wszValue == L"file")
					{
						pToolkit->loggingMode = LoggingMode::LoggingModeFile;
						i++;
					}
					else if (wszValue == L"consoleandfile")
					{
						pToolkit->loggingMode = LoggingMode::LoggingModeConsoleAndFile;
						i++;
					}
					else
					{
						Logger::Write(LogLevel::logLevelFailed, L"The specified value is not a logging mode. Valid options are 'none', 'console', 'file', and 'consoleandfile'");
						return false;
					}
				}
			}
			else return false;
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
						pToolkit->m_serviceWorker = new ExchangeAccountWorker();
					pToolkit->m_serviceWorker->m_serviceType = ServiceType::ServiceType_Mailbox;
						i++;
					}
					else if (wszValue == L"pst")
					{
						pToolkit->m_serviceWorker = new DataFileWorker();
						pToolkit->m_serviceWorker->m_serviceType = ServiceType::ServiceType_Pst;
						i++;
					}
					else if (wszValue == L"addressbook")
					{
						pToolkit->m_serviceWorker = new AddressBookWorker();
						pToolkit->m_serviceWorker->m_serviceType = ServiceType::ServiceType_AddressBook;
						i++;
					}
					else
					{
						pToolkit->m_serviceWorker = NULL;
					}
				}
			}
		}

		// profile worker
		pToolkit->m_profileWorker = new ProfileWorker();

		for (int i = 1; i < argc; i++)
		{
			std::wstring wsArg = argv[i];
			std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

			if ((wsArg == L"-profilename") || (wsArg == L"-pn"))
			{
				if (i + 1 < argc)
				{
					pToolkit->m_profileWorker->profileName = argv[i + 1];
					i++;
				}
			}
			else return false;
		}
	}

	//void ParamsArrayToMap(_TCHAR* argv[], int argc, Toolkit * pToolkit)
	//{
	//	 empty map container 
	//	std::map<std::wstring, std::wstring>parameterMap;

	//	for (int i = 1; i < argc; i++)
	//	{
	//		std::wstring wsArg = argv[i];
	//		std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

	//		if ((wsArg == L"-exportpath") || (wsArg == L"-ep"))
	//		{
	//			std::wstring wszExportPath = argv[i + 1];
	//			std::transform(wszExportPath.begin(), wszExportPath.end(), wszExportPath.begin(), ::tolower);
	//			pToolkit->wszExportPath = wszExportPath;
	//			i++;
	//		}
	//		else if ((wsArg == L"-addressbookdisplayname") || (wsArg == L"-abdn"))
	//		{
	//			if (i + 1 < argc)
	//			{

	//				parameterMap.insert(std::pair<std::wstring, std::wstring>(L"addressbookdisplayname", argv[i + 1]));
	//				i++;
	//			}
	//		}
	//		else if ((wsArg == L"-addressbookservername") || (wsArg == L"-absn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				parameterMap.insert(std::pair<std::wstring, std::wstring>(L"addressbookservername", argv[i + 1]));
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-addressbookconfigfilepath") || (wsArg == L"-abcfp"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-profile") || (wsArg == L"-p"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"add")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_ADD;
	//					i++;
	//				}
	//				else if (wszValue == L"update")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_UPDATE;
	//					i++;
	//				}
	//				else if (wszValue == L"remove")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_REMOVE;
	//					i++;
	//				}
	//				else if (wszValue == L"removeall")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_REMOVEALL;
	//					i++;
	//				}
	//				else if (wszValue == L"list")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_LIST;
	//					i++;
	//				}
	//				else if (wszValue == L"listall")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_LISTALL;
	//					i++;
	//				}
	//				else if (wszValue == L"clone")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_CLONE;
	//					i++;
	//				}
	//				else if (wszValue == L"promotedelegates")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_PROMOTEDELEGATES;
	//					i++;
	//				}
	//				else if (wszValue == L"setdefault")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_SETDEFAULT;
	//					i++;
	//				}
	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-profilemode") || (wsArg == L"-pm"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"default")
	//				{
	//					pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Default;
	//					i++;

	//				}
	//				else if (wszValue == L"specific")
	//				{
	//					pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Specific;
	//					i++;

	//				}
	//				else if (wszValue == L"all")
	//				{
	//					pRunOpts->profileOptions->profileMode = ProfileMode::Mode_All;
	//					i++;

	//				}
	//				else return false;
	//			}
	//		}
	//		else if ((wsArg == L"-profilename") || (wsArg == L"-pn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->wszProfileName = argv[i + 1];
	//				pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Specific;
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-setdefaultprofile") || (wsArg == L"-sdp"))
	//		{
	//			pRunOpts->profileOptions->bSetDefaultProfile = true;

	//		}
	//		else if ((wsArg == L"-service") || (wsArg == L"-s"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"add")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_ADD;
	//					i++;
	//				}
	//				else if (wszValue == L"update")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_UPDATE;
	//					i++;
	//				}
	//				else if (wszValue == L"remove")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_REMOVE;
	//					i++;
	//				}
	//				else if (wszValue == L"removeall")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_REMOVEALL;
	//					i++;
	//				}
	//				else if (wszValue == L"list")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_LIST;
	//					i++;
	//				}
	//				else if (wszValue == L"listall")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_LISTALL;
	//					i++;
	//				}
	//				else if (wszValue == L"setcachedmode")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_SETCACHEDMODE;
	//					i++;
	//				}
	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-servicetype") || (wsArg == L"-st"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"mailbox")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceType = ServiceType::ServiceType_Mailbox;
	//					i++;
	//				}
	//				else if (wszValue == L"pst")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceType = ServiceType::ServiceType_Pst;
	//					i++;
	//				}
	//				else if (wszValue == L"addressbook")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceType = ServiceType::ServiceType_AddressBook;
	//					i++;
	//				}
	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-servicemode") || (wsArg == L"-sm"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"default")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_Default;
	//					i++;

	//				}
	//				else if (wszValue == L"specific")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_Specific;
	//					i++;

	//				}
	//				else if (wszValue == L"all")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_All;
	//					i++;

	//				}
	//				else return false;
	//			}
	//		}
	//		else if ((wsArg == L"-mailbox") || (wsArg == L"-m"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"add")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_ADD;
	//					i++;
	//				}
	//				else if (wszValue == L"update")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_UPDATE;
	//					i++;
	//				}
	//				else if (wszValue == L"remove")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_REMOVE;
	//					i++;
	//				}
	//				else if (wszValue == L"removeall")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_REMOVEALL;
	//					i++;
	//				}
	//				else if (wszValue == L"list")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_LIST;
	//					i++;
	//				}
	//				else if (wszValue == L"listall")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_LIST;
	//					i++;
	//				}

	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-mailboxtype") || (wsArg == L"-mt"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"primary")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->providerOptions->providerType = ProviderType::PrimaryMailbox;
	//					i++;
	//				}
	//				else if (wszValue == L"delegate")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->providerOptions->providerType = ProviderType::Delegate;
	//					i++;
	//				}
	//				else if (wszValue == L"publicfolder")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->providerOptions->providerType = ProviderType::PublicFolder;
	//					i++;
	//				}
	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-setdefaultservice") || (wsArg == L"-sds"))
	//		{
	//			pRunOpts->profileOptions->serviceOptions->bSetDefaultservice = true;

	//		}
	//		else if ((wsArg == L"-cachedmodemonths") || (wsArg == L"-cmm"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->iCachedModeMonths = _wtoi(argv[i + 1]);
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-serviceindex") || (wsArg == L"-si"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->iServiceIndex = _wtoi(argv[i + 1]);
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-abexternalurl") || (wsArg == L"-abeu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszAddressBookExternalUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-abinternalurl") || (wsArg == L"-abiu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszAddressBookInternalUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-autodiscoverurl") || (wsArg == L"-au"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszAutodiscoverUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-mailboxdisplayname") || (wsArg == L"-mdn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszMailboxDisplayName = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailboxDisplayName = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-mailboxlegacydn") || (wsArg == L"-mldn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszMailboxLegacyDN = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailboxLegacyDN = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-mailstoreexternalurl") || (wsArg == L"-mseu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszMailStoreExternalUrl = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailStoreExternalUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-mailstoreinternalurl") || (wsArg == L"-msiu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszMailStoreInternalUrl = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailStoreExternalUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-rohproxyserver") || (wsArg == L"-rps"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszRohProxyServer = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszRohProxyServer = argv[i + 1];
	//				i++;
	//			}
	//		}
	//		else if ((wsArg == L"-rohproxyserverflags") || (wsArg == L"-rpsf"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->ulRohProxyServerFlags = _wtoi(argv[i + 1]);
	//				i++;
	//			}
	//		}
	//		else if ((wsArg == L"-rohproxyserverauthpackage") || (wsArg == L"-mrpsap"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->ulRohProxyServerAuthPackage = _wtoi(argv[i + 1]);
	//				i++;
	//			}
	//		}
	//		else if ((wsArg == L"-serverdisplayname") || (wsArg == L"-sdn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszServerDisplayName = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszServerDisplayName = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-serverlegacydn") || (wsArg == L"-sldn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszServerLegacyDN = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszServerLegacyDN = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-smtpaddress") || (wsArg == L"-sa"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszSmtpAddress = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszSmtpAddress = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-unresolvedserver") || (wsArg == L"-us"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszUnresolvedServer = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-unresolveduser") || (wsArg == L"-uu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszUnresolvedUser = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-cachedmodeowner") || (wsArg == L"-cmo"))
	//		{
	//			pRunOpts->profileOptions->serviceOptions->cachedModeOwner = CachedMode::Enabled;

	//		}
	//		else if ((wsArg == L"-cachedmodepublicfolder") || (wsArg == L"-cmpf"))
	//		{
	//			pRunOpts->profileOptions->serviceOptions->cachedModePublicFolders = CachedMode::Enabled;

	//		}
	//		else if ((wsArg == L"-cachedmodeshared") || (wsArg == L"-cms"))
	//		{
	//			pRunOpts->profileOptions->serviceOptions->cachedModeShared = CachedMode::Enabled;

	//		}
	//		else if ((wsArg == L"-configflags") || (wsArg == L"-cf"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->ulConfigFlags = _wtol(argv[i + 1]);
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-connectmode") || (wsArg == L"-cm"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"roh")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->connectMode = ConnectMode::ConnectMode_RpcOverHttp;
	//					i++;

	//				}
	//				if (wszValue == L"moh")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->connectMode = ConnectMode::ConnectMode_MapiOverHttp;
	//					i++;

	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-resourceflags") || (wsArg == L"-rf"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->ulResourceFlags = _wtol(argv[i + 1]);
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-cachedmodeowner") || (wsArg == L"-cmo"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"enable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModeOwner = CachedMode::Enabled;
	//					i++;

	//				}
	//				if (wszValue == L"disable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModeOwner = CachedMode::Disabled;
	//					i++;

	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-cachedmodeshared") || (wsArg == L"-cms"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"enable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModeShared = CachedMode::Enabled;
	//					i++;

	//				}
	//				if (wszValue == L"disable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModeShared = CachedMode::Disabled;
	//					i++;

	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-cachedmodepublicfolders") || (wsArg == L"-cmpf"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"enable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModePublicFolders = CachedMode::Enabled;
	//					i++;

	//				}
	//				if (wszValue == L"disable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModePublicFolders = CachedMode::Disabled;
	//					i++;

	//				}
	//			}
	//		}
	//		else return false;
	//	}
	//}

	//BOOL ValidateScenario(int argc, _TCHAR* argv[], RuntimeOptions* pRunOpts)
	//{
	//	std::vector<std::string> wszDiscardedArgs;
	//	if (!pRunOpts) return FALSE;
	//	ZeroMemory(pRunOpts, sizeof(RuntimeOptions));
	//	pRunOpts->action = ACTION_UNSPECIFIED;
	//	int iThreeParam = 0;
	//	pRunOpts->iOutlookVersion = GetOutlookVersion();
	//	pRunOpts->loggingMode = LoggingMode::LoggingModeConsole;

	//	pRunOpts->profileOptions = new ProfileOptions();
	//	pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Default;
	//	pRunOpts->profileOptions->serviceOptions = new ServiceOptions();
	//	pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_Default;
	//	pRunOpts->profileOptions->serviceOptions->connectMode = ConnectMode::ConnectMode_RpcOverHttp;
	//	pRunOpts->profileOptions->serviceOptions->providerOptions = new ProviderOptions();
	//	pRunOpts->profileOptions->serviceOptions->addressBookOptions = new AddressBookOptions();

	//	for (int i = 1; i < argc; i++)
	//	{
	//		std::wstring wsArg = argv[i];
	//		std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

	//		if ((wsArg == L"-exportpath") || (wsArg == L"-ep"))
	//		{
	//			std::wstring wszExportPath = argv[i + 1];
	//			std::transform(wszExportPath.begin(), wszExportPath.end(), wszExportPath.begin(), ::tolower);
	//			pRunOpts->wszExportPath = wszExportPath;
	//			pRunOpts->exportMode = ExportMode::Export;
	//			i++;
	//		}
	//		else if ((wsArg == L"-addressbookdisplayname") || (wsArg == L"-abdn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszABDisplayName = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-addressbookservername") || (wsArg == L"-absn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszABServerName = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-addressbookconfigfilepath") || (wsArg == L"-abcfp"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-profile") || (wsArg == L"-p"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"add")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_ADD;
	//					i++;
	//				}
	//				else if (wszValue == L"update")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_UPDATE;
	//					i++;
	//				}
	//				else if (wszValue == L"remove")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_REMOVE;
	//					i++;
	//				}
	//				else if (wszValue == L"removeall")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_REMOVEALL;
	//					i++;
	//				}
	//				else if (wszValue == L"list")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_LIST;
	//					i++;
	//				}
	//				else if (wszValue == L"listall")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_LISTALL;
	//					i++;
	//				}
	//				else if (wszValue == L"clone")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_CLONE;
	//					i++;
	//				}
	//				else if (wszValue == L"promotedelegates")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_PROMOTEDELEGATES;
	//					i++;
	//				}
	//				else if (wszValue == L"setdefault")
	//				{
	//					pRunOpts->action |= ACTION_PROFILE_SETDEFAULT;
	//					i++;
	//				}
	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-profilemode") || (wsArg == L"-pm"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"default")
	//				{
	//					pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Default;
	//					i++;

	//				}
	//				else if (wszValue == L"specific")
	//				{
	//					pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Specific;
	//					i++;

	//				}
	//				else if (wszValue == L"all")
	//				{
	//					pRunOpts->profileOptions->profileMode = ProfileMode::Mode_All;
	//					i++;

	//				}
	//				else return false;
	//			}
	//		}
	//		else if ((wsArg == L"-profilename") || (wsArg == L"-pn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->wszProfileName = argv[i + 1];
	//				pRunOpts->profileOptions->profileMode = ProfileMode::Mode_Specific;
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-setdefaultprofile") || (wsArg == L"-sdp"))
	//		{
	//			pRunOpts->profileOptions->bSetDefaultProfile = true;

	//		}
	//		else if ((wsArg == L"-service") || (wsArg == L"-s"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"add")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_ADD;
	//					i++;
	//				}
	//				else if (wszValue == L"update")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_UPDATE;
	//					i++;
	//				}
	//				else if (wszValue == L"remove")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_REMOVE;
	//					i++;
	//				}
	//				else if (wszValue == L"removeall")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_REMOVEALL;
	//					i++;
	//				}
	//				else if (wszValue == L"list")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_LIST;
	//					i++;
	//				}
	//				else if (wszValue == L"listall")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_LISTALL;
	//					i++;
	//				}
	//				else if (wszValue == L"setcachedmode")
	//				{
	//					pRunOpts->action |= ACTION_SERVICE_SETCACHEDMODE;
	//					i++;
	//				}
	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-servicetype") || (wsArg == L"-st"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"mailbox")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceType = ServiceType::ServiceType_Mailbox;
	//					i++;
	//				}
	//				else if (wszValue == L"pst")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceType = ServiceType::ServiceType_Pst;
	//					i++;
	//				}
	//				else if (wszValue == L"addressbook")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceType = ServiceType::ServiceType_AddressBook;
	//					i++;
	//				}
	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-servicemode") || (wsArg == L"-sm"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"default")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_Default;
	//					i++;

	//				}
	//				else if (wszValue == L"specific")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_Specific;
	//					i++;

	//				}
	//				else if (wszValue == L"all")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->serviceMode = ServiceMode::Mode_All;
	//					i++;

	//				}
	//				else return false;
	//			}
	//		}
	//		else if ((wsArg == L"-mailbox") || (wsArg == L"-m"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"add")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_ADD;
	//					i++;
	//				}
	//				else if (wszValue == L"update")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_UPDATE;
	//					i++;
	//				}
	//				else if (wszValue == L"remove")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_REMOVE;
	//					i++;
	//				}
	//				else if (wszValue == L"removeall")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_REMOVEALL;
	//					i++;
	//				}
	//				else if (wszValue == L"list")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_LIST;
	//					i++;
	//				}
	//				else if (wszValue == L"listall")
	//				{
	//					pRunOpts->action |= ACTION_PROVIDER_LIST;
	//					i++;
	//				}

	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-mailboxtype") || (wsArg == L"-mt"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"primary")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->providerOptions->providerType = ProviderType::PrimaryMailbox;
	//					i++;
	//				}
	//				else if (wszValue == L"delegate")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->providerOptions->providerType = ProviderType::Delegate;
	//					i++;
	//				}
	//				else if (wszValue == L"publicfolder")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->providerOptions->providerType = ProviderType::PublicFolder;
	//					i++;
	//				}
	//				else
	//				{
	//					return false;
	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-setdefaultservice") || (wsArg == L"-sds"))
	//		{
	//			pRunOpts->profileOptions->serviceOptions->bSetDefaultservice = true;

	//		}
	//		else if ((wsArg == L"-cachedmodemonths") || (wsArg == L"-cmm"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->iCachedModeMonths = _wtoi(argv[i + 1]);
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-serviceindex") || (wsArg == L"-si"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->iServiceIndex = _wtoi(argv[i + 1]);
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-abexternalurl") || (wsArg == L"-abeu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszAddressBookExternalUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-abinternalurl") || (wsArg == L"-abiu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszAddressBookInternalUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-autodiscoverurl") || (wsArg == L"-au"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszAutodiscoverUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-mailboxdisplayname") || (wsArg == L"-mdn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszMailboxDisplayName = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailboxDisplayName = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-mailboxlegacydn") || (wsArg == L"-mldn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszMailboxLegacyDN = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailboxLegacyDN = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-mailstoreexternalurl") || (wsArg == L"-mseu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszMailStoreExternalUrl = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailStoreExternalUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-mailstoreinternalurl") || (wsArg == L"-msiu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszMailStoreInternalUrl = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszMailStoreExternalUrl = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-rohproxyserver") || (wsArg == L"-rps"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszRohProxyServer = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszRohProxyServer = argv[i + 1];
	//				i++;
	//			}
	//		}
	//		else if ((wsArg == L"-rohproxyserverflags") || (wsArg == L"-rpsf"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->ulRohProxyServerFlags = _wtoi(argv[i + 1]);
	//				i++;
	//			}
	//		}
	//		else if ((wsArg == L"-rohproxyserverauthpackage") || (wsArg == L"-mrpsap"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->ulRohProxyServerAuthPackage = _wtoi(argv[i + 1]);
	//				i++;
	//			}
	//		}
	//		else if ((wsArg == L"-serverdisplayname") || (wsArg == L"-sdn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszServerDisplayName = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszServerDisplayName = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-serverlegacydn") || (wsArg == L"-sldn"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszServerLegacyDN = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszServerLegacyDN = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-smtpaddress") || (wsArg == L"-sa"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszSmtpAddress = argv[i + 1];
	//				pRunOpts->profileOptions->serviceOptions->providerOptions->wszSmtpAddress = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-unresolvedserver") || (wsArg == L"-us"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszUnresolvedServer = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-unresolveduser") || (wsArg == L"-uu"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->wszUnresolvedUser = argv[i + 1];
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-cachedmodeowner") || (wsArg == L"-cmo"))
	//		{
	//			pRunOpts->profileOptions->serviceOptions->cachedModeOwner = CachedMode::Enabled;

	//		}
	//		else if ((wsArg == L"-cachedmodepublicfolder") || (wsArg == L"-cmpf"))
	//		{
	//			pRunOpts->profileOptions->serviceOptions->cachedModePublicFolders = CachedMode::Enabled;

	//		}
	//		else if ((wsArg == L"-cachedmodeshared") || (wsArg == L"-cms"))
	//		{
	//			pRunOpts->profileOptions->serviceOptions->cachedModeShared = CachedMode::Enabled;

	//		}
	//		else if ((wsArg == L"-configflags") || (wsArg == L"-cf"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->ulConfigFlags = _wtol(argv[i + 1]);
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-connectmode") || (wsArg == L"-cm"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"roh")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->connectMode = ConnectMode::ConnectMode_RpcOverHttp;
	//					i++;

	//				}
	//				if (wszValue == L"moh")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->connectMode = ConnectMode::ConnectMode_MapiOverHttp;
	//					i++;

	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-resourceflags") || (wsArg == L"-rf"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				pRunOpts->profileOptions->serviceOptions->ulResourceFlags = _wtol(argv[i + 1]);
	//				i++;

	//			}
	//		}
	//		else if ((wsArg == L"-cachedmodeowner") || (wsArg == L"-cmo"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"enable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModeOwner = CachedMode::Enabled;
	//					i++;

	//				}
	//				if (wszValue == L"disable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModeOwner = CachedMode::Disabled;
	//					i++;

	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-cachedmodeshared") || (wsArg == L"-cms"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"enable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModeShared = CachedMode::Enabled;
	//					i++;

	//				}
	//				if (wszValue == L"disable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModeShared = CachedMode::Disabled;
	//					i++;

	//				}
	//			}
	//		}
	//		else if ((wsArg == L"-cachedmodepublicfolders") || (wsArg == L"-cmpf"))
	//		{
	//			if (i + 1 < argc)
	//			{
	//				std::wstring wszValue = argv[i + 1];
	//				std::transform(wszValue.begin(), wszValue.end(), wszValue.begin(), ::tolower);
	//				if (wszValue == L"enable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModePublicFolders = CachedMode::Enabled;
	//					i++;

	//				}
	//				if (wszValue == L"disable")
	//				{
	//					pRunOpts->profileOptions->serviceOptions->cachedModePublicFolders = CachedMode::Disabled;
	//					i++;

	//				}
	//			}
	//		}
	//		else return false;
	//	}

	//	// Address Book specific validation
	//	if VCHK(pRunOpts->profileOptions->serviceOptions->serviceType, ServiceType::ServiceType_AddressBook)
	//	{
	//		if FCHK(pRunOpts->action, ACTION_SERVICE_ADD)
	//		{
	//			if (pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszConfigFilePath.empty())
	//			{
	//				return false;
	//			}
	//			else if (FCHK(pRunOpts->action, ACTION_SERVICE_UPDATE) ||
	//				FCHK(pRunOpts->action, ACTION_SERVICE_LIST) ||
	//				FCHK(pRunOpts->action, ACTION_SERVICE_REMOVE))
	//			{
	//				if (pRunOpts->profileOptions->serviceOptions->addressBookOptions->wszABDisplayName.empty())
	//				{
	//					return false;
	//				}

	//			}
	//		}
	//	}
	//	return true;
	//}

	// TODO: This is an example of a library function
	void fnMAPIToolkit()
	{
	}


}