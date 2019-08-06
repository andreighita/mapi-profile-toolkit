#pragma once
#include <Windows.h>
#include <tchar.h>

#include "ProfileWorker.h"
#include "ToolkitTypeDefs.h"
#include "AddressBookWorker.h"
#include "DataFileWorker.h"
#include "ExchangeAccountWorker.h"
#include "ProviderWorker.h"

#include <MAPIX.h>
namespace MAPIToolkit
{
	class Toolkit
	{

	protected:
		static void DisplayUsage();
		static BOOL Is64BitProcess(void);
		static int GetOutlookVersion();
		static BOOL IsCorrectBitness();

		static VOID RunAction();
		static BOOL ParseParams(int argc, wchar_t* argv[]);

		static BOOL Initialise();
		static VOID Uninitialise();

		static std::map<std::wstring, ULONG> g_actionsMap;
		static std::map<std::wstring, ULONG> g_profileModeMap;
		static std::map<std::wstring, ULONG> g_serviceModeMap;
		static std::map<std::wstring, ULONG> g_serviceTypeMap;
		static ULONG m_action;
		static int m_OutlookVersion;
		static ULONG m_loggingMode;
		static ServiceWorker* m_serviceWorker;
		static ProviderWorker* m_providerWorker;
		static ProfileWorker* m_profileWorker;
		static ULONG m_profileCount;
		static std::wstring m_wszExportPath;
		static ULONG m_exportMode; // 0 = no export; 1 = export;
		static std::wstring m_wszLogFilePath;
		static ULONG m_profileMode; // pm
		static LPPROFADMIN m_pProfAdmin;

	public:
		 static VOID Run(int argc, wchar_t* argv[]);
	};
}