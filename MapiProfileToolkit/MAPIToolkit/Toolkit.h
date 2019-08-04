#pragma once

#include "ToolkitTypeDefs.h"
#include "ProfileWorker.h"
namespace MAPIToolkit
{
	class Toolkit
	{
	public:
		ProfileWorker* m_profileWorker;
		ServiceWorker* m_serviceWorker;
		ProviderWorker* m_providerWorker;

		ULONG action;
		LoggingMode loggingMode;
		std::wstring wszExportPath;
		ExportMode exportMode; // 0 = no export; 1 = export;
		std::wstring wszLogFilePath;
		int iOutlookVersion;

	public:
		Toolkit();

	private:
		~Toolkit();
	};
}
