#pragma once
#include <Windows.h>
#include "ToolkitTypeDefs.h"
#include <string.h>

namespace MAPIToolkit
{
	class ProviderWorker
	{
	public:
		std::wstring wszSmtpAddress;			// msa		| 
		std::wstring wszMailboxLegacyDN;		// mmldn
		std::wstring wszMailboxDisplayName;		// mmdn
		std::wstring wszServerLegacyDN;			// msldn
		std::wstring wszServerDisplayName;		// msdn
		std::wstring wszRohProxyServer;			// mrps
		std::wstring wszMailStoreExternalUrl;	// mmse
		std::wstring wszMailStoreInternalUrl;	// mmsi
		ProviderType providerType;
		ULONG ulRohProxyServerFlags;			// mrpsf
		ULONG ulRohProxyServerAuthPackage;		// mrpsap
	};
}

