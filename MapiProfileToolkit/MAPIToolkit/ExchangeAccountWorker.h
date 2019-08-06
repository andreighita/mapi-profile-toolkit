#pragma once
#include "ServiceWorker.h"
namespace MAPIToolkit
{
	class ExchangeAccountWorker : public ServiceWorker
	{
	public:
		std::wstring wszAddressBookExternalUrl; // sabe 
		std::wstring wszAddressBookInternalUrl; // sabi
		std::wstring wszAutodiscoverUrl;		// sau
		std::wstring wszMailboxDisplayName;		// smdn
		std::wstring wszMailboxLegacyDN;		// smldn
		std::wstring wszMailStoreExternalUrl;	// smse
		std::wstring wszMailStoreInternalUrl;	// smsi
		std::wstring wszProfileName;			// spn
		std::wstring wszRohProxyServer;			// srps
		std::wstring wszServerDisplayName;		// ssdn
		std::wstring wszServerLegacyDN;			// ssldn
		std::wstring wszSmtpAddress;			// ssa
		std::wstring wszUnresolvedServer;		// sus
		std::wstring wszUnresolvedUser;			// suu

		ULONG connectMode;					// scnctm		| ROH or MOH
		ULONG cachedModeOwner;				// scmo		| 1 = disabled; 2 = enabled; 
		ULONG cachedModePublicFolders;			// scmpf	| 1 = disabled; 2 = enabled; 
		ULONG cachedModeShared;				// scms		| 1 = disabled; 2 = enabled; 

		// ACTION_SERVICE_ADD
		void AddExchangeAccount();
		// ACTION_SERVICE_SETCACHEDMODE
		void SetCachedMode();
	public:
		ExchangeAccountWorker();
	};

}