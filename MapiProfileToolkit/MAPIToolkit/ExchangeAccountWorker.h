#pragma once
#include "ServiceWorker.h"
class ExchangeAccountWorker: ServiceWorker
{
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

	ConnectMode connectMode;					// scnctm		| ROH or MOH
	CachedMode cachedModeOwner;				// scmo		| 1 = disabled; 2 = enabled; 
	CachedMode cachedModePublicFolders;			// scmpf	| 1 = disabled; 2 = enabled; 
	CachedMode cachedModeShared;				// scms		| 1 = disabled; 2 = enabled; 

	// ACTION_SERVICE_SETCACHEDMODE
	void SetCachedMode();
};

