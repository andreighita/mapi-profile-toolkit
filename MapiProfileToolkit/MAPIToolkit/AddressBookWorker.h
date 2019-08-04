#pragma once
#include "ServiceWorker.h"
class AddressBookWorker : public ServiceWorker
{
private:
	AddressBookWorker();
	// ACTION_SERVICE_ADD	
	void AddAddressBookService();

	// ACTION_SERVICE_UPDATE	
	void UpdateAddressBookService();

	// ACTION_SERVICE_LIST
	void ListAddressBookService();

	// ACTION_SERVICE_LISTALL
	void ListAllAddressBookServices();

	// ACTION_SERVICE_REMOVE
	void RemoveAddressBookService();

	// ACTION_SERVICE_REMOVEALL
	void RemoveAllAddressBookServices();
public:
	std::wstring wszProfileName;
	std::wstring wszABDisplayName;
	std::wstring wszConfigFilePath;
	std::wstring wszABServerName;
};

