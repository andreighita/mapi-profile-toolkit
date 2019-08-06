#pragma once
#include "ServiceWorker.h"
namespace MAPIToolkit
{
	class AddressBookWorker : public ServiceWorker
	{
	public:
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
		BOOL enableBrowsing;
		BOOL requireSPA;
		BOOL useSSL;
		std::wstring configFilePath;
		std::wstring customSearchBase;
		std::wstring displayName;
		std::wstring maxResults;
		std::wstring serverName;
		std::wstring serverPort;
		std::wstring timeout;
		std::wstring username;
		std::wstring password;
		ULONG defaultSearchBase;

	};
}

