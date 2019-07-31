#pragma once
#include "ServiceWorker.h"
class AddressBookWorker : public ServiceWorker
{
private:
	AddressBookWorker();

public:
	std::wstring wszProfileName;
	std::wstring wszABDisplayName;
	std::wstring wszConfigFilePath;
	std::wstring wszABServerName;
};

