#pragma once
#include "ToolkitTypeDefs.h"
#include "ProviderWorker.h"

class ServiceWorker
{
public:
	ServiceWorker();

#define ACTION_SERVICE_CHANGEDATAFILEPATH	0x00800000
	// ACTION_SERVICE_ADD	
	void AddService();

	// ACTION_SERVICE_UPDATE	
	void UpdateService();

	// ACTION_SERVICE_LIST
	void ListService();
	void ListDefaultService();

	// ACTION_SERVICE_LISTALL
	void ListAllServices();

	// ACTION_SERVICE_REMOVE
	void RemoveService();

	// ACTION_SERVICE_REMOVEALL
	void RemoveAllServices();
protected:
	LPMAPIUID m_pServiceUid;
	int m_iServiceIndex;
	ULONG m_ulConfigFlags;					// scfgf		| PR_PROFILE_CONFIG_FLAGS
	ULONG ulResourceFlags;					// srf		| PR_RESOURCES_FLAGS
	ServiceType m_serviceType;
	ServiceMode m_serviceMode;
};

