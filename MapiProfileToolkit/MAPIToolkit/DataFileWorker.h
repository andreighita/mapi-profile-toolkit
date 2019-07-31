#pragma once
#include "ServiceWorker.h"
class DataFileWorker : ServiceWorker
{
	int iPstIndex;
	ULONG ulPstType;
	std::wstring wszPstPath;
	std::wstring wszDisplayName;
	bool bMovePst;
	std::wstring wszPstOldPath;
	std::wstring wszPstNewPath;

	// ACTION_SERVICE_CHANGEDATAFILEPATH
	void ChangeDataFilePath();
};

