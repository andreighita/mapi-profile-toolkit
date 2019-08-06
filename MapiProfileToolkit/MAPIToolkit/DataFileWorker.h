#pragma once
#include "ServiceWorker.h"
namespace MAPIToolkit
{
	class DataFileWorker : public ServiceWorker
	{
	public:
		int iPstIndex;
		ULONG ulPstType;
		std::wstring wszPstPath;
		std::wstring wszDisplayName;
		bool bMovePst;
		std::wstring wszPstOldPath;
		std::wstring wszPstNewPath;

		// ACTION_SERVICE_ADD
		void AddDataFile();
		// ACTION_SERVICE_CHANGEDATAFILEPATH
		void ChangeDataFilePath();
	};
}

