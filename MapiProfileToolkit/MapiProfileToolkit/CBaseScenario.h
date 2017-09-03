#pragma once
#include "stdafx.h"

class CBaseScenario
{
	ULONG ulScenario;
	ULONG ulActionType;
	ULONG ulAction;
	ULONG ulLoggingMode;
	std::wstring wszExportPath;
	BOOL bExportMode; // 0 = no export; 1 = export;
	BOOL bNoHeader;
	std::wstring wszLogFilePath;
public:
	void Init();
};
