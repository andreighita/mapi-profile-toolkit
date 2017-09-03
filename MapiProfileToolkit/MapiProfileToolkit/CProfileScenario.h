#pragma once
#include "stdafx.h"
#include "CBaseScenario.h"

class CProfileScenario : CBaseScenario
{
	ULONG ulProfileMode;
	std::wstring wszProfileName;
	bool bSetDefaultProfile;
	ULONG ulProfileConnectMode;
};	