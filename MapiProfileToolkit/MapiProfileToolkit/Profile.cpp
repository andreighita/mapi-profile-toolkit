#include "Profile.h"

LPWSTR GetDefaultProfileNameLP()
{
	return (LPWSTR)GetDefaultProfileName().c_str();
}

// GetDefaultProfileName
// returns a std::wstring value with the name of the default Outlook profile
std::wstring GetDefaultProfileName()
{
	std::wstring szDefaultProfileName;
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPSRestriction lpProfRes = NULL;
	LPSRestriction lpProfResLvl1 = NULL;
	LPSPropValue lpProfPropVal = NULL;
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	HRESULT hRes = S_OK;
	Logger::Write(logLevelInfo, L"Attempting to retrieve the default MAPI profile name");

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME_A };

	EC_HRES_MSG(MAPIAdminProfiles(0, &lpProfAdmin), L"Calling MAPIAdminProfiles" );

	EC_HRES_MSG(lpProfAdmin->GetProfileTable(0, &lpProfTable), L"Calling GetProfileTable");

	// Allocate memory for the restriction
	EC_HRES_MSG(MAPIAllocateBuffer(sizeof(SRestriction), (LPVOID*)&lpProfRes), L"Calling MAPIAllocateBuffer");

	EC_HRES_MSG(MAPIAllocateBuffer(sizeof(SRestriction) * 2, (LPVOID*)&lpProfResLvl1), L"Calling MAPIAllocateBuffer");

	EC_HRES_MSG(MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)&lpProfPropVal), L"Calling MAPIAllocateBuffer");

	// Set up restriction to query the profile table
	lpProfRes->rt = RES_AND;
	lpProfRes->res.resAnd.cRes = 0x00000002;
	lpProfRes->res.resAnd.lpRes = lpProfResLvl1;

	lpProfResLvl1[0].rt = RES_EXIST;
	lpProfResLvl1[0].res.resExist.ulPropTag = PR_DEFAULT_PROFILE;
	lpProfResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
	lpProfResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
	lpProfResLvl1[1].rt = RES_PROPERTY;
	lpProfResLvl1[1].res.resProperty.relop = RELOP_EQ;
	lpProfResLvl1[1].res.resProperty.ulPropTag = PR_DEFAULT_PROFILE;
	lpProfResLvl1[1].res.resProperty.lpProp = lpProfPropVal;

	lpProfPropVal->ulPropTag = PR_DEFAULT_PROFILE;
	lpProfPropVal->Value.b = true;

	// Query the table to get the the default profile only
	EC_HRES_MSG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		lpProfRes,
		NULL,
		0,
		&lpProfRows), L"Calling HrQueryAllRows");

	if (lpProfRows->cRows == 0)
	{
		Logger::Write(logLevelFailed, L"No default profile set.");
	}
	else if (lpProfRows->cRows == 1)
	{

		szDefaultProfileName = ConvertMultiByteToWideChar(lpProfRows->aRow->lpProps[iDisplayName].Value.lpszA);
	}
	else
	{
		Logger::Write(logLevelError, L"Query resulted in incosinstent results");
	}

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	if (lpProfRows) FreeProws(lpProfRows);
	if (lpProfTable) lpProfTable->Release();
	if (lpProfRes) MAPIFreeBuffer(lpProfRes);
	if (lpProfResLvl1) MAPIFreeBuffer(lpProfResLvl1);
	if (lpProfAdmin) lpProfAdmin->Release();
	return szDefaultProfileName;
}

// GetProfileCount
// returns the number of mapi profiles for the current user
ULONG GetProfileCount()
{
	std::string szDefaultProfileName;
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPMAPITABLE lpProfTable = NULL;
	ULONG ulRowCount = 0;
	HRESULT hRes = S_OK;

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin
									 // Get an IProfAdmin interface.

	EC_HRES_MSG(lpProfAdmin->GetProfileTable(0, &lpProfTable), L"Calling GetProfileTable");

	EC_HRES_MSG(lpProfTable->GetRowCount(0, &ulRowCount), L"Calling GetRowCount");

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	if (lpProfTable) lpProfTable->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	return ulRowCount;
}

HRESULT HrGetProfiles(ULONG ulProfileCount, ProfileInfo * profileInfo)
{
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	HRESULT hRes = S_OK;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME_A };

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin
									 // Get an IProfAdmin interface.

	EC_HRES_MSG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), L"Calling GetProfileTable");

	// Query the table to get the the default profile only
	EC_HRES_MSG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		NULL,
		NULL,
		0,
		&lpProfRows), L"Calling HrQueryAllRows");

	if (lpProfRows->cRows == ulProfileCount)
	{
		for (unsigned int i = 0; i < lpProfRows->cRows; i++)
		{
			EC_HRES_MSG(HrGetProfile(ConvertMultiByteToWideChar(lpProfRows->aRow[i].lpProps[iDisplayName].Value.lpszA), &profileInfo[i]), L"Calling HrGetProfile");
		}
	}

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	if (lpProfRows) FreeProws(lpProfRows);
	if (lpProfTable) lpProfTable->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	return hRes;
}

HRESULT HrDeleteProfile(LPWSTR lpszProfileName)
{
	HRESULT				hRes = S_OK;            // Result from MAPI calls.
	LPPROFADMIN			lpProfAdmin = NULL;     // Profile Admin object.
	LPSERVICEADMIN		lpSvcAdmin = NULL;      // Service Admin object.
	LPSERVICEADMIN2		lpSvcAdmin2 = NULL;

	// This indicates columns we want returned from HrQueryAllRows.
	enum { iSvcName, iSvcUID, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME, PR_SERVICE_UID };

	// Get an IProfAdmin interface.

	EC_HRES_MSG(MAPIAdminProfiles(0,              // Flags.
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin.

													   // Create a new profile.
		EC_HRES_MSG(lpProfAdmin->DeleteProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName), NULL), L"Calling DeleteProfile");
		// Create a new profile.

Error:
	goto Cleanup;

Cleanup:
	// Clean up
	if (lpProfAdmin) lpProfAdmin->Release();

	return 0;

}

HRESULT HrCreateProfile(LPWSTR lpszProfileName)
{
	HRESULT				hRes = S_OK;            // Result from MAPI calls.
	LPPROFADMIN			lpProfAdmin = NULL;     // Profile Admin object.
	LPSERVICEADMIN		lpSvcAdmin = NULL;      // Service Admin object.
	LPSERVICEADMIN2		lpSvcAdmin2 = NULL;

	// This indicates columns we want returned from HrQueryAllRows.
	enum { iSvcName, iSvcUID, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME, PR_SERVICE_UID };

	// Get an IProfAdmin interface.

	EC_HRES_MSG(MAPIAdminProfiles(0,              // Flags.
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin.

													   // Create a new profile.
	hRes = lpProfAdmin->CreateProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
		nullptr,          // Password for profile.
		0,          // Handle to parent window.
		0);        // Flags.

	if (hRes == E_ACCESSDENIED)
	{
		EC_HRES_MSG(lpProfAdmin->DeleteProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName), NULL), L"Calling DeleteProfile");
		// Create a new profile.

		EC_HRES_MSG(lpProfAdmin->CreateProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
			nullptr,          // Password for profile.
			0,          // Handle to parent window.
			0), L"Calling CreateProfile.");        // Flags.
	}

Error:
	goto Cleanup;

Cleanup:
	// Clean up
	if (lpProfAdmin) lpProfAdmin->Release();

	return 0;

}

HRESULT HrCreateProfile(LPWSTR lpszProfileName, LPSERVICEADMIN2 *lppSvcAdmin2)
{
	HRESULT				hRes = S_OK;            // Result from MAPI calls.
	LPPROFADMIN			lpProfAdmin = NULL;     // Profile Admin object.
	LPSERVICEADMIN		lpSvcAdmin = NULL;      // Service Admin object.
	LPSERVICEADMIN2		lpSvcAdmin2 = NULL;

	// This indicates columns we want returned from HrQueryAllRows.
	enum { iSvcName, iSvcUID, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME, PR_SERVICE_UID };

	// Get an IProfAdmin interface.

	EC_HRES_MSG(MAPIAdminProfiles(0,              // Flags.
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin.

													   // Create a new profile.
	hRes = lpProfAdmin->CreateProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
		nullptr,          // Password for profile.
		0,          // Handle to parent window.
		0);        // Flags.

	if (hRes == E_ACCESSDENIED)
	{
		EC_HRES_MSG(lpProfAdmin->DeleteProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName), NULL), L"Calling DeleteProfile.");

		EC_HRES_MSG(lpProfAdmin->CreateProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
			nullptr,          // Password for profile.
			0,          // Handle to parent window.
			0), L"Calling CreateProfile.");        // Flags.
	}

	// Get an IMsgServiceAdmin interface off of the IProfAdmin interface.
	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Profile that we want to modify.
		nullptr,          // Password for that profile.
		0,          // Handle to parent window.
		0,             // Flags.
		&lpSvcAdmin), L"Calling AdminServices."); // Pointer to new IMsgServiceAdmin.

												  // Create the new message service for Exchange.
	if (lpSvcAdmin) EC_HRES_MSG(lpSvcAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)&lpSvcAdmin2), L"Calling QueryInterface");

	*lppSvcAdmin2 = lpSvcAdmin2;

	goto Cleanup;

Error:
	goto Cleanup;

Cleanup:
	// Clean up
	if (lpSvcAdmin) lpSvcAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();

	return 0;

}

HRESULT HrSetDefaultProfile(LPWSTR lpszProfileName)
{
	HRESULT				hRes = S_OK;            // Result from MAPI calls.
	LPPROFADMIN			lpProfAdmin = NULL;     // Profile Admin object.

												// This indicates columns we want returned from HrQueryAllRows.
	enum { iSvcName, iSvcUID, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME, PR_SERVICE_UID };

	// Get an IProfAdmin interface.

	EC_HRES_MSG(MAPIAdminProfiles(0,              // Flags.
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin.

													   // Create a new profile.
	EC_HRES_MSG(lpProfAdmin->SetDefaultProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
		0), L"Calling SetDefaultProfile.");        // Flags.

Error:
	goto Cleanup;

Cleanup:
	// Clean up
	if (lpProfAdmin) lpProfAdmin->Release();

	return 0;

}

// Outlook 2016
HRESULT HrCloneProfile(ProfileInfo * profileInfo)
{
	HRESULT hRes = S_OK;
	LPSERVICEADMIN2 lpServiceAdmin = NULL;
	unsigned int uiServiceIndex = 0;
	profileInfo->wszProfileName = profileInfo->wszProfileName + L"_Clone";
	Logger::Write(logLevelInfo, L"Creating new profile named: " + profileInfo->wszProfileName);
	EC_HRES_MSG(HrCreateProfile((LPWSTR)profileInfo->wszProfileName.c_str(), &lpServiceAdmin), L"Calling HrCreateProfile.");
	if (lpServiceAdmin)
	{
		for (unsigned int i = 0; i < profileInfo->ulServiceCount; i++)
		{
			MAPIUID uidService = { 0 };
			LPMAPIUID lpServiceUid = &uidService;
			if (profileInfo->profileServices[i].ulServiceType == SERVICETYPE_MAILBOX)
			{
				Logger::Write(logLevelInfo, L"Adding exchange mailbox: " + profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress);
				EC_HRES_MSG(HrCreateMsemsServiceModernExt(false, // sort this out later
					(LPWSTR)profileInfo->wszProfileName.c_str(),
					profileInfo->profileServices[i].ulResourceFlags,
					profileInfo->profileServices[i].exchangeAccountInfo->ulProfileConfigFlags,
					profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths,
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszDisplayName.c_str()), L"Calling HrCreateMsemsServiceModernExt.");

				MAPIUID uidService = { 0 };
				memcpy((LPVOID)&uidService, lpServiceUid, sizeof(MAPIUID));
				for (unsigned int j = 0; j < profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount; j++)
				{
					if (profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulProfileType == PROFILE_DELEGATE)
					{
						Logger::Write(logLevelInfo, L"Adding additional mailbox: " + profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress);
						// this should not add online archives
						if (TRUE != profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].bIsOnlineArchive)
							EC_HRES_MSG(HrAddDelegateMailboxModern(false,
							(LPWSTR)profileInfo->wszProfileName.c_str(),
								FALSE,
								uiServiceIndex,
								(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName.c_str(),
								(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress.c_str()), L"Calling HrAddDelegateMailboxModern.");
					}
				}
				uiServiceIndex++;
			}
			else if (profileInfo->profileServices[i].ulServiceType == SERVICETYPE_PST)
			{
				Logger::Write(logLevelInfo, L"Adding PST file: " + profileInfo->profileServices[i].pstInfo->wszPstPath);
				EC_HRES_MSG(HrCreatePstService(lpServiceAdmin,
					&lpServiceUid,
					(LPWSTR)profileInfo->profileServices[i].wszServiceName.c_str(),
					profileInfo->profileServices[i].ulResourceFlags,
					profileInfo->profileServices[i].pstInfo->ulPstConfigFlags,
					(LPWSTR)profileInfo->profileServices[i].pstInfo->wszPstPath.c_str(),
					(LPWSTR)profileInfo->profileServices[i].pstInfo->wszDisplayName.c_str()), L"Calling HrCreatePstService.");
				uiServiceIndex++;
			}

		}

		Logger::Write(logLevelInfo, L"Setting profile as default.");
		EC_HRES_MSG(HrSetDefaultProfile((LPWSTR)profileInfo->wszProfileName.c_str()), L"Calling HrSetDefaultProfile.");
	}
	goto Cleanup;

Error:
	goto Cleanup;
Cleanup:
	return hRes;
}

// Outlook 2013
HRESULT HrSimpleCloneProfile(ProfileInfo * profileInfo, bool bSetDefaultProfile)
{
	HRESULT hRes = S_OK;
	LPSERVICEADMIN2 lpServiceAdmin = NULL;
	unsigned int uiServiceIndex = 0;
	profileInfo->wszProfileName = profileInfo->wszProfileName + L"_Clone";
	Logger::Write(logLevelInfo, L"Creating new profile named: " + profileInfo->wszProfileName);
	EC_HRES_MSG(HrCreateProfile((LPWSTR)profileInfo->wszProfileName.c_str(), &lpServiceAdmin), L"Calling HrCreateProfile.");
	if (lpServiceAdmin)
	{
		for (unsigned int i = 0; i < profileInfo->ulServiceCount; i++)
		{
			MAPIUID uidService = { 0 };
			LPMAPIUID lpServiceUid = &uidService;
			if (profileInfo->profileServices[i].ulServiceType == SERVICETYPE_MAILBOX)
			{
				Logger::Write(logLevelInfo, L"Adding exchange mailbox: " + profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress);
				
				EC_HRES_MSG(HrCreateMsemsServiceMOH(false,
					(LPWSTR)profileInfo->wszProfileName.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszMailboxDN.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerName.c_str(),
					NULL,
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszMailStoreExternalUrl.c_str(),
					NULL,
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszAddressBookExternalUrl.c_str()), L"HrCreateMsemsServiceMOH");

				uiServiceIndex++;
			}
		}
		if (bSetDefaultProfile)
		{
			Logger::Write(logLevelInfo, L"Setting profile as default.");
			EC_HRES_MSG(HrSetDefaultProfile((LPWSTR)profileInfo->wszProfileName.c_str()), L"Calling HrSetDefaultProfile.");
		}
	}
	goto Cleanup;

Error:
	goto Cleanup;
Cleanup:
	return hRes;
}

VOID PrintProfile(ProfileInfo * profileInfo)
{
	if (profileInfo)
	{
		wprintf(L"Profile name: %ls\n", profileInfo->wszProfileName.c_str());
		wprintf(L"Service count: %i\n", profileInfo->ulServiceCount);
		for (unsigned int i = 0; i < profileInfo->ulServiceCount; i++)
		{
			wprintf(L" -> Service #%i\n", i);
			wprintf(L" -> [%i] Display name: %ls\n", i, profileInfo->profileServices[i].wszDisplayName.c_str());
			wprintf(L" -> [%i] Service name: %ls\n", i, profileInfo->profileServices[i].wszServiceName.c_str());
			wprintf(L" -> [%i] Service resource flags: %#x\n", i, profileInfo->profileServices[i].ulResourceFlags);
			MAPIUID uidService = { 0 };
			LPMAPIUID lpServiceUid = &uidService;
			if (profileInfo->profileServices[i].ulServiceType == SERVICETYPE_MAILBOX)
			{
				wprintf(L" -> [%i] Service type: %ls\n", i, L"Exchange Mailbox");
				wprintf(L" -> [%i] E-mail address: %ls\n", i, profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress.c_str());
				wprintf(L" -> [%i] User display name: %ls\n", i, profileInfo->profileServices[i].exchangeAccountInfo->wszDisplayName.c_str());
				wprintf(L" -> [%i] OST path: %ls\n", i, profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath.c_str());
				wprintf(L" -> [%i] Config flags: %#x\n", i, profileInfo->profileServices[i].exchangeAccountInfo->ulProfileConfigFlags);
				wprintf(L" -> [%i] Cached mode months: %i\n", i, profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths);
				wprintf(L" -> [%i] Mailbox count: %i\n", i, profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount);
				for (unsigned int j = 0; j < profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount; j++)
				{
					wprintf(L" -> [%i] -> Mailbox #%i\n", i, j);
					if (profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulProfileType == PROFILE_DELEGATE)
					{
						wprintf(L" -> [%i] -> [%i] (Delegate) -> E-mail address: %ls\n", i, j, profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress.c_str());
						wprintf(L" -> [%i] -> [%i] (Delegate) -> User display name: %ls\n", i, j, profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName.c_str());
					}
					else
					{
						wprintf(L" -> [%i] -> [%i] (Other)-> E-mail address:%ls\n", i, j, profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress.c_str());
						wprintf(L" -> [%i] -> [%i] (Other) -> User display name:%ls\n", i, j, profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName.c_str());
					}
				}
			}
			else if (profileInfo->profileServices[i].ulServiceType == SERVICETYPE_PST)
			{
				wprintf(L" -> [%i] Service type: %ls\n", i, L"PST");
				wprintf(L" -> [%i] Display name: %ls\n", i, profileInfo->profileServices[i].pstInfo->wszDisplayName.c_str());
				wprintf(L" -> [%i] PST path: %ls\n", i, profileInfo->profileServices[i].pstInfo->wszPstPath.c_str());
				wprintf(L" -> [%i] Config flags: %#x\n", i, profileInfo->profileServices[i].pstInfo->ulPstConfigFlags);
			}
		}
	}

}

HRESULT HrGetProfile(LPWSTR lpszProfileName, ProfileInfo * profileInfo)
{
	HRESULT hRes = S_OK;
	profileInfo->wszProfileName = ConvertWideCharToStdWstring(lpszProfileName);

	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPSRestriction lpProfRes = NULL;
	LPSRestriction lpProfResLvl1 = NULL;
	LPSPropValue lpProfPropVal = NULL;
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, iDefaultProfile, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME, PR_DEFAULT_PROFILE };

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin
													   // Get an IProfAdmin interface.

	EC_HRES_MSG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), L"Calling GetProfileTable.");

	// Allocate memory for the restriction
	EC_HRES_MSG(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)&lpProfRes), L"Calling MAPIAllocateBuffer.");

	EC_HRES_MSG(MAPIAllocateBuffer(
		sizeof(SRestriction) * 2,
		(LPVOID*)&lpProfResLvl1), L"Calling MAPIAllocateBuffer");

	EC_HRES_MSG(MAPIAllocateBuffer(
		sizeof(SPropValue),
		(LPVOID*)&lpProfPropVal), L"Calling MAPIAllocateBuffer");

	// Set up restriction to query the profile table
	lpProfRes->rt = RES_AND;
	lpProfRes->res.resAnd.cRes = 0x00000002;
	lpProfRes->res.resAnd.lpRes = lpProfResLvl1;

	lpProfResLvl1[0].rt = RES_EXIST;
	lpProfResLvl1[0].res.resExist.ulPropTag = PR_DISPLAY_NAME_A;
	lpProfResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
	lpProfResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
	lpProfResLvl1[1].rt = RES_PROPERTY;
	lpProfResLvl1[1].res.resProperty.relop = RELOP_EQ;
	lpProfResLvl1[1].res.resProperty.ulPropTag = PR_DISPLAY_NAME_A;
	lpProfResLvl1[1].res.resProperty.lpProp = lpProfPropVal;

	lpProfPropVal->ulPropTag = PR_DISPLAY_NAME_A;
	lpProfPropVal->Value.lpszA = ConvertWideCharToMultiByte(lpszProfileName);

	// Query the table to get the the default profile only
	EC_HRES_MSG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		lpProfRes,
		NULL,
		0,
		&lpProfRows), L"Calling HrQueryAllRows.");

	if (lpProfRows->cRows == 0)
	{
		return MAPI_E_NOT_FOUND;
	}
	else if (lpProfRows->cRows == 1)
	{
		profileInfo->bDefaultProfile = lpProfRows->aRow->lpProps[iDefaultProfile].Value.b;
	}
	else
	{
		return MAPI_E_CALL_FAILED;
	}

	// Begin process services
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPMAPITABLE lpServiceTable = NULL;
	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		MAPI_UNICODE,                    // Flags.
		&lpServiceAdmin), L"Calling AdminServices.");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		lpServiceAdmin->GetMsgServiceTable(0,
			&lpServiceTable);
		LPSRestriction lpSvcRes = NULL;
		LPSRestriction lpsvcResLvl1 = NULL;
		LPSPropValue lpSvcPropVal = NULL;
		LPSRowSet lpSvcRows = NULL;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceName, iDisplayNameW, iEmsMdbSectUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID ,PR_SERVICE_NAME_A, PR_DISPLAY_NAME_W, PR_EMSMDB_SECTION_UID, PR_RESOURCE_FLAGS };

		// Query the table to get the the default profile only
		EC_HRES_MSG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			NULL,
			NULL,
			0,
			&lpSvcRows), L"Calling HrQueryAllRows.");

		if (lpSvcRows->cRows > 0)
		{
			profileInfo->ulServiceCount = lpSvcRows->cRows;
			profileInfo->profileServices = new ServiceInfo[lpSvcRows->cRows];


			// Start loop services
			for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
			{
				ZeroMemory(&profileInfo->profileServices[i], sizeof(ServiceInfo));
				profileInfo->profileServices[i].wszServiceName = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA));
				profileInfo->profileServices[i].ulResourceFlags = lpSvcRows->aRow[i].lpProps[iServiceResFlags].Value.l;
				profileInfo->profileServices[i].wszDisplayName = lpSvcRows->aRow[i].lpProps[iDisplayNameW].Value.lpszW;
				profileInfo->profileServices[i].ulServiceType = SERVICETYPE_OTHER;;
				if (profileInfo->profileServices[i].ulResourceFlags & SERVICE_DEFAULT_STORE)
				{
					profileInfo->profileServices[i].bDefaultStore = true;
				}
				// Exchange account
				if (0 == strcmp(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA, "MSEMS"))
				{
					profileInfo->profileServices[i].exchangeAccountInfo = new ExchangeAccountInfo();
					profileInfo->profileServices[i].ulServiceType = SERVICETYPE_MAILBOX;
					LPPROVIDERADMIN lpProvAdmin = NULL;

					if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb,
						0,
						&lpProvAdmin)))
					{

						// Read the EMSMDB section
						LPPROFSECT lpProfSect = NULL;
						if (SUCCEEDED(lpProvAdmin->OpenProfileSection((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iEmsMdbSectUid].Value.bin.lpb,
							NULL,
							0L,
							&lpProfSect)))
						{
							LPMAPIPROP pMAPIProp = NULL;
							if (SUCCEEDED(lpProfSect->QueryInterface(IID_IMAPIProp, (void**)&pMAPIProp)))
							{
								// bind to the PR_RULE_ACTION_TYPE property to get the ammount to sync
								LPSPropValue profilePrRuleActionType = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_RULE_ACTION_TYPE, &profilePrRuleActionType)))
								{
									if (profilePrRuleActionType)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths = profilePrRuleActionType->Value.i;
										if (profilePrRuleActionType) MAPIFreeBuffer(profilePrRuleActionType);
									}

								}
								else
								{
									profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths = 0;
								}

								// PR_PROFILE_OFFLINE_STORE_PATH_W
								LPSPropValue profileOfflineStorePath = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_OFFLINE_STORE_PATH_W, &profileOfflineStorePath)))
								{
									if (profileOfflineStorePath)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath = profileOfflineStorePath->Value.lpszW;
										if (profileOfflineStorePath) MAPIFreeBuffer(profileOfflineStorePath);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath = L"";
									}

								}
								else
								{
									profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths = 0;
								}
								// bind to the PR_PROFILE_CONFIG_FLAGS property
								LPSPropValue profileConfigFlags = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_CONFIG_FLAGS, &profileConfigFlags)))
								{
									if (profileConfigFlags)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->ulProfileConfigFlags = profileConfigFlags->Value.l;
									}
									MAPIFreeBuffer(profileConfigFlags);
								}
								// bind to the PR_PROFILE_USER_SMTP_EMAIL_ADDRESS property
								LPSPropValue profileUserSmtpEmailAddress = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_USER_SMTP_EMAIL_ADDRESS, &profileUserSmtpEmailAddress)))
								{
									if (profileUserSmtpEmailAddress)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress = SubstringToEnd(L"smtp:", ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUserSmtpEmailAddress->Value.lpszA)));
										if (profileUserSmtpEmailAddress) MAPIFreeBuffer(profileUserSmtpEmailAddress);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress = std::wstring(L" ");
									}
								}
								LPSPropValue profileDisplayName = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_DISPLAY_NAME_A, &profileDisplayName)))
								{
									if (profileDisplayName)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDisplayName = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileDisplayName->Value.lpszA));
										if (profileDisplayName) MAPIFreeBuffer(profileDisplayName);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDisplayName = std::wstring(L" ");
									}
								}

								// PR_PROFILE_USER
								LPSPropValue profileUser = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_USER, &profileUser)))
								{
									if (profileUser)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszMailboxDN = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUser->Value.lpszA));
										if (profileUser) MAPIFreeBuffer(profileUser);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszMailboxDN = std::wstring(L" ");
									}
								}

								// PR_PROFILE_HOME_SERVER_DN
								LPSPropValue profileHomeServerDN = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_HOME_SERVER_DN, &profileHomeServerDN)))
								{
									if (profileHomeServerDN)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileHomeServerDN->Value.lpszA));
										if (profileHomeServerDN) MAPIFreeBuffer(profileHomeServerDN);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN = std::wstring(L" ");
									}
								}


								// PR_PROFILE_HOME_SERVER
								LPSPropValue profileHomeServer = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_HOME_SERVER, &profileHomeServer)))
								{
									if (profileHomeServer)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerName = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileHomeServer->Value.lpszA));
										if (profileHomeServer) MAPIFreeBuffer(profileHomeServer);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN = std::wstring(L" ");
									}
								}

								// PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL
								LPSPropValue profileMapiHttpMailStoreExternal = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL, &profileMapiHttpMailStoreExternal)))
								{
									if (profileMapiHttpMailStoreExternal)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszMailStoreExternalUrl = profileMapiHttpMailStoreExternal->Value.lpszW;
										if (profileMapiHttpMailStoreExternal) MAPIFreeBuffer(profileMapiHttpMailStoreExternal);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath = L"";
									}

								}

								// PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL
								LPSPropValue profileMapiHttpAddressbookExternal = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL, &profileMapiHttpAddressbookExternal)))
								{
									if (profileMapiHttpAddressbookExternal)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszAddressBookExternalUrl = profileMapiHttpAddressbookExternal->Value.lpszW;
										if (profileMapiHttpAddressbookExternal) MAPIFreeBuffer(profileMapiHttpAddressbookExternal);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath = L"";
									}
								}

								// PR_SERVICE_UID
								LPSPropValue serviceUid = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_SERVICE_UID, &serviceUid)))
								{
									if (serviceUid)
									{
										LPMAPIUID lpMuidServiceUid = NULL;
										lpMuidServiceUid = &profileInfo->profileServices[i].muidServiceUid;
										memcpy(lpMuidServiceUid, serviceUid->Value.bin.lpb, sizeof(MAPIUID));
										if (serviceUid) MAPIFreeBuffer(serviceUid);
									}
								}
							}
							if (lpProfSect) lpProfSect->Release();
						}

						// End read the EMSMDB section


						// Loop providers
						LPMAPITABLE lpProvTable = NULL;
						LPSRestriction lpProvRes = NULL;
						LPSRestriction lpProvResLvl1 = NULL;
						LPSPropValue lpProvPropVal = NULL;
						LPSRowSet lpProvRows = NULL;

						// Setting up an enum and a prop tag array with the props we'll use
						enum { iProvInstanceKey, cptaProvProps };
						SizedSPropTagArray(cptaProvProps, sptaProvProps) = { cptaProvProps, PR_INSTANCE_KEY };

						// Allocate memory for the restriction
						EC_HRES_MSG(MAPIAllocateBuffer(
							sizeof(SRestriction),
							(LPVOID*)&lpProvRes), L"Calling MAPIAllocateBuffer.");

						EC_HRES_MSG(MAPIAllocateBuffer(
							sizeof(SRestriction) * 2,
							(LPVOID*)&lpProvResLvl1), L"Calling MAPIAllocateBuffer.");

						EC_HRES_MSG(MAPIAllocateBuffer(
							sizeof(SPropValue),
							(LPVOID*)&lpProvPropVal), L"Calling MAPIAllocateBuffer.");

						// Set up restriction to query the provider table
						lpProvRes->rt = RES_AND;
						lpProvRes->res.resAnd.cRes = 0x00000002;
						lpProvRes->res.resAnd.lpRes = lpProvResLvl1;

						lpProvResLvl1[0].rt = RES_EXIST;
						lpProvResLvl1[0].res.resExist.ulPropTag = PR_RESOURCE_TYPE;
						lpProvResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
						lpProvResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
						lpProvResLvl1[1].rt = RES_PROPERTY;
						lpProvResLvl1[1].res.resProperty.ulPropTag = PR_RESOURCE_TYPE;
						lpProvResLvl1[1].res.resProperty.lpProp = lpProvPropVal;
						lpProvResLvl1[1].res.resProperty.relop = RELOP_EQ;

						lpProvPropVal->ulPropTag = PR_RESOURCE_TYPE;
						lpProvPropVal->Value.l = MAPI_STORE_PROVIDER;

						lpProvAdmin->GetProviderTable(0,
							&lpProvTable);
						// Query the table to get the the default profile only
						EC_HRES_MSG(HrQueryAllRows(lpProvTable,
							(LPSPropTagArray)&sptaProvProps,
							lpProvRes,
							NULL,
							0,
							&lpProvRows), L"Calling HrQueryAllRows.");

						if (lpProvRows->cRows > 0)
						{
							profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount = lpProvRows->cRows;
							profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes = new MailboxInfo[lpProvRows->cRows];

							for (unsigned int j = 0; j < lpProvRows->cRows; j++)
							{
								ZeroMemory(&profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j], sizeof(MailboxInfo));
								profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName = std::wstring(L" ");
								LPPROFSECT lpProfSection = NULL;
								if (SUCCEEDED(lpServiceAdmin->OpenProfileSection((LPMAPIUID)lpProvRows->aRow[j].lpProps[iProvInstanceKey].Value.bin.lpb, NULL, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSection)))
								{

									LPMAPIPROP lpMAPIProp = NULL;
									if (SUCCEEDED(lpProfSection->QueryInterface(IID_IMAPIProp, (void**)&lpMAPIProp)))
									{
										// PR_DISPLAY_NAME
										LPSPropValue prDisplayName = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_DISPLAY_NAME, &prDisplayName)))
										{
											profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName = ConvertWideCharToStdWstring(prDisplayName->Value.lpszW);
											if (prDisplayName) MAPIFreeBuffer(prDisplayName);
										}
										else
										{
											profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName = std::wstring(L" ");
										}

										// PR_PROFILE_TYPE
										LPSPropValue prProfileType = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_TYPE, &prProfileType)))
										{
											profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulProfileType = prProfileType->Value.l;
										}

										// PR_PROFILE_USER_SMTP_EMAIL_ADDRESS
										LPSPropValue profileUserSmtpEmailAddress = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_USER_SMTP_EMAIL_ADDRESS, &profileUserSmtpEmailAddress)))
										{
											if (profileUserSmtpEmailAddress)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUserSmtpEmailAddress->Value.lpszA));
												if (profileUserSmtpEmailAddress) MAPIFreeBuffer(profileUserSmtpEmailAddress);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress = std::wstring(L" ");
											}
										}

										// PR_PROFILE_MAILBOX
										LPSPropValue profileMailbox = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_MAILBOX, &profileMailbox)))
										{
											if (profileMailbox)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileMailbox = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileMailbox->Value.lpszA));
												if (profileMailbox) MAPIFreeBuffer(profileMailbox);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileMailbox = std::wstring(L" ");
											}
										}

										// PR_PROFILE_SERVER_DN
										LPSPropValue profileServerDN = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_SERVER_DN, &profileServerDN)))
										{
											if (profileMailbox)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServerDN = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileServerDN->Value.lpszA));
												if (profileServerDN) MAPIFreeBuffer(profileServerDN);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServerDN = std::wstring(L" ");
											}
										}

										// PR_ROH_PROXY_SERVER
										LPSPropValue rohProxyServer = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_ROH_PROXY_SERVER, &rohProxyServer)))
										{
											if (rohProxyServer)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszRohProxyServer = ConvertWideCharToStdWstring(rohProxyServer->Value.lpszW);
												if (rohProxyServer) MAPIFreeBuffer(rohProxyServer);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszRohProxyServer = std::wstring(L" ");
											}
										}

										// PR_PROFILE_SERVER
										LPSPropValue profileServer = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_SERVER, &profileServer)))
										{
											if (profileServer)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServer = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileServer->Value.lpszA));
												if (profileServer) MAPIFreeBuffer(profileServer);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServer = std::wstring(L" ");
											}
										}

										// PR_PROFILE_SERVER_FQDN_W
										LPSPropValue profileServerFqdnW = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_SERVER_FQDN_W, &profileServerFqdnW)))
										{
											if (profileServerFqdnW)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServerFqdnW = ConvertWideCharToStdWstring(profileServerFqdnW->Value.lpszW);
												if (profileServerFqdnW) MAPIFreeBuffer(profileServerFqdnW);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServerFqdnW = std::wstring(L" ");
											}
										}

										// PR_PROFILE_LKG_AUTODISCOVER_URL
										LPSPropValue profileAutodiscoverUrl = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_LKG_AUTODISCOVER_URL, &profileAutodiscoverUrl)))
										{
											if (profileAutodiscoverUrl)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszAutodiscoverUrl = ConvertWideCharToStdWstring(profileAutodiscoverUrl->Value.lpszW);
												if (profileServerFqdnW) MAPIFreeBuffer(profileAutodiscoverUrl);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszAutodiscoverUrl = std::wstring(L" ");
											}
										}

										// PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL
										LPSPropValue profileMailStoreInternalUrl = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL, &profileMailStoreInternalUrl)))
										{
											if (profileMailStoreInternalUrl)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszMailStoreInternalUrl = ConvertWideCharToStdWstring(profileMailStoreInternalUrl->Value.lpszW);
												if (profileMailStoreInternalUrl) MAPIFreeBuffer(profileMailStoreInternalUrl);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszMailStoreInternalUrl = std::wstring(L" ");
											}
										}

										// PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL
										LPSPropValue profileMailStoreExternalUrl = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL, &profileMailStoreExternalUrl)))
										{
											if (profileMailStoreExternalUrl)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszMailStoreExternalUrl = ConvertWideCharToStdWstring(profileMailStoreExternalUrl->Value.lpszW);
												if (profileMailStoreExternalUrl) MAPIFreeBuffer(profileMailStoreExternalUrl);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszMailStoreExternalUrl = std::wstring(L" ");
											}
										}

										// PR_RESOURCE_FLAGS
										LPSPropValue resourceFlags = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_RESOURCE_FLAGS, &resourceFlags)))
										{
											if (resourceFlags)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulResourceFlags = resourceFlags->Value.l;
												if (resourceFlags) MAPIFreeBuffer(resourceFlags);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulResourceFlags = 0;
											}
										}

										// PR_PROFILE_RPC_PROXY_SERVER_AUTH_PACKAGE
										LPSPropValue rohAuthPackage = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_RPC_PROXY_SERVER_AUTH_PACKAGE, &rohAuthPackage)))
										{
											if (rohAuthPackage)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulRohProxyAuthScheme = rohAuthPackage->Value.l;
												if (rohAuthPackage) MAPIFreeBuffer(rohAuthPackage);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulRohProxyAuthScheme = 0;
											}
										}

										// PR_ROH_FLAGS
										LPSPropValue rohFlags = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_ROH_FLAGS, &rohFlags)))
										{
											if (rohFlags)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulRohFlags = rohFlags->Value.l;
												if (rohFlags) MAPIFreeBuffer(rohFlags);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulRohFlags = 0;
											}
										}

										// PR_SERVICE_UID
										LPSPropValue serviceUid = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_SERVICE_UID, &serviceUid)))
										{
											if (serviceUid)
											{
												LPMAPIUID lpMuidServiceUid = NULL;
												lpMuidServiceUid = &profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].muidServiceUid;
												memcpy(lpMuidServiceUid, (LPMAPIUID)serviceUid->Value.bin.lpb, sizeof(MAPIUID));
												if (serviceUid) MAPIFreeBuffer(serviceUid);
											}
										}
										
										// PR_PROFILE_ALTERNATE_STORE_TYPE
										LPSPropValue alternateStoreType = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_ALTERNATE_STORE_TYPE, &alternateStoreType)))
										{
											if (alternateStoreType)
											{
												if (ConvertWideCharToStdWstring(alternateStoreType->Value.lpszW) == L"Archive")
												{
													profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].bIsOnlineArchive = true;
												}
												else
												{
													profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].bIsOnlineArchive = false;
												}
												if (alternateStoreType) MAPIFreeBuffer(alternateStoreType);
											}
										}

										// PR_PROVIDER_UID
										LPMAPIUID lpMuidProviderUid = NULL;
										lpMuidProviderUid = &profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].muidProviderUid;
										memcpy(lpMuidProviderUid, (LPMAPIUID)lpProvRows->aRow[j].lpProps[iProvInstanceKey].Value.bin.lpb, sizeof(MAPIUID));

										//LPSPropValue providerUid = NULL;
										//hRes = HrGetOneProp(lpMAPIProp, PR_INSTANCE_KEY, &providerUid);
										//if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROVIDER_UID, &providerUid)))
										//{
										//	if (providerUid)
										//	{
										//		LPMAPIUID lpMuidProviderUid = NULL;
										//		lpMuidProviderUid = &profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].muidProviderUid;
										//		memcpy(lpMuidProviderUid, (LPMAPIUID)providerUid->Value.bin.lpb, sizeof(MAPIUID));
										//		if (providerUid) MAPIFreeBuffer(providerUid);
										//	}
										//}


									}
								}
							}
							if (lpProvRows) FreeProws(lpProvRows);
						}
						else
						{
							profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount = lpProvRows->cRows;
							profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes = new MailboxInfo();
						}
						if (lpProvPropVal) MAPIFreeBuffer(lpProvPropVal);
						if (lpProvResLvl1) MAPIFreeBuffer(lpProvResLvl1);
						if (lpProvRes) MAPIFreeBuffer(lpProvRes);
						if (lpProvTable) lpProvTable->Release();
						//End Loop Providers
						if (lpProvAdmin) lpProvAdmin->Release();
					}

				}

				else if ((0 == strcmp(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA, "MSPST MS")) || (0 == strcmp(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA, "MSUPST MS")))
				{
					profileInfo->profileServices[i].pstInfo = new PstInfo();
					profileInfo->profileServices[i].pstInfo->wszDisplayName = std::wstring(L" ");
					profileInfo->profileServices[i].pstInfo->wszPstPath = std::wstring(L" ");
					profileInfo->profileServices[i].pstInfo->ulPstConfigFlags = 0;
					profileInfo->profileServices[i].ulServiceType = SERVICETYPE_PST;

					LPPROVIDERADMIN lpProvAdmin = NULL;

					if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb,
						0,
						&lpProvAdmin)))
					{
						// Loop providers
						LPMAPITABLE lpProvTable = NULL;
						LPSRestriction lpProvRes = NULL;
						LPSRestriction lpProvResLvl1 = NULL;
						LPSPropValue lpProvPropVal = NULL;
						LPSRowSet lpProvRows = NULL;

						// Setting up an enum and a prop tag array with the props we'll use
						enum { iProvInstanceKey, cptaProvProps };
						SizedSPropTagArray(cptaProvProps, sptaProvProps) = { cptaProvProps, PR_INSTANCE_KEY };

						// Allocate memory for the restriction
						EC_HRES_MSG(MAPIAllocateBuffer(
							sizeof(SRestriction),
							(LPVOID*)&lpProvRes), L"Calling MAPIAllocateBuffer");

						EC_HRES_MSG(MAPIAllocateBuffer(
							sizeof(SRestriction) * 2,
							(LPVOID*)&lpProvResLvl1), L"Calling MAPIAllocateBuffer");

						EC_HRES_MSG(MAPIAllocateBuffer(
							sizeof(SPropValue),
							(LPVOID*)&lpProvPropVal), L"Calling MAPIAllocateBuffer");

						// Set up restriction to query the provider table
						lpProvRes->rt = RES_AND;
						lpProvRes->res.resAnd.cRes = 0x00000002;
						lpProvRes->res.resAnd.lpRes = lpProvResLvl1;

						lpProvResLvl1[0].rt = RES_EXIST;
						lpProvResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_UID;
						lpProvResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
						lpProvResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
						lpProvResLvl1[1].rt = RES_PROPERTY;
						lpProvResLvl1[1].res.resProperty.relop = RELOP_EQ;
						lpProvResLvl1[1].res.resProperty.ulPropTag = PR_SERVICE_UID;
						lpProvResLvl1[1].res.resProperty.lpProp = lpProvPropVal;

						lpProvPropVal->ulPropTag = PR_SERVICE_UID;
						lpProvPropVal->Value = lpSvcRows->aRow[i].lpProps[iServiceUid].Value;

						lpProvAdmin->GetProviderTable(0,
							&lpProvTable);
						// Query the table to get the the default profile only
						EC_HRES_MSG(HrQueryAllRows(lpProvTable,
							(LPSPropTagArray)&sptaProvProps,
							lpProvRes,
							NULL,
							0,
							&lpProvRows), L"HrGetProfile");

						if (lpProvRows->cRows > 0)
						{

							LPPROFSECT lpProfSection = NULL;
							if (SUCCEEDED(lpServiceAdmin->OpenProfileSection((LPMAPIUID)lpProvRows->aRow->lpProps[iProvInstanceKey].Value.bin.lpb, NULL, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSection)))
							{
								LPMAPIPROP lpMAPIProp = NULL;
								if (SUCCEEDED(lpProfSection->QueryInterface(IID_IMAPIProp, (void**)&lpMAPIProp)))
								{
									LPSPropValue prDisplayName = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_DISPLAY_NAME_W, &prDisplayName)))
									{
										profileInfo->profileServices[i].pstInfo->wszDisplayName = ConvertWideCharToStdWstring(prDisplayName->Value.lpszW);
										if (prDisplayName) MAPIFreeBuffer(prDisplayName);
									}
									else
									{
										profileInfo->profileServices[i].pstInfo->wszDisplayName = std::wstring(L" ");
									}
									// bind to the PR_PST_PATH_W property
									LPSPropValue pstPathW = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_PATH_W, &pstPathW)))
									{
										if (pstPathW)
										{
											profileInfo->profileServices[i].pstInfo->wszPstPath = ConvertWideCharToStdWstring(pstPathW->Value.lpszW);
											if (pstPathW) MAPIFreeBuffer(pstPathW);
										}
										else
										{
											profileInfo->profileServices[i].pstInfo->wszPstPath = std::wstring(L" ");
										}
									}
									// bind to the PR_PST_CONFIG_FLAGS property to get the ammount to sync
									LPSPropValue pstConfigFlags = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_CONFIG_FLAGS, &pstConfigFlags)))
									{
										if (pstConfigFlags)
										{
											profileInfo->profileServices[i].pstInfo->ulPstConfigFlags = pstConfigFlags->Value.l;
											if (pstConfigFlags) MAPIFreeBuffer(pstConfigFlags);
										}
									}
								}
							}

							if (lpProvRows) FreeProws(lpProvRows);
						}
						if (lpProvPropVal) MAPIFreeBuffer(lpProvPropVal);
						if (lpProvResLvl1) MAPIFreeBuffer(lpProvResLvl1);
						if (lpProvRes) MAPIFreeBuffer(lpProvRes);
						if (lpProvTable) lpProvTable->Release();
						//End Loop Providers
						if (lpProvAdmin) lpProvAdmin->Release();
					}

				}

			}
			if (lpSvcRows) FreeProws(lpSvcRows);
			// End loop services


		}

		if (lpSvcPropVal) MAPIFreeBuffer(lpSvcPropVal);
		if (lpsvcResLvl1) MAPIFreeBuffer(lpsvcResLvl1);
		if (lpSvcRes) MAPIFreeBuffer(lpSvcRes);
		if (lpServiceTable) lpServiceTable->Release();
		if (lpServiceAdmin) lpServiceAdmin->Release();

	}
	// End process services

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	if (lpProfRows) FreeProws(lpProfRows);
	if (lpProfTable) lpProfTable->Release();
	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;
}

// HrDeleteProvider
// Deletes the provider with the specified UID from the service with the specified UID in a given profile
HRESULT HrDeleteProvider(LPWSTR lpwszProfileName, LPMAPIUID lpServiceUid, LPMAPIUID lpProviderUid)
{
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPPROVIDERADMIN lpProviderAdmin = NULL;

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin

	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling AdminServices");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		EC_HRES_MSG(lpServiceAdmin->AdminProviders(lpServiceUid, NULL, &lpProviderAdmin), L"Calling AdminProviders");
		if (lpProviderAdmin)
		{
			EC_HRES_MSG(lpProviderAdmin->DeleteProvider(lpProviderUid), L"Calling DeleteProvider");
		}
	}

Error:
	return hRes;
}

// HrGetSections
// Returns the EMSMDB and StoreProvider sections of a service
HRESULT HrGetSections(LPSERVICEADMIN2 lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT* lppEmsMdbSection, LPPROFSECT* lppStoreProviderSection)
{
	HRESULT hRes = S_OK;
	SizedSPropTagArray(2, sptaUids) = { 2,{ PR_STORE_PROVIDERS, PR_EMSMDB_SECTION_UID } };
	ULONG				cValues = 0;
	LPSPropValue		lpProps = nullptr;
	LPPROFSECT			lpSvcProfSect = nullptr;
	LPPROFSECT			lpEmsMdbProfSect = nullptr;
	LPPROFSECT			lpStoreProvProfSect = nullptr;

	if (!lpSvcAdmin || !lpServiceUid || !lppStoreProviderSection)
		return E_INVALIDARG;

	if (NULL != lppStoreProviderSection)
	{
		*lppStoreProviderSection = nullptr;
	}
	if (NULL != lppEmsMdbSection)
	{
		*lppEmsMdbSection = nullptr;
	}

	EC_HRES_MSG(lpSvcAdmin->OpenProfileSection(lpServiceUid, NULL, MAPI_FORCE_ACCESS | MAPI_MODIFY, &lpSvcProfSect), L"Calling OpenProfileSection.");

	EC_HRES_MSG(lpSvcProfSect->GetProps(
		(LPSPropTagArray)& sptaUids,
		0,
		&cValues,
		&lpProps), L"Calling GetProps.");

	if (cValues != 2)
		return E_FAIL;


	if (lpProps[0].ulPropTag != sptaUids.aulPropTag[0])
		EC_HRES_MSG(lpProps[0].Value.err, L"Cheking Value.err");
	if (NULL != lppEmsMdbSection)
	{
		if (lpProps[1].ulPropTag != sptaUids.aulPropTag[1])
			EC_HRES_MSG(lpProps[1].Value.err, L"Cheking Value.err");
	}

	if (NULL != lppStoreProviderSection)
	{
		EC_HRES_MSG(lpSvcAdmin->OpenProfileSection(
			(LPMAPIUID)lpProps[0].Value.bin.lpb,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpStoreProvProfSect), L"Calling OpenProfileSection.");
	}

	if (NULL != lppEmsMdbSection)
	{
		EC_HRES_MSG(lpSvcAdmin->OpenProfileSection(
			(LPMAPIUID)lpProps[1].Value.bin.lpb,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpEmsMdbProfSect), L"Calling OpenProfileSection.");
	}

	if (NULL != lppEmsMdbSection)
		* lppEmsMdbSection = lpEmsMdbProfSect;

	if (NULL != lppStoreProviderSection)
		* lppStoreProviderSection = lpStoreProvProfSect;
Error:
	goto Cleanup;

Cleanup:
	if (lpSvcProfSect) lpSvcProfSect->Release();
	if (lpProps)
	{
		MAPIFreeBuffer(lpProps);
		lpProps = nullptr;
	}
	return hRes;
}

// HrGetSections
// Returns the EMSMDB and StoreProvider sections of a service
HRESULT HrGetSections(LPSERVICEADMIN lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT * lppEmsMdbSection, LPPROFSECT * lppStoreProviderSection)
{
	HRESULT hRes = S_OK;
	SizedSPropTagArray(2, sptaUids) = { 2,{ PR_STORE_PROVIDERS, PR_EMSMDB_SECTION_UID } };
	ULONG				cValues = 0;
	LPSPropValue		lpProps = nullptr;
	LPPROFSECT			lpSvcProfSect = nullptr;
	LPPROFSECT			lpEmsMdbProfSect = nullptr;
	LPPROFSECT			lpStoreProvProfSect = nullptr;

	if (!lpSvcAdmin || !lpServiceUid || !lppStoreProviderSection)
		return E_INVALIDARG;

	if (NULL != lppStoreProviderSection)
	{
		*lppStoreProviderSection = nullptr;
	}
	if (NULL != lppEmsMdbSection)
	{
		*lppEmsMdbSection = nullptr;
	}

	EC_HRES_MSG(lpSvcAdmin->OpenProfileSection(lpServiceUid, NULL, MAPI_FORCE_ACCESS | MAPI_MODIFY, &lpSvcProfSect), L"Calling OpenProfileSection.");

	EC_HRES_MSG(lpSvcProfSect->GetProps(
		(LPSPropTagArray)& sptaUids,
		0,
		&cValues,
		&lpProps), L"Calling GetProps.");

	if (cValues != 2)
		return E_FAIL;


	if (lpProps[0].ulPropTag != sptaUids.aulPropTag[0])
		EC_HRES_MSG(lpProps[0].Value.err, L"Cheking Value.err");
	if (NULL != lppEmsMdbSection)
	{
		if (lpProps[1].ulPropTag != sptaUids.aulPropTag[1])
			EC_HRES_MSG(lpProps[1].Value.err, L"Cheking Value.err");
	}

	if (NULL != lpStoreProvProfSect)
	{
		EC_HRES_MSG(lpSvcAdmin->OpenProfileSection(
			(LPMAPIUID)lpProps[0].Value.bin.lpb,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpStoreProvProfSect), L"Calling OpenProfileSection.");
	}

	if (NULL != lppEmsMdbSection)
	{
		EC_HRES_MSG(lpSvcAdmin->OpenProfileSection(
			(LPMAPIUID)lpProps[1].Value.bin.lpb,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpEmsMdbProfSect), L"Calling OpenProfileSection.");
	}

	if (NULL != lppEmsMdbSection)
		* lppEmsMdbSection = lpEmsMdbProfSect;

	if (NULL != lppStoreProviderSection)
		* lppStoreProviderSection = lpStoreProvProfSect;
Error:
	goto Cleanup;

Cleanup:
	if (lpSvcProfSect) lpSvcProfSect->Release();
	if (lpProps)
	{
		MAPIFreeBuffer(lpProps);
		lpProps = nullptr;
	}
	return hRes;
}