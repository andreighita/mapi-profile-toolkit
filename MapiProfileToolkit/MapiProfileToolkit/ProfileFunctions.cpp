/*
* © 2015 Microsoft Corporation
*
* written by Andrei Ghita
*
* Microsoft provides programming examples for illustration only, without warranty either expressed or implied.
* This includes, but is not limited to, the implied warranties of merchantability or fitness for a particular purpose.
* This article assumes that you are familiar with the programming language that is being demonstrated and with
* the tools that are used to create and to debug procedures. Microsoft support engineers can help explain the
* functionality of a particular procedure, but they will not modify these examples to provide added functionality
* or construct procedures to meet your specific requirements.
*/

#include "stdafx.h"
#include "ProfileFunctions.h"
#include "StringOperations.h"
#include "ToolkitObjects.h"

#define MAPI_FORCE_ACCESS 0x00080000
#define PR_EMSMDB_SECTION_UID					PROP_TAG(PT_BINARY, 0x3D15)
#define PR_PROFILE_USER_SMTP_EMAIL_ADDRESS		PROP_TAG(PT_STRING8, pidProfileMin+0x41)
#define PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W	PROP_TAG(PT_UNICODE, pidProfileMin+0x41)
#define PR_ROH_PROXY_SERVER						PROP_TAG(PT_UNICODE, pidProfileMin+0x22)
#define PR_PROFILE_RPC_PROXY_SERVER_FLAGS		PROP_TAG(PT_LONG,	pidProfileMin+0x23)
#define PR_ROH_PROXY_PRINCIPAL_NAME				PROP_TAG( PT_UNICODE, pidProfileMin+0x25)
#define PR_PROFILE_RPC_PROXY_SERVER_AUTH_PACKAGE	PROP_TAG(PT_LONG,	pidProfileMin+0x27)
#define PR_PROFILE_RPC_PROXY_SERVER_W			PROP_TAG( PT_UNICODE, pidProfileMin+0x22)
#define PR_PROFILE_HOME_SERVER_FQDN				PROP_TAG(PT_UNICODE, pidProfileMin+0x2A)
#define	PR_PROFILE_SERVER_FQDN_W				PROP_TAG( PT_UNICODE, pidProfileMin+0x2b)
#define PR_PROFILE_ACCT_NAME					PROP_TAG( PT_STRING8, pidProfileMin+0x20)  
#define PR_PROFILE_ACCT_NAME_W					PROP_TAG( PT_UNICODE, pidProfileMin+0x20) 
#define PR_PROFILE_USER_EMAIL_W					PROP_TAG(PT_UNICODE, pidProfileMin+0x3d) 

#define	PR_PROFILE_UNRESOLVED_NAME_W			PROP_TAG( PT_UNICODE, pidProfileMin+0x07)  
#define PR_PROFILE_OFFLINE_STORE_PATH_W	PROP_TAG( PT_UNICODE, pidProfileMin+0x10) 
#define PR_PROFILE_LKG_AUTODISCOVER_URL			PROP_TAG(PT_UNICODE, pidProfileMin+0x4A)

#define PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL PROP_TAG(PT_UNICODE, pidProfileMin+0x52)
#define PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL PROP_TAG(PT_UNICODE, pidProfileMin+0x53)
#define PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL PROP_TAG(PT_UNICODE, pidProfileMin+0x54)
#define PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL PROP_TAG(PT_UNICODE, pidProfileMin+0x55)

#define PR_PROFILE_ALTERNATE_STORE_TYPE PROP_TAG(PT_UNICODE, 0x65D0)
#ifndef CONFIG_OST_CACHE_PRIVATE
#define CONFIG_OST_CACHE_PRIVATE			((ULONG)0x00000180)
#endif
#ifndef CONFIG_OST_CACHE_DELEGATE_PIM
#define CONFIG_OST_CACHE_DELEGATE_PIM		((ULONG)0x00000800)
#endif
#ifndef CONFIG_OST_CACHE_PUBLIC
#define CONFIG_OST_CACHE_PUBLIC				((ULONG)0x00000400)
#endif

#pragma region // Profile Methods //

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



HRESULT HrSetCachedModeOneService(LPSTR lpszProfileName, LPMAPIUID lpServiceUid, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths)
{
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin
									 // Get an IProfAdmin interface.

									 // Begin process services
	LPSERVICEADMIN lpServiceAdmin = NULL;
	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling AdminServices");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		LPPROVIDERADMIN lpProvAdmin = NULL;
		LPPROFSECT lpEmsMdbProfSect, lpStoreProviderProfSect = NULL;
		if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpServiceUid,
			0,
			&lpProvAdmin)))
		{
			EC_HRES_MSG(HrGetSections(lpServiceAdmin, lpServiceUid, &lpEmsMdbProfSect, &lpStoreProviderProfSect), L"Calling HrGetSections");
			// Access the EMSMDB section
			if (lpEmsMdbProfSect)
			{
				LPMAPIPROP pMAPIProp = NULL;
				if (SUCCEEDED(lpEmsMdbProfSect->QueryInterface(IID_IMAPIProp, (void**)&pMAPIProp)))
				{
					// bind to the PR_PROFILE_CONFIG_FLAGS property
					LPSPropValue profileConfigFlags = NULL;
					if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_CONFIG_FLAGS, &profileConfigFlags)))
					{
						if (profileConfigFlags)
						{
							if (bCachedModeOwner)
							{
								if (!(profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PRIVATE))
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PRIVATE;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									printf("Cached mode owner enabled.\n");
								}
								else
								{
									printf("Cached mode owner already enabled on service.\n");
								}
							}
							else
							{
								if (profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PRIVATE)
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PRIVATE;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									printf("Cached mode owner disabled.\n");
								}
								else
								{
									printf("Cached mode owner already disabled on service.\n");
								}
							}


							if (bCachedModeShared)
							{
								if (!(profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_DELEGATE_PIM))
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_DELEGATE_PIM;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									printf("Cached mode shared enabled.\n");
								}
								else
								{
									printf("Cached mode shared already enabled on service.\n");
								}
							}
							else
							{
								if (profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_DELEGATE_PIM)
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_DELEGATE_PIM;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									printf("Cached mode shared disabled.\n");
								}
								else
								{
									printf("Cached mode shared already disabled on service.\n");
								}
							}


							if (bCachedModePublicFolders)
							{
								if (!(profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PUBLIC))
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PUBLIC;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									printf("Cached mode public folders enabled.\n");
								}
								else
								{
									printf("Cached mode public folders already enabled on service.\n");
								}
							}
							else
							{
								if (profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PUBLIC)
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PUBLIC;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									printf("Cached mode public folders disabled.\n");
								}
								else
								{
									printf("Cached mode public folders already disabled on service.\n");
								}
							}

							EC_HRES_MSG(lpEmsMdbProfSect->SaveChanges(0), L"Calling #");
							if (profileConfigFlags) MAPIFreeBuffer(profileConfigFlags);
						}
					}
					// bind to the PR_RULE_ACTION_TYPE property for setting the amout of mail to cache
					LPSPropValue profileRuleActionType = NULL;
					if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_RULE_ACTION_TYPE, &profileRuleActionType)))
					{
						if (profileRuleActionType)
						{

							profileRuleActionType[0].Value.i = iCachedModeMonths;
							EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileRuleActionType, NULL), L"Calling SetProps");
							printf("Cached mode amount to sync set.\n");

							EC_HRES_MSG(lpEmsMdbProfSect->SaveChanges(0), L"Calling SaveChanges");
							if (profileRuleActionType) MAPIFreeBuffer(profileConfigFlags);
						}
					}
				}
				if (lpEmsMdbProfSect) lpEmsMdbProfSect->Release();
			}

			if (lpProvAdmin) lpProvAdmin->Release();

		}

		if (lpServiceAdmin) lpServiceAdmin->Release();

	}
	// End process services

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;
	return S_OK;
}




HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszOldPath, LPWSTR lpszNewPath, bool bMoveFiles)
{
	HRESULT hRes = S_OK;

	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPSRestriction lpProfRes = NULL;
	LPSRestriction lpProfResLvl1 = NULL;
	LPSPropValue lpProfPropVal = NULL;
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME };

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin
									 // Get an IProfAdmin interface.

	EC_HRES_MSG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), L"Calling GetProfileTable");

	// Allocate memory for the restriction
	EC_HRES_MSG(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)&lpProfRes), L"Calling MAPIAllocateBuffer");

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
		&lpProfRows), L"Calling #");

	if (lpProfRows->cRows == 0)
	{
		return MAPI_E_NOT_FOUND;
	}
	else if (lpProfRows->cRows != 1)
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
		&lpServiceAdmin), L"Calling #");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		lpServiceAdmin->GetMsgServiceTable(0,
			&lpServiceTable);
		LPSRestriction lpSvcRes = NULL;
		LPSRestriction lpsvcResLvl1 = NULL;
		LPSRestriction lpsvcResLvl2 = NULL;
		LPSPropValue lpSvcPropVal1 = NULL;
		LPSPropValue lpSvcPropVal2 = NULL;
		LPSRowSet lpSvcRows = NULL;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceName, iEmsMdbSectUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID,PR_SERVICE_NAME_A, PR_EMSMDB_SECTION_UID, PR_RESOURCE_FLAGS };

		// Allocate memory for the restriction
		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpSvcRes), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl2), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal2), L"Calling MAPIAllocateBuffer");

		// Set up restriction to query the profile table
		lpSvcRes->rt = RES_AND;
		lpSvcRes->res.resAnd.cRes = 0x00000002;
		lpSvcRes->res.resAnd.lpRes = lpsvcResLvl1;

		lpsvcResLvl1[0].rt = RES_EXIST;
		lpsvcResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
		lpsvcResLvl1[0].res.resExist.ulReserved2 = 0x00000000;

		lpsvcResLvl1[1].rt = RES_OR;
		lpsvcResLvl1[1].res.resOr.cRes = 0x00000002;
		lpsvcResLvl1[1].res.resOr.lpRes = lpsvcResLvl2;

		lpsvcResLvl2[0].rt = RES_CONTENT;
		lpsvcResLvl2[0].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
		lpsvcResLvl2[0].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl2[0].res.resContent.lpProp = lpSvcPropVal1;

		lpSvcPropVal1->ulPropTag = PR_SERVICE_NAME_A;
		lpSvcPropVal1->Value.lpszA = "MSPST MS";

		lpsvcResLvl2[1].rt = RES_CONTENT;
		lpsvcResLvl2[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
		lpsvcResLvl2[1].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl2[1].res.resContent.lpProp = lpSvcPropVal2;

		lpSvcPropVal2->ulPropTag = PR_SERVICE_NAME_A;
		lpSvcPropVal2->Value.lpszA = "MSUPST MS";

		// Query the table to get the the default profile only
		EC_HRES_MSG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			lpSvcRes,
			NULL,
			0,
			&lpSvcRows), L"Calling HrQueryAllRows");

		if (lpSvcRows->cRows > 0)
		{
			wprintf(L"Found %i PST services in profile %s\n", lpSvcRows->cRows, lpszProfileName);
			// Start loop services
			for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
			{

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
						&lpProvRows), L"Calling HrQueryAllRows");

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
									// bind to the PR_PST_PATH_W property
									LPSPropValue pstPathW = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_PATH_W, &pstPathW)))
									{
										if (pstPathW)
										{
											std::wstring szCurrentPath = ConvertWideCharToStdWstring(pstPathW->Value.lpszW);
											if (WStringReplace(&szCurrentPath, ConvertWideCharToStdWstring(lpszOldPath), ConvertWideCharToStdWstring(lpszNewPath)))
											{
												if (bMoveFiles)
												{
													wprintf(L"Moving file %s to new location %s\n", pstPathW->Value.lpszW, szCurrentPath.c_str());
													BOOL bFileMoved = false;
													bFileMoved = MoveFileExW(pstPathW->Value.lpszW, (LPCWSTR)szCurrentPath.c_str(), MOVEFILE_COPY_ALLOWED | MOVEFILE_WRITE_THROUGH);
													if (bFileMoved)
													{
														wprintf(L"Updating path for data file named %s\n", pstPathW->Value.lpszW);
														pstPathW[0].Value.lpszW = (LPWSTR)szCurrentPath.c_str();
														EC_HRES_MSG(lpProfSection->SetProps(1, pstPathW, NULL), L"Calling SetProps");
													}
													else
													{
														wprintf(L"Unable to move file\n");
													}
												}
												else
												{
													wprintf(L"Updating path for data file named %s\n", pstPathW->Value.lpszW);
													pstPathW[0].Value.lpszW = (LPWSTR)szCurrentPath.c_str();
													EC_HRES_MSG(lpProfSection->SetProps(1, pstPathW, NULL), L"Calling SetProps");
												}
											}
											if (pstPathW) MAPIFreeBuffer(pstPathW);
										}
									}
									if (prDisplayName) MAPIFreeBuffer(prDisplayName);
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
			if (lpSvcRows) FreeProws(lpSvcRows);
			// End loop services
		}
		if (lpSvcPropVal1) MAPIFreeBuffer(lpSvcPropVal1);
		if (lpSvcPropVal2) MAPIFreeBuffer(lpSvcPropVal2);
		if (lpsvcResLvl1) MAPIFreeBuffer(lpsvcResLvl1);
		if (lpsvcResLvl2) MAPIFreeBuffer(lpsvcResLvl2);
		if (lpSvcRes) MAPIFreeBuffer(lpSvcRes);
		if (lpServiceTable) lpServiceTable->Release();
		if (lpServiceAdmin) lpServiceAdmin->Release();
	}
	// End process services

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	//if (lpProfPropVal) MAPIFreeBuffer(lpProfPropVal);
	//if (lpProfResLvl1) MAPIFreeBuffer(lpProfResLvl1);
	//if (lpProfRes) MAPIFreeBuffer(lpProfRes);
	if (lpProfRows) FreeProws(lpProfRows);
	if (lpProfTable) lpProfTable->Release();
	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;
}

HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszNewPath, bool bMoveFiles)
{
	HRESULT hRes = S_OK;

	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPSRestriction lpProfRes = NULL;
	LPSRestriction lpProfResLvl1 = NULL;
	LPSPropValue lpProfPropVal = NULL;
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME };

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin
									 // Get an IProfAdmin interface.

	EC_HRES_MSG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), L"Calling GetProfileTable");

	// Allocate memory for the restriction
	EC_HRES_MSG(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)&lpProfRes), L"Calling MAPIAllocateBuffer");

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
		&lpProfRows), L"Calling HrQueryAllRows");

	if (lpProfRows->cRows == 0)
	{
		return MAPI_E_NOT_FOUND;
	}
	else if (lpProfRows->cRows != 1)
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
		&lpServiceAdmin), L"Calling AdminServices");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		lpServiceAdmin->GetMsgServiceTable(0,
			&lpServiceTable);
		LPSRestriction lpSvcRes = NULL;
		LPSRestriction lpsvcResLvl1 = NULL;
		LPSRestriction lpsvcResLvl2 = NULL;
		LPSPropValue lpSvcPropVal1 = NULL;
		LPSPropValue lpSvcPropVal2 = NULL;
		LPSRowSet lpSvcRows = NULL;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceName, iEmsMdbSectUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID,PR_SERVICE_NAME_A, PR_EMSMDB_SECTION_UID, PR_RESOURCE_FLAGS };

		// Allocate memory for the restriction
		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpSvcRes), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl2), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal2), L"Calling MAPIAllocateBuffer");

		// Set up restriction to query the profile table
		lpSvcRes->rt = RES_AND;
		lpSvcRes->res.resAnd.cRes = 0x00000002;
		lpSvcRes->res.resAnd.lpRes = lpsvcResLvl1;

		lpsvcResLvl1[0].rt = RES_EXIST;
		lpsvcResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
		lpsvcResLvl1[0].res.resExist.ulReserved2 = 0x00000000;

		lpsvcResLvl1[1].rt = RES_OR;
		lpsvcResLvl1[1].res.resOr.cRes = 0x00000002;
		lpsvcResLvl1[1].res.resOr.lpRes = lpsvcResLvl2;

		lpsvcResLvl2[0].rt = RES_CONTENT;
		lpsvcResLvl2[0].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
		lpsvcResLvl2[0].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl2[0].res.resContent.lpProp = lpSvcPropVal1;

		lpSvcPropVal1->ulPropTag = PR_SERVICE_NAME_A;
		lpSvcPropVal1->Value.lpszA = "MSPST MS";

		lpsvcResLvl2[1].rt = RES_CONTENT;
		lpsvcResLvl2[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
		lpsvcResLvl2[1].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl2[1].res.resContent.lpProp = lpSvcPropVal2;

		lpSvcPropVal2->ulPropTag = PR_SERVICE_NAME_A;
		lpSvcPropVal2->Value.lpszA = "MSUPST MS";

		// Query the table to get the the default profile only
		EC_HRES_MSG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			lpSvcRes,
			NULL,
			0,
			&lpSvcRows), L"Calling HrQueryAllRows");

		if (lpSvcRows->cRows > 0)
		{
			wprintf(L"Found %i PST services in profile %s\n", lpSvcRows->cRows, lpszProfileName);
			// Start loop services
			for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
			{

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
						&lpProvRows), L"Calling HrQueryAllRows");

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
									// bind to the PR_PST_PATH_W property
									LPSPropValue pstPathW = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_PATH_W, &pstPathW)))
									{
										if (pstPathW)
										{
											std::wstring szCurrentPath = ConvertWideCharToStdWstring(pstPathW->Value.lpszW);
											std::wstring szOldPath = szCurrentPath;
											LPWSTR lpszOldPath = (LPWSTR)szOldPath.c_str();
											if SUCCEEDED(PathRemoveFileSpec(lpszOldPath))
											{
												if (WStringReplace(&szCurrentPath, lpszOldPath, ConvertWideCharToStdWstring(lpszNewPath)))
												{
													if (bMoveFiles)
													{
														wprintf(L"Moving file %s to new location %s\n", pstPathW->Value.lpszW, szCurrentPath.c_str());
														BOOL bFileMoved = false;
														bFileMoved = MoveFileExW(pstPathW->Value.lpszW, (LPCWSTR)szCurrentPath.c_str(), MOVEFILE_COPY_ALLOWED | MOVEFILE_WRITE_THROUGH);
														if (bFileMoved)
														{
															wprintf(L"Updating path for data file named %s\n", pstPathW->Value.lpszW);
															pstPathW[0].Value.lpszW = (LPWSTR)szCurrentPath.c_str();
															EC_HRES_MSG(lpProfSection->SetProps(1, pstPathW, NULL), L"Calling SetProps");
														}
														else
														{
															wprintf(L"Unable to move file\n");
														}
													}
													else
													{
														wprintf(L"Updating path for data file named %s\n", pstPathW->Value.lpszW);
														pstPathW[0].Value.lpszW = (LPWSTR)szCurrentPath.c_str();
														EC_HRES_MSG(lpProfSection->SetProps(1, pstPathW, NULL), L"Calling SetProps");
													}
												}
											}
											if (pstPathW) MAPIFreeBuffer(pstPathW);
										}
									}
									if (prDisplayName) MAPIFreeBuffer(prDisplayName);
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
			if (lpSvcRows) FreeProws(lpSvcRows);
			// End loop services
		}
		if (lpSvcPropVal1) MAPIFreeBuffer(lpSvcPropVal1);
		if (lpSvcPropVal2) MAPIFreeBuffer(lpSvcPropVal2);
		if (lpsvcResLvl1) MAPIFreeBuffer(lpsvcResLvl1);
		if (lpsvcResLvl2) MAPIFreeBuffer(lpsvcResLvl2);
		if (lpSvcRes) MAPIFreeBuffer(lpSvcRes);
		if (lpServiceTable) lpServiceTable->Release();
		if (lpServiceAdmin) lpServiceAdmin->Release();
	}
	// End process services

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	//if (lpProfPropVal) MAPIFreeBuffer(lpProfPropVal);
	//if (lpProfResLvl1) MAPIFreeBuffer(lpProfResLvl1);
	//if (lpProfRes) MAPIFreeBuffer(lpProfRes);
	if (lpProfRows) FreeProws(lpProfRows);
	if (lpProfTable) lpProfTable->Release();
	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;
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
	for (unsigned int i = 0; i < profileInfo->ulServiceCount; i++)
	{
		MAPIUID uidService = { 0 };
		LPMAPIUID lpServiceUid = &uidService;
		if (profileInfo->profileServices[i].ulServiceType == SERVICETYPE_MAILBOX)
		{
			Logger::Write(logLevelInfo, L"Adding exchange mailbox: " + profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress);
			EC_HRES_MSG(HrCreateMsemsServiceModernExt(TRUE,
				(LPWSTR)GetDefaultProfileName().c_str(),
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
					EC_HRES_MSG(HrAddDelegateMailboxModern(TRUE,
						(LPWSTR)GetDefaultProfileName().c_str(),
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
										profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUserSmtpEmailAddress->Value.lpszA));
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
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServer = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileServerDN->Value.lpszA));
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
							&lpProvRows), L"HrGetProfile", L"0044");

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


#pragma endregion

#pragma region // PST Methods //

HRESULT HrCreatePstService(LPSERVICEADMIN2 lpServiceAdmin2, LPMAPIUID * lppServiceUid, LPWSTR lpszServiceName, ULONG ulResourceFlags, ULONG ulPstConfigFlag, LPWSTR lpszPstPathW, LPWSTR lpszDisplayName)
{
	HRESULT			hRes = S_OK; // Result code returned from MAPI calls.
	SPropValue		rgvalStoreProvider[3];
	MAPIUID			uidService = { 0 };
	LPMAPIUID		lpServiceUid = &uidService;
	LPPROFSECT		lpProfSect = NULL;
	LPPROFSECT		lpStoreProviderSect = nullptr;

	// Adds a message service to the current profile and returns that newly added service UID.
	EC_HRES_MSG(lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)ConvertWideCharToMultiByte(lpszServiceName),
		(LPTSTR)ConvertWideCharToMultiByte(lpszDisplayName),
		NULL,
		0,
		&uidService), L"Calling CreateMsgServiceEx.");

	EC_HRES_MSG(lpServiceAdmin2->OpenProfileSection(&uidService,
		0,
		MAPI_FORCE_ACCESS | MAPI_MODIFY,
		&lpProfSect), L"Calling OpenProfileSection.");


	LPMAPIPROP lpMapiProp = NULL;
	EC_HRES_MSG(lpProfSect->QueryInterface(IID_IMAPIProp, (LPVOID*)&lpMapiProp), L"Calling QueryInterface.");

	if (lpMapiProp)
	{
		LPSPropValue prResourceFlags;
		MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)&prResourceFlags);

		prResourceFlags->ulPropTag = PR_RESOURCE_FLAGS;
		prResourceFlags->Value.l = ulResourceFlags;
		EC_HRES_MSG(lpMapiProp->SetProps(1, prResourceFlags, NULL), L"Calling SetProps.");

		EC_HRES_MSG(lpMapiProp->SaveChanges(FORCE_SAVE), L"Calling SaveChanges.");
		MAPIFreeBuffer(prResourceFlags);
		lpMapiProp->Release();
	}

	MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)&lpStoreProviderSect);
	ZeroMemory(lpStoreProviderSect, sizeof(LPPROFSECT));

	EC_HRES_MSG(HrGetSections(lpServiceAdmin2, lpServiceUid, NULL, &lpStoreProviderSect), L"Calling HrGetSections.");

	// Set up a SPropValue array for the properties you need to configure.
	/*
	PR_PST_CONFIG_FLAGS
	PR_PST_PATH_W
	PR_DISPLAY_NAME_W
	*/

	ZeroMemory(&rgvalStoreProvider[0], sizeof(SPropValue));
	rgvalStoreProvider[0].ulPropTag = PR_PST_CONFIG_FLAGS;
	rgvalStoreProvider[0].Value.l = ulPstConfigFlag;

	ZeroMemory(&rgvalStoreProvider[1], sizeof(SPropValue));
	rgvalStoreProvider[1].ulPropTag = PR_PST_PATH_W;
	rgvalStoreProvider[1].Value.lpszW = lpszPstPathW;

	ZeroMemory(&rgvalStoreProvider[2], sizeof(SPropValue));
	rgvalStoreProvider[2].ulPropTag = PR_DISPLAY_NAME_W;
	rgvalStoreProvider[2].Value.lpszW = lpszDisplayName;

	EC_HRES_MSG(lpStoreProviderSect->SetProps(
		2,
		rgvalStoreProvider,
		nullptr), L"Calling SetProps.");

	EC_HRES_MSG(lpStoreProviderSect->SaveChanges(KEEP_OPEN_READWRITE), L"Calling SaveChanges.");

	goto Cleanup;
Error:
	goto Cleanup;

Cleanup:
	// Clean up
	if (lpStoreProviderSect) lpStoreProviderSect->Release();
	if (lpProfSect) lpProfSect->Release();
	return hRes;
}

#pragma endregion

#pragma region // Delegate Mailbox Methods //

// HrAddDelegateMailboxModern
// Adds a delegate mailbox to a given service. The property set is one for Outlook 2016 where all is needed is:
// - the SMTP address of the mailbox
// - the Display Name for the mailbox
HRESULT HrAddDelegateMailboxModern(
	BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	BOOL bDefaultService,
	int iServiceIndex,
	LPWSTR lpszwDisplayName,
	LPWSTR lpszwSMTPAddress)
{

	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling #"); // Pointer to new IProfAdmin

									 // Begin process services
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPMAPITABLE lpServiceTable = NULL;

	if (bDefaultProfile)
	{
		lpwszProfileName = (LPWSTR)GetDefaultProfileName().c_str();
	}

	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling AdminServices");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		lpServiceAdmin->GetMsgServiceTable(0,
			&lpServiceTable);
		LPSRestriction lpSvcRes = NULL;
		LPSRestriction lpsvcResLvl1 = NULL;
		LPSPropValue lpSvcPropVal = NULL;
		LPSRowSet lpSvcRows = NULL;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID, PR_RESOURCE_FLAGS };

		// Allocate memory for the restriction
		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpSvcRes), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal), L"Calling MAPIAllocateBuffer");

		// Set up restriction to query the profile table
		lpSvcRes->rt = RES_AND;
		lpSvcRes->res.resAnd.cRes = 0x00000002;
		lpSvcRes->res.resAnd.lpRes = lpsvcResLvl1;

		lpsvcResLvl1[0].rt = RES_EXIST;
		lpsvcResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
		lpsvcResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
		lpsvcResLvl1[1].rt = RES_PROPERTY;
		lpsvcResLvl1[1].res.resProperty.relop = RELOP_EQ;
		lpsvcResLvl1[1].res.resProperty.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[1].res.resProperty.lpProp = lpSvcPropVal;

		lpSvcPropVal->ulPropTag = PR_SERVICE_NAME_A;
		lpSvcPropVal->Value.lpszA = "MSEMS";

		// Query the table to get the the default profile only
		EC_HRES_MSG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			lpSvcRes,
			NULL,
			0,
			&lpSvcRows), L"Calling HrQueryAllRows");

		if (bDefaultService && lpSvcRows->cRows > 0)
		{
			for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
			{
				if (lpSvcRows->aRow[i].lpProps[iServiceResFlags].Value.l & SERVICE_DEFAULT_STORE)
				{
					LPPROVIDERADMIN lpProvAdmin = NULL;
					if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb,
						0,
						&lpProvAdmin)))
					{
						std::wstring wszSmtpAddress = ConvertWideCharToStdWstring(lpszwSMTPAddress);
						wszSmtpAddress = L"SMTP:" + wszSmtpAddress;

						SPropValue		rgval[2]; // Property value structure to hold configuration info.

						ZeroMemory(&rgval[0], sizeof(SPropValue));
						rgval[0].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
						rgval[0].Value.lpszA = ConvertWideCharToMultiByte((LPWSTR)wszSmtpAddress.c_str());;

						ZeroMemory(&rgval[1], sizeof(SPropValue));
						rgval[1].ulPropTag = PR_DISPLAY_NAME_W;
						rgval[1].Value.lpszW = lpszwDisplayName;

						// Create the message service with the above properties.
						EC_HRES_MSG(lpProvAdmin->CreateProvider(LPWSTR("EMSDelegate"),
							2,
							rgval,
							0,
							0,
							(LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb), L"Calling CreateProvider");
						if (FAILED(hRes)) goto Error;
						if (lpProvAdmin) lpProvAdmin->Release();
						break;
					}
				}
			}
			if (lpSvcRows) FreeProws(lpSvcRows);
		}
		else if (lpSvcRows->cRows >= iServiceIndex)
		{
			LPPROVIDERADMIN lpProvAdmin = NULL;
			if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[iServiceIndex].lpProps[iServiceUid].Value.bin.lpb,
				0,
				&lpProvAdmin)))
			{
				std::wstring wszSmtpAddress = ConvertWideCharToStdWstring(lpszwSMTPAddress);
				wszSmtpAddress = L"SMTP:" + wszSmtpAddress;

				SPropValue		rgval[2]; // Property value structure to hold configuration info.

				ZeroMemory(&rgval[0], sizeof(SPropValue));
				rgval[0].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
				rgval[0].Value.lpszA = ConvertWideCharToMultiByte((LPWSTR)wszSmtpAddress.c_str());;

				ZeroMemory(&rgval[1], sizeof(SPropValue));
				rgval[1].ulPropTag = PR_DISPLAY_NAME_W;
				rgval[1].Value.lpszW = lpszwDisplayName;

				// Create the message service with the above properties.
				EC_HRES_MSG(lpProvAdmin->CreateProvider(LPWSTR("EMSDelegate"),
					2,
					rgval,
					0,
					0,
					(LPMAPIUID)lpSvcRows->aRow[iServiceIndex].lpProps[iServiceUid].Value.bin.lpb), L"Calling CreateProvider");
				if (FAILED(hRes)) goto Error;
				if (lpProvAdmin) lpProvAdmin->Release();
			}
		}
		else
		{
			wprintf(L"No service found.\n");
		}

		if (lpServiceTable) lpServiceTable->Release();
		if (lpServiceAdmin) lpServiceAdmin->Release();

	}
	// End process services

Error:
	goto Cleanup;

Cleanup:
	// Clean up.
	// Free up memory
	if (lpProfAdmin) lpProfAdmin->Release();
	return hRes;
}

// HrAddDelegateMailbox
// Adds a delegate mailbox to a given service. The property set is one for Outlook 2010 and 2013 where all is needed is:
// - the Display Name for the mailbox
// - the mailbox distinguished name
// - the server NETBIOS or FQDN
// - the server DN
// - the SMTP address of the mailbox
HRESULT HrAddDelegateMailbox(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	BOOL bDefaultService,
	int iServiceIndex,
	LPWSTR lpszwMailboxDisplay,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer,
	LPWSTR lpszwServerDN,
	LPWSTR lpszwSMTPAddress,
	LPWSTR lpRohProxyserver,
	ULONG ulRohProxyServerFlags,
	ULONG ulRohProxyServerAuthPackage,
	LPWSTR lpwszMapiHttpMailStoreInternalUrl)
{
	HRESULT hRes = S_OK; // Result code returned from MAPI calls.
	SPropValue rgval[5]; // Property value structure to hold configuration info.
	LPPROFADMIN lpProfAdmin = NULL;
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPMAPITABLE lpServiceTable = NULL;
	// Enumeration for convenience.
	enum { iDispName, iSvcName, iSvcUID, iResourceFlags, iEmsMdbSectionUid, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_DISPLAY_NAME, PR_SERVICE_NAME, PR_SERVICE_UID, PR_RESOURCE_FLAGS, PR_EMSMDB_SECTION_UID };

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin

	if (bDefaultProfile)
	{
		lpwszProfileName = (LPWSTR)GetDefaultProfileName().c_str();
	}

	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
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
		enum { iServiceUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID, PR_RESOURCE_FLAGS };

		// Allocate memory for the restriction
		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpSvcRes), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal), L"Calling MAPIAllocateBuffer");

		// Set up restriction to query the profile table
		lpSvcRes->rt = RES_AND;
		lpSvcRes->res.resAnd.cRes = 0x00000002;
		lpSvcRes->res.resAnd.lpRes = lpsvcResLvl1;

		lpsvcResLvl1[0].rt = RES_EXIST;
		lpsvcResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
		lpsvcResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
		lpsvcResLvl1[1].rt = RES_PROPERTY;
		lpsvcResLvl1[1].res.resProperty.relop = RELOP_EQ;
		lpsvcResLvl1[1].res.resProperty.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[1].res.resProperty.lpProp = lpSvcPropVal;

		lpSvcPropVal->ulPropTag = PR_SERVICE_NAME_A;
		lpSvcPropVal->Value.lpszA = "MSEMS";

		// Query the table to get the the default profile only
		EC_HRES_MSG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			lpSvcRes,
			NULL,
			0,
			&lpSvcRows), L"Calling HrQueryAllRows");

		if (bDefaultService && lpSvcRows->cRows > 0)
		{
			for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
			{
				if (lpSvcRows->aRow[i].lpProps[iServiceResFlags].Value.l & SERVICE_DEFAULT_STORE)
				{
					LPPROVIDERADMIN lpProvAdmin = NULL;
					if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb,
						0,
						&lpProvAdmin)))
					{

						std::wstring wszSmtpAddress = ConvertWideCharToStdWstring(lpszwSMTPAddress);
						wszSmtpAddress = L"SMTP:" + wszSmtpAddress;

						// Set up a SPropValue array for the properties you need to configure.
						ZeroMemory(&rgval[0], sizeof(SPropValue));
						rgval[0].ulPropTag = PR_DISPLAY_NAME_W;
						rgval[0].Value.lpszW = lpszwMailboxDisplay;

						ZeroMemory(&rgval[1], sizeof(SPropValue));
						rgval[1].ulPropTag = PR_PROFILE_MAILBOX;
						rgval[1].Value.lpszA = ConvertWideCharToMultiByte(lpszwMailboxDN);

						ZeroMemory(&rgval[2], sizeof(SPropValue));
						rgval[2].ulPropTag = PR_PROFILE_SERVER;
						rgval[2].Value.lpszA = ConvertWideCharToMultiByte(lpszwServer);

						ZeroMemory(&rgval[3], sizeof(SPropValue));
						rgval[3].ulPropTag = PR_PROFILE_SERVER_DN;
						rgval[3].Value.lpszA = ConvertWideCharToMultiByte(lpszwServerDN);

						ZeroMemory(&rgval[4], sizeof(SPropValue));
						rgval[4].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
						rgval[4].Value.lpszA = ConvertWideCharToMultiByte((LPWSTR)wszSmtpAddress.c_str());

						printf("Creating EMSDelegate provider.\n");
						// Create the message service with the above properties.
						hRes = lpProvAdmin->CreateProvider(LPWSTR("EMSDelegate"),
							5,
							rgval,
							0,
							0,
							(LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb);
						if (FAILED(hRes)) goto Error;
					}
				}
			}
			if (lpSvcRows) FreeProws(lpSvcRows);
		}
		else if (lpSvcRows->cRows >= iServiceIndex)
		{
			LPPROVIDERADMIN lpProvAdmin = NULL;
			if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[iServiceIndex].lpProps[iServiceUid].Value.bin.lpb,
				0,
				&lpProvAdmin)))
			{

				std::wstring wszSmtpAddress = ConvertWideCharToStdWstring(lpszwSMTPAddress);
				wszSmtpAddress = L"SMTP:" + wszSmtpAddress;

				// Set up a SPropValue array for the properties you need to configure.
				ZeroMemory(&rgval[0], sizeof(SPropValue));
				rgval[0].ulPropTag = PR_DISPLAY_NAME_W;
				rgval[0].Value.lpszW = lpszwMailboxDisplay;

				ZeroMemory(&rgval[1], sizeof(SPropValue));
				rgval[1].ulPropTag = PR_PROFILE_MAILBOX;
				rgval[1].Value.lpszA = ConvertWideCharToMultiByte(lpszwMailboxDN);

				ZeroMemory(&rgval[2], sizeof(SPropValue));
				rgval[2].ulPropTag = PR_PROFILE_SERVER;
				rgval[2].Value.lpszA = ConvertWideCharToMultiByte(lpszwServer);

				ZeroMemory(&rgval[3], sizeof(SPropValue));
				rgval[3].ulPropTag = PR_PROFILE_SERVER_DN;
				rgval[3].Value.lpszA = ConvertWideCharToMultiByte(lpszwServerDN);

				ZeroMemory(&rgval[4], sizeof(SPropValue));
				rgval[4].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
				rgval[4].Value.lpszA = ConvertWideCharToMultiByte((LPWSTR)wszSmtpAddress.c_str());

				printf("Creating EMSDelegate provider.\n");
				// Create the message service with the above properties.
				hRes = lpProvAdmin->CreateProvider(LPWSTR("EMSDelegate"),
					5,
					rgval,
					0,
					0,
					(LPMAPIUID)lpSvcRows->aRow[iServiceIndex].lpProps[iServiceUid].Value.bin.lpb);
				if (FAILED(hRes)) goto Error;

			}
		}
		else
		{
			wprintf(L"No service found.\n");
		}

		if (lpServiceTable) lpServiceTable->Release();
		if (lpServiceAdmin) lpServiceAdmin->Release();

	}
	// End process services
	goto cleanup;

Error:
	printf("ERROR: hRes = %0x\n", hRes);

cleanup:
	// Clean up.
	if (lpProfAdmin) lpProfAdmin->Release();
	printf("Done cleaning up.\n");
	return hRes;
}

// HrAddDelegateMailbox
// Adds a delegate mailbox to a given service. The property set is one for Outlook 2007 where all is needed is:
// - the Display Name for the mailbox
// - the mailbox distinguished name
// - the server NETBIOS or FQDN
// - the server DN
HRESULT HrAddDelegateMailboxLegacy(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	BOOL bDefaultService,
	int iServiceIndex,
	LPWSTR lpszwMailboxDisplay,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer,
	LPWSTR lpszwServerDN)
{
	HRESULT hRes = S_OK; // Result code returned from MAPI calls.
	SPropValue rgval[4]; // Property value structure to hold configuration info.
	LPPROFADMIN lpProfAdmin = NULL;
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPMAPITABLE lpServiceTable = NULL;

	// Enumeration for convenience.
	enum { iDispName, iSvcName, iSvcUID, iResourceFlags, iEmsMdbSectionUid, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_DISPLAY_NAME, PR_SERVICE_NAME, PR_SERVICE_UID, PR_RESOURCE_FLAGS, PR_EMSMDB_SECTION_UID };

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin

													  // Begin process services


	if (bDefaultProfile)
	{
		lpwszProfileName = (LPWSTR)GetDefaultProfileName().c_str();
	}

	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling MAPIAdminProfiles");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		lpServiceAdmin->GetMsgServiceTable(0,
			&lpServiceTable);
		LPSRestriction lpSvcRes = NULL;
		LPSRestriction lpsvcResLvl1 = NULL;
		LPSPropValue lpSvcPropVal = NULL;
		LPSRowSet lpSvcRows = NULL;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID, PR_RESOURCE_FLAGS };

		// Allocate memory for the restriction
		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpSvcRes), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal), L"Calling MAPIAllocateBuffer");

		// Set up restriction to query the profile table
		lpSvcRes->rt = RES_AND;
		lpSvcRes->res.resAnd.cRes = 0x00000002;
		lpSvcRes->res.resAnd.lpRes = lpsvcResLvl1;

		lpsvcResLvl1[0].rt = RES_EXIST;
		lpsvcResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
		lpsvcResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
		lpsvcResLvl1[1].rt = RES_PROPERTY;
		lpsvcResLvl1[1].res.resProperty.relop = RELOP_EQ;
		lpsvcResLvl1[1].res.resProperty.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[1].res.resProperty.lpProp = lpSvcPropVal;

		lpSvcPropVal->ulPropTag = PR_SERVICE_NAME_A;
		lpSvcPropVal->Value.lpszA = "MSEMS";

		// Query the table to get the the default profile only
		EC_HRES_MSG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			lpSvcRes,
			NULL,
			0,
			&lpSvcRows), L"Calling HrQueryAllRows");

		if (bDefaultService && lpSvcRows->cRows > 0)
		{
			for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
			{
				if (lpSvcRows->aRow[i].lpProps[iServiceResFlags].Value.l & SERVICE_DEFAULT_STORE)
				{
					LPPROVIDERADMIN lpProvAdmin = NULL;
					if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb,
						0,
						&lpProvAdmin)))
					{
						// Set up a SPropValue array for the properties you need to configure.
						ZeroMemory(&rgval[0], sizeof(SPropValue));
						rgval[0].ulPropTag = PR_DISPLAY_NAME_W;
						rgval[0].Value.lpszW = lpszwMailboxDisplay;

						ZeroMemory(&rgval[1], sizeof(SPropValue));
						rgval[1].ulPropTag = PR_PROFILE_MAILBOX;
						rgval[1].Value.lpszA = ConvertWideCharToMultiByte(lpszwMailboxDN);

						ZeroMemory(&rgval[2], sizeof(SPropValue));
						rgval[2].ulPropTag = PR_PROFILE_SERVER;
						rgval[2].Value.lpszA = ConvertWideCharToMultiByte(lpszwServer);

						ZeroMemory(&rgval[3], sizeof(SPropValue));
						rgval[3].ulPropTag = PR_PROFILE_SERVER_DN;
						rgval[3].Value.lpszA = ConvertWideCharToMultiByte(lpszwServerDN);

						printf("Creating EMSDelegate provider.\n");
						// Create the message service with the above properties.
						hRes = lpProvAdmin->CreateProvider(LPWSTR("EMSDelegate"),
							4,
							rgval,
							0,
							0,
							(LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb);
						if (FAILED(hRes)) goto Error;
					}
				}
			}
			if (lpSvcRows) FreeProws(lpSvcRows);
		}
		else if (lpSvcRows->cRows >= iServiceIndex)
		{
			LPPROVIDERADMIN lpProvAdmin = NULL;
			if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[iServiceIndex].lpProps[iServiceUid].Value.bin.lpb,
				0,
				&lpProvAdmin)))
			{
				// Set up a SPropValue array for the properties you need to configure.
				ZeroMemory(&rgval[0], sizeof(SPropValue));
				rgval[0].ulPropTag = PR_DISPLAY_NAME_W;
				rgval[0].Value.lpszW = lpszwMailboxDisplay;

				ZeroMemory(&rgval[1], sizeof(SPropValue));
				rgval[1].ulPropTag = PR_PROFILE_MAILBOX;
				rgval[1].Value.lpszA = ConvertWideCharToMultiByte(lpszwMailboxDN);

				ZeroMemory(&rgval[2], sizeof(SPropValue));
				rgval[2].ulPropTag = PR_PROFILE_SERVER;
				rgval[2].Value.lpszA = ConvertWideCharToMultiByte(lpszwServer);

				ZeroMemory(&rgval[3], sizeof(SPropValue));
				rgval[3].ulPropTag = PR_PROFILE_SERVER_DN;
				rgval[3].Value.lpszA = ConvertWideCharToMultiByte(lpszwServerDN);

				printf("Creating EMSDelegate provider.\n");
				// Create the message service with the above properties.
				hRes = lpProvAdmin->CreateProvider(LPWSTR("EMSDelegate"),
					4,
					rgval,
					0,
					0,
					(LPMAPIUID)lpSvcRows->aRow[iServiceIndex].lpProps[iServiceUid].Value.bin.lpb);
				if (FAILED(hRes)) goto Error;
			}
		}
		else
		{
			wprintf(L"No service found.\n");
		}

		if (lpServiceTable) lpServiceTable->Release();
		if (lpServiceAdmin) lpServiceAdmin->Release();

	}
	// End process services
	goto cleanup;

Error:
	printf("ERROR: hRes = %0x\n", hRes);

cleanup:
	// Clean up.
	if (lpProfAdmin) lpProfAdmin->Release();
	printf("Done cleaning up.\n");
	return hRes;
}

#pragma endregion

#pragma region // Service Methods //

// HrGetDefaultMsemsServiceAdminProviderPtr
// Returns the provider admin interface pointer for the default service in a given profile
HRESULT HrGetDefaultMsemsServiceAdminProviderPtr(LPWSTR lpwszProfileName, LPPROVIDERADMIN * lppProvAdmin, LPMAPIUID * lppServiceUid)
{
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling #"); // Pointer to new IProfAdmin
									 // Get an IProfAdmin interface.

									 // Begin process services
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPMAPITABLE lpServiceTable = NULL;
	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling #");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		lpServiceAdmin->GetMsgServiceTable(0,
			&lpServiceTable);
		LPSRestriction lpSvcRes = NULL;
		LPSRestriction lpsvcResLvl1 = NULL;
		LPSPropValue lpSvcPropVal = NULL;
		LPSRowSet lpSvcRows = NULL;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID, PR_RESOURCE_FLAGS };

		// Allocate memory for the restriction
		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpSvcRes), L"Calling #");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl1), L"Calling #");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal), L"Calling #");

		// Set up restriction to query the profile table
		lpSvcRes->rt = RES_AND;
		lpSvcRes->res.resAnd.cRes = 0x00000002;
		lpSvcRes->res.resAnd.lpRes = lpsvcResLvl1;

		lpsvcResLvl1[0].rt = RES_EXIST;
		lpsvcResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
		lpsvcResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
		lpsvcResLvl1[1].rt = RES_PROPERTY;
		lpsvcResLvl1[1].res.resProperty.relop = RELOP_EQ;
		lpsvcResLvl1[1].res.resProperty.ulPropTag = PR_SERVICE_NAME_A;
		lpsvcResLvl1[1].res.resProperty.lpProp = lpSvcPropVal;

		lpSvcPropVal->ulPropTag = PR_SERVICE_NAME_A;
		lpSvcPropVal->Value.lpszA = "MSEMS";

		// Query the table to get the the default profile only
		EC_HRES_MSG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			lpSvcRes,
			NULL,
			0,
			&lpSvcRows), L"Calling #");

		if (lpSvcRows->cRows > 0)
		{
			for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
			{
				if (lpSvcRows->aRow[i].lpProps[iServiceResFlags].Value.l & SERVICE_DEFAULT_STORE)
				{
					LPPROVIDERADMIN lpProvAdmin = NULL;
					if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb,
						0,
						&lpProvAdmin)))
					{
						*lppServiceUid = (LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb;
						*lppProvAdmin = lpProvAdmin;
						if (lpProvAdmin) lpProvAdmin->Release();
						break;
					}
				}
			}
			if (lpSvcRows) FreeProws(lpSvcRows);
		}
		else
		{
			wprintf(L"No service found.\n");
		}

		if (lpServiceTable) lpServiceTable->Release();
		if (lpServiceAdmin) lpServiceAdmin->Release();

	}
	// End process services

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;
}


// HrGetSections
// Returns the EMSMDB and StoreProvider sections of a service
HRESULT HrGetSections(LPSERVICEADMIN2 lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT * lppEmsMdbSection, LPPROFSECT * lppStoreProviderSection)
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
		(LPSPropTagArray)&sptaUids,
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
		*lppEmsMdbSection = lpEmsMdbProfSect;

	if (NULL != lppStoreProviderSection)
		*lppStoreProviderSection = lpStoreProvProfSect;
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
		(LPSPropTagArray)&sptaUids,
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
		*lppEmsMdbSection = lpEmsMdbProfSect;

	if (NULL != lppStoreProviderSection)
	*lppStoreProviderSection = lpStoreProvProfSect;
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

// HrCrateMsemsServiceModernExt
// Crates a new message store service and configures the following properties:
// - PR_PROFILE_CONFIG_FLAGS
// - PR_RULE_ACTION_TYPE
// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
// - PR_DISPLAY_NAME_W
// - PR_PROFILE_ACCT_NAME_W
// - PR_PROFILE_UNRESOLVED_NAME_W
// - PR_PROFILE_USER_EMAIL_W
// Also updates the store provider section with the two following properties:
// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
// - PR_DISPLAY_NAME_W
// This implementation is Outlook 2016 specific
HRESULT HrCreateMsemsServiceModernExt(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	ULONG ulResourceFlags,
	ULONG ulProfileConfigFlags,
	ULONG ulCachedModeMonths,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszDisplayName)
{
	HRESULT			hRes = S_OK; // Result code returned from MAPI calls.
	SPropValue		rgvalEmsMdbSect[7]; // Property value structure to hold configuration info.
	SPropValue		rgvalStoreProvider[2];
	//	SPropValue		rgvalService[1];
	MAPIUID			uidService = { 0 };
	LPMAPIUID		lpServiceUid = &uidService;
	LPPROFSECT		lpProfSect = NULL;
	LPPROFSECT		lpEmsMdbProfSect = nullptr;
	LPPROFSECT		lpStoreProviderSect = nullptr;

	LPPROFADMIN lpProfAdmin = NULL;
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPSERVICEADMIN2 lpServiceAdmin2 = NULL;
	LPMAPITABLE lpServiceTable = NULL;

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling #"); // Pointer to new IProfAdmin

									 // Begin process services
	if (bDefaultProfile)
	{
		lpwszProfileName = (LPWSTR)GetDefaultProfileName().c_str();
	}

	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling #");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)&lpServiceAdmin2), L"Calling QueryInterface.");

		// Adds a message service to the current profile and returns that newly added service UID.
		EC_HRES_MSG(lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSEMS", (LPTSTR)"Microsoft Exchange", NULL, 0, &uidService), L"Calling CreateMsgServiceEx.");

		EC_HRES_MSG(lpServiceAdmin2->OpenProfileSection(&uidService,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpProfSect), L"Calling OpenProfileSection.");


		LPMAPIPROP lpMapiProp = NULL;
		EC_HRES_MSG(lpProfSect->QueryInterface(IID_IMAPIProp, (LPVOID*)&lpMapiProp), L"Calling QueryInterface.");

		if (lpMapiProp)
		{
			LPSPropValue prResourceFlags;
			MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)&prResourceFlags);

			prResourceFlags->ulPropTag = PR_RESOURCE_FLAGS;
			prResourceFlags->Value.l = ulResourceFlags;
			EC_HRES_MSG(lpMapiProp->SetProps(1, prResourceFlags, NULL), L"Calling SetProps.");

			EC_HRES_MSG(lpMapiProp->SaveChanges(FORCE_SAVE), L"Calling SaveChanges.");
			MAPIFreeBuffer(prResourceFlags);
			lpMapiProp->Release();
		}

		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)&lpEmsMdbProfSect);
		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)&lpStoreProviderSect);
		ZeroMemory(lpEmsMdbProfSect, sizeof(LPPROFSECT));
		ZeroMemory(lpStoreProviderSect, sizeof(LPPROFSECT));

		EC_HRES_MSG(HrGetSections(lpServiceAdmin2, lpServiceUid, &lpEmsMdbProfSect, &lpStoreProviderSect), L"Calling HrGetSections.");

		// Set up a SPropValue array for the properties you need to configure.
		/*
		PR_PROFILE_CONFIG_FLAGS
		PR_RULE_ACTION_TYPE
		PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
		PR_DISPLAY_NAME_W
		PR_PROFILE_ACCT_NAME_W
		PR_PROFILE_UNRESOLVED_NAME_W
		PR_PROFILE_USER_EMAIL_W
		*/
		std::wstring wszSmtpAddress = ConvertWideCharToStdWstring(lpszSmtpAddress);
		if ((wszSmtpAddress.find(L"SMTP:") == std::string::npos) || (wszSmtpAddress.find(L"smtp:") == std::string::npos))
		{
			wszSmtpAddress = L"SMTP:" + wszSmtpAddress;
		}

		ZeroMemory(&rgvalEmsMdbSect[0], sizeof(SPropValue));
		rgvalEmsMdbSect[0].ulPropTag = PR_PROFILE_CONFIG_FLAGS;
		rgvalEmsMdbSect[0].Value.l = ulProfileConfigFlags;

		ZeroMemory(&rgvalEmsMdbSect[1], sizeof(SPropValue));
		rgvalEmsMdbSect[1].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
		rgvalEmsMdbSect[1].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();

		ZeroMemory(&rgvalEmsMdbSect[2], sizeof(SPropValue));
		rgvalEmsMdbSect[2].ulPropTag = PR_DISPLAY_NAME_W;
		rgvalEmsMdbSect[2].Value.lpszW = lpszDisplayName;

		ZeroMemory(&rgvalEmsMdbSect[3], sizeof(SPropValue));
		rgvalEmsMdbSect[3].ulPropTag = PR_PROFILE_ACCT_NAME_W;
		rgvalEmsMdbSect[3].Value.lpszW = lpszDisplayName;

		ZeroMemory(&rgvalEmsMdbSect[4], sizeof(SPropValue));
		rgvalEmsMdbSect[4].ulPropTag = PR_PROFILE_UNRESOLVED_NAME_W;
		rgvalEmsMdbSect[4].Value.lpszW = lpszDisplayName;

		ZeroMemory(&rgvalEmsMdbSect[5], sizeof(SPropValue));
		rgvalEmsMdbSect[5].ulPropTag = PR_PROFILE_USER_EMAIL_W;
		rgvalEmsMdbSect[5].Value.lpszW = lpszDisplayName;

		ZeroMemory(&rgvalEmsMdbSect[6], sizeof(SPropValue));
		rgvalEmsMdbSect[6].ulPropTag = PR_RULE_ACTION_TYPE;
		rgvalEmsMdbSect[6].Value.l = ulCachedModeMonths;

		EC_HRES_MSG(lpEmsMdbProfSect->SetProps(
			7,
			rgvalEmsMdbSect,
			nullptr), L"Calling SetProps.");

		EC_HRES_MSG(lpEmsMdbProfSect->SaveChanges(KEEP_OPEN_READWRITE), L"Calling SaveChanges.");

		//Updating store provider 
		/*
		PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
		PR_DISPLAY_NAME_W
		*/
		ZeroMemory(&rgvalStoreProvider[0], sizeof(SPropValue));
		rgvalStoreProvider[0].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
		rgvalStoreProvider[0].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();

		ZeroMemory(&rgvalStoreProvider[1], sizeof(SPropValue));
		rgvalStoreProvider[1].ulPropTag = PR_DISPLAY_NAME_W;
		rgvalStoreProvider[1].Value.lpszW = lpszDisplayName;

		EC_HRES_MSG(lpStoreProviderSect->SetProps(
			2,
			rgvalStoreProvider,
			nullptr), L"Calling SetProps.");

		EC_HRES_MSG(lpStoreProviderSect->SaveChanges(KEEP_OPEN_READWRITE), L"Calling SaveChanges.");
	}

	goto Cleanup;
Error:
	return hRes;

Cleanup:
	// Clean up
	if (lpStoreProviderSect) lpStoreProviderSect->Release();
	if (lpEmsMdbProfSect) lpEmsMdbProfSect->Release();
	if (lpProfSect) lpProfSect->Release();
	if (lpServiceAdmin2) lpServiceAdmin2->Release();
	if (lpServiceAdmin) lpServiceAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;
}

// HrCrateMsemsServiceModern
// Crates a new message store service and configures the following properties:
// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
// - PR_DISPLAY_NAME_W
// - PR_PROFILE_ACCT_NAME_W
// - PR_PROFILE_UNRESOLVED_NAME_W
// - PR_PROFILE_USER_EMAIL_W
// Also updates the store provider section with the two following properties:
// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
// - PR_DISPLAY_NAME_W
// This implementation is Outlook 2016 specific
HRESULT HrCreateMsemsServiceModern(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszDisplayName	)
{
	HRESULT			hRes = S_OK; // Result code returned from MAPI calls.
	SPropValue		rgvalEmsMdbSect[5]; // Property value structure to hold configuration info.
	SPropValue		rgvalStoreProvider[2];
	//	SPropValue		rgvalService[1];
	MAPIUID			uidService = { 0 };
	LPMAPIUID		lpServiceUid = &uidService;
	LPPROFSECT		lpProfSect = NULL;
	LPPROFSECT		lpEmsMdbProfSect = nullptr;
	LPPROFSECT		lpStoreProviderSect = nullptr;

	LPPROFADMIN lpProfAdmin = NULL;
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPSERVICEADMIN2 lpServiceAdmin2 = NULL;
	LPMAPITABLE lpServiceTable = NULL;

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling #"); // Pointer to new IProfAdmin

									 // Begin process services

	if (bDefaultProfile)
	{
		lpwszProfileName = (LPWSTR)GetDefaultProfileName().c_str();
	}

	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling #");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{

		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)&lpServiceAdmin2), L"Calling QueryInterface.");
		// Adds a message service to the current profile and returns that newly added service UID.
		EC_HRES_MSG(lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSEMS", (LPTSTR)"Microsoft Exchange", NULL, 0, &uidService), L"Calling CreateMsgServiceEx.");

		EC_HRES_MSG(lpServiceAdmin2->OpenProfileSection(&uidService,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpProfSect), L"Calling OpenProfileSection.");

		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)&lpEmsMdbProfSect);
		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)&lpStoreProviderSect);
		ZeroMemory(lpEmsMdbProfSect, sizeof(LPPROFSECT));
		ZeroMemory(lpStoreProviderSect, sizeof(LPPROFSECT));

		EC_HRES_MSG(HrGetSections(lpServiceAdmin2, &uidService, &lpEmsMdbProfSect, &lpStoreProviderSect), L"Calling HrGetSections.");

		// Set up a SPropValue array for the properties you need to configure.
		/*
		PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
		PR_DISPLAY_NAME_W
		PR_PROFILE_ACCT_NAME_W
		PR_PROFILE_UNRESOLVED_NAME_W
		PR_PROFILE_USER_EMAIL_W
		*/

		std::wstring wszSmtpAddress = ConvertWideCharToStdWstring(lpszSmtpAddress);
		if ((wszSmtpAddress.find(L"SMTP:") == std::string::npos) || (wszSmtpAddress.find(L"smtp:") == std::string::npos))
		{
			wszSmtpAddress = L"SMTP:" + wszSmtpAddress;
		}

		ZeroMemory(&rgvalEmsMdbSect[0], sizeof(SPropValue));
		rgvalEmsMdbSect[0].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
		rgvalEmsMdbSect[0].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();

		ZeroMemory(&rgvalEmsMdbSect[1], sizeof(SPropValue));
		rgvalEmsMdbSect[1].ulPropTag = PR_DISPLAY_NAME_W;
		rgvalEmsMdbSect[1].Value.lpszW = lpszDisplayName;

		ZeroMemory(&rgvalEmsMdbSect[2], sizeof(SPropValue));
		rgvalEmsMdbSect[2].ulPropTag = PR_PROFILE_ACCT_NAME_W;
		rgvalEmsMdbSect[2].Value.lpszW = lpszDisplayName;

		ZeroMemory(&rgvalEmsMdbSect[3], sizeof(SPropValue));
		rgvalEmsMdbSect[3].ulPropTag = PR_PROFILE_UNRESOLVED_NAME_W;
		rgvalEmsMdbSect[3].Value.lpszW = lpszDisplayName;

		ZeroMemory(&rgvalEmsMdbSect[4], sizeof(SPropValue));
		rgvalEmsMdbSect[4].ulPropTag = PR_PROFILE_USER_EMAIL_W;
		rgvalEmsMdbSect[4].Value.lpszW = lpszDisplayName;

		EC_HRES_MSG(lpEmsMdbProfSect->SetProps(
			5,
			rgvalEmsMdbSect,
			nullptr), L"Calling SetProps.");

		EC_HRES_MSG(lpEmsMdbProfSect->SaveChanges(KEEP_OPEN_READWRITE), L"Calling SaveChanges.");

		//Updating store provider 
		/*
		PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
		PR_DISPLAY_NAME_W
		*/
		ZeroMemory(&rgvalStoreProvider[0], sizeof(SPropValue));
		rgvalStoreProvider[0].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
		rgvalStoreProvider[0].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();;

		ZeroMemory(&rgvalStoreProvider[1], sizeof(SPropValue));
		rgvalStoreProvider[1].ulPropTag = PR_DISPLAY_NAME_W;
		rgvalStoreProvider[1].Value.lpszW = lpszDisplayName;

		EC_HRES_MSG(lpStoreProviderSect->SetProps(
			2,
			rgvalStoreProvider,
			nullptr), L"Calling SetProps.");

		EC_HRES_MSG(lpStoreProviderSect->SaveChanges(KEEP_OPEN_READWRITE), L"Calling SaveChanges.");
	}

	goto Cleanup;
Error:
	return hRes;

Cleanup:
	// Clean up
	if (lpStoreProviderSect) lpStoreProviderSect->Release();
	if (lpEmsMdbProfSect) lpEmsMdbProfSect->Release();
	if (lpProfSect) lpProfSect->Release();
	if (lpServiceAdmin2) lpServiceAdmin2->Release();
	if (lpServiceAdmin) lpServiceAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	return hRes;
}

// HrCreateMsemsServiceLegacyUnresolved
// Crates a new message store service and configures the following properties it with a default property set. 
// This is the legacy implementation where Outlook resolves the mailbox based on "unresolved" mailbox and server names. I use this for Outlook 2007.
HRESULT HrCreateMsemsServiceLegacyUnresolved(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	LPWSTR lpszwMailboxDN,
	LPWSTR lpszwServer)
{
	HRESULT hRes = S_OK; // Result code returned from MAPI calls.
	LPPROFADMIN lpProfAdmin = NULL; // Profile Admin pointer.
	SPropValue rgval[2]; // Property value structure to hold configuration info.
	ULONG ulProps = 0; // Count of props.
	ULONG cbNewBuffer = 0; // Count of bytes for new buffer.
	LPPROVIDERADMIN lpProvAdmin = NULL;
	LPMAPIUID lpServiceUid = NULL;
	LPMAPIUID lpEmsMdbSectionUid = NULL;
	MAPIUID				uidService = { 0 };
	LPMAPIUID			lpuidService = &uidService;
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPSERVICEADMIN2 lpServiceAdmin2 = NULL;
	LPMAPITABLE lpServiceTable = NULL;

	// Enumeration for convenience.
	enum { iDispName, iSvcName, iSvcUID, iResourceFlags, iEmsMdbSectionUid, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_DISPLAY_NAME, PR_SERVICE_NAME, PR_SERVICE_UID, PR_RESOURCE_FLAGS, PR_EMSMDB_SECTION_UID };



	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling #"); // Pointer to new IProfAdmin

									 // Begin process services


	if (bDefaultProfile)
	{
		lpwszProfileName = (LPWSTR)GetDefaultProfileName().c_str();
	}

	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling #");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{

		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)&lpServiceAdmin2), L"Calling QueryInterface.");

		printf("Creating MsgService.\n");
		// Adds a message service to the current profile and returns that newly added service UID.
		hRes = lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSEMS", (LPTSTR)"Microsoft Exchange", NULL, 0, &uidService);
		if (FAILED(hRes)) goto Error;

		// Set up a SPropValue array for the properties you need to configure.
		// First, the server name.
		ZeroMemory(&rgval[0], sizeof(SPropValue));
		rgval[0].ulPropTag = PR_PROFILE_UNRESOLVED_SERVER;
		rgval[0].Value.lpszA = ConvertWideCharToMultiByte(L"1d07c6cd-9c3f-42f8-93d3-7781e20725d9@adelaide.lab");
		// Next, the DN of the mailbox.
		ZeroMemory(&rgval[1], sizeof(SPropValue));
		rgval[1].ulPropTag = PR_PROFILE_UNRESOLVED_NAME;
		rgval[1].Value.lpszA = ConvertWideCharToMultiByte(lpszwMailboxDN);

		printf("Configuring MsgService.\n");
		// Create the message service with the above properties.
		hRes = lpServiceAdmin2->ConfigureMsgService(&uidService,
			NULL,
			0,
			2,
			rgval);
		if (FAILED(hRes)) goto Error;


	}
	goto cleanup;

Error:
	printf("ERROR: hRes = %0x\n", hRes);

cleanup:
	// Clean up
	printf("Done cleaning up.\n");
	if (lpServiceAdmin2) lpServiceAdmin2->Release();
	if (lpServiceAdmin) lpServiceAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	return hRes;
}

// HrCreateMsemsServiceROH
// Creates a new message store service and sets it for RPC / HTTP with the following properties:
// - PR_DISPLAY_NAME_A
// - PR_PROFILE_HOME_SERVER
// - PR_PROFILE_USER
// - PR_PROFILE_HOME_SERVER_DN
// - PR_PROFILE_CONFIG_FLAGS
// - PR_ROH_PROXY_SERVER
// - PR_ROH_FLAGS
// - PR_ROH_PROXY_AUTH_SCHEME
// - PR_PROFILE_AUTH_PACKAGE
// - PR_PROFILE_SERVER_FQDN_W
// Configures the Store Provider with the following properties:
// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
// - PR_DIPLAY_NAME_W
HRESULT HrCreateMsemsServiceROH(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszMailboxLegacyDn,
	LPWSTR lpszUnresolvedServer,
	LPWSTR lpszRohProxyServer,
	LPWSTR lpszProfileServerDn,
	LPWSTR lpszAutodiscoverUrl)
{
	HRESULT hRes = S_OK; // Result code returned from MAPI calls.
	SPropValue rgvalSvc[10];
	SPropValue rgvalEmsMdbSect[14]; // Property value structure to hold configuration info.
	SPropValue rgvalStoreProvider[2];
	LPPROVIDERADMIN lpProvAdmin = NULL;
	LPMAPIUID lpServiceUid = NULL;
	LPMAPIUID lpEmsMdbSectionUid = NULL;
	MAPIUID				uidService = { 0 };
	LPMAPIUID			lpuidService = &uidService;
	LPPROFSECT lpProfSect = NULL;
	LPPROFSECT		lpEmsMdbProfSect = nullptr;
	LPPROFSECT lpStoreProviderSect = nullptr;
	ULONG			cValues = 0;
	LPSPropValue	lpProps = nullptr;

	// Enumeration for convenience.
	enum { iDispName, iSvcName, iSvcUID, iResourceFlags, iEmsMdbSectionUid, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_DISPLAY_NAME, PR_SERVICE_NAME, PR_SERVICE_UID, PR_RESOURCE_FLAGS, PR_EMSMDB_SECTION_UID };
	std::wstring wszSmtpAddress = ConvertWideCharToStdWstring(lpszSmtpAddress);
	wszSmtpAddress = L"SMTP:" + wszSmtpAddress;
	//// This structure tells our GetProps call what properties to get from the global profile section.
	//SizedSPropTagArray(1, sptGlobal) = { 1, PR_STORE_PROVIDERS };

	LPPROFADMIN lpProfAdmin = NULL;
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPSERVICEADMIN2 lpServiceAdmin2 = NULL;
	LPMAPITABLE lpServiceTable = NULL;

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin

									 // Begin process services

	if (bDefaultProfile)
	{
		lpwszProfileName = (LPWSTR)GetDefaultProfileName().c_str();
	}

	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling AdminServices");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{

		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)&lpServiceAdmin2), L"Calling QueryInterface.");


		printf("Creating MsgService.\n");
		// Adds a message service to the current profile and returns that newly added service UID.
		hRes = lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSEMS", (LPTSTR)"Microsoft Exchange", NULL, 0, &uidService);
		if (FAILED(hRes)) goto Error;

		printf("Configuring MsgService.\n");

		int paramC = 0;
		if (lpszSmtpAddress)
		{
			ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
			rgvalSvc[paramC].ulPropTag = PR_DISPLAY_NAME_A;
			rgvalSvc[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszSmtpAddress);
			paramC++;
		}

		if (lpszUnresolvedServer)
		{
			ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
			rgvalSvc[paramC].ulPropTag = PR_PROFILE_HOME_SERVER;
			rgvalSvc[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszUnresolvedServer);
			paramC++;
		}

		if (lpszMailboxLegacyDn)
		{
			ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
			rgvalSvc[paramC].ulPropTag = PR_PROFILE_USER;
			rgvalSvc[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszMailboxLegacyDn);
			paramC++;
		}

		if (lpszProfileServerDn)
		{
			ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
			rgvalSvc[paramC].ulPropTag = PR_PROFILE_HOME_SERVER_DN;
			rgvalSvc[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszProfileServerDn);
			paramC++;
		}

		ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
		rgvalSvc[paramC].ulPropTag = PR_PROFILE_CONFIG_FLAGS;
		rgvalSvc[paramC].Value.l = CONFIG_SHOW_CONNECT_UI;
		paramC++;

		if (lpszRohProxyServer)
		{
			ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
			rgvalSvc[paramC].ulPropTag = PR_ROH_PROXY_SERVER;
			rgvalSvc[paramC].Value.lpszW = lpszRohProxyServer;
			paramC++;
		}

		ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
		rgvalSvc[paramC].ulPropTag = PR_ROH_FLAGS;
		rgvalSvc[paramC].Value.l = ROHFLAGS_USE_ROH | ROHFLAGS_HTTP_FIRST_ON_FAST | ROHFLAGS_HTTP_FIRST_ON_SLOW;
		paramC++;

		ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
		rgvalSvc[paramC].ulPropTag = PR_ROH_PROXY_AUTH_SCHEME;
		rgvalSvc[paramC].Value.l = RPC_C_HTTP_AUTHN_SCHEME_NTLM;
		paramC++;

		ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
		rgvalSvc[paramC].ulPropTag = PR_PROFILE_AUTH_PACKAGE;
		rgvalSvc[paramC].Value.l = RPC_C_AUTHN_WINNT;
		paramC++;


		if (lpszUnresolvedServer)
		{
			ZeroMemory(&rgvalSvc[paramC], sizeof(SPropValue));
			rgvalSvc[paramC].ulPropTag = PR_PROFILE_SERVER_FQDN_W;
			rgvalSvc[paramC].Value.lpszW = lpszUnresolvedServer;
			paramC++;
		}

		// Create the message service with the above properties.
		hRes = lpServiceAdmin2->ConfigureMsgService(&uidService,
			NULL,
			0,
			paramC,
			rgvalSvc);
		if (FAILED(hRes)) goto Error;

		printf("Accessing MsgService.\n");

		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)&lpEmsMdbProfSect);
		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)&lpStoreProviderSect);
		ZeroMemory(lpEmsMdbProfSect, sizeof(LPPROFSECT));
		ZeroMemory(lpStoreProviderSect, sizeof(LPPROFSECT));

		EC_HRES_MSG(HrGetSections(lpServiceAdmin2, &uidService, &lpEmsMdbProfSect, &lpStoreProviderSect), L"Calling HrGetSections.");

		paramC = 0;
		// Set up a SPropValue array for the properties you need to configure.
		if (lpszMailboxLegacyDn)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_USER;
			rgvalEmsMdbSect[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszMailboxLegacyDn);
			paramC++;
		}

		if (lpszUnresolvedServer)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_HOME_SERVER;
			rgvalEmsMdbSect[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszUnresolvedServer);
			paramC++;
		}

		if (lpszRohProxyServer)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_ROH_PROXY_SERVER;
			rgvalEmsMdbSect[paramC].Value.lpszW = lpszRohProxyServer;
			paramC++;
		}

		ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
		rgvalEmsMdbSect[paramC].ulPropTag = PR_ROH_FLAGS;
		rgvalEmsMdbSect[paramC].Value.l = ROHFLAGS_USE_ROH | ROHFLAGS_HTTP_FIRST_ON_FAST | ROHFLAGS_HTTP_FIRST_ON_SLOW;
		paramC++;

		ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
		rgvalEmsMdbSect[paramC].ulPropTag = PR_ROH_PROXY_AUTH_SCHEME;
		rgvalEmsMdbSect[paramC].Value.l = RPC_C_HTTP_AUTHN_SCHEME_NTLM;
		paramC++;

		ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
		rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_AUTH_PACKAGE;
		rgvalEmsMdbSect[paramC].Value.l = RPC_C_AUTHN_WINNT;
		paramC++;

		if (lpszSmtpAddress)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_DISPLAY_NAME_W;
			rgvalEmsMdbSect[paramC].Value.lpszW = lpszSmtpAddress;
			paramC++;
		}

		if (lpszProfileServerDn)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_HOME_SERVER_DN;
			rgvalEmsMdbSect[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszProfileServerDn);
			paramC++;
		}

		if (lpszUnresolvedServer)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_HOME_SERVER_FQDN;
			rgvalEmsMdbSect[paramC].Value.lpszW = lpszUnresolvedServer;
			paramC++;
		}

		if (lpszSmtpAddress)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_UNRESOLVED_NAME;
			rgvalEmsMdbSect[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszSmtpAddress);
			paramC++;
		}

		if (lpszUnresolvedServer)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_UNRESOLVED_SERVER;
			rgvalEmsMdbSect[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszUnresolvedServer);
			paramC++;
		}

		if (lpszSmtpAddress)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_ACCT_NAME_W;
			rgvalEmsMdbSect[paramC].Value.lpszW = lpszSmtpAddress;
			paramC++;
		}

		ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
		rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
		rgvalEmsMdbSect[paramC].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();
		paramC++;

		if (lpszAutodiscoverUrl)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_LKG_AUTODISCOVER_URL;
			rgvalEmsMdbSect[paramC].Value.lpszW = lpszAutodiscoverUrl;
			paramC++;
		}

		hRes = lpEmsMdbProfSect->SetProps(
			paramC,
			rgvalEmsMdbSect,
			nullptr);

		if (FAILED(hRes))
		{
			goto Error;
		}

		printf("Saving changes.\n");

		hRes = lpEmsMdbProfSect->SaveChanges(KEEP_OPEN_READWRITE);

		if (FAILED(hRes))
		{
			goto Error;
		}

		//Updating store provider 
		if (lpStoreProviderSect)
		{
			ZeroMemory(&rgvalStoreProvider[0], sizeof(SPropValue));
			rgvalStoreProvider[0].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
			rgvalStoreProvider[0].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();

			ZeroMemory(&rgvalStoreProvider[1], sizeof(SPropValue));
			rgvalStoreProvider[1].ulPropTag = PR_DISPLAY_NAME_W;
			rgvalStoreProvider[1].Value.lpszW = lpszSmtpAddress;

			hRes = lpStoreProviderSect->SetProps(
				2,
				rgvalStoreProvider,
				nullptr);

			if (FAILED(hRes))
			{
				goto Error;
			}

			printf("Saving changes.\n");
			hRes = lpStoreProviderSect->SaveChanges(KEEP_OPEN_READWRITE);

			if (FAILED(hRes))
			{
				goto Error;
			}

		}
	}
	goto cleanup;


Error:
	printf("ERROR: hRes = %0x\n", hRes);

cleanup:
	// Clean up
	if (lpStoreProviderSect) lpStoreProviderSect->Release();
	if (lpEmsMdbProfSect) lpEmsMdbProfSect->Release();
	if (lpProfSect) lpProfSect->Release();
	if (lpServiceAdmin2) lpServiceAdmin2->Release();
	if (lpServiceAdmin) lpServiceAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	printf("Done cleaning up.\n");
	return hRes;
}

// HrCreateMsemsServiceMOH
// Creates a new message store service and sets it for MAPI / HTTP with the following properties:
// - PR_PROFILE_CONFIG_FLAGS
// - PR_PROFILE_AUTH_PACKAGE
// - PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL
// - PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL
// - PR_PROFILE_USER
// Configures the Store Provider with the following properties:
// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
// - PR_DIPLAY_NAME_W
// - PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL
// - PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL
HRESULT HrCreateMsemsServiceMOH(BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	LPWSTR lpszSmtpAddress,
	LPWSTR lpszMailboxDn,
	LPWSTR lpszMailStoreInternalUrl,
	LPWSTR lpszMailStoreExternalUrl,
	LPWSTR lpszAddressBookInternalUrl,
	LPWSTR lpszAddressBookExternalUrl,
	LPWSTR lpszRohProxyServer)
{
	HRESULT hRes = S_OK; // Result code returned from MAPI calls.
	SPropValue rgvalSvc[5];
	//	SPropValue rgvalEmsMdbSect[14]; // Property value structure to hold configuration info.
	SPropValue rgvalStoreProvider[5];
	LPPROVIDERADMIN lpProvAdmin = NULL;
	LPMAPIUID lpServiceUid = NULL;
	LPMAPIUID lpEmsMdbSectionUid = NULL;
	MAPIUID				uidService = { 0 };
	LPMAPIUID			lpuidService = &uidService;
	LPPROFSECT lpProfSect = NULL;
	LPPROFSECT		lpEmsMdbProfSect = nullptr;
	LPPROFSECT lpStoreProviderSect = nullptr;
	ULONG			cValues = 0;
	LPSPropValue	lpProps = nullptr;

	// Enumeration for convenience.
	enum { iDispName, iSvcName, iSvcUID, iResourceFlags, iEmsMdbSectionUid, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_DISPLAY_NAME, PR_SERVICE_NAME, PR_SERVICE_UID, PR_RESOURCE_FLAGS, PR_EMSMDB_SECTION_UID };
	std::wstring wszSmtpAddress = ConvertWideCharToStdWstring(lpszSmtpAddress);
	if ((wszSmtpAddress.find(L"SMTP:") == std::string::npos) || (wszSmtpAddress.find(L"smtp:") == std::string::npos))
	{
		wszSmtpAddress = L"SMTP:" + wszSmtpAddress;
	}

	//// This structure tells our GetProps call what properties to get from the global profile section.
	//SizedSPropTagArray(1, sptGlobal) = { 1, PR_STORE_PROVIDERS };
	LPPROFADMIN lpProfAdmin = NULL;

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling #"); // Pointer to new IProfAdmin

									 // Begin process services
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPSERVICEADMIN2 lpServiceAdmin2 = NULL;
	LPMAPITABLE lpServiceTable = NULL;

	if (bDefaultProfile)
	{
		lpwszProfileName = (LPWSTR)GetDefaultProfileName().c_str();
	}

	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling #");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{

		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)&lpServiceAdmin2), L"Calling QueryInterface.");

		printf("Creating MsgService.\n");

		// Adds a message service to the current profile and returns that newly added service UID.
		hRes = lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSEMS", (LPTSTR)"Microsoft Exchange", NULL, 0, &uidService);
		if (FAILED(hRes)) goto Error;

		EC_HRES_MSG(HrGetSections(lpServiceAdmin2, &uidService, &lpEmsMdbProfSect, &lpStoreProviderSect), L"Calling HrGetSections");

		int paramC = 0;
		std::vector<SPropValue> rgvalVector;
		SPropValue sPropValue;
		
	
		//Updating emsmdb section 
		if (lpEmsMdbProfSect)
		{

			rgvalVector.resize(0);

			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
			sPropValue.Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();
			rgvalVector.push_back(sPropValue);
			paramC++;

			if (lpszAddressBookInternalUrl)
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL;
				sPropValue.Value.lpszW = lpszAddressBookInternalUrl;
				rgvalVector.push_back(sPropValue);
				paramC++;
			}

			if (lpszAddressBookInternalUrl)
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL;
				sPropValue.Value.lpszW = lpszAddressBookExternalUrl;
				rgvalVector.push_back(sPropValue);
				paramC++;
			}

			if (lpszSmtpAddress)
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_DISPLAY_NAME_W;
				sPropValue.Value.lpszW = lpszSmtpAddress;
				rgvalVector.push_back(sPropValue);
				paramC++;
			}

			if (lpszSmtpAddress)
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_PROFILE_UNRESOLVED_NAME;
				sPropValue.Value.lpszA = ConvertWideCharToMultiByte(lpszSmtpAddress);
				rgvalVector.push_back(sPropValue);
				paramC++;
			}

			if (lpszSmtpAddress)
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_PROFILE_MAILBOX;
				sPropValue.Value.lpszA = ConvertWideCharToMultiByte(lpszMailboxDn);
				rgvalVector.push_back(sPropValue);
				paramC++;
			}

			if (lpszRohProxyServer)
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_ROH_PROXY_SERVER;
				sPropValue.Value.lpszW = lpszRohProxyServer;
				rgvalVector.push_back(sPropValue);
				paramC++;
			}
			
			
			hRes = lpEmsMdbProfSect->SetProps(
				rgvalVector.size(),
				rgvalVector.data(),
				nullptr);

			if (FAILED(hRes))
			{
				goto Error;
			}

			printf("Saving changes.\n");
			hRes = lpEmsMdbProfSect->SaveChanges(KEEP_OPEN_READWRITE);

			if (FAILED(hRes))
			{
				goto Error;
			}

		}

		//Updating store provider 
		if (lpStoreProviderSect)
		{
			
			rgvalVector.resize(0);

			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
			sPropValue.Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();
			rgvalVector.push_back(sPropValue);
			paramC++;

			if (lpszAddressBookExternalUrl)
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL;
				sPropValue.Value.lpszW = lpszMailStoreExternalUrl;
				rgvalVector.push_back(sPropValue);
				paramC++;
			}

			if (lpszSmtpAddress)
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_DISPLAY_NAME_W;
				sPropValue.Value.lpszW = lpszSmtpAddress;
				rgvalVector.push_back(sPropValue);
				paramC++;
			}

			if (lpszMailStoreInternalUrl)
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL;
				sPropValue.Value.lpszW = lpszMailStoreInternalUrl;
				rgvalVector.push_back(sPropValue);
				paramC++;
			}

			//ZeroMemory(&rgvalStoreProvider[4], sizeof(SPropValue));
			//rgvalStoreProvider[4].ulPropTag = PR_PROFILE_USER;
			//rgvalStoreProvider[4].Value.lpszA = ConvertWideCharToMultiByte(lpszMailboxDn);;

			hRes = lpStoreProviderSect->SetProps(
				rgvalVector.size(),
				rgvalVector.data(),
				nullptr);

			if (FAILED(hRes))
			{
				goto Error;
			}

			printf("Saving changes.\n");
			hRes = lpStoreProviderSect->SaveChanges(KEEP_OPEN_READWRITE);

			if (FAILED(hRes))
			{
				goto Error;
			}

		}
	}
	goto cleanup;


Error:
	printf("ERROR: hRes = %0x\n", hRes);

cleanup:
	// Clean up
	if (lpStoreProviderSect) lpStoreProviderSect->Release();
	if (lpEmsMdbProfSect) lpEmsMdbProfSect->Release();
	if (lpProfSect) lpProfSect->Release();
	if (lpServiceAdmin2) lpServiceAdmin2->Release();
	if (lpServiceAdmin) lpServiceAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	printf("Done cleaning up.\n");
	return hRes;
}

#pragma endregion

HRESULT HrSetCachedMode(LPWSTR lpwszProfileName, BOOL bDefaultProfile, BOOL bAllProfiles, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths)
{
	HRESULT hRes = S_OK;

	if (bDefaultProfile)
	{
		ProfileInfo * profileInfo = new ProfileInfo();
		EC_HRES_MSG(HrGetProfile((LPWSTR)GetDefaultProfileName().c_str(), profileInfo), L"Calling GetProfile");
		EC_HRES_MSG(HrSetCachedModeOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultProfile, bAllServices, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths), L"Calling HrPromoteDelegatesInProfile");

	}
	else if (bAllProfiles)
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo * profileInfo = new ProfileInfo[ulProfileCount];
		hRes = HrGetProfiles(ulProfileCount, profileInfo);
		for (int i = 0; i <= ulProfileCount; i++)
		{
			EC_HRES_MSG(HrSetCachedModeOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultProfile, bAllServices, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths), L"Calling HrPromoteDelegatesInProfile");
		}
	}
	else
	{
		if (lpwszProfileName)
		{
			ProfileInfo * profileInfo = new ProfileInfo();
			hRes = HrGetProfile(lpwszProfileName, profileInfo);
			EC_HRES_MSG(HrSetCachedModeOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultProfile, bAllServices, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths), L"Calling HrPromoteDelegatesInProfile");
		}
		else
			wprintf(L"The specified profile name is invalid or no profile name was specified.\n");
	}

Error:
Cleanup:
	return hRes;
}

HRESULT HrSetCachedModeOneProfile(LPWSTR lpwszProfileName, ProfileInfo * pProfileInfo, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths)
{
	HRESULT hRes = S_OK;

	for (int i = 0; i <= pProfileInfo->ulServiceCount; i++)
	{
		if (bDefaultService)
		{
			if (pProfileInfo->profileServices[i].bDefaultStore)
			{
				if (pProfileInfo->profileServices[i].ulServiceType == SERVICETYPE_MAILBOX)
				{
					EC_HRES_MSG(HrSetCachedModeOneService(ConvertWideCharToMultiByte(lpwszProfileName), &pProfileInfo->profileServices[i].muidServiceUid, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths), L"Calling HrSetCachedModeOneService on service");
				}
			}
		}
		else if (iServiceIndex != -1)
		{
			if (pProfileInfo->profileServices[iServiceIndex].ulServiceType == SERVICETYPE_MAILBOX)
			{
				EC_HRES_MSG(HrSetCachedModeOneService(ConvertWideCharToMultiByte(lpwszProfileName), &pProfileInfo->profileServices[iServiceIndex].muidServiceUid, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths), L"Calling HrSetCachedModeOneService on service");

			}
		}
		else if (bAllServices)
		{
			if (pProfileInfo->profileServices[i].ulServiceType == SERVICETYPE_MAILBOX)
			{
				EC_HRES_MSG(HrSetCachedModeOneService(ConvertWideCharToMultiByte(lpwszProfileName), &pProfileInfo->profileServices[i].muidServiceUid, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths), L"Calling HrSetCachedModeOneService on service");
			}
		}
	}
Error:
Cleanup:

	return hRes;
}


HRESULT HrPromoteDelegates(LPWSTR lpwszProfileName, BOOL bDefaultProfile, BOOL bAllProfiles, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ULONG ulConnectMode)
{
	HRESULT hRes = S_OK;

	if (bDefaultProfile)
	{
		ProfileInfo * profileInfo = new ProfileInfo();
		EC_HRES_MSG(HrGetProfile((LPWSTR)GetDefaultProfileName().c_str(), profileInfo), L"Calling GetProfile");
		EC_HRES_MSG(HrPromoteDelegatesOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultService, bAllServices, iOutlookVersion, ulConnectMode), L"Calling HrPromoteDelegatesOneProfile");

	}
	else if (bAllProfiles)
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo * profileInfo = new ProfileInfo[ulProfileCount];
		hRes = HrGetProfiles(ulProfileCount, profileInfo);
		for (int i = 0; i <= ulProfileCount; i++)
		{
			EC_HRES_MSG(HrPromoteDelegatesOneProfile((LPWSTR)profileInfo[i].wszProfileName.c_str(), &profileInfo[i], iServiceIndex, bDefaultService, bAllServices, iOutlookVersion, ulConnectMode), L"Calling HrPromoteDelegatesOneProfile");
		}
	}
	else
	{
		if (lpwszProfileName)
		{
			ProfileInfo * profileInfo = new ProfileInfo();
			hRes = HrGetProfile(lpwszProfileName, profileInfo);
			EC_HRES_MSG(HrPromoteDelegatesOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultService, bAllServices, iOutlookVersion, ulConnectMode), L"Calling HrPromoteDelegatesOneProfile");

		}
		else
			Logger::Write(logLevelError, L"The specified profile name is invalid or no profile name was specified.\n");
	}

Error:
	Cleanup:
	return hRes;
}

HRESULT HrPromoteDelegatesOneProfile(LPWSTR lpwszProfileName, ProfileInfo * pProfileInfo, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ULONG ulConnectMode)
{
	HRESULT hRes = S_OK;

	for (int i = 0; i <= pProfileInfo->ulServiceCount; i++)
	{
		if (bDefaultService)
		{
			if (pProfileInfo->profileServices[i].bDefaultStore)
			{
				if (pProfileInfo->profileServices[i].ulServiceType == SERVICETYPE_MAILBOX)
				{
					for (int j = 0; j <= pProfileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount; j++)
					{
						if ((pProfileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulProfileType == PROFILE_DELEGATE) && (pProfileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].bIsOnlineArchive == false))
						{
							EC_HRES_MSG(HrPromoteOneDelegate(lpwszProfileName, iOutlookVersion, ulConnectMode, pProfileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j]), L"Calling HrPromoteDelegate");
						}
					}
				}
			}
		}
		else if (iServiceIndex != -1)
		{
			if (pProfileInfo->profileServices[iServiceIndex].ulServiceType == SERVICETYPE_MAILBOX)
			{
				for (int j = 0; j <= pProfileInfo->profileServices[iServiceIndex].exchangeAccountInfo->ulMailboxCount; j++)
				{
					if ((pProfileInfo->profileServices[iServiceIndex].exchangeAccountInfo->accountMailboxes[j].ulProfileType == PROFILE_DELEGATE) && (pProfileInfo->profileServices[iServiceIndex].exchangeAccountInfo->accountMailboxes[j].bIsOnlineArchive == false))
					{
						EC_HRES_MSG(HrPromoteOneDelegate(lpwszProfileName, iOutlookVersion, ulConnectMode, pProfileInfo->profileServices[iServiceIndex].exchangeAccountInfo->accountMailboxes[j]), L"Calling HrPromoteDelegate");
					}
				}
			}
		}
		else if (bAllServices)
		{
			if (pProfileInfo->profileServices[i].ulServiceType == SERVICETYPE_MAILBOX)
			{
				for (int j = 0; j <= pProfileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount; j++)
				{
					if ((pProfileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulProfileType == PROFILE_DELEGATE) && (pProfileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].bIsOnlineArchive == false))
					{
						EC_HRES_MSG(HrPromoteOneDelegate(lpwszProfileName, iOutlookVersion, ulConnectMode, pProfileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j]), L"Calling HrPromoteDelegate");
					}
				}
			}
		}


	}
Error:
Cleanup:

	return hRes;
}

HRESULT HrPromoteOneDelegate(LPWSTR lpwszProfileName, int iOutlookVersion, ULONG ulConnectMode, MailboxInfo mailboxInfo)
{
	HRESULT hRes = S_OK;
	switch (iOutlookVersion)
	{
	case 2007:
		// I haven't tested this so best not use it
		//if (SUCCEEDED(HrCreateMsemsServiceLegacyUnresolved(FALSE,
		//	profileName,
		//	(LPWSTR)mailboxInfo.wszProfileMailbox.c_str(),
		//	(LPWSTR)mailboxInfo.wszProfileServer.c_str(),
		//	loggingMode)))
		//{
		//	EC_HRES_MSG(HrDeleteProvider(profileName, &pProfileInfo->profileServices[i].muidServiceUid, &mailboxInfo.muidProviderUid, loggingMode), L"Calling HrDeleteProvider");
		//}
		Logger::Write(logLevelError, L"This client version is not currently supported.");
		break;
	case 2010:
	case 2013:
		if (ulConnectMode == CONNECT_ROH)
		{
			// This id a bit of a hack since delegate mailboxes don't need to have the personalised server name in the delegate provider
			// I'm just creating these based on the legacyDN and the MailStore so best check that those have value
			Logger::Write(logLevelError, L"Validating delegate information.");
			if ((mailboxInfo.wszMailStoreInternalUrl != L"") && (mailboxInfo.wszProfileMailbox != L""))
			{
				std::wstring wszParsedSmtpAddress = SubstringToEnd(L"smtp:", mailboxInfo.wszSmtpAddress);
				std::wstring wszPersonalisedServerName = SubstringToEnd(L"MailboxId=", mailboxInfo.wszMailStoreInternalUrl);
				std::wstring wszServerDN = SubstringFromStart(L"cn=Recipients", mailboxInfo.wszProfileMailbox) + L"/cn=Configuration/cn=Servers/cn=" + wszPersonalisedServerName;

				Logger::Write(logLevelInfo, L"Creating and configuring new ROH service.");
				if (SUCCEEDED(HrCreateMsemsServiceROH(FALSE,
					lpwszProfileName,
					(LPWSTR)wszParsedSmtpAddress.c_str(),
					(LPWSTR)mailboxInfo.wszProfileMailbox.c_str(),
					(LPWSTR)wszPersonalisedServerName.c_str(),
					(LPWSTR)mailboxInfo.wszRohProxyServer.c_str(),
					(LPWSTR)wszServerDN.c_str(),
					(LPWSTR)NULL)))
				{
					EC_HRES_MSG(HrDeleteProvider(lpwszProfileName, &mailboxInfo.muidServiceUid, &mailboxInfo.muidProviderUid), L"Calling HrDeleteProvider");
				}
			}
			else
				Logger::Write(logLevelError, L"Not enough information in the profile for ROH mailbox.");
		}
		// best not be used for now as I haven't sorted it out
		else if (ulConnectMode == CONNECT_MOH)
		{
			Logger::Write(logLevelError, L"MOH logic is not currently available.");
			//if (SUCCEEDED(HrCreateMsemsServiceMOH(FALSE,
			//	profileName,
			//	(LPWSTR)SubstringToEnd(L"smtp:", mailboxInfo.wszSmtpAddress).c_str(),
			//	(LPWSTR)mailboxInfo.wszProfileMailbox.c_str(),
			//	(LPWSTR)mailboxInfo.wszMailStoreInternalUrl.c_str(),
			//	NULL,
			//	NULL,
			//	NULL,
			//	(LPWSTR)mailboxInfo.wszRohProxyServer.c_str(),
			//	loggingMode)))
			//{
			//	EC_HRES_MSG(HrDeleteProvider(profileName, &pProfileInfo->profileServices[i].muidServiceUid, &mailboxInfo.muidProviderUid, loggingMode), L"Calling HrDeleteProvider");
			//}
		}

		break;
	case 2016:
		Logger::Write(logLevelInfo, L"Creating and configuring new service.");
		if (SUCCEEDED(HrCreateMsemsServiceModern(FALSE,
			lpwszProfileName,
			(LPWSTR)SubstringToEnd(L"smtp:", mailboxInfo.wszSmtpAddress).c_str(),
			(LPWSTR)SubstringToEnd(L"smtp:", mailboxInfo.wszSmtpAddress).c_str())))
		{
			EC_HRES_MSG(HrDeleteProvider(lpwszProfileName, &mailboxInfo.muidServiceUid, &mailboxInfo.muidProviderUid), L"Calling HrDeleteProvider");
		}

		break;
	}

Error:
Cleanup:
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
Cleanup:

	return hRes;
}