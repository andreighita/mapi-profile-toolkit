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
#include <initguid.h>
#define USES_IID_IMAPIProp 
#include "ProfileFunctions.h"
#include "MAPIObjects.h"
#include <EdkMdb.h>
#include <MAPIGuid.h>
#include <MSPST.h>
#include "Logger.h"
#include <WinBase.h>
#include <Shlwapi.h>
#include "StringOperations.h"


#define MAPI_FORCE_ACCESS 0x00080000
#define PR_EMSMDB_SECTION_UID	PROP_TAG(PT_BINARY, 0x3D15)
#define PR_PROFILE_USER_SMTP_EMAIL_ADDRESS	PROP_TAG(PT_STRING8, pidProfileMin+0x41)
#define PR_ROH_PROXY_SERVER	PROP_TAG(PT_STRING8, 0x6622)

#ifndef CONFIG_OST_CACHE_PRIVATE
#define CONFIG_OST_CACHE_PRIVATE			((ULONG)0x00000180)
#endif
#ifndef CONFIG_OST_CACHE_DELEGATE_PIM
#define CONFIG_OST_CACHE_DELEGATE_PIM		((ULONG)0x00000800)
#endif
#ifndef CONFIG_OST_CACHE_PUBLIC
#define CONFIG_OST_CACHE_PUBLIC				((ULONG)0x00000400)
#endif



std::wstring GetDefaultProfileName(LoggingMode loggingMode)
{
	std::wstring szDefaultProfileName;
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPSRestriction lpProfRes = NULL;
	LPSRestriction lpProfResLvl1 = NULL;
	LPSPropValue lpProfPropVal = NULL;
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	HRESULT hRes = S_OK;
	std::wstring errormessage = std::wstring(L" ");
	Logger::Write(logLevelInfo, L"Attempting to retrieve the default MAPI profile name", loggingMode);

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, iProviderUid, iServiceUid, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME_A, PR_PROVIDER_UID, PR_SERVICE_UID };
	EC_HRES_LOG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), loggingMode); // Pointer to new IProfAdmin
						// Get an IProfAdmin interface.

	EC_HRES_LOG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), loggingMode);

	// Allocate memory for the restriction
	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)&lpProfRes), loggingMode);

	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SRestriction) * 2,
		(LPVOID*)&lpProfResLvl1), loggingMode);

	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SPropValue),
		(LPVOID*)&lpProfPropVal), loggingMode);

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
	EC_HRES_LOG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		lpProfRes,
		NULL,
		0,
		&lpProfRows), loggingMode);

	if (lpProfRows->cRows == 0)
	{
		Logger::Write(logLevelFailed, L"No default profile set.", loggingMode);
	}
	else if (lpProfRows->cRows == 1)
	{

		szDefaultProfileName = ConvertMultiByteToWideChar(lpProfRows->aRow->lpProps[iDisplayName].Value.lpszA);
	}
	else
	{
		Logger::Write(logLevelError, L"Query resulted in incosinstent results", loggingMode);
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

ULONG GetProfileCount(LoggingMode loggingMode)
{
	std::string szDefaultProfileName;
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPMAPITABLE lpProfTable = NULL;
	ULONG ulRowCount = 0;
	HRESULT hRes = S_OK;

	EC_HRES_LOG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), loggingMode); // Pointer to new IProfAdmin
						// Get an IProfAdmin interface.

	EC_HRES_LOG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), loggingMode);

	EC_HRES_LOG(lpProfTable->GetRowCount(0,
		&ulRowCount), loggingMode);

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	if (lpProfTable) lpProfTable->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	return ulRowCount;
}

HRESULT GetProfiles(ULONG ulProfileCount, ProfileInfo * profileInfo, LoggingMode loggingMode)
{
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	HRESULT hRes = S_OK;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME_A };

	EC_HRES_LOG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), loggingMode); // Pointer to new IProfAdmin
						// Get an IProfAdmin interface.

	EC_HRES_LOG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), loggingMode);

	// Query the table to get the the default profile only
	EC_HRES_LOG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		NULL,
		NULL,
		0,
		&lpProfRows), loggingMode);

	if (lpProfRows->cRows == ulProfileCount)
	{
		for (unsigned int i = 0; i < lpProfRows->cRows; i++)
		{
			GetProfile(ConvertMultiByteToWideChar(lpProfRows->aRow[i].lpProps[iDisplayName].Value.lpszA), &profileInfo[i], loggingMode);
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

HRESULT GetProfile(LPWSTR lpszProfileName, ProfileInfo * profileInfo, LoggingMode loggingMode)
{
	HRESULT hRes = S_OK;
	profileInfo->szProfileName = ConvertWideCharToStdWstring(lpszProfileName);

	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPSRestriction lpProfRes = NULL;
	LPSRestriction lpProfResLvl1 = NULL;
	LPSPropValue lpProfPropVal = NULL;
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, iDefaultProfile, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME, PR_DEFAULT_PROFILE };

	EC_HRES_LOG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), loggingMode); // Pointer to new IProfAdmin
						// Get an IProfAdmin interface.

	EC_HRES_LOG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), loggingMode);

	// Allocate memory for the restriction
	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)&lpProfRes), loggingMode);

	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SRestriction) * 2,
		(LPVOID*)&lpProfResLvl1), loggingMode);

	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SPropValue),
		(LPVOID*)&lpProfPropVal), loggingMode);

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
	EC_HRES_LOG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		lpProfRes,
		NULL,
		0,
		&lpProfRows), loggingMode);

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
	EC_HRES_LOG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		MAPI_UNICODE,                    // Flags.
		&lpServiceAdmin), loggingMode);        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		lpServiceAdmin->GetMsgServiceTable(0,
			&lpServiceTable);
		LPSRestriction lpSvcRes = NULL;
		LPSRestriction lpsvcResLvl1 = NULL;
		LPSPropValue lpSvcPropVal = NULL;
		LPSRowSet lpSvcRows = NULL;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceName, iEmsMdbSectUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID,PR_SERVICE_NAME_A, PR_EMSMDB_SECTION_UID, PR_RESOURCE_FLAGS };

		//// Allocate memory for the restriction
		//EC_HRES_LOG(MAPIAllocateBuffer(
		//	sizeof(SRestriction),
		//	(LPVOID*)&lpSvcRes));

		//EC_HRES_LOG(MAPIAllocateBuffer(
		//	sizeof(SRestriction) * 2,
		//	(LPVOID*)&lpsvcResLvl1));

		//EC_HRES_LOG(MAPIAllocateBuffer(
		//	sizeof(SPropValue),
		//	(LPVOID*)&lpSvcPropVal));

		//// Set up restriction to query the profile table
		//lpSvcRes->rt = RES_AND;
		//lpSvcRes->res.resAnd.cRes = 0x00000002;
		//lpSvcRes->res.resAnd.lpRes = lpsvcResLvl1;

		//lpsvcResLvl1[0].rt = RES_EXIST;
		//lpsvcResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;
		//lpsvcResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
		//lpsvcResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
		//lpsvcResLvl1[1].rt = RES_PROPERTY;
		//lpsvcResLvl1[1].res.resProperty.relop = RELOP_EQ;
		//lpsvcResLvl1[1].res.resProperty.ulPropTag = PR_SERVICE_NAME_A;
		//lpsvcResLvl1[1].res.resProperty.lpProp = lpSvcPropVal;

		//lpSvcPropVal->ulPropTag = PR_SERVICE_NAME_A;
		//lpSvcPropVal->Value.lpszA = "MSEMS";

		// Query the table to get the the default profile only
		EC_HRES_LOG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			NULL,
			NULL,
			0,
			&lpSvcRows), loggingMode);

		if (lpSvcRows->cRows > 0)
		{
			profileInfo->ulServiceCount = lpSvcRows->cRows;
			profileInfo->profileServices = new ServiceInfo[lpSvcRows->cRows];;

			// Start loop services
			for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
			{
				profileInfo->profileServices[i].szServiceName = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA));
				profileInfo->profileServices[i].bDefaultStore = (lpSvcRows->aRow[i].lpProps[iServiceResFlags].Value.l & SERVICE_DEFAULT_STORE);
				profileInfo->profileServices[i].ulServiceType = SERVICETYPE_OTHER;
				// Exchange account
				if (0 == strcmp(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA, "MSEMS"))
				{
					profileInfo->profileServices[i].ulServiceType = SERVICETYPE_EXCHANGEACCOUNT;
					profileInfo->profileServices[i].exchangeAccountInfo = new ExchangeAccountInfo;
					profileInfo->profileServices[i].exchangeAccountInfo->szDatafilePath = std::wstring(L" ");
					profileInfo->profileServices[i].exchangeAccountInfo->szServiceDisplayName = std::wstring(L" ");
					profileInfo->profileServices[i].exchangeAccountInfo->szUserEmailSmtpAddress = std::wstring(L" ");
					profileInfo->profileServices[i].exchangeAccountInfo->szUserName = std::wstring(L" ");
					profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN = std::wstring(L" ");
					profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerName = std::wstring(L" ");
					profileInfo->profileServices[i].exchangeAccountInfo->wszRohProxyServer = std::wstring(L" ");
					profileInfo->profileServices[i].exchangeAccountInfo->wszUnresolvedServer = std::wstring(L" ");

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

								// bind to the PR_PROFILE_CONFIG_FLAGS property
								LPSPropValue profileUnresolvedName = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_UNRESOLVED_NAME, &profileUnresolvedName)))
								{
									if (profileUnresolvedName)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->szUserName = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUnresolvedName->Value.lpszA));
										if (profileUnresolvedName) MAPIFreeBuffer(profileUnresolvedName);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->szUserName = std::wstring(L" ");
									}

								}
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
								// bind to the PR_PROFILE_CONFIG_FLAGS property
								LPSPropValue profileConfigFlags = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_CONFIG_FLAGS, &profileConfigFlags)))
								{
									if (profileConfigFlags)
									{
										if (profileConfigFlags->Value.l & CONFIG_OST_CACHE_PRIVATE)
										{
											profileInfo->profileServices[i].exchangeAccountInfo->bCachedModeEnabledOwner = 1;
										}
										else
										{
											profileInfo->profileServices[i].exchangeAccountInfo->bCachedModeEnabledOwner = 0;
										}
										if (profileConfigFlags->Value.l & CONFIG_OST_CACHE_DELEGATE_PIM)
										{
											profileInfo->profileServices[i].exchangeAccountInfo->bCachedModeEnabledShared = 1;
										}
										else
										{
											profileInfo->profileServices[i].exchangeAccountInfo->bCachedModeEnabledShared = 0;
										}
										if (profileConfigFlags->Value.l & CONFIG_OST_CACHE_PUBLIC)
										{
											profileInfo->profileServices[i].exchangeAccountInfo->bCachedModeEnabledPublicFolders = 1;
										}
										else
										{
											profileInfo->profileServices[i].exchangeAccountInfo->bCachedModeEnabledPublicFolders = 0;
										}
										if (profileConfigFlags) MAPIFreeBuffer(profileConfigFlags);
									}
								}
								// bind to the PR_PROFILE_USER_SMTP_EMAIL_ADDRESS property
								LPSPropValue profileUserSmtpEmailAddress = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_USER_SMTP_EMAIL_ADDRESS, &profileUserSmtpEmailAddress)))
								{
									if (profileUserSmtpEmailAddress)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->szUserEmailSmtpAddress = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUserSmtpEmailAddress->Value.lpszA));
										if (profileUserSmtpEmailAddress) MAPIFreeBuffer(profileUserSmtpEmailAddress);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->szUserEmailSmtpAddress = std::wstring(L" ");
									}
								}
								// bind to the PR_PROFILE_HOME_SERVER_DN property
								LPSPropValue profileHomeServerDn = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_HOME_SERVER_DN, &profileHomeServerDn)))
								{
									if (profileHomeServerDn)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileHomeServerDn->Value.lpszA));
										if (profileHomeServerDn) MAPIFreeBuffer(profileHomeServerDn);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN = std::wstring(L" ");
									}
								}
								// bind to the PR_PROFILE_UNRESOLVED_SERVER property
								LPSPropValue profileUnresolvedServer = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_UNRESOLVED_SERVER, &profileUnresolvedServer)))
								{
									if (profileUnresolvedServer)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszUnresolvedServer = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUnresolvedServer->Value.lpszA));
										if (profileUnresolvedServer) MAPIFreeBuffer(profileUnresolvedServer);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszUnresolvedServer = std::wstring(L" ");
									}
								}
								// bind to the PR_PROFILE_HOME_SERVER property
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
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerName = std::wstring(L" ");
									}
								}
								// bind to the PR_ROH_PROXY_SERVER property
								LPSPropValue profileRohProxyServer = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_ROH_PROXY_SERVER, &profileRohProxyServer)))
								{
									if (profileUserSmtpEmailAddress)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszRohProxyServer = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileRohProxyServer->Value.lpszA));
										if (profileRohProxyServer) MAPIFreeBuffer(profileRohProxyServer);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszRohProxyServer = std::wstring(L" ");
									}
								}
								LPSPropValue profileOfflineStorePath = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_OFFLINE_STORE_PATH, &profileOfflineStorePath)))
								{
									if (profileOfflineStorePath)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->szDatafilePath = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileOfflineStorePath->Value.lpszA));
										if (profileOfflineStorePath) MAPIFreeBuffer(profileOfflineStorePath);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->szDatafilePath = std::wstring(L" ");
									}
								}
								LPSPropValue profileDisplayName = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_DISPLAY_NAME_A, &profileDisplayName)))
								{
									if (profileDisplayName)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->szServiceDisplayName = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileDisplayName->Value.lpszA));
										if (profileDisplayName) MAPIFreeBuffer(profileDisplayName);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->szServiceDisplayName = std::wstring(L" ");
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
						EC_HRES_LOG(MAPIAllocateBuffer(
							sizeof(SRestriction),
							(LPVOID*)&lpProvRes), loggingMode);

						EC_HRES_LOG(MAPIAllocateBuffer(
							sizeof(SRestriction) * 2,
							(LPVOID*)&lpProvResLvl1), loggingMode);

						EC_HRES_LOG(MAPIAllocateBuffer(
							sizeof(SPropValue),
							(LPVOID*)&lpProvPropVal), loggingMode);

						// Set up restriction to query the provider table
						lpProvRes->rt = RES_AND;
						lpProvRes->res.resAnd.cRes = 0x00000002;
						lpProvRes->res.resAnd.lpRes = lpProvResLvl1;

						lpProvResLvl1[0].rt = RES_EXIST;
						lpProvResLvl1[0].res.resExist.ulPropTag = PR_PROVIDER_DISPLAY_A;
						lpProvResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
						lpProvResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
						lpProvResLvl1[1].rt = RES_CONTENT;
						lpProvResLvl1[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
						lpProvResLvl1[1].res.resContent.ulPropTag = PR_PROVIDER_DISPLAY_A;
						lpProvResLvl1[1].res.resContent.lpProp = lpProvPropVal;

						lpProvPropVal->ulPropTag = PR_PROVIDER_DISPLAY_A;
						lpProvPropVal->Value.lpszA = "Microsoft Exchange Message Store";

						lpProvAdmin->GetProviderTable(0,
							&lpProvTable);
						// Query the table to get the the default profile only
						EC_HRES_LOG(HrQueryAllRows(lpProvTable,
							(LPSPropTagArray)&sptaProvProps,
							lpProvRes,
							NULL,
							0,
							&lpProvRows), loggingMode);

						if (lpProvRows->cRows > 0)
						{
							profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount = lpProvRows->cRows;
							profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes = new MailboxInfo[lpProvRows->cRows];

							for (unsigned int j = 0; j < lpProvRows->cRows; j++)
							{
								profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].szDisplayName = std::wstring(L" ");
								LPPROFSECT lpProfSection = NULL;
								if (SUCCEEDED(lpServiceAdmin->OpenProfileSection((LPMAPIUID)lpProvRows->aRow[j].lpProps[iProvInstanceKey].Value.bin.lpb, NULL, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSection)))
								{
									LPMAPIPROP lpMAPIProp = NULL;
									if (SUCCEEDED(lpProfSection->QueryInterface(IID_IMAPIProp, (void**)&lpMAPIProp)))
									{
										LPSPropValue prDisplayName = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_DISPLAY_NAME, &prDisplayName)))
										{
											profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].szDisplayName = ConvertWideCharToStdWstring(prDisplayName->Value.lpszW);
											if (prDisplayName) MAPIFreeBuffer(prDisplayName);
										}
										else
										{
											profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].szDisplayName = std::wstring(L" ");
										}

										LPSPropValue prProfileType = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_TYPE, &prProfileType)))
										{
											profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].bDefaultMailbox = (prProfileType->Value.l == PROFILE_PRIMARY_USER);

											if (prProfileType->Value.l == PROFILE_PRIMARY_USER)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulEntryType = ENTRYTYPE_PRIMARY;
											}
											else if (prProfileType->Value.l == PROFILE_DELEGATE)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulEntryType = ENTRYTYPE_DELEGATE;
											}
											else if (prProfileType->Value.l == PROFILE_PUBLIC_STORE)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulEntryType = ENTRYTYPE_PUBLIC_FOLDERS;
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulEntryType = ENTRYTYPE_UNKNOWN;
											}
											if (prDisplayName) MAPIFreeBuffer(prProfileType);
										}
									}
								}
							}
							if (lpProvRows) FreeProws(lpProvRows);
						}
						else
						{
							profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount = lpProvRows->cRows;
							profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes = new MailboxInfo[1];
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
					profileInfo->profileServices[i].ulServiceType = SERVICETYPE_PST;
					profileInfo->profileServices[i].pstInfo = new PstInfo;
					profileInfo->profileServices[i].pstInfo->szDisplayName = std::wstring(L" ");
					profileInfo->profileServices[i].pstInfo->szPstPath = std::wstring(L" ");

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
						EC_HRES_LOG(MAPIAllocateBuffer(
							sizeof(SRestriction),
							(LPVOID*)&lpProvRes), loggingMode);

						EC_HRES_LOG(MAPIAllocateBuffer(
							sizeof(SRestriction) * 2,
							(LPVOID*)&lpProvResLvl1), loggingMode);

						EC_HRES_LOG(MAPIAllocateBuffer(
							sizeof(SPropValue),
							(LPVOID*)&lpProvPropVal), loggingMode);

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
						EC_HRES_LOG(HrQueryAllRows(lpProvTable,
							(LPSPropTagArray)&sptaProvProps,
							lpProvRes,
							NULL,
							0,
							&lpProvRows), loggingMode);

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
										profileInfo->profileServices[i].pstInfo->szDisplayName = ConvertWideCharToStdWstring(prDisplayName->Value.lpszW);
										if (prDisplayName) MAPIFreeBuffer(prDisplayName);
									}
									else
									{
										profileInfo->profileServices[i].pstInfo->szDisplayName = std::wstring(L" ");
									}
									// bind to the PR_PST_PATH_W property
									LPSPropValue pstPathW = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_PATH_W, &pstPathW)))
									{
										if (pstPathW)
										{
											profileInfo->profileServices[i].pstInfo->szPstPath = ConvertWideCharToStdWstring(pstPathW->Value.lpszW);
											if (pstPathW) MAPIFreeBuffer(pstPathW);
										}
										else
										{
											profileInfo->profileServices[i].pstInfo->szPstPath = std::wstring(L" ");
										}
									}
									// bind to the PR_PST_CONFIG_FLAGS property to get the ammount to sync
									LPSPropValue pstConfigFlags = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_CONFIG_FLAGS, &pstConfigFlags)))
									{
										if (pstConfigFlags)
										{
											profileInfo->profileServices[i].pstInfo->ulPstType = pstConfigFlags->Value.l;
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

HRESULT UpdateCachedModeConfig(LPSTR lpszProfileName, ULONG ulSectionIndex, ULONG ulCachedModeOwner, ULONG ulCachedModeShared, ULONG ulCachedModePublicFolders, int iCachedModeMonths, LoggingMode loggingMode)
{
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer

	EC_HRES_LOG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), loggingMode); // Pointer to new IProfAdmin
						// Get an IProfAdmin interface.

						// Begin process services
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPMAPITABLE lpServiceTable = NULL;
	EC_HRES_LOG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), loggingMode);        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		lpServiceAdmin->GetMsgServiceTable(0,
			&lpServiceTable);
		LPSRestriction lpSvcRes = NULL;
		LPSRestriction lpsvcResLvl1 = NULL;
		LPSPropValue lpSvcPropVal = NULL;
		LPSRowSet lpSvcRows = NULL;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceName, iEmsMdbSectUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID,PR_SERVICE_NAME_A, PR_EMSMDB_SECTION_UID, PR_RESOURCE_FLAGS };

		// Allocate memory for the restriction
		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpSvcRes), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl1), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal), loggingMode);

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
		EC_HRES_LOG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			lpSvcRes,
			NULL,
			0,
			&lpSvcRows), loggingMode);

		if (lpSvcRows->cRows >= ulSectionIndex)
		{
			LPPROVIDERADMIN lpProvAdmin = NULL;

			if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[ulSectionIndex - 1].lpProps[iServiceUid].Value.bin.lpb,
				0,
				&lpProvAdmin)))
			{
				// Access the EMSMDB section
				LPPROFSECT lpProfSect = NULL;
				if (SUCCEEDED(lpProvAdmin->OpenProfileSection((LPMAPIUID)lpSvcRows->aRow[ulSectionIndex - 1].lpProps[iEmsMdbSectUid].Value.bin.lpb,
					NULL,
					MAPI_MODIFY,
					&lpProfSect)))
				{
					LPMAPIPROP pMAPIProp = NULL;
					if (SUCCEEDED(lpProfSect->QueryInterface(IID_IMAPIProp, (void**)&pMAPIProp)))
					{
						// bind to the PR_PROFILE_CONFIG_FLAGS property
						LPSPropValue profileConfigFlags = NULL;
						if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_CONFIG_FLAGS, &profileConfigFlags)))
						{
							if (profileConfigFlags)
							{
								if (ulCachedModeOwner > 0)
								{
									if (ulCachedModeOwner == CACHEDMODE_ENABLED)
									{
										if (!(profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PRIVATE))
										{
											profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PRIVATE;
											EC_HRES_LOG(lpProfSect->SetProps(1, profileConfigFlags, NULL), loggingMode);
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
											EC_HRES_LOG(lpProfSect->SetProps(1, profileConfigFlags, NULL), loggingMode);
											printf("Cached mode owner disabled.\n");
										}
										else
										{
											printf("Cached mode owner already disabled on service.\n");
										}
									}
								}
								if (ulCachedModeShared > 0)
								{
									if (ulCachedModeShared == CACHEDMODE_ENABLED)
									{
										if (!(profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_DELEGATE_PIM))
										{
											profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_DELEGATE_PIM;
											EC_HRES_LOG(lpProfSect->SetProps(1, profileConfigFlags, NULL), loggingMode);
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
											EC_HRES_LOG(lpProfSect->SetProps(1, profileConfigFlags, NULL), loggingMode);
											printf("Cached mode shared disabled.\n");
										}
										else
										{
											printf("Cached mode shared already disabled on service.\n");
										}
									}
								}
								if (ulCachedModePublicFolders > 0)
								{
									if (ulCachedModePublicFolders == CACHEDMODE_ENABLED)
									{
										if (!(profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PUBLIC))
										{
											profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PUBLIC;
											EC_HRES_LOG(lpProfSect->SetProps(1, profileConfigFlags, NULL), loggingMode);
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
											EC_HRES_LOG(lpProfSect->SetProps(1, profileConfigFlags, NULL), loggingMode);
											printf("Cached mode public folders disabled.\n");
										}
										else
										{
											printf("Cached mode public folders already disabled on service.\n");
										}
									}
								}
								EC_HRES_LOG(lpProfSect->SaveChanges(0), loggingMode);
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
								EC_HRES_LOG(lpProfSect->SetProps(1, profileRuleActionType, NULL), loggingMode);
								printf("Cached mode amount to sync set.\n");

								EC_HRES_LOG(lpProfSect->SaveChanges(0), loggingMode);
								if (profileRuleActionType) MAPIFreeBuffer(profileConfigFlags);
							}
						}
					}
					if (lpProfSect) lpProfSect->Release();
				}

				if (lpProvAdmin) lpProvAdmin->Release();
			}

			if (lpSvcRows) FreeProws(lpSvcRows);
		}
		else
		{
			printf("Invalid service index specified %u.\n", ulSectionIndex);
			printf("Highest possible index is %u.\n", lpSvcRows->cRows);
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
	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;
	return S_OK;
}


HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszOldPath, LPWSTR lpszNewPath, bool bMoveFiles, LoggingMode loggingMode)
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

	EC_HRES_LOG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), loggingMode); // Pointer to new IProfAdmin
						// Get an IProfAdmin interface.

	EC_HRES_LOG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), loggingMode);

	// Allocate memory for the restriction
	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)&lpProfRes), loggingMode);

	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SRestriction) * 2,
		(LPVOID*)&lpProfResLvl1), loggingMode);

	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SPropValue),
		(LPVOID*)&lpProfPropVal), loggingMode);

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
	EC_HRES_LOG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		lpProfRes,
		NULL,
		0,
		&lpProfRows), loggingMode);

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
	EC_HRES_LOG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		MAPI_UNICODE,                    // Flags.
		&lpServiceAdmin), loggingMode);        // Pointer to new IMsgServiceAdmin.

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
		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpSvcRes), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl1), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl2), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal1), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal2), loggingMode);

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
		EC_HRES_LOG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			lpSvcRes,
			NULL,
			0,
			&lpSvcRows), loggingMode);

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
						EC_HRES_LOG(MAPIAllocateBuffer(
							sizeof(SRestriction),
							(LPVOID*)&lpProvRes), loggingMode);

						EC_HRES_LOG(MAPIAllocateBuffer(
							sizeof(SRestriction) * 2,
							(LPVOID*)&lpProvResLvl1), loggingMode);

						EC_HRES_LOG(MAPIAllocateBuffer(
							sizeof(SPropValue),
							(LPVOID*)&lpProvPropVal), loggingMode);

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
						EC_HRES_LOG(HrQueryAllRows(lpProvTable,
							(LPSPropTagArray)&sptaProvProps,
							lpProvRes,
							NULL,
							0,
							&lpProvRows), loggingMode);

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
															EC_HRES_LOG(lpProfSection->SetProps(1, pstPathW, NULL), loggingMode);
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
														EC_HRES_LOG(lpProfSection->SetProps(1, pstPathW, NULL), loggingMode);
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

HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszNewPath, bool bMoveFiles, LoggingMode loggingMode)
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

	EC_HRES_LOG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), loggingMode); // Pointer to new IProfAdmin
						// Get an IProfAdmin interface.

	EC_HRES_LOG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), loggingMode);

	// Allocate memory for the restriction
	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)&lpProfRes), loggingMode);

	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SRestriction) * 2,
		(LPVOID*)&lpProfResLvl1), loggingMode);

	EC_HRES_LOG(MAPIAllocateBuffer(
		sizeof(SPropValue),
		(LPVOID*)&lpProfPropVal), loggingMode);

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
	EC_HRES_LOG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		lpProfRes,
		NULL,
		0,
		&lpProfRows), loggingMode);

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
	EC_HRES_LOG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		MAPI_UNICODE,                    // Flags.
		&lpServiceAdmin), loggingMode);        // Pointer to new IMsgServiceAdmin.

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
		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpSvcRes), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl1), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)&lpsvcResLvl2), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal1), loggingMode);

		EC_HRES_LOG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)&lpSvcPropVal2), loggingMode);

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
		EC_HRES_LOG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			lpSvcRes,
			NULL,
			0,
			&lpSvcRows), loggingMode);

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
					EC_HRES_LOG(MAPIAllocateBuffer(
						sizeof(SRestriction),
						(LPVOID*)&lpProvRes), loggingMode);

					EC_HRES_LOG(MAPIAllocateBuffer(
						sizeof(SRestriction) * 2,
						(LPVOID*)&lpProvResLvl1), loggingMode);

					EC_HRES_LOG(MAPIAllocateBuffer(
						sizeof(SPropValue),
						(LPVOID*)&lpProvPropVal), loggingMode);

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
					EC_HRES_LOG(HrQueryAllRows(lpProvTable,
						(LPSPropTagArray)&sptaProvProps,
						lpProvRes,
						NULL,
						0,
						&lpProvRows), loggingMode);

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
															EC_HRES_LOG(lpProfSection->SetProps(1, pstPathW, NULL), loggingMode);
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
														EC_HRES_LOG(lpProfSection->SetProps(1, pstPathW, NULL), loggingMode);
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

