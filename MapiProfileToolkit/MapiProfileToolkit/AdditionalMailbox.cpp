
#include "AdditionalMailbox.h"


// HrAddDelegateMailboxModern
// Adds a delegate mailbox to a given service. The property set is one for Outlook 2016 where all is needed is:
// - the SMTP address of the mailbox
// - the Display Name for the mailbox
HRESULT HrAddDelegateMailboxModern(
	BOOL bDefaultProfile,
	LPWSTR lpwszProfileName,
	BOOL bDefaultService,
	ULONG iServiceIndex,
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
		MAPIUID muidProviderUid = { 0 };
		LPMAPIUID lpProviderUid = &muidProviderUid;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID, PR_RESOURCE_FLAGS };

		// Allocate memory for the restriction
		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)& lpSvcRes), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)& lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)& lpSvcPropVal), L"Calling MAPIAllocateBuffer");

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
			(LPSPropTagArray)& sptaSvcProps,
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
						rgval[0].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();

						ZeroMemory(&rgval[1], sizeof(SPropValue));
						rgval[1].ulPropTag = PR_DISPLAY_NAME_W;
						rgval[1].Value.lpszW = lpszwDisplayName;

						// Create the message service with the above properties.
						EC_HRES_MSG(lpProvAdmin->CreateProvider(LPWSTR("EMSDelegate"),
							2,
							rgval,
							0,
							0,
							lpProviderUid), L"Calling CreateProvider");

						EC_HRES_MSG(HrUpdatePrStoreProviders(lpServiceAdmin, (LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb, &muidProviderUid), L"Calling HrComputePrStoreProviders");

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
				rgval[0].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();

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



HRESULT HrAddDelegateMailbox(ULONG ulProifileMode, LPWSTR lpwszProfileName, ULONG ulServiceMode, int iServiceIndex, int iOutlookVersion, MailboxOptions* pMailboxOptions)
{
	HRESULT hRes = S_OK;

	if (ulProifileMode == PROFILEMODE_DEFAULT)
	{
		EC_HRES_MSG(HrAddDelegateMailboxOneProfile((LPWSTR)GetDefaultProfileName().c_str(), iOutlookVersion, ulServiceMode, iServiceIndex, pMailboxOptions), L"Calling HrAddDelegateMailboxOneProfile");

	}
	else if (ulProifileMode == PROFILEMODE_ALL)
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo* profileInfo = new ProfileInfo[ulProfileCount];
		hRes = HrGetProfiles(ulProfileCount, profileInfo);
		for (ULONG i = 0; i <= ulProfileCount; i++)
		{
			EC_HRES_MSG(HrAddDelegateMailboxOneProfile((LPWSTR)profileInfo[i].wszProfileName.c_str(), iOutlookVersion, ulServiceMode, iServiceIndex, pMailboxOptions), L"Calling HrAddDelegateMailboxOneProfile");
		}
	}
	else
	{
		if (ulProifileMode == PROFILEMODE_ONE)
		{
			EC_HRES_MSG(HrAddDelegateMailboxOneProfile(lpwszProfileName, iOutlookVersion, ulServiceMode, iServiceIndex, pMailboxOptions), L"Calling HrAddDelegateMailboxOneProfile");
		}
		else
			Logger::Write(logLevelError, L"The specified profile name is invalid or no profile name was specified.\n");
	}

Error:
	return hRes;
}

HRESULT HrAddDelegateMailboxOneProfile(LPWSTR lpwszProfileName, int iOutlookVersion, ULONG ulServiceMode, int iServiceIndex, MailboxOptions* pMailboxOptions)
{
	HRESULT hRes = S_OK;
	switch (iOutlookVersion)
	{
	case 2007:
		EC_HRES_MSG(HrAddDelegateMailboxLegacy(FALSE,
			lpwszProfileName,
			SERVICEMODE_DEFAULT,
			iServiceIndex,
			(LPWSTR)pMailboxOptions->wszMailboxDisplayName.c_str(),
			(LPWSTR)pMailboxOptions->wszMailboxLegacyDN.c_str(),
			(LPWSTR)pMailboxOptions->wszServerDisplayName.c_str(),
			(LPWSTR)pMailboxOptions->wszServerLegacyDN.c_str()), L"Calling HrAddDelegateMailboxLegacy");
		break;
	case 2010:
	case 2013:
		EC_HRES_MSG(HrAddDelegateMailbox(FALSE,
			lpwszProfileName,
			ulServiceMode == SERVICEMODE_DEFAULT,
			iServiceIndex,
			(LPWSTR)pMailboxOptions->wszMailboxDisplayName.c_str(),
			(LPWSTR)pMailboxOptions->wszMailboxLegacyDN.c_str(),
			(LPWSTR)pMailboxOptions->wszServerDisplayName.c_str(),
			(LPWSTR)pMailboxOptions->wszServerLegacyDN.c_str(),
			(LPWSTR)pMailboxOptions->wszSmtpAddress.c_str(),
			(LPWSTR)pMailboxOptions->wszRohProxyServer.c_str(),
			pMailboxOptions->ulRohProxyServerFlags,
			pMailboxOptions->ulRohProxyServerAuthPackage,
			(LPWSTR)pMailboxOptions->wszMailStoreInternalUrl.c_str()), L"Calling HrAddDelegateMailbox");
		break;
	case 2016:
		EC_HRES_MSG(HrAddDelegateMailboxModern(FALSE,
			lpwszProfileName,
			ulServiceMode == SERVICEMODE_DEFAULT,
			iServiceIndex,
			(LPWSTR)pMailboxOptions->wszMailboxDisplayName.c_str(),
			(LPWSTR)pMailboxOptions->wszSmtpAddress.c_str()), L"Calling HrCreateMsemsServiceModern");
		break;
	}

Error:
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
	ULONG ulServiceIndex,
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
			(LPVOID*)& lpSvcRes), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)& lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)& lpSvcPropVal), L"Calling MAPIAllocateBuffer");

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
			(LPSPropTagArray)& sptaSvcProps,
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
						rgval[4].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();

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
		else if (lpSvcRows->cRows >= ulServiceIndex)
		{
			LPPROVIDERADMIN lpProvAdmin = NULL;
			if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[ulServiceIndex].lpProps[iServiceUid].Value.bin.lpb,
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
				rgval[4].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();

				printf("Creating EMSDelegate provider.\n");
				// Create the message service with the above properties.
				hRes = lpProvAdmin->CreateProvider(LPWSTR("EMSDelegate"),
					5,
					rgval,
					0,
					0,
					(LPMAPIUID)lpSvcRows->aRow[ulServiceIndex].lpProps[iServiceUid].Value.bin.lpb);
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
	ULONG ulServiceIndex,
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
		MAPIUID muidProviderUid = { 0 };
		LPMAPIUID lpProviderUid = &muidProviderUid;
		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID, PR_RESOURCE_FLAGS };

		// Allocate memory for the restriction
		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)& lpSvcRes), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)& lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)& lpSvcPropVal), L"Calling MAPIAllocateBuffer");

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
			(LPSPropTagArray)& sptaSvcProps,
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
						EC_HRES_MSG(lpProvAdmin->CreateProvider(LPWSTR("EMSDelegate"),
							4,
							rgval,
							0,
							0,
							lpProviderUid), L"Calling CreateProvider");
						if (FAILED(hRes)) goto Error;

						EC_HRES_MSG(HrUpdatePrStoreProviders(lpServiceAdmin, (LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb, &muidProviderUid), L"Calling HrComputePrStoreProviders");
					}
				}
			}
			if (lpSvcRows) FreeProws(lpSvcRows);
		}
		else if (lpSvcRows->cRows >= ulServiceIndex)
		{
			LPPROVIDERADMIN lpProvAdmin = NULL;
			if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[ulServiceIndex].lpProps[iServiceUid].Value.bin.lpb,
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
					(LPMAPIUID)lpSvcRows->aRow[ulServiceIndex].lpProps[iServiceUid].Value.bin.lpb);
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

HRESULT HrPromoteDelegates(LPWSTR lpwszProfileName, BOOL bDefaultProfile, BOOL bAllProfiles, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ULONG ulConnectMode)
{
	HRESULT hRes = S_OK;

	if (bDefaultProfile)
	{
		ProfileInfo* profileInfo = new ProfileInfo();
		EC_HRES_MSG(HrGetProfile((LPWSTR)GetDefaultProfileName().c_str(), profileInfo), L"Calling GetProfile");
		EC_HRES_MSG(HrPromoteDelegatesOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultService, bAllServices, iOutlookVersion, ulConnectMode), L"Calling HrPromoteDelegatesOneProfile");

	}
	else if (bAllProfiles)
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo* profileInfo = new ProfileInfo[ulProfileCount];
		hRes = HrGetProfiles(ulProfileCount, profileInfo);
		for (ULONG i = 0; i <= ulProfileCount; i++)
		{
			EC_HRES_MSG(HrPromoteDelegatesOneProfile((LPWSTR)profileInfo[i].wszProfileName.c_str(), &profileInfo[i], iServiceIndex, bDefaultService, bAllServices, iOutlookVersion, ulConnectMode), L"Calling HrPromoteDelegatesOneProfile");
		}
	}
	else
	{
		if (lpwszProfileName)
		{
			ProfileInfo* profileInfo = new ProfileInfo();
			hRes = HrGetProfile(lpwszProfileName, profileInfo);
			EC_HRES_MSG(HrPromoteDelegatesOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultService, bAllServices, iOutlookVersion, ulConnectMode), L"Calling HrPromoteDelegatesOneProfile");

		}
		else
			Logger::Write(logLevelError, L"The specified profile name is invalid or no profile name was specified.\n");
	}

Error:
	return hRes;
}

HRESULT HrPromoteDelegatesOneProfile(LPWSTR lpwszProfileName, ProfileInfo* pProfileInfo, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, int iOutlookVersion, ULONG ulConnectMode)
{
	HRESULT hRes = S_OK;

	for (ULONG i = 0; i <= pProfileInfo->ulServiceCount; i++)
	{
		if (bDefaultService)
		{
			if (pProfileInfo->profileServices[i].bDefaultStore)
			{
				if (pProfileInfo->profileServices[i].ulServiceType == SERVICETYPE_MAILBOX)
				{
					for (ULONG j = 0; j <= pProfileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount; j++)
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
				for (ULONG j = 0; j <= pProfileInfo->profileServices[iServiceIndex].exchangeAccountInfo->ulMailboxCount; j++)
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
				for (ULONG j = 0; j <= pProfileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount; j++)
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
			if (((mailboxInfo.wszMailStoreInternalUrl != L"") || (mailboxInfo.wszMailStoreExternalUrl != L"")) && (mailboxInfo.wszProfileMailbox != L""))
			{
				std::wstring wszParsedSmtpAddress = SubstringToEnd(L"smtp:", mailboxInfo.wszSmtpAddress);
				std::wstring wszPersonalisedServerName;
				if ((mailboxInfo.wszMailStoreInternalUrl != L""))
					wszPersonalisedServerName = SubstringToEnd(L"MailboxId=", mailboxInfo.wszMailStoreInternalUrl);
				else
					wszPersonalisedServerName = SubstringToEnd(L"MailboxId=", mailboxInfo.wszMailStoreExternalUrl);

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
			Logger::Write(logLevelError, L"Validating delegate information.");
			if (((mailboxInfo.wszMailStoreInternalUrl != L"") || (mailboxInfo.wszMailStoreExternalUrl != L"")) && (mailboxInfo.wszProfileMailbox != L""))
			{
				//Logger::Write(logLevelError, L"MOH logic is not currently available.");
				std::wstring wszParsedSmtpAddress = SubstringToEnd(L"smtp:", mailboxInfo.wszSmtpAddress);
				std::wstring wszPersonalisedServerName;
				if (mailboxInfo.wszMailStoreInternalUrl != L"")
					wszPersonalisedServerName = SubstringToEnd(L"MailboxId=", mailboxInfo.wszMailStoreInternalUrl);
				else
					wszPersonalisedServerName = SubstringToEnd(L"MailboxId=", mailboxInfo.wszMailStoreExternalUrl);

				if ((mailboxInfo.wszAddressBookInternalUrl == L"") && (mailboxInfo.wszMailStoreInternalUrl != L""))
				{
					mailboxInfo.wszAddressBookInternalUrl = SubstringFromStart(L"emsmdb", mailboxInfo.wszMailStoreInternalUrl) + L"/nspi" + SubstringToEnd(L"emsmdb", mailboxInfo.wszMailStoreInternalUrl);
				}
				if ((mailboxInfo.wszAddressBookExternalUrl == L"") && (mailboxInfo.wszMailStoreExternalUrl != L""))
				{
					mailboxInfo.wszAddressBookExternalUrl = SubstringFromStart(L"emsmdb", mailboxInfo.wszMailStoreExternalUrl) + L"/nspi" + SubstringToEnd(L"emsmdb", mailboxInfo.wszMailStoreExternalUrl);
				}
				std::wstring wszServerDN = SubstringFromStart(L"cn=Recipients", mailboxInfo.wszProfileMailbox) + L"/cn=Configuration/cn=Servers/cn=" + wszPersonalisedServerName;


				if (SUCCEEDED(HrCreateMsemsServiceMOH(FALSE,
					lpwszProfileName,
					(LPWSTR)wszParsedSmtpAddress.c_str(),
					(LPWSTR)mailboxInfo.wszProfileMailbox.c_str(),
					(LPWSTR)wszServerDN.c_str(),
					(LPWSTR)mailboxInfo.wszProfileServerFqdnW.c_str(),
					(LPWSTR)mailboxInfo.wszMailStoreInternalUrl.c_str(),
					(LPWSTR)mailboxInfo.wszMailStoreExternalUrl.c_str(),
					(LPWSTR)mailboxInfo.wszAddressBookInternalUrl.c_str(),
					(LPWSTR)mailboxInfo.wszAddressBookExternalUrl.c_str())))
				{
					EC_HRES_MSG(HrDeleteProvider(lpwszProfileName, &mailboxInfo.muidServiceUid, &mailboxInfo.muidProviderUid), L"Calling HrDeleteProvider");
				}
			}
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
	return hRes;
}