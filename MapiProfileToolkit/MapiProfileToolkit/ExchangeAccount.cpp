
#include "ExchangeAccount.h"


HRESULT HrCreateMsemsService(ULONG ulProifileMode, LPWSTR lpwszProfileName, int iOutlookVersion, ServiceOptions* pServiceOptions)
{
	HRESULT hRes = S_OK;

	if (ulProifileMode == PROFILEMODE_DEFAULT)
	{
		EC_HRES_MSG(HrCreateMsemsServiceOneProfile((LPWSTR)GetDefaultProfileName().c_str(), iOutlookVersion, pServiceOptions), L"Calling HrCreateMsemsServiceOneProfile");

	}
	else if (ulProifileMode == PROFILEMODE_ALL)
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo* profileInfo = new ProfileInfo[ulProfileCount];
		hRes = HrGetProfiles(ulProfileCount, profileInfo);
		for (ULONG i = 0; i <= ulProfileCount; i++)
		{
			EC_HRES_MSG(HrCreateMsemsServiceOneProfile((LPWSTR)profileInfo[i].wszProfileName.c_str(), iOutlookVersion, pServiceOptions), L"Calling HrCreateMsemsServiceOneProfile");
		}
	}
	else
	{
		if (ulProifileMode == PROFILEMODE_ONE)
		{
			EC_HRES_MSG(HrCreateMsemsServiceOneProfile(lpwszProfileName, iOutlookVersion, pServiceOptions), L"Calling HrCreateMsemsServiceOneProfile");
		}
		else
			Logger::Write(logLevelError, L"The specified profile name is invalid or no profile name was specified.\n");
	}

Error:
	return hRes;
}

HRESULT HrCreateMsemsServiceOneProfile(LPWSTR lpwszProfileName, int iOutlookVersion, ServiceOptions* pServiceOptions)
{
	HRESULT hRes = S_OK;
	switch (iOutlookVersion)
	{
	case 2007:
		Logger::Write(logLevelError, L"This client version is not currently supported.");
		break;
	case 2010:

		if (pServiceOptions->ulConnectMode == CONNECT_ROH)
		{
			// This id a bit of a hack since delegate mailboxes don't need to have the personalised server name in the delegate provider
			// I'm just creating these based on the legacyDN and the MailStore so best check that those have value
			Logger::Write(logLevelError, L"Validating delegate information.");

			Logger::Write(logLevelInfo, L"Creating and configuring new ROH service.");
			EC_HRES_MSG(HrCreateMsemsServiceROH(FALSE,
				lpwszProfileName,
				(LPWSTR)pServiceOptions->wszSmtpAddress.c_str(),
				(LPWSTR)pServiceOptions->wszMailboxLegacyDN.c_str(),
				(LPWSTR)pServiceOptions->wszUnresolvedServer.c_str(),
				(LPWSTR)pServiceOptions->wszRohProxyServer.c_str(),
				(LPWSTR)pServiceOptions->wszServerLegacyDN.c_str(),
				(LPWSTR)pServiceOptions->wszAutodiscoverUrl.c_str()), L"Calling HrCreateMsemsServiceROH");
		}
		// best not be used for now as I haven't sorted it out
		else if (pServiceOptions->ulConnectMode == CONNECT_MOH)
		{
			Logger::Write(logLevelError, L"Validating delegate information.");
			if (((pServiceOptions->wszMailStoreInternalUrl != L"") || (pServiceOptions->wszMailStoreExternalUrl != L"")) && (pServiceOptions->wszMailboxLegacyDN != L""))
			{
				//Logger::Write(logLevelError, L"MOH logic is not currently available.");
				std::wstring wszParsedSmtpAddress = SubstringToEnd(L"smtp:", pServiceOptions->wszSmtpAddress);
				std::wstring wszPersonalisedServerName;
				if (pServiceOptions->wszMailStoreInternalUrl != L"")
					wszPersonalisedServerName = SubstringToEnd(L"MailboxId=", pServiceOptions->wszMailStoreInternalUrl);
				else
					wszPersonalisedServerName = SubstringToEnd(L"MailboxId=", pServiceOptions->wszMailStoreExternalUrl);

				if ((pServiceOptions->wszAddressBookInternalUrl == L"") && (pServiceOptions->wszMailStoreInternalUrl != L""))
				{
					pServiceOptions->wszAddressBookInternalUrl = SubstringFromStart(L"emsmdb", pServiceOptions->wszMailStoreInternalUrl) + L"/nspi" + SubstringToEnd(L"emsmdb", pServiceOptions->wszMailStoreInternalUrl);
				}
				if ((pServiceOptions->wszAddressBookExternalUrl == L"") && (pServiceOptions->wszMailStoreExternalUrl != L""))
				{
					pServiceOptions->wszAddressBookExternalUrl = SubstringFromStart(L"emsmdb", pServiceOptions->wszMailStoreExternalUrl) + L"/nspi" + SubstringToEnd(L"emsmdb", pServiceOptions->wszMailStoreExternalUrl);
				}
				std::wstring wszServerDN = SubstringFromStart(L"cn=Recipients", pServiceOptions->wszMailboxLegacyDN) + L"/cn=Configuration/cn=Servers/cn=" + wszPersonalisedServerName;

				EC_HRES_MSG(HrCreateMsemsServiceMOH(FALSE,
					lpwszProfileName,
					(LPWSTR)pServiceOptions->wszSmtpAddress.c_str(),
					(LPWSTR)pServiceOptions->wszMailboxLegacyDN.c_str(),
					(LPWSTR)pServiceOptions->wszServerLegacyDN.c_str(),
					(LPWSTR)pServiceOptions->wszServerDisplayName.c_str(),
					(LPWSTR)pServiceOptions->wszMailStoreInternalUrl.c_str(),
					(LPWSTR)pServiceOptions->wszMailStoreExternalUrl.c_str(),
					(LPWSTR)pServiceOptions->wszAddressBookInternalUrl.c_str(),
					(LPWSTR)pServiceOptions->wszAddressBookExternalUrl.c_str()), L"Calling HrCreateMsemsServiceMOH");
			}
		}

		break;
	case 2013:
		if (pServiceOptions->ulConnectMode == CONNECT_ROH)
		{
			// This id a bit of a hack since delegate mailboxes don't need to have the personalised server name in the delegate provider
			// I'm just creating these based on the legacyDN and the MailStore so best check that those have value
			Logger::Write(logLevelError, L"Validating delegate information.");

			Logger::Write(logLevelInfo, L"Creating and configuring new ROH service.");
			EC_HRES_MSG(HrCreateMsemsServiceROH(FALSE,
				lpwszProfileName,
				(LPWSTR)pServiceOptions->wszSmtpAddress.c_str(),
				(LPWSTR)pServiceOptions->wszMailboxLegacyDN.c_str(),
				(LPWSTR)pServiceOptions->wszUnresolvedServer.c_str(),
				(LPWSTR)pServiceOptions->wszRohProxyServer.c_str(),
				(LPWSTR)pServiceOptions->wszServerLegacyDN.c_str(),
				(LPWSTR)pServiceOptions->wszAutodiscoverUrl.c_str()), L"Calling HrCreateMsemsServiceROH");
		}
		// best not be used for now as I haven't sorted it out
		else if (pServiceOptions->ulConnectMode == CONNECT_MOH)
		{
			Logger::Write(logLevelError, L"Validating delegate information.");
			if (((pServiceOptions->wszMailStoreInternalUrl != L"") || (pServiceOptions->wszMailStoreExternalUrl != L"")) && (pServiceOptions->wszMailboxLegacyDN != L""))
			{
				//Logger::Write(logLevelError, L"MOH logic is not currently available.");
				std::wstring wszParsedSmtpAddress = SubstringToEnd(L"smtp:", pServiceOptions->wszSmtpAddress);
				std::wstring wszPersonalisedServerName;
				if (pServiceOptions->wszMailStoreInternalUrl != L"")
					wszPersonalisedServerName = SubstringToEnd(L"MailboxId=", pServiceOptions->wszMailStoreInternalUrl);
				else
					wszPersonalisedServerName = SubstringToEnd(L"MailboxId=", pServiceOptions->wszMailStoreExternalUrl);

				if ((pServiceOptions->wszAddressBookInternalUrl == L"") && (pServiceOptions->wszMailStoreInternalUrl != L""))
				{
					pServiceOptions->wszAddressBookInternalUrl = SubstringFromStart(L"emsmdb", pServiceOptions->wszMailStoreInternalUrl) + L"/nspi" + SubstringToEnd(L"emsmdb", pServiceOptions->wszMailStoreInternalUrl);
				}
				if ((pServiceOptions->wszAddressBookExternalUrl == L"") && (pServiceOptions->wszMailStoreExternalUrl != L""))
				{
					pServiceOptions->wszAddressBookExternalUrl = SubstringFromStart(L"emsmdb", pServiceOptions->wszMailStoreExternalUrl) + L"/nspi" + SubstringToEnd(L"emsmdb", pServiceOptions->wszMailStoreExternalUrl);
				}
				std::wstring wszServerDN = SubstringFromStart(L"cn=Recipients", pServiceOptions->wszMailboxLegacyDN) + L"/cn=Configuration/cn=Servers/cn=" + wszPersonalisedServerName;


				EC_HRES_MSG(HrCreateMsemsServiceMOH(FALSE,
					lpwszProfileName,
					(LPWSTR)pServiceOptions->wszSmtpAddress.c_str(),
					(LPWSTR)pServiceOptions->wszMailboxLegacyDN.c_str(),
					(LPWSTR)pServiceOptions->wszServerLegacyDN.c_str(),
					(LPWSTR)pServiceOptions->wszServerDisplayName.c_str(),
					(LPWSTR)pServiceOptions->wszMailStoreInternalUrl.c_str(),
					(LPWSTR)pServiceOptions->wszMailStoreExternalUrl.c_str(),
					(LPWSTR)pServiceOptions->wszAddressBookInternalUrl.c_str(),
					(LPWSTR)pServiceOptions->wszAddressBookExternalUrl.c_str()), L"Calling HrCreateMsemsServiceMOH");
			}
		}

		break;
	case 2016:
		Logger::Write(logLevelInfo, L"Creating and configuring new service.");
		EC_HRES_MSG(HrCreateMsemsServiceModern(FALSE,
			lpwszProfileName,
			(LPWSTR)pServiceOptions->wszSmtpAddress.c_str(),
			(LPWSTR)pServiceOptions->wszSmtpAddress.c_str()), L"Calling HrCreateMsemsServiceModern");

		break;
	}

Error:
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
		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)& lpServiceAdmin2), L"Calling QueryInterface.");

		// Adds a message service to the current profile and returns that newly added service UID.
		EC_HRES_MSG(lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSEMS", (LPTSTR)"Microsoft Exchange", NULL, 0, &uidService), L"Calling CreateMsgServiceEx.");

		EC_HRES_MSG(lpServiceAdmin2->OpenProfileSection(&uidService,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpProfSect), L"Calling OpenProfileSection.");


		LPMAPIPROP lpMapiProp = NULL;
		EC_HRES_MSG(lpProfSect->QueryInterface(IID_IMAPIProp, (LPVOID*)& lpMapiProp), L"Calling QueryInterface.");

		if (lpMapiProp)
		{
			LPSPropValue prResourceFlags;
			MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)& prResourceFlags);

			prResourceFlags->ulPropTag = PR_RESOURCE_FLAGS;
			prResourceFlags->Value.l = ulResourceFlags;
			EC_HRES_MSG(lpMapiProp->SetProps(1, prResourceFlags, NULL), L"Calling SetProps.");

			EC_HRES_MSG(lpMapiProp->SaveChanges(FORCE_SAVE), L"Calling SaveChanges.");
			MAPIFreeBuffer(prResourceFlags);
			lpMapiProp->Release();
		}

		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)& lpEmsMdbProfSect);
		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)& lpStoreProviderSect);
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
	LPWSTR lpszDisplayName)
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

		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)& lpServiceAdmin2), L"Calling QueryInterface.");
		// Adds a message service to the current profile and returns that newly added service UID.
		EC_HRES_MSG(lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSEMS", (LPTSTR)"Microsoft Exchange", NULL, 0, &uidService), L"Calling CreateMsgServiceEx.");

		EC_HRES_MSG(lpServiceAdmin2->OpenProfileSection(&uidService,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpProfSect), L"Calling OpenProfileSection.");

		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)& lpEmsMdbProfSect);
		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)& lpStoreProviderSect);
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

		//ZeroMemory(&rgvalEmsMdbSect[1], sizeof(SPropValue));
		//rgvalEmsMdbSect[1].ulPropTag = PR_DISPLAY_NAME_W;
		//rgvalEmsMdbSect[1].Value.lpszW = lpszDisplayName;

		//ZeroMemory(&rgvalEmsMdbSect[2], sizeof(SPropValue));
		//rgvalEmsMdbSect[2].ulPropTag = PR_PROFILE_ACCT_NAME_W;
		//rgvalEmsMdbSect[2].Value.lpszW = lpszDisplayName;

		//ZeroMemory(&rgvalEmsMdbSect[3], sizeof(SPropValue));
		//rgvalEmsMdbSect[3].ulPropTag = PR_PROFILE_UNRESOLVED_NAME_W;
		//rgvalEmsMdbSect[3].Value.lpszW = lpszDisplayName;

		//ZeroMemory(&rgvalEmsMdbSect[4], sizeof(SPropValue));
		//rgvalEmsMdbSect[4].ulPropTag = PR_PROFILE_USER_EMAIL_W;
		//rgvalEmsMdbSect[4].Value.lpszW = lpszDisplayName;

		EC_HRES_MSG(lpEmsMdbProfSect->SetProps(
			1,
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

		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)& lpServiceAdmin2), L"Calling QueryInterface.");

		wprintf(L"Creating MsgService.\n");
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

		wprintf(L"Configuring MsgService.\n");
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
	wprintf(L"ERROR: hRes = %0x\n", hRes);

cleanup:
	// Clean up
	wprintf(L"Done cleaning up.\n");
	if (lpServiceAdmin2) lpServiceAdmin2->Release();
	if (lpServiceAdmin) lpServiceAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	return hRes;
}

// HrCreateMsemsServiceROH (Outlook 2010 and 2013)
//Creates a new message store service and sets it for RPC / HTTP with the following properties:
//	PR_PROFILE_USER
//	PR_DISPLAY_NAME_W
//	PR_PROFILE_UNRESOLVED_NAME_W
//	PR_PROFILE_HOME_SERVER
//	PR_PROFILE_HOME_SERVER_FQDN
//	PR_PROFILE_HOME_SERVER_DN
//	PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
//	PR_PROFILE_HOME_SERVER_ADDRS
//	PR_PROFILE_ACCT_NAME_W
//	PR_PROFILE_CONFIG_FLAGS
//	PR_PROFILE_TRANSPORT_FLAGS
//	PR_PROFILE_CONNECT_FLAGS
//	PR_PROFILE_UI_STATE
//	PR_PROFILE_AUTH_PACKAGE
//Configures the Store Provider with the following properties:
//	PR_PROFILE_SERVER
//	PR_PROFILE_SERVER_FQDN
//	PR_PROFILE_SERVER_DN
//	PR_PROFILE_MAILBOX
//	PR_DISPLAY_NAME_W
//	PR_PROFILE_DISPLAYNAME_SET
//	PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W 
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
	SPropValue rgvalEmsMdbSect[18]; // Property value structure to hold configuration info.
	SPropValue rgvalStoreProvider[7];
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

	// Validate parameters
	if (!lpszSmtpAddress || !lpszMailboxLegacyDn || !lpszUnresolvedServer || !lpszRohProxyServer || !lpszProfileServerDn || !lpszAutodiscoverUrl)
	{
		Logger::Write(LogLevel::logLevelFailed, L"Please provide a value for all of the following parameters: autodiscoverurl, mailboxlegacydn, rohproxyserver, serverlegacydn, smtpaddress, unresolvedserver");
		return MAPI_E_CANCEL;
	}
	// Enumeration for convenience.
	enum { iDispName, iSvcName, iSvcUID, iResourceFlags, iEmsMdbSectionUid, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_DISPLAY_NAME, PR_SERVICE_NAME, PR_SERVICE_UID, PR_RESOURCE_FLAGS, PR_EMSMDB_SECTION_UID };
	std::wstring wszSmtpAddress = ConvertWideCharToStdWstring(lpszSmtpAddress);
	wszSmtpAddress = L"SMTP:" + wszSmtpAddress;

	std::wstring wszServerName = ConvertWideCharToStdWstring(lpszUnresolvedServer);
	std::wstring wszncacn_http;
	std::wstring wszncacn_ip_tcp;
	wszncacn_http = L"ncacn_http:" + wszServerName;
	wszncacn_ip_tcp = L"wszncacn_ip_tcp:" + wszServerName;
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
		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)& lpServiceAdmin2), L"Calling QueryInterface.");

		wprintf(L"Creating MsgService.\n");
		// Adds a message service to the current profile and returns that newly added service UID.
		hRes = lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSEMS", (LPTSTR)"Microsoft Exchange", NULL, 0, &uidService);
		if (FAILED(hRes)) goto Error;

		wprintf(L"Configuring MsgService.\n");

		int paramC = 0;

		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)& lpEmsMdbProfSect);
		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)& lpStoreProviderSect);
		ZeroMemory(lpEmsMdbProfSect, sizeof(LPPROFSECT));
		ZeroMemory(lpStoreProviderSect, sizeof(LPPROFSECT));

		EC_HRES_MSG(HrGetSections(lpServiceAdmin2, &uidService, &lpEmsMdbProfSect, &lpStoreProviderSect), L"Calling HrGetSections.");

		// Set up a SPropValue array for the properties you need to configure.

		if (lpszMailboxLegacyDn)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_USER; // 0x6603
			rgvalEmsMdbSect[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszMailboxLegacyDn);
			paramC++;
		}

		if (lpszSmtpAddress)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_DISPLAY_NAME_W; // 0x3001
			rgvalEmsMdbSect[paramC].Value.lpszW = lpszSmtpAddress;
			paramC++;

			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_UNRESOLVED_NAME_W; // 0x66
			rgvalEmsMdbSect[paramC].Value.lpszW = lpszSmtpAddress;
			paramC++;

			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_ACCT_NAME_W;
			rgvalEmsMdbSect[paramC].Value.lpszW = lpszSmtpAddress;
			paramC++;

			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
			rgvalEmsMdbSect[paramC].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();
			paramC++;
		}

		if (lpszUnresolvedServer)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_HOME_SERVER;
			rgvalEmsMdbSect[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszUnresolvedServer);
			paramC++;

			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_HOME_SERVER_FQDN;
			rgvalEmsMdbSect[paramC].Value.lpszW = lpszUnresolvedServer;
			paramC++;

			LPSTR * lpszHomeServerValues = NULL;

			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_HOME_SERVER_ADDRS; // 6613
			rgvalEmsMdbSect[paramC].Value.MVszA.cValues = 2;

			MAPIAllocateBuffer(sizeof(LPSTR) * 2, (LPVOID*)&lpszHomeServerValues);
			lpszHomeServerValues[0] = ConvertWideCharToMultiByte((LPWSTR)wszncacn_http.c_str());
			lpszHomeServerValues[1] = ConvertWideCharToMultiByte((LPWSTR)wszncacn_ip_tcp.c_str());

			rgvalEmsMdbSect[paramC].Value.MVszA.lppszA = lpszHomeServerValues;
			paramC++;
		}

		if (lpszProfileServerDn)
		{
			ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
			rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_HOME_SERVER_DN;
			rgvalEmsMdbSect[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszProfileServerDn);
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
		rgvalEmsMdbSect[paramC].Value.l = RPC_C_HTTP_AUTHN_SCHEME_NEGOTIATE;
		paramC++;

		ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
		rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_AUTH_PACKAGE;
		rgvalEmsMdbSect[paramC].Value.l = RPC_C_AUTHN_GSS_NEGOTIATE;
		paramC++;

		ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
		rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_UI_STATE;
		rgvalEmsMdbSect[paramC].Value.l = 0;
		paramC++;

		ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
		rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_CONNECT_FLAGS;
		rgvalEmsMdbSect[paramC].Value.l = 0;
		paramC++;

		ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
		rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_TRANSPORT_FLAGS;
		rgvalEmsMdbSect[paramC].Value.l = TRANSPORT_DOWNLOAD | TRANSPORT_UPLOAD;
		paramC++;

		ZeroMemory(&rgvalEmsMdbSect[paramC], sizeof(SPropValue));
		rgvalEmsMdbSect[paramC].ulPropTag = PR_PROFILE_CONFIG_FLAGS;
		rgvalEmsMdbSect[paramC].Value.l = CONFIG_SHOW_CONNECT_UI | CONFIG_OST_CACHE_PRIVATE | CONFIG_OST_CACHE_DELEGATE_PIM;
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

		wprintf(L"Saving changes.\n");

		hRes = lpEmsMdbProfSect->SaveChanges(KEEP_OPEN_READWRITE);

		if (FAILED(hRes))
		{
			goto Error;
		}

		//Updating store provider 

		paramC = 0;

		if (lpStoreProviderSect)
		{
			ZeroMemory(&rgvalStoreProvider[paramC], sizeof(SPropValue));
			rgvalStoreProvider[paramC].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
			rgvalStoreProvider[paramC].Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();
			paramC++;

			ZeroMemory(&rgvalStoreProvider[paramC], sizeof(SPropValue));
			rgvalStoreProvider[paramC].ulPropTag = PR_DISPLAY_NAME_W;
			rgvalStoreProvider[paramC].Value.lpszW = lpszSmtpAddress;
			paramC++;

			ZeroMemory(&rgvalStoreProvider[paramC], sizeof(SPropValue));
			rgvalStoreProvider[paramC].ulPropTag = PR_PROFILE_DISPLAYNAME_SET;
			rgvalStoreProvider[paramC].Value.l = 1;
			paramC++;

			ZeroMemory(&rgvalStoreProvider[paramC], sizeof(SPropValue));
			rgvalStoreProvider[paramC].ulPropTag = PR_PROFILE_SERVER;
			rgvalStoreProvider[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszUnresolvedServer);
			paramC++;

			ZeroMemory(&rgvalStoreProvider[paramC], sizeof(SPropValue));
			rgvalStoreProvider[paramC].ulPropTag = PR_PROFILE_SERVER_FQDN;
			rgvalStoreProvider[paramC].Value.lpszW = lpszUnresolvedServer;
			paramC++;

			ZeroMemory(&rgvalStoreProvider[paramC], sizeof(SPropValue));
			rgvalStoreProvider[paramC].ulPropTag = PR_PROFILE_SERVER_DN;
			rgvalStoreProvider[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszProfileServerDn);
			paramC++;

			ZeroMemory(&rgvalStoreProvider[paramC], sizeof(SPropValue));
			rgvalStoreProvider[paramC].ulPropTag = PR_PROFILE_MAILBOX;
			rgvalStoreProvider[paramC].Value.lpszA = ConvertWideCharToMultiByte(lpszMailboxLegacyDn);
			paramC++;

			hRes = lpStoreProviderSect->SetProps(
				paramC,
				rgvalStoreProvider,
				nullptr);

			if (FAILED(hRes))
			{
				goto Error;
			}

			wprintf(L"Saving changes.\n");
			hRes = lpStoreProviderSect->SaveChanges(KEEP_OPEN_READWRITE);

			if (FAILED(hRes))
			{
				goto Error;
			}

		}

		goto cleanup;

	Error:
		wprintf(L"ERROR: hRes = %0x\n", hRes);
	}

cleanup:
	// Clean up
	if (lpStoreProviderSect) lpStoreProviderSect->Release();
	if (lpEmsMdbProfSect) lpEmsMdbProfSect->Release();
	if (lpProfSect) lpProfSect->Release();
	if (lpServiceAdmin2) lpServiceAdmin2->Release();
	if (lpServiceAdmin) lpServiceAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	wprintf(L"Done cleaning up.\n");
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
	LPWSTR lpszServerDn,
	LPWSTR lpszServerName,
	LPWSTR lpszMailStoreInternalUrl,
	LPWSTR lpszMailStoreExternalUrl,
	LPWSTR lpszAddressBookInternalUrl,
	LPWSTR lpszAddressBookExternalUrl)
{
	HRESULT hRes = S_OK; // Result code returned from MAPI calls.
	//	SPropValue rgvalEmsMdbSect[14]; // Property value structure to hold configuration info.
	// SPropValue rgvalStoreProvider[5];
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
	std::wstring wszSmtpAddress = SubstringToEnd(L"smtp:", ConvertWideCharToStdWstring(lpszSmtpAddress));
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

		EC_HRES_MSG(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)& lpServiceAdmin2), L"Calling QueryInterface.");

		wprintf(L"Creating MsgService.\n");

		// Adds a message service to the current profile and returns that newly added service UID.
		hRes = lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSEMS", (LPTSTR)"Microsoft Exchange", NULL, 0, &uidService);
		if (FAILED(hRes)) goto Error;

		EC_HRES_MSG(HrGetSections(lpServiceAdmin2, &uidService, &lpEmsMdbProfSect, &lpStoreProviderSect), L"Calling HrGetSections");

		int paramC = 0;
		std::vector<SPropValue> rgvalVector;
		SPropValue sPropValue;



		//Updating store provider 
		if (lpStoreProviderSect)
		{

			rgvalVector.resize(0);

			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
			sPropValue.Value.lpszW = (LPWSTR)wszSmtpAddress.c_str();
			rgvalVector.push_back(sPropValue);
			paramC++;

			if (lpszMailStoreExternalUrl && (lpszMailStoreExternalUrl != L""))
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL;
				sPropValue.Value.lpszW = lpszMailStoreExternalUrl;
				rgvalVector.push_back(sPropValue);
				paramC++;
			}

			if (lpszMailStoreInternalUrl && (lpszMailStoreInternalUrl != L""))
			{
				ZeroMemory(&sPropValue, sizeof(SPropValue));
				sPropValue.ulPropTag = PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL;
				sPropValue.Value.lpszW = lpszMailStoreInternalUrl;
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

			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PR_PROFILE_USER;
			sPropValue.Value.lpszA = ConvertWideCharToMultiByte(lpszMailboxDn);
			rgvalVector.push_back(sPropValue);
			paramC++;

			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PR_PROFILE_HOME_SERVER_DN;
			sPropValue.Value.lpszA = ConvertWideCharToMultiByte(lpszServerDn);
			rgvalVector.push_back(sPropValue);
			paramC++;

			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PR_PROFILE_HOME_SERVER;
			sPropValue.Value.lpszA = ConvertWideCharToMultiByte(lpszServerName);
			rgvalVector.push_back(sPropValue);
			paramC++;

			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PR_PROFILE_UNRESOLVED_SERVER;
			sPropValue.Value.lpszA = ConvertWideCharToMultiByte(lpszServerName);
			rgvalVector.push_back(sPropValue);
			paramC++;

			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PR_PROFILE_CONFIG_FLAGS;
			sPropValue.Value.l = CONFIG_PROMPT_FOR_CREDENTIALS | CONFIG_SHOW_CONNECT_UI;
			rgvalVector.push_back(sPropValue);
			paramC++;

			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PR_PROFILE_AUTH_PACKAGE;
			sPropValue.Value.l = RPC_C_AUTHN_GSS_NEGOTIATE;
			rgvalVector.push_back(sPropValue);
			paramC++;

			hRes = lpStoreProviderSect->SetProps(
				(ULONG)rgvalVector.size(),
				rgvalVector.data(),
				nullptr);

			if (FAILED(hRes))
			{
				goto Error;
			}

			wprintf(L"Saving changes.\n");
			hRes = lpStoreProviderSect->SaveChanges(KEEP_OPEN_READWRITE);

			if (FAILED(hRes))
			{
				goto Error;
			}


			//Updating emsmdb section 
			if (lpEmsMdbProfSect)
			{

				if (lpszAddressBookInternalUrl && (lpszAddressBookInternalUrl != L""))
				{
					ZeroMemory(&sPropValue, sizeof(SPropValue));
					sPropValue.ulPropTag = PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL;
					sPropValue.Value.lpszW = lpszAddressBookInternalUrl;
					rgvalVector.push_back(sPropValue);
					paramC++;
				}

				if (lpszAddressBookExternalUrl && (lpszAddressBookExternalUrl != L""))
				{
					ZeroMemory(&sPropValue, sizeof(SPropValue));
					sPropValue.ulPropTag = PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL;
					sPropValue.Value.lpszW = lpszAddressBookExternalUrl;
					rgvalVector.push_back(sPropValue);
					paramC++;
				}

				//ZeroMemory(&sPropValue, sizeof(SPropValue));
				//sPropValue.ulPropTag = PR_PROFILE_MAILBOX;
				//sPropValue.Value.lpszA = ConvertWideCharToMultiByte(lpszMailboxDn);
				//rgvalVector.push_back(sPropValue);
				//paramC++;

				hRes = lpServiceAdmin2->ConfigureMsgService(&uidService,
					NULL,
					0,
					(ULONG)rgvalVector.size(),
					rgvalVector.data());

				if (FAILED(hRes))
				{
					goto Error;
				}

				if (FAILED(hRes))
				{
					goto Error;
				}

			}
		}
	}
	goto cleanup;


Error:
	wprintf(L"ERROR: hRes = %0x\n", hRes);

cleanup:
	// Clean up
	if (lpStoreProviderSect) lpStoreProviderSect->Release();
	if (lpEmsMdbProfSect) lpEmsMdbProfSect->Release();
	if (lpProfSect) lpProfSect->Release();
	if (lpServiceAdmin2) lpServiceAdmin2->Release();
	if (lpServiceAdmin) lpServiceAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	wprintf(L"Done cleaning up.\n");
	return hRes;
}

// HrGetDefaultMsemsServiceAdminProviderPtr
// Returns the provider admin interface pointer for the default service in a given profile
HRESULT HrGetDefaultMsemsServiceAdminProviderPtr(LPWSTR lpwszProfileName, LPPROVIDERADMIN* lppProvAdmin, LPMAPIUID* lppServiceUid)
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
			(LPVOID*)& lpSvcRes), L"Calling #");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)& lpsvcResLvl1), L"Calling #");

		EC_HRES_MSG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)& lpSvcPropVal), L"Calling #");

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

HRESULT HrUpdatePrStoreProviders(LPSERVICEADMIN lpServiceAdmin, LPMAPIUID lpServiceUid, LPMAPIUID lpProviderUid)
{
	HRESULT hRes = S_OK;

	SizedSPropTagArray(1, sptGlobal) = { 1, PR_STORE_PROVIDERS };
	LPPROFSECT lpEmsMdbSection = NULL;
	LPPROFSECT lpStoreProvSection = NULL;
	LPSPropValue lpGlobalVals = NULL; // Property value struct pointer for global profile section.
	ULONG ulProps = 0; // Count of props.
	ULONG cbNewBuffer = 0;
	SPropValue NewVals;

	EC_HRES_MSG(HrGetSections(lpServiceAdmin, lpServiceUid, &lpEmsMdbSection, &lpStoreProvSection), L"Calling HrGetSections");

	if (lpEmsMdbSection)
	{
		LPSPropValue lpPrStoreProviders = NULL;

		// Get the list of store providers in PR_STORE_PROVIDERS.
		EC_HRES_MSG(lpEmsMdbSection->GetProps((LPSPropTagArray)& sptGlobal,
			0,
			&ulProps,
			&lpGlobalVals), L"Calling GetProps");

		wprintf(L"Got the list of mailboxes being opened.\n");

		// Now we set up an SPropValue structure with the original
		// list + the UID of the new service.

		// Compute the new byte count
		cbNewBuffer = sizeof(MAPIUID) + lpGlobalVals->Value.bin.cb;

		// Allocate space for the new list of UIDs.
		hRes = MAPIAllocateBuffer(cbNewBuffer,
			(LPVOID*)& NewVals.Value.bin.lpb);

		wprintf(L"Allocated buffer to hold new list of mailboxes to be opened.\n");

		// Copy the bits into the list.
		// First, copy the existing list.
		memcpy(NewVals.Value.bin.lpb,
			lpGlobalVals->Value.bin.lpb,
			lpGlobalVals->Value.bin.cb);

		// Next, copy the new UID onto the end of the list.
		memcpy(NewVals.Value.bin.lpb + lpGlobalVals->Value.bin.cb,
			lpProviderUid,
			sizeof(MAPIUID));
		wprintf(L"Concatenated list of mailboxes and new mailbox.\n");

		// Set the count of bytes on the SPropValue variable.
		NewVals.Value.bin.cb = cbNewBuffer;
		// Initialize dwAlignPad.
		NewVals.dwAlignPad = 0;
		// Set the prop tag.
		NewVals.ulPropTag = PR_STORE_PROVIDERS;

		// Set the property on the global profile section.
		hRes = lpEmsMdbSection->SetProps(ulProps,
			&NewVals,
			NULL);
	}


Error:
	goto Cleanup;

Cleanup:
	// Clean up.
	// Free up memory
	// To do: free up memory here
	if (lpEmsMdbSection) lpEmsMdbSection->Release();
	if (lpStoreProvSection) lpStoreProvSection->Release();
	return hRes;
}
