#include "stdafx.h"

#include "AccountFunctions.h"
#include "ADHelper.h"

HRESULT GetDefaultAccount(LPMAPISESSION lpSession, LPWSTR lpwszProfile, OlkAccount * pOlkAccount)
{
	HRESULT hRes = S_OK;
	LPOLKACCOUNTMANAGER lpAcctMgr = NULL;
	LPUNKNOWN lpUnk = NULL;
	LPOLKACCOUNT lpAccount = NULL;
	hRes = CoCreateInstance(CLSID_OlkAccountManager,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		(LPVOID*)&lpAcctMgr);
	if (SUCCEEDED(hRes) && lpAcctMgr)
	{
		CAccountHelper* pMyAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if (pMyAcctHelper)
		{
			LPOLKACCOUNTHELPER lpAcctHelper = NULL;
			hRes = pMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, (LPVOID*)&lpAcctHelper);
			if (SUCCEEDED(hRes) && lpAcctHelper)
			{
				hRes = lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS);
				if (SUCCEEDED(hRes))
				{
					LPOLKENUM lpAcctEnum = NULL;
					hRes = lpAcctMgr->EnumerateAccounts(&CLSID_OlkMail,
						NULL,
						OLK_ACCOUNT_NO_FLAGS,
						&lpAcctEnum);
					if (SUCCEEDED(hRes) && lpAcctEnum)
					{
						DWORD cAccounts = 0;

						hRes = lpAcctEnum->GetCount(&cAccounts);
						if (SUCCEEDED(hRes) && cAccounts)
						{
							hRes = lpAcctEnum->Reset();
							if (SUCCEEDED(hRes))
							{
								DWORD i = 0;
								for (i = 0; i < cAccounts; i++)
								{
									hRes = lpAcctEnum->GetNext(&lpUnk);
									if (SUCCEEDED(hRes) && lpUnk)
									{
										hRes = lpUnk->QueryInterface(IID_IOlkAccount, (LPVOID*)&lpAccount);
										if (SUCCEEDED(hRes) && lpAccount)
										{
											ACCT_VARIANT pProp = { 0 };
											//PROP_ACCT_IS_DEFAULT_MAIL
											hRes = lpAccount->GetProp(PROP_ACCT_IS_DEFAULT_MAIL, &pProp);
											if (SUCCEEDED(hRes) && (pProp.Val.dw == 1))
											{
												GetAccountData(&lpAccount, pOlkAccount);
												break;
											}
											else
												hRes = S_OK;
										}
									}
								}
							}
						}
					}
					if (lpAccount)
						lpAccount->Release();
					lpAccount = NULL;

					if (lpUnk)
						lpUnk->Release();
					lpUnk = NULL;

					if (lpAcctEnum)
						lpAcctEnum->Release();
				}
			}

			if (lpAcctHelper)
				lpAcctHelper->Release();
		}

		if (pMyAcctHelper)
			pMyAcctHelper->Release();
	}

	if (lpAcctMgr)
		lpAcctMgr->Release();

	return hRes;
}

HRESULT UpdateAcctName(LPMAPISESSION lpSession, LPWSTR lpwszProfile, long lAcctId, LPWSTR lpszNewAcctName)
{
	HRESULT hRes = S_OK;
	ACCT_VARIANT pProp;
	pProp.dwType = PT_UNICODE;
	pProp.Val.pwsz = lpszNewAcctName;

	ACCT_VARIANT pAcctId;
	pAcctId.dwType = PT_LONG;
	pAcctId.Val.dw = lAcctId;

	LPOLKACCOUNTMANAGER lpAcctMgr = NULL;
	LPUNKNOWN lpUnk = NULL;
	LPOLKACCOUNT lpAccount = NULL;
	hRes = CoCreateInstance(CLSID_OlkAccountManager,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		(LPVOID*)&lpAcctMgr);
	if (SUCCEEDED(hRes) && lpAcctMgr)
	{
		CAccountHelper* pMyAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if (pMyAcctHelper)
		{
			LPOLKACCOUNTHELPER lpAcctHelper = NULL;
			hRes = pMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, (LPVOID*)&lpAcctHelper);
			if (SUCCEEDED(hRes) && lpAcctHelper)
			{
				hRes = lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS);
				if (SUCCEEDED(hRes))
				{
					LPOLKACCOUNT lpAccount = NULL;
					hRes = lpAcctMgr->FindAccount(PROP_ACCT_ID, &pAcctId, &lpAccount);
					if (pProp.Val.pwsz)
					{
						wprintf(L"Updating account name.\n");
						EC_HRES(lpAccount->SetPropW(PROP_ACCT_NAME, &pProp));
						EC_HRES(lpAccount->SaveChanges(OLK_ACCOUNT_NO_FLAGS));
					}

					if (lpAccount)
						lpAccount->Release();
					lpAccount = NULL;

				}

				if (lpUnk)
					lpUnk->Release();
				lpUnk = NULL;

			}

			if (lpAcctHelper)
				lpAcctHelper->Release();
		}

		if (pMyAcctHelper)
			pMyAcctHelper->Release();
	}

	if (lpAcctMgr)
		lpAcctMgr->Release();

Error:
	goto Cleanup;
Cleanup:
	return hRes;

}

HRESULT GetAccountData(LPOLKACCOUNT* lpAccount, OlkAccount* pOlkAccount)
{

	HRESULT hRes = S_OK;

	if (lpAccount)
	{
		ACCT_VARIANT pProp = { 0 };
		//Account ID PROP_ACCT_ID
		hRes = (*lpAccount)->GetProp(PROP_ACCT_ID, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.dw)
		{
			pOlkAccount->lAcctId = pProp.Val.dw;
		}
		else
			hRes = S_OK;

		//Account Name PROP_ACCT_NAME
		hRes = (*lpAccount)->GetProp(PROP_ACCT_NAME, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szAcctName = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_MINI_UID
		hRes = (*lpAccount)->GetProp(PROP_ACCT_MINI_UID, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.dw)
		{
			pOlkAccount->lAcctMiniUid = pProp.Val.dw;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_TYPE
		hRes = (*lpAccount)->GetProp(PROP_ACCT_TYPE, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szAcctType = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_IDENTITY
		hRes = (*lpAccount)->GetProp(PROP_ACCT_IDENTITY, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szAcctIdentity = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_FLAVOR
		hRes = (*lpAccount)->GetProp(PROP_ACCT_FLAVOR, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szAcctFlavor = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_IS_DEFAULT_MAIL
		hRes = (*lpAccount)->GetProp(PROP_ACCT_IS_DEFAULT_MAIL, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.dw)
		{
			pOlkAccount->lAcctIsDefaultMail = pProp.Val.dw;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_USER_DISPLAY_NAME
		hRes = (*lpAccount)->GetProp(PROP_ACCT_USER_DISPLAY_NAME, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szAcctUserDisplayName = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_USER_EMAIL_ADDR
		hRes = (*lpAccount)->GetProp(PROP_ACCT_USER_EMAIL_ADDR, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szAcctUserEmailAddr = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_STAMP
		hRes = (*lpAccount)->GetProp(PROP_ACCT_STAMP, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szAcctStamp = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_SEND_STAMP
		hRes = (*lpAccount)->GetProp(PROP_ACCT_SEND_STAMP, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szAcctSendStamp = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_IS_EXCH
		hRes = (*lpAccount)->GetProp(PROP_ACCT_IS_EXCH, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.dw)
		{
			pOlkAccount->lAcctIsExch = pProp.Val.dw;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_DISABLED
		hRes = (*lpAccount)->GetProp(PROP_ACCT_DISABLED, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.dw)
		{
			pOlkAccount->lAcctDisabled = pProp.Val.dw;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_DISABLED
		hRes = (*lpAccount)->GetProp(PROP_ACCT_DISABLED, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.dw)
		{
			pOlkAccount->lAcctDisabled = pProp.Val.dw;
		}
		else
			hRes = S_OK;

		//PROP_ACCT_PREFERENCES_UID
		hRes = (*lpAccount)->GetProp(PROP_ACCT_PREFERENCES_UID, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.dw)
		{
			pOlkAccount->lAcctPreferencesUid = pProp.Val.dw;
		}
		else
			hRes = S_OK;

		/******************************************/

		//PROP_MAPI_SERVICE_UID
		hRes = (*lpAccount)->GetProp(PROP_MAPI_SERVICE_UID, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.bin.pb)
		{
			pOlkAccount->bMapiServiceUid.cb = pProp.Val.bin.cb;
			pOlkAccount->bMapiServiceUid.pb = pProp.Val.bin.pb;
		}
		else
			hRes = S_OK;

		//PROP_MAPI_PROVIDER_TYPE
		hRes = (*lpAccount)->GetProp(PROP_MAPI_PROVIDER_TYPE, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.dw)
		{
			pOlkAccount->lMapiProviderType = pProp.Val.dw;
		}
		else
			hRes = S_OK;

		//PROP_MAPI_IDENTITY_ENTRYID
		hRes = (*lpAccount)->GetProp(PROP_MAPI_IDENTITY_ENTRYID, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.bin.pb)
		{
			pOlkAccount->bMapiIdentityEntryId.cb = pProp.Val.bin.cb;
			pOlkAccount->bMapiIdentityEntryId.pb = pProp.Val.bin.pb;
		}
		else
			hRes = S_OK;

		/***************************************************************/

		//PR_PRIMARY_SEND_ACCT
		hRes = (*lpAccount)->GetProp(PR_PRIMARY_SEND_ACCT, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szPrimarySendAcct = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

		//PR_NEXT_SEND_ACCT
		hRes = (*lpAccount)->GetProp(PR_NEXT_SEND_ACCT, &pProp);
		if (SUCCEEDED(hRes) && pProp.Val.pwsz)
		{
			pOlkAccount->szNextSendAcct = pProp.Val.pwsz;
		}
		else
			hRes = S_OK;

	}

	return hRes;
}

HRESULT GetAccounts(LPWSTR lpwszProfile, DWORD* pcAccounts, OlkAccount** ppAccounts)
{
	HRESULT hRes = S_OK;
	LPMAPISESSION lpSession;
	LPOLKACCOUNTMANAGER lpAcctMgr = NULL;

	hRes = MAPILogonEx(0,
		(LPTSTR)lpwszProfile,
		NULL,
		fMapiUnicode | MAPI_EXTENDED | MAPI_EXPLICIT_PROFILE |
		MAPI_NEW_SESSION | MAPI_NO_MAIL | MAPI_LOGON_UI,
		&lpSession);
	if (FAILED(hRes))
	{
		MessageBox(NULL, L"Failed to login to the selected profile", L"Error", 0);
	}

	hRes = CoCreateInstance(CLSID_OlkAccountManager,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IOlkAccountManager,
		(LPVOID*)&lpAcctMgr);
	if (SUCCEEDED(hRes) && lpAcctMgr)
	{
		CAccountHelper* pMyAcctHelper = new CAccountHelper(lpwszProfile, lpSession);
		if (pMyAcctHelper)
		{
			LPOLKACCOUNTHELPER lpAcctHelper = NULL;
			hRes = pMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, (LPVOID*)&lpAcctHelper);
			if (SUCCEEDED(hRes) && lpAcctHelper)
			{
				hRes = lpAcctMgr->Init(lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS);
				if (SUCCEEDED(hRes))
				{
					LPOLKENUM lpAcctEnum = NULL;

					hRes = lpAcctMgr->EnumerateAccounts(&CLSID_OlkMail,
						NULL,
						OLK_ACCOUNT_NO_FLAGS,
						&lpAcctEnum);
					if (SUCCEEDED(hRes) && lpAcctEnum)
					{
						DWORD cAccounts = 0;

						hRes = lpAcctEnum->GetCount(&cAccounts);
						if (SUCCEEDED(hRes) && cAccounts)
						{
							OlkAccount* pAccounts = new OlkAccount[cAccounts];

							hRes = lpAcctEnum->Reset();
							if (SUCCEEDED(hRes))
							{
								DWORD i = 0;
								for (i = 0; i < cAccounts; i++)
								{
									LPUNKNOWN lpUnk = NULL;

									hRes = lpAcctEnum->GetNext(&lpUnk);
									if (SUCCEEDED(hRes) && lpUnk)
									{
										LPOLKACCOUNT lpAccount = NULL;

										hRes = lpUnk->QueryInterface(IID_IOlkAccount, (LPVOID*)&lpAccount);
										if (SUCCEEDED(hRes) && lpAccount)
										{
											ACCT_VARIANT pProp = { 0 };
											//Account ID PROP_ACCT_ID
											hRes = lpAccount->GetProp(PROP_ACCT_ID, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.dw)
											{
												pAccounts[i].lAcctId = pProp.Val.dw;
											}
											else
												hRes = S_OK;

											//Account Name PROP_ACCT_NAME
											hRes = lpAccount->GetProp(PROP_ACCT_NAME, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szAcctName = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_MINI_UID
											hRes = lpAccount->GetProp(PROP_ACCT_MINI_UID, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.dw)
											{
												pAccounts[i].lAcctMiniUid = pProp.Val.dw;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_TYPE
											hRes = lpAccount->GetProp(PROP_ACCT_TYPE, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szAcctType = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_IDENTITY
											hRes = lpAccount->GetProp(PROP_ACCT_IDENTITY, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szAcctIdentity = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_FLAVOR
											hRes = lpAccount->GetProp(PROP_ACCT_FLAVOR, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szAcctFlavor = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_IS_DEFAULT_MAIL
											hRes = lpAccount->GetProp(PROP_ACCT_IS_DEFAULT_MAIL, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.dw)
											{
												pAccounts[i].lAcctIsDefaultMail = pProp.Val.dw;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_USER_DISPLAY_NAME
											hRes = lpAccount->GetProp(PROP_ACCT_USER_DISPLAY_NAME, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szAcctUserDisplayName = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_USER_EMAIL_ADDR
											hRes = lpAccount->GetProp(PROP_ACCT_USER_EMAIL_ADDR, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szAcctUserEmailAddr = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_STAMP
											hRes = lpAccount->GetProp(PROP_ACCT_STAMP, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szAcctStamp = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_SEND_STAMP
											hRes = lpAccount->GetProp(PROP_ACCT_SEND_STAMP, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szAcctSendStamp = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_IS_EXCH
											hRes = lpAccount->GetProp(PROP_ACCT_IS_EXCH, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.dw)
											{
												pAccounts[i].lAcctIsExch = pProp.Val.dw;
											}
											else
												hRes = S_OK;

											//PROP_ACCT_DISABLED
											hRes = lpAccount->GetProp(PROP_ACCT_DISABLED, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.dw)
											{
												pAccounts[i].lAcctDisabled = pProp.Val.dw;
											}
											else
												hRes = S_OK;

											/*********************************************************/

											//PROP_MAPI_SERVICE_UID
											hRes = lpAccount->GetProp(PROP_MAPI_SERVICE_UID, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.bin.pb)
											{
												pAccounts[i].bMapiServiceUid.cb = pProp.Val.bin.cb;
												pAccounts[i].bMapiServiceUid.pb = pProp.Val.bin.pb;
											}
											else
												hRes = S_OK;

											//PROP_MAPI_PROVIDER_TYPE
											hRes = lpAccount->GetProp(PROP_MAPI_PROVIDER_TYPE, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.dw)
											{
												pAccounts[i].lMapiProviderType = pProp.Val.dw;
											}
											else
												hRes = S_OK;

											//PROP_MAPI_IDENTITY_ENTRYID
											hRes = lpAccount->GetProp(PROP_MAPI_IDENTITY_ENTRYID, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.bin.pb)
											{
												pAccounts[i].bMapiIdentityEntryId.cb = pProp.Val.bin.cb;
												pAccounts[i].bMapiIdentityEntryId.pb = pProp.Val.bin.pb;
											}
											else
												hRes = S_OK;

											/***************************************************************/

											//PR_PRIMARY_SEND_ACCT
											hRes = lpAccount->GetProp(PR_PRIMARY_SEND_ACCT, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szPrimarySendAcct = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

											//PR_NEXT_SEND_ACCT
											hRes = lpAccount->GetProp(PR_NEXT_SEND_ACCT, &pProp);
											if (SUCCEEDED(hRes) && pProp.Val.pwsz)
											{
												pAccounts[i].szNextSendAcct = pProp.Val.pwsz;
											}
											else
												hRes = S_OK;

										}

										if (lpAccount)
											lpAccount->Release();
										lpAccount = NULL;
									}

									if (lpUnk)
										lpUnk->Release();
									lpUnk = NULL;
								}

								*pcAccounts = cAccounts;
								*ppAccounts = pAccounts;
							}
						}
					}

					if (lpAcctEnum)
						lpAcctEnum->Release();
				}
			}

			if (lpAcctHelper)
				lpAcctHelper->Release();
		}

		if (pMyAcctHelper)
			pMyAcctHelper->Release();
	}

	if (lpAcctMgr)
		lpAcctMgr->Release();

	return hRes;
}

std::wstring GetDefaultAccountNameW(LPMAPISESSION lpSession, LPWSTR lpszProfileName)
{
	HRESULT hRes = S_OK;
	std::wstring szAccountName;
	OlkAccount * pOlkAccount = new OlkAccount;
	EC_HRES(GetDefaultAccount(lpSession,
		lpszProfileName,
		pOlkAccount));

	szAccountName = pOlkAccount->szAcctName;

Error:
	goto Cleanup;
Cleanup:
	return szAccountName;
}

HRESULT ProcessAccounts(LPMAPISESSION lpSession, LPWSTR lpProfileName, RuntimeOptions runtimeOptions)
{
	HRESULT hRes = S_OK;
	DWORD dwCount = 0;
	OlkAccount * pOlkAccounts = NULL;

	hRes = GetAccounts(lpProfileName, &dwCount, &pOlkAccounts);

	if (SUCCEEDED(hRes))
	{
		if ((NULL != pOlkAccounts) && (dwCount > 0))
		{
			for (unsigned int i = 0; i < dwCount; i++)
			{
				std::wstring szAccountName;
				szAccountName = pOlkAccounts[i].szAcctName;

				if (szAccountName.find(L"@") != std::wstring::npos)
				{
					//if (INPUTMODE_ACTIVEDIRECTORY == runtimeOptions.ulRunningMode)
					//{
					//	LPWSTR lpszPrimarySMTPAddress = new OLECHAR[MAX_PATH * 2];
					//	// Assuming that the account name is the primary or a secondary SMTP address of the user, 
					//	// query Active Directory for the primary SMTP address 
					//	EC_HRES(FindPrimarySMTPAddress(pOlkAccounts[i].szAcctName,
					//		&lpszPrimarySMTPAddress,
					//		runtimeOptions.ulAdTimeout,
					//		(LPWSTR)runtimeOptions.szADsPAth.c_str()));
					//	if (lpszPrimarySMTPAddress)
					//	{
					//		std::wstring szPrimarySmtpAddress(lpszPrimarySMTPAddress);
					//		if (szPrimarySmtpAddress.find(L"@") != std::wstring::npos)
					//		{
					//			wprintf(L"Primary SMTP address is: %s.\n", (LPWSTR)szPrimarySmtpAddress.c_str());
					//			if (wcsncmp(pOlkAccounts[i].szAcctName, (LPWSTR)szPrimarySmtpAddress.c_str(), lstrlenW((LPWSTR)szPrimarySmtpAddress.c_str())) == 0)
					//			{
					//				wprintf(L"Account name is up to date: %s.\n", (LPWSTR)szPrimarySmtpAddress.c_str());
					//			}
					//			else
					//			{
					//				wprintf(L"Account name is not up to date: %s.\n", pOlkAccounts[i].szAcctName);
					//				wprintf(L"Updating account name to: %s.\n", (LPWSTR)szPrimarySmtpAddress.c_str());
					//				// If the current account name is not the primary SMTP address we then attempt to 
					//				// update the account name with the correct value
					//				EC_HRES(UpdateAcctName(lpSession, (LPWSTR)runtimeOptions.szProfileName.c_str(), pOlkAccounts[i].lAcctId, (LPWSTR)szPrimarySmtpAddress.c_str()));
					//			}
					//		}
					//		if (lpszPrimarySMTPAddress) MAPIFreeBuffer(lpszPrimarySMTPAddress);
					//	}
					//}
					//else 
					//if (INPUTMODE_USERINPUT == runtimeOptions.ulRunningMode)
					//{

					//	if (szAccountName.find(runtimeOptions.szOldDomainName) != std::wstring::npos)
					//	{
					//		std::wstring szNewSmtpAddress = szAccountName;
					//		int pos = szNewSmtpAddress.find(runtimeOptions.szOldDomainName);
					//		int len = runtimeOptions.szOldDomainName.length();
					//		szNewSmtpAddress.replace(pos, len, runtimeOptions.szNewDomainName);
					//		wprintf(L"Updating account name to: %s.\n", (LPWSTR)szNewSmtpAddress.c_str());
					//		// If the current account name is not the primary SMTP address we then attempt to 
					//		// update the account name with the correct value
					//		EC_HRES(UpdateAcctName(lpSession, (LPWSTR)runtimeOptions.szProfileName.c_str(), pOlkAccounts[i].lAcctId, (LPWSTR)szNewSmtpAddress.c_str()));
					//	}
					/*}*/
				}
				else
				{
					wprintf(L"Account name %s is not a valid SMTP address.\n", pOlkAccounts[i].szAcctName);
				}
			}
		}
	}
	else
	{
		goto Error;
	}

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	return hRes;
}