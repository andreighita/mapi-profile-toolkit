#pragma once

#include "../stdafx.h"
#include "AccountHelper.h"
#include "AccountMgmt.h"
#include "AccountObjects.h"
#include "../Misc/ADSI/ADHelper.h"

HRESULT GetDefaultAccount(LPMAPISESSION lpSession, LPWSTR lpwszProfile, OlkAccount * pOlkAccount);
HRESULT UpdateAcctName(LPMAPISESSION lpSession, LPWSTR lpwszProfile, long lAcctId, LPWSTR lpszNewAcctName);
HRESULT GetAccountData(LPOLKACCOUNT *lpAccount, OlkAccount * pOlkAccount);
HRESULT GetAccounts(LPWSTR lpwszProfile, DWORD* pcAccounts, OlkAccount** ppAccounts);
std::wstring GetDefaultAccountNameW(LPMAPISESSION lpSession, LPWSTR lpszProfileName);
HRESULT ProcessAccounts(LPMAPISESSION lpSession, LPWSTR lpProfileName, RuntimeOptions runtimeOptions);
