#pragma once

#include "../stdafx.h"
#include <MAPIX.h>
#define USES_IID_IMAPISession
#include <windows.h>
#include <strsafe.h>
#include "AccountHelper.h"
#include "AccountMgmt.h"


CAccountHelper::CAccountHelper(LPWSTR lpwszProfName, LPMAPISESSION lpSession)
{
	m_cRef = 1;
	m_lpUnkSession = NULL;
	m_lpwszProfile = NULL;
	m_cchProfile = 0;

	HRESULT hRes = S_OK;

	if (lpwszProfName)
	{
		hRes = StringCchLengthW(lpwszProfName, STRSAFE_MAX_CCH, &m_cchProfile);
		if (SUCCEEDED(hRes) && m_cchProfile)
		{
			m_cchProfile++;

			m_lpwszProfile = (LPWSTR)malloc(m_cchProfile * sizeof(WCHAR));

			if (m_lpwszProfile)
			{
				hRes = StringCchCopyW(m_lpwszProfile, m_cchProfile, lpwszProfName);
			}
		}
	}

	if (lpSession)
	{
		hRes = lpSession->QueryInterface(IID_IUnknown, (LPVOID*)&m_lpUnkSession);
	}
}

CAccountHelper::~CAccountHelper()
{
	if (m_lpUnkSession)
		m_lpUnkSession->Release();

	if (m_lpwszProfile)
		free(m_lpwszProfile);
}

STDMETHODIMP CAccountHelper::QueryInterface(REFIID riid,
	LPVOID * ppvObj)
{
	*ppvObj = 0;
	if (riid == IID_IOlkAccountHelper ||
		riid == IID_IUnknown)
	{
		*ppvObj = (LPVOID)this;
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CAccountHelper::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	return lCount;
}

STDMETHODIMP_(ULONG) CAccountHelper::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	if (!lCount)  delete this;
	return lCount;
}

STDMETHODIMP CAccountHelper::GetIdentity(LPWSTR pwszIdentity, DWORD * pcch)
{
	if (!pcch || !m_lpwszProfile)
		return E_INVALIDARG;

	HRESULT hRes = S_OK;

	if (m_cchProfile > *pcch)
	{
		*pcch = m_cchProfile;
		return E_OUTOFMEMORY;
	}

	hRes = StringCchCopyW(pwszIdentity, *pcch, m_lpwszProfile);

	*pcch = m_cchProfile;

	return hRes;
}

STDMETHODIMP CAccountHelper::GetMapiSession(LPUNKNOWN * ppmsess)
{
	if (!ppmsess)
		return E_INVALIDARG;

	if (m_lpUnkSession)
	{
		return m_lpUnkSession->QueryInterface(IID_IMAPISession, (LPVOID*)ppmsess);
	}

	return E_NOTIMPL;
}

STDMETHODIMP CAccountHelper::HandsOffSession()
{
	return E_NOTIMPL;
}