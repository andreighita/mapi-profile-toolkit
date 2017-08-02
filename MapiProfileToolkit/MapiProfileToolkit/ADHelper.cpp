
#include "stdafx.h"
#include "ADHelper.h"

std::wstring GetUserDn()
{
	std::wstring wszUserDn = L"";
	HRESULT hRes = S_OK;

	IADsADSystemInfo *pADsys;
	EC_HRES(CoCreateInstance(CLSID_ADSystemInfo,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IADsADSystemInfo,
		(void**)&pADsys));

	if (pADsys)
	{
		BSTR bstrUserName = NULL;
		EC_HRES(pADsys->get_UserName(&bstrUserName));
		if (bstrUserName)
		{
			wszUserDn = std::wstring(bstrUserName);
			::SysFreeString(bstrUserName);
		}
		pADsys->Release();
	}
Error:
	return wszUserDn;
}

std::wstring FindPrimarySMTPAddress(std::wstring wszUserDn)
{

	std::wstring wszSmtpAddress = L"";

	//Intialize COM



	HRESULT hr = S_OK;

	//Get rootDSE and the config container's DN.

	LPADS lpAds = NULL;

	wszUserDn = L"LDAP://" + wszUserDn;
	hr = ADsOpenObject((LPCWSTR)wszUserDn.c_str(),
		NULL,
		NULL,
		ADS_SECURE_AUTHENTICATION,
		//Use Secure Authentication
		IID_IADs,
		(void**)&lpAds);

	if ((S_OK == hr) && lpAds)
	{

		VARIANT varPropValue;
		BSTR bstrProperty = BSTR(L"proxyAddresses");
		hr = lpAds->Get(bstrProperty, &varPropValue);
		if ((SUCCEEDED(hr)) && (VT_VARIANT ^ varPropValue.vt))
		{
			LONG cElements, lLBound, lUBound;

			if (SafeArrayGetDim(varPropValue.parray) == 1)
			{
				// Get array bounds.
				hr = SafeArrayGetLBound(varPropValue.parray, 1, &lLBound);
				if (FAILED(hr))
					goto Error;
				hr = SafeArrayGetUBound(varPropValue.parray, 1, &lUBound);
				if (FAILED(hr))
					goto Error;

				cElements = lUBound - lLBound + 1;

				VARIANT propVal;
				VariantInit(&propVal);
				for (LONG i = 0; i < cElements - 1; i++)
				{
					hr = SafeArrayGetElement(varPropValue.parray, &i, &propVal);
					if (propVal.vt == VT_BSTR)
					{
						std::wstring wszAddress = std::wstring(propVal.bstrVal);
						size_t pos = wszAddress.find(L"SMTP:");
						if (pos != std::wstring::npos)
						{
							pos = wszAddress.find(L":");
							wszSmtpAddress = wszAddress.substr(pos + 1);
							break;
						}
					}
				}


			}
		}
		lpAds->Release();
	}
Error:
	return wszSmtpAddress;
}


