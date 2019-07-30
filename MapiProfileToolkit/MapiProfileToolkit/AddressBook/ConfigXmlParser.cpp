/*
* © 2016 Microsoft Corporation
*
* written by Andrei Ghita
*
* Microsoft provides programming examples for illustration only, without warranty either expressed or implied.
* This includes, but is not limited to, the implied warranties of merchantability or fitness for a particular purpose.
* This sample assumes that you are familiar with the programming language that is being demonstrated and with
* the tools that are used to create and to debug procedures. Microsoft support engineers can help explain the
* functionality of a particular procedure, but they will not modify these examples to provide added functionality
* or construct procedures to meet your specific requirements.
*/

#include "../stdafx.h"
#include "../Profile/Profile.h"
#include "ConfigXmlParser.h"


HRESULT ParseConfigXml(LPTSTR lpszABConfigurationPath, ABProvider* pABProvider)
{
	HRESULT hRes = S_OK;
	CComPtr<IXMLDOMDocument> pXMLDoc;
	CComPtr<IXMLDOMElement> xmlElement;

	EC_HRES(pXMLDoc.CoCreateInstance(__uuidof(DOMDocument)));

	VARIANT_BOOL bSuccess = false;
	// load the xml from the specified filename
	EC_HRES(pXMLDoc->load(variant_t(lpszABConfigurationPath), &bSuccess));

	if (bSuccess)
	{
		// Get a pointer to the root element
		CComPtr<IXMLDOMElement> pRootElm;
		EC_HRES(pXMLDoc->get_documentElement(&pRootElm));

		CComPtr<IXMLDOMNodeList> pXMLNodes;
		EC_HRES(pRootElm->get_childNodes(&pXMLNodes));

		long lCount;
		EC_HRES(pXMLNodes->get_length(&lCount));

		for (int i = 0; i < lCount; i++)
		{
			CComPtr<IXMLDOMNode> pXMLNode;
			EC_HRES(pXMLNodes->get_item(i, &pXMLNode));

			BSTR bstrNodeName;
			EC_HRES(pXMLNode->get_nodeName(&bstrNodeName));

			VARIANT pNodeTypedValue;
			EC_HRES(pXMLNode->get_nodeTypedValue(&pNodeTypedValue));

			if (0 == _wcsicmp(bstrNodeName, L"DisplayName"))
			{
				if (NULL != pNodeTypedValue.bstrVal)
					pABProvider->lpszDisplayName = pNodeTypedValue.bstrVal;
				else
					pABProvider->lpszDisplayName = L"";
			}
			if (0 == _wcsicmp(bstrNodeName, L"ServerName"))
			{
				if (NULL != pNodeTypedValue.bstrVal)
					pABProvider->lpszServerName = pNodeTypedValue.bstrVal;
				else
					pABProvider->lpszServerName = L"";
			}
			if (0 == _wcsicmp(bstrNodeName, L"ServerPort"))
			{
				if (NULL != pNodeTypedValue.bstrVal)
					pABProvider->lpszServerPort = pNodeTypedValue.bstrVal;
				else
					pABProvider->lpszServerPort = L"";
			}
			if (0 == _wcsicmp(bstrNodeName, L"UseSSL"))
			{
				if (0 == _wcsicmp(pNodeTypedValue.bstrVal, L"true"))
				{
					pABProvider->bUseSSL = true;
				}
				else
				{
					if (0 == _wcsicmp(pNodeTypedValue.bstrVal, L"false"))
					{
						pABProvider->bUseSSL = false;
					}
					else
					{
						pABProvider->bUseSSL = false;
					}
				}
			}
			if (0 == _wcsicmp(bstrNodeName, L"Username"))
			{
				if (NULL != pNodeTypedValue.bstrVal)
					pABProvider->lpszUsername = pNodeTypedValue.bstrVal;
				else
					pABProvider->lpszUsername = L"";
			}
			if (0 == _wcsicmp(bstrNodeName, L"Password"))
			{
				if (NULL != pNodeTypedValue.bstrVal)
					pABProvider->lpszPassword = pNodeTypedValue.bstrVal;
				else
					pABProvider->lpszPassword = L"";
			}
			if (0 == _wcsicmp(bstrNodeName, L"RequireSecurePasswordAuth"))
			{
				if (0 == _wcsicmp(pNodeTypedValue.bstrVal, L"true"))
				{
					pABProvider->bRequireSPA = true;
				}
				else
				{
					if (0 == _wcsicmp(pNodeTypedValue.bstrVal, L"false"))
					{
						pABProvider->bRequireSPA = false;
					}
					else
					{
						pABProvider->bRequireSPA = false;
					}
				}
			}
			if (0 == _wcsicmp(bstrNodeName, L"SearchTimeoutSeconds"))
			{
				if (NULL != pNodeTypedValue.bstrVal)
					pABProvider->lpszTimeout = pNodeTypedValue.bstrVal;
				else
					pABProvider->lpszTimeout = L"";
			}
			if (0 == _wcsicmp(bstrNodeName, L"MaxEntriesReturned"))
			{
				if (NULL != pNodeTypedValue.bstrVal)
					pABProvider->lpszMaxResults = pNodeTypedValue.bstrVal;
				else
					pABProvider->lpszMaxResults = L"";;
			}
			if (0 == _wcsicmp(bstrNodeName, L"DefaultSearchBase"))
			{
				if (0 == _wcsicmp(pNodeTypedValue.bstrVal, L"true"))
				{
					pABProvider->ulDefaultSearchBase = 1;
				}
				else
				{
					if (0 == _wcsicmp(pNodeTypedValue.bstrVal, L"false"))
					{
						pABProvider->ulDefaultSearchBase = 0;
					}
					else
					{
						pABProvider->ulDefaultSearchBase = 0;
					}
				}
			}
			if (0 == _wcsicmp(bstrNodeName, L"CustomSearchBase"))
			{
				if (NULL != pNodeTypedValue.bstrVal)
					pABProvider->lpszCustomSearchBase = pNodeTypedValue.bstrVal;
				else
					pABProvider->lpszCustomSearchBase = L"";
			}
			if (0 == _wcsicmp(bstrNodeName, L"EnableBrowsing"))
			{
				if (0 == _wcsicmp(pNodeTypedValue.bstrVal, L"true"))
				{
					pABProvider->bEnableBrowsing = true;
				}
				else
				{
					if (0 == _wcsicmp(pNodeTypedValue.bstrVal, L"false"))
					{
						pABProvider->bEnableBrowsing = false;
					}
					else
					{
						pABProvider->bEnableBrowsing = false;
					}
				}
			}

			if (bstrNodeName) SysFreeString(bstrNodeName);
		}

	}
Error:

	return hRes;
}
