#pragma once
#include "stdafx.h"
#include "ABProviderObjects.h"
#include <windows.h>
#include <msxml.h>
#include <objsafe.h>
#include <objbase.h>
#include <atlbase.h>
#include "wchar.h"

HRESULT ParseXml(LPTSTR lpszABConfigurationPath, ABProvider* abProvider)
{
	HRESULT hRes = S_OK;
	CComPtr<IXMLDOMDocument> pXMLDoc;
	EC_HRES(pXMLDoc.CoCreateInstance(__uuidof(DOMDocument)));
	CComPtr<IXMLDOMElement> xmlElement;

	// Load the file. 
	VARIANT_BOOL bSuccess = false;
	// Can load it from a url/filename...
	//iXMLDoc->load(CComVariant(url),&bSuccess);
	// or from a BSTR...
	EC_HRES(pXMLDoc->loadXML(CComBSTR(lpszABConfigurationPath), &bSuccess));

	if (bSuccess)
	{
		// Get a pointer to the root
		CComPtr<IXMLDOMElement> pRootElm;
		EC_HRES(pXMLDoc->get_documentElement(&pRootElm));
		CComPtr<IXMLDOMNodeList> pXMLNodes;
		// Thanks to the magic of CComPtr, we never need call
		// Release() -- that gets done automatically.
		EC_HRES(pRootElm->get_childNodes(&pXMLNodes));
		long lCount;
		EC_HRES(pXMLNodes->get_length(&lCount));
		for (int i = 0; i < lCount; i++)
		{
			CComPtr<IXMLDOMNode> pXMLNode;
			EC_HRES(pXMLNodes->get_item(i, &pXMLNode));
			BSTR bstrNodeName;
			EC_HRES(pXMLNode->get_nodeName(&bstrNodeName));
			VARIANT pNodeValue;
			EC_HRES(pXMLNode->get_nodeValue(&pNodeValue));
			if (0 == _wcsicmp(bstrNodeName, L"DisplayName"))
			{
				abProvider->lpszDisplayName = pNodeValue.bstrVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"ServerName"))
			{
				abProvider->lpszServerName = pNodeValue.bstrVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"ServerPort"))
			{
				abProvider->lpszServerPort = pNodeValue.bstrVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"UseSSL"))
			{
				abProvider->bUseSSL = pNodeValue.boolVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"Username"))
			{
				abProvider->lpszUsername = pNodeValue.bstrVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"Password"))
			{
				abProvider->lpszPassword = pNodeValue.bstrVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"RequireSecurePasswordAuth"))
			{
				abProvider->bRequireSPA = pNodeValue.boolVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"SearchTimeoutSeconds"))
			{
				abProvider->lpszTimeout = pNodeValue.bstrVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"MaxEntriesReturned"))
			{
				abProvider->lpszMaxResults = pNodeValue.bstrVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"DefaultSearchBase"))
			{
				abProvider->uDefaultSearchBase = pNodeValue.ulVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"CustomSearchBase"))
			{
				abProvider->lpszCustomSearchBase = pNodeValue.bstrVal;
				break;
			}
			if (0 == _wcsicmp(bstrNodeName, L"EnableBrowsing"))
			{
				abProvider->bEnableBrowsing = pNodeValue.boolVal;
				break;
			}
		}
	}
Error:
	return hRes;
}