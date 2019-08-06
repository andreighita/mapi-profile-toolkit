#pragma once
#include <windows.h>
#include <msxml.h>
#include <objsafe.h>
#include <objbase.h>
#include <atlbase.h>
#include "wchar.h"
#include "../../ToolkitTypeDefs.h"
#include "../../InlineAndMacros.h"
#include "../../Toolkit.h"

namespace MAPIToolkit
{
	HRESULT ParseXml(LPTSTR lpszABConfigurationPath, AddressBookWorker* addressBookWorker)
	{
		HRESULT hRes = S_OK;
		CComPtr<IXMLDOMDocument> pXMLDoc;

		CComPtr<IXMLDOMElement> xmlElement;
		VARIANT_BOOL bSuccess = true;

		HCK(pXMLDoc.CoCreateInstance(__uuidof(DOMDocument)));

		// Load the file. 

		// Can load it from a url/filename...
		//iXMLDoc->load(CComVariant(url),&bSuccess);
		// or from a BSTR...
		HCK(pXMLDoc->loadXML(CComBSTR(lpszABConfigurationPath), &bSuccess));

		if (bSuccess)
		{
			// Get a pointer to the root
			CComPtr<IXMLDOMElement> pRootElm;
			HCK(pXMLDoc->get_documentElement(&pRootElm));
			CComPtr<IXMLDOMNodeList> pXMLNodes;
			// Thanks to the magic of CComPtr, we never need call
			// Release() -- that gets done automatically.
			HCK(pRootElm->get_childNodes(&pXMLNodes));
			long lCount;
			HCK(pXMLNodes->get_length(&lCount));
			for (int i = 0; i < lCount; i++)
			{
				CComPtr<IXMLDOMNode> pXMLNode;
				HCK(pXMLNodes->get_item(i, &pXMLNode));
				BSTR bstrNodeName;
				HCK(pXMLNode->get_nodeName(&bstrNodeName));
				VARIANT pNodeValue;
				HCK(pXMLNode->get_nodeValue(&pNodeValue));
				if (0 == _wcsicmp(bstrNodeName, L"DisplayName"))
				{
					addressBookWorker->displayName = pNodeValue.bstrVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"ServerName"))
				{
					addressBookWorker->serverName = pNodeValue.bstrVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"ServerPort"))
				{
					addressBookWorker->serverPort = pNodeValue.bstrVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"UseSSL"))
				{
					addressBookWorker->useSSL = pNodeValue.boolVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"Username"))
				{
					addressBookWorker->username = pNodeValue.bstrVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"Password"))
				{
					addressBookWorker->password = pNodeValue.bstrVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"RequireSecurePasswordAuth"))
				{
					addressBookWorker->requireSPA = pNodeValue.boolVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"SearchTimeoutSeconds"))
				{
					addressBookWorker->timeout = pNodeValue.bstrVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"MaxEntriesReturned"))
				{
					addressBookWorker->maxResults = pNodeValue.bstrVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"DefaultSearchBase"))
				{
					addressBookWorker->defaultSearchBase = pNodeValue.ulVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"CustomSearchBase"))
				{
					addressBookWorker->customSearchBase = pNodeValue.bstrVal;
					break;
				}
				if (0 == _wcsicmp(bstrNodeName, L"EnableBrowsing"))
				{
					addressBookWorker->enableBrowsing = pNodeValue.boolVal;
					break;
				}
			}
		}
	Error:
		return hRes;
	}
}