/*
* � 2015 Microsoft Corporation
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
#include "pch.h"
#include "XMLHelper.h"

// Macro that calls a COM method returning HRESULT value.
#define CHK_HR(stmt)        do { hr=(stmt); if (FAILED(hr)) goto CleanUp; } while(0)

// Macro to verify memory allcation.
#define CHK_ALLOC(p)        do { if (!(p)) { hr = E_OUTOFMEMORY; goto CleanUp; } } while(0)

// Macro that releases a COM object if not NULL.
#define SAFE_RELEASE(p)     do { if ((p)) { (p)->Release(); (p) = NULL; } } while(0)

template <class T>
inline std::wstring ConvertTToString(const T& t)
{
	std::wstringstream wss;
	wss << t;
	return wss.str();
}


inline std::wstring ConvertIntToString(int t)
{
	std::wstringstream wss;
	wss << t;
	return wss.str();
}

// Helper function to create a VT_BSTR variant from a null terminated string. 
HRESULT VariantFromString(PCWSTR wszValue, VARIANT& Variant)
{
	HRESULT hr = S_OK;
	BSTR bstr = SysAllocString(wszValue);
	CHK_ALLOC(bstr);

	V_VT(&Variant) = VT_BSTR;
	V_BSTR(&Variant) = bstr;

CleanUp:
	return hr;
}

// Helper function to create a DOM instance. 
HRESULT CreateAndInitDOM(IXMLDOMDocument** ppDoc)
{
	HRESULT hr = CoCreateInstance(__uuidof(DOMDocument60), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(ppDoc));
	if (SUCCEEDED(hr))
	{
		// these methods should not fail so don't inspect result
		(*ppDoc)->put_async(VARIANT_FALSE);
		(*ppDoc)->put_validateOnParse(VARIANT_FALSE);
		(*ppDoc)->put_resolveExternals(VARIANT_FALSE);
		(*ppDoc)->put_preserveWhiteSpace(VARIANT_TRUE);
	}
	return hr;
}

// Helper that allocates the BSTR param for the caller.
HRESULT CreateElement(IXMLDOMDocument* pXMLDom, PCWSTR wszName, IXMLDOMElement** ppElement)
{
	HRESULT hr = S_OK;
	*ppElement = NULL;

	BSTR bstrName = SysAllocString(wszName);
	CHK_ALLOC(bstrName);
	CHK_HR(pXMLDom->createElement(bstrName, ppElement));

CleanUp:
	SysFreeString(bstrName);
	return hr;
}

// Helper function to append a child to a parent node.
HRESULT AppendChildToParent(IXMLDOMNode* pChild, IXMLDOMNode* pParent)
{
	HRESULT hr = S_OK;
	IXMLDOMNode* pChildOut = NULL;
	CHK_HR(pParent->appendChild(pChild, &pChildOut));

CleanUp:
	SAFE_RELEASE(pChildOut);
	return hr;
}

// Helper function to create and add a processing instruction to a document node.
HRESULT CreateAndAddPINode(IXMLDOMDocument* pDom, PCWSTR wszTarget, PCWSTR wszData)
{
	HRESULT hr = S_OK;
	IXMLDOMProcessingInstruction* pPI = NULL;

	BSTR bstrTarget = SysAllocString(wszTarget);
	BSTR bstrData = SysAllocString(wszData);
	CHK_ALLOC(bstrTarget && bstrData);

	CHK_HR(pDom->createProcessingInstruction(bstrTarget, bstrData, &pPI));
	CHK_HR(AppendChildToParent(pPI, pDom));

CleanUp:
	SAFE_RELEASE(pPI);
	SysFreeString(bstrTarget);
	SysFreeString(bstrData);
	return hr;
}

// Helper function to create and add a comment to a document node.
HRESULT CreateAndAddCommentNode(IXMLDOMDocument* pDom, PCWSTR wszComment)
{
	HRESULT hr = S_OK;
	IXMLDOMComment* pComment = NULL;

	BSTR bstrComment = SysAllocString(wszComment);
	CHK_ALLOC(bstrComment);

	CHK_HR(pDom->createComment(bstrComment, &pComment));
	CHK_HR(AppendChildToParent(pComment, pDom));

CleanUp:
	SAFE_RELEASE(pComment);
	SysFreeString(bstrComment);
	return hr;
}

// Helper function to create and add an attribute to a parent node.
HRESULT CreateAndAddAttributeNode(IXMLDOMDocument* pDom, PCWSTR wszName, PCWSTR wszValue, IXMLDOMElement* pParent)
{
	HRESULT hr = S_OK;
	IXMLDOMAttribute* pAttribute = NULL;
	IXMLDOMAttribute* pAttributeOut = NULL; // Out param that is not used

	BSTR bstrName = NULL;
	VARIANT varValue;
	VariantInit(&varValue);

	bstrName = SysAllocString(wszName);
	CHK_ALLOC(bstrName);
	CHK_HR(VariantFromString(wszValue, varValue));

	CHK_HR(pDom->createAttribute(bstrName, &pAttribute));
	CHK_HR(pAttribute->put_value(varValue));
	CHK_HR(pParent->setAttributeNode(pAttribute, &pAttributeOut));

CleanUp:
	SAFE_RELEASE(pAttribute);
	SAFE_RELEASE(pAttributeOut);
	SysFreeString(bstrName);
	VariantClear(&varValue);
	return hr;
}

// Helper function to create and append a text node to a parent node.
HRESULT CreateAndAddTextNode(IXMLDOMDocument* pDom, PCWSTR wszText, IXMLDOMNode* pParent)
{
	HRESULT hr = S_OK;
	IXMLDOMText* pText = NULL;

	BSTR bstrText = SysAllocString(wszText);
	CHK_ALLOC(bstrText);

	CHK_HR(pDom->createTextNode(bstrText, &pText));
	CHK_HR(AppendChildToParent(pText, pParent));

CleanUp:
	SAFE_RELEASE(pText);
	SysFreeString(bstrText);
	return hr;
}

// Helper function to create and append a CDATA node to a parent node.
HRESULT CreateAndAddCDATANode(IXMLDOMDocument* pDom, PCWSTR wszCDATA, IXMLDOMNode* pParent)
{
	HRESULT hr = S_OK;
	IXMLDOMCDATASection* pCDATA = NULL;

	BSTR bstrCDATA = SysAllocString(wszCDATA);
	CHK_ALLOC(bstrCDATA);

	CHK_HR(pDom->createCDATASection(bstrCDATA, &pCDATA));
	CHK_HR(AppendChildToParent(pCDATA, pParent));

CleanUp:
	SAFE_RELEASE(pCDATA);
	SysFreeString(bstrCDATA);
	return hr;
}

// Helper function to create and append an element node to a parent node, and pass the newly created
// element node to caller if it wants.
HRESULT CreateAndAddElementNode(IXMLDOMDocument* pDom, PCWSTR wszName, PCWSTR wszNewline, IXMLDOMNode* pParent, IXMLDOMElement** ppElement = NULL)
{
	HRESULT hr = S_OK;
	IXMLDOMElement* pElement = NULL;

	CHK_HR(CreateElement(pDom, wszName, &pElement));
	// Add NEWLINE+TAB for identation before this element.
	CHK_HR(CreateAndAddTextNode(pDom, wszNewline, pParent));
	// Append this element to parent.
	CHK_HR(AppendChildToParent(pElement, pParent));

CleanUp:
	if (ppElement)
		* ppElement = pElement;  // Caller is repsonsible to release this element.
	else
		SAFE_RELEASE(pElement); // Caller is not interested on this element, so release it.
	return hr;
}

void ExportXML(ULONG cProfileInfo, ProfileInfo* profileInfo, std::wstring szExportPath)
{
	HRESULT hr = S_OK;
	IXMLDOMDocument* pXMLDom = NULL;
	IXMLDOMElement* pRoot = NULL;

	BSTR bstrXML = NULL;
	VARIANT varFileName;
	VariantInit(&varFileName);

	CHK_HR(CreateAndInitDOM(&pXMLDom));

	// Create a processing instruction element.
	CHK_HR(CreateAndAddPINode(pXMLDom, L"xml", L"version='1.0'"));

	// Create the root element.
	CHK_HR(CreateElement(pXMLDom, L"Profiles", &pRoot));

	for (unsigned int i = 0; i < cProfileInfo; i++)
	{
		IXMLDOMElement* pProfileNode = NULL;
		CHK_HR(CreateAndAddElementNode(pXMLDom, L"Profile", L"\n\t", pRoot, &pProfileNode));
		// Add ProfileName node and value
		IXMLDOMElement* pProfileNameNode = NULL;
		CHK_HR(CreateAndAddElementNode(pXMLDom, L"ProfileName", L"\n\t", pProfileNode, &pProfileNameNode));
		CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].wszProfileName.c_str(), pProfileNameNode));
		SAFE_RELEASE(pProfileNameNode);
		// Add ProfileIndex node and value
		IXMLDOMElement* pProfileIndexNode = NULL;
		CHK_HR(CreateAndAddElementNode(pXMLDom, L"ProfileIndex", L"\n\t", pProfileNode, &pProfileIndexNode));
		CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(i + 1)), pProfileIndexNode));
		SAFE_RELEASE(pProfileIndexNode);
		// Add DefaultProfile node and value
		IXMLDOMElement* pDefaultProfileNode = NULL;
		CHK_HR(CreateAndAddElementNode(pXMLDom, L"DefaultProfile", L"\n\t", pProfileNode, &pDefaultProfileNode));
		CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(profileInfo[i].bDefaultProfile)), pDefaultProfileNode));
		SAFE_RELEASE(pDefaultProfileNode);
		// Add ServiceCount node and value
		IXMLDOMElement* pServiceCountNode = NULL;
		CHK_HR(CreateAndAddElementNode(pXMLDom, L"ServicesCount", L"\n\t", pProfileNode, &pServiceCountNode));
		CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(profileInfo[i].ulServiceCount)), pServiceCountNode));
		SAFE_RELEASE(pServiceCountNode);

		if (profileInfo[i].ulServiceCount > 0)
		{
			// create root node for Services
			IXMLDOMElement* pServicesRootNode = NULL;
			CHK_HR(CreateAndAddElementNode(pXMLDom, L"Services", L"\n\t", pProfileNode, &pServicesRootNode));

			for (unsigned int j = 0; j < profileInfo[i].ulServiceCount; j++)
			{
				if (ServiceType::ServiceType_Mailbox == profileInfo[i].profileServices[j].serviceType)
				{
					// create child node for each service
					IXMLDOMElement* pServiceNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"ExchangeAccount", L"\n\t", pServicesRootNode, &pServiceNode));
					// Add ServiceName node and value
					IXMLDOMElement* pServiceNameNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"AccountName", L"\n\t", pServiceNode, &pServiceNameNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].wszServiceName.c_str(), pServiceNameNode));
					SAFE_RELEASE(pServiceNameNode);
					// Add ServiceIndex node and value
					IXMLDOMElement* pServiceIndexNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"AccountIndex", L"\n\t", pServiceNode, &pServiceIndexNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(j + 1)), pServiceIndexNode));
					SAFE_RELEASE(pServiceIndexNode);
					// Add ServiceDisplayName node and value
					IXMLDOMElement* pServiceDisplayNameNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"AccountDisplayName", L"\n\t", pServiceNode, &pServiceDisplayNameNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].exchangeAccountInfo->wszDisplayName.c_str(), pServiceDisplayNameNode));
					SAFE_RELEASE(pServiceDisplayNameNode);
					// Add DefaultService node and value
					IXMLDOMElement* pDefaultServiceNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"DefaultAccount", L"\n\t", pServiceNode, &pDefaultServiceNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(profileInfo[i].profileServices[j].bDefaultStore)), pDefaultServiceNode));
					SAFE_RELEASE(pDefaultServiceNode);
					// Add DatafilePath node and value
					IXMLDOMElement* pDatafilePathNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"DatafilePath", L"\n\t", pServiceNode, &pDatafilePathNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].exchangeAccountInfo->wszDatafilePath.c_str(), pDatafilePathNode));
					SAFE_RELEASE(pDatafilePathNode);
					// Add UnresolvedServer node and value
					IXMLDOMElement* pUnresolvedServerNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"UnresolvedServer", L"\n\t", pServiceNode, &pUnresolvedServerNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].exchangeAccountInfo->wszUnresolvedServer.c_str(), pUnresolvedServerNode));
					SAFE_RELEASE(pUnresolvedServerNode);
					// Add HomeServer node and value
					IXMLDOMElement* pHomeServerNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"HomeServer", L"\n\t", pServiceNode, &pHomeServerNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].exchangeAccountInfo->wszHomeServerName.c_str(), pHomeServerNode));
					SAFE_RELEASE(pHomeServerNode);
					// Add HomeServerDN node and value
					IXMLDOMElement* pHomeServerDNNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"HomeServerDN", L"\n\t", pServiceNode, &pHomeServerDNNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].exchangeAccountInfo->wszHomeServerDN.c_str(), pHomeServerDNNode));
					SAFE_RELEASE(pHomeServerDNNode);
					// Add RohProxyServer node and value
					IXMLDOMElement* pRohProxyServerNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"RohProxyServer", L"\n\t", pServiceNode, &pRohProxyServerNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].exchangeAccountInfo->wszRohProxyServer.c_str(), pRohProxyServerNode));
					SAFE_RELEASE(pRohProxyServerNode);
					// Add CachedModeEnabledOwner node and value
					IXMLDOMElement* pCachedModeEnabledOwnerNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"CachedModeEnabledOwner", L"\n\t", pServiceNode, &pCachedModeEnabledOwnerNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(profileInfo[i].profileServices[j].exchangeAccountInfo->bCachedModeEnabledOwner)), pCachedModeEnabledOwnerNode));
					SAFE_RELEASE(pCachedModeEnabledOwnerNode);
					// Add CachedModeEnabledShared node and value
					IXMLDOMElement* pCachedModeEnabledSharedNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"CachedModeEnabledShared", L"\n\t", pServiceNode, &pCachedModeEnabledSharedNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(profileInfo[i].profileServices[j].exchangeAccountInfo->bCachedModeEnabledShared)), pCachedModeEnabledSharedNode));
					SAFE_RELEASE(pCachedModeEnabledSharedNode);
					// Add CachedModeEnabledPublicFolders node and value
					IXMLDOMElement* pCachedModeEnabledPublicFoldersNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"CachedModeEnabledPublicFolders", L"\n\t", pServiceNode, &pCachedModeEnabledPublicFoldersNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(profileInfo[i].profileServices[j].exchangeAccountInfo->bCachedModeEnabledPublicFolders)), pCachedModeEnabledPublicFoldersNode));
					SAFE_RELEASE(pCachedModeEnabledPublicFoldersNode);
					// Add CachedModeMonths node and value
					IXMLDOMElement* pCachedModeMonthsNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"CachedModeMonths", L"\n\t", pServiceNode, &pCachedModeMonthsNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(profileInfo[i].profileServices[j].exchangeAccountInfo->iCachedModeMonths)), pCachedModeMonthsNode));
					SAFE_RELEASE(pCachedModeMonthsNode);
					// Add UserName node and value
					IXMLDOMElement* pUserNameNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"UserName", L"\n\t", pServiceNode, &pUserNameNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].exchangeAccountInfo->szUserName.c_str(), pUserNameNode));
					SAFE_RELEASE(pUserNameNode);
					// Add UserEmailSmtpAddress node and value
					IXMLDOMElement* pUserEmailSmtpAddressNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"UserEmailSmtpAddress", L"\n\t", pServiceNode, &pUserEmailSmtpAddressNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].exchangeAccountInfo->szUserEmailSmtpAddress.c_str(), pUserEmailSmtpAddressNode));
					SAFE_RELEASE(pUserEmailSmtpAddressNode);
					// Add MailboxCount node and value
					IXMLDOMElement* pMailboxCountNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"MailboxCount", L"\n\t", pServiceNode, &pMailboxCountNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(profileInfo[i].profileServices[j].exchangeAccountInfo->ulMailboxCount)), pMailboxCountNode));
					SAFE_RELEASE(pMailboxCountNode);

					if (profileInfo[i].profileServices[j].exchangeAccountInfo->ulMailboxCount > 0)
					{
						// create root node for Mailboxes
						IXMLDOMElement* pMailboxesRootNode = NULL;
						CHK_HR(CreateAndAddElementNode(pXMLDom, L"Mailboxes", L"\n\t", pServiceNode, &pMailboxesRootNode));

						for (unsigned int k = 0; k < profileInfo[i].profileServices[j].exchangeAccountInfo->ulMailboxCount; k++)
						{
							// create child node for each mailbox
							IXMLDOMElement* pMailboxNode = NULL;
							CHK_HR(CreateAndAddElementNode(pXMLDom, L"Mailbox", L"\n\t", pMailboxesRootNode, &pMailboxNode));
							// Add DisplayName node and value
							IXMLDOMElement* pDisplayNameNode = NULL;
							CHK_HR(CreateAndAddElementNode(pXMLDom, L"DisplayName", L"\n\t", pMailboxNode, &pDisplayNameNode));
							CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].exchangeAccountInfo->accountMailboxes[k].wszDisplayName.c_str(), pDisplayNameNode));
							SAFE_RELEASE(pDisplayNameNode);
							// Add MailboxIndex node and value
							IXMLDOMElement* pMailboxIndexNode = NULL;
							CHK_HR(CreateAndAddElementNode(pXMLDom, L"MailboxIndex", L"\n\t", pMailboxNode, &pMailboxIndexNode));
							CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(k + 1)), pMailboxIndexNode));
							SAFE_RELEASE(pMailboxIndexNode);
							// Add DefaultMailbox node and value
							IXMLDOMElement* pDefaultMailboxNode = NULL;
							CHK_HR(CreateAndAddElementNode(pXMLDom, L"DefaultMailbox", L"\n\t", pMailboxNode, &pDefaultMailboxNode));
							CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(profileInfo[i].profileServices[j].exchangeAccountInfo->accountMailboxes[k].bPrimaryMailbox)), pDefaultMailboxNode));
							SAFE_RELEASE(pDefaultMailboxNode);
							// Add EntryType node and value
							IXMLDOMElement* pEntryTypeNode = NULL;
							CHK_HR(CreateAndAddElementNode(pXMLDom, L"EntryType", L"\n\t", pMailboxNode, &pEntryTypeNode));
							switch (profileInfo[i].profileServices[j].exchangeAccountInfo->accountMailboxes[k].ulProfileType)
							{
							case PROFILE_PRIMARY_USER:
								CHK_HR(CreateAndAddTextNode(pXMLDom, L"Primary", pEntryTypeNode));
								break;
							case PROFILE_DELEGATE:
								CHK_HR(CreateAndAddTextNode(pXMLDom, L"Delegate", pEntryTypeNode));
								break;
							case PROFILE_PUBLIC_STORE:
								CHK_HR(CreateAndAddTextNode(pXMLDom, L"Public Store", pEntryTypeNode));
								break;
							case 0:
							default:
								CHK_HR(CreateAndAddTextNode(pXMLDom, L"Unknown", pEntryTypeNode));
								break;
							}
							SAFE_RELEASE(pEntryTypeNode);
							SAFE_RELEASE(pMailboxNode);
						}
						SAFE_RELEASE(pMailboxesRootNode);
					}

					SAFE_RELEASE(pServiceNode);
				}
				if (profileInfo[i].profileServices[j].serviceType == ServiceType::ServiceType_Pst)
				{
					// create child node for each service
					IXMLDOMElement* pServiceNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"Pst", L"\n\t", pServicesRootNode, &pServiceNode));
					// Add ServiceName node and value
					IXMLDOMElement* pServiceNameNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"ServiceName", L"\n\t", pServiceNode, &pServiceNameNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].wszServiceName.c_str(), pServiceNameNode));
					SAFE_RELEASE(pServiceNameNode);
					// Add ServiceIndex node and value
					IXMLDOMElement* pServiceIndexNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"ServiceIndex", L"\n\t", pServiceNode, &pServiceIndexNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, ConvertStdStringToWideChar(ConvertIntToString(j + 1)), pServiceIndexNode));
					SAFE_RELEASE(pServiceIndexNode);
					// Add PstName node and value
					IXMLDOMElement* pPstNameNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"PstName", L"\n\t", pServiceNode, &pPstNameNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].pstInfo->wszDisplayName.c_str(), pPstNameNode));
					SAFE_RELEASE(pPstNameNode);
					// Add DatafilePath node and value
					IXMLDOMElement* pDatafilePathNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"DatafilePath", L"\n\t", pServiceNode, &pDatafilePathNode));
					CHK_HR(CreateAndAddTextNode(pXMLDom, profileInfo[i].profileServices[j].pstInfo->wszPstPath.c_str(), pDatafilePathNode));
					SAFE_RELEASE(pDatafilePathNode);
					// Add EntryType node and value
					IXMLDOMElement* pEntryTypeNode = NULL;
					CHK_HR(CreateAndAddElementNode(pXMLDom, L"PstType", L"\n\t", pServiceNode, &pEntryTypeNode));
					switch (profileInfo[i].profileServices[j].pstInfo->ulPstType)
					{
					case PSTTYPE_ANSI:
						CHK_HR(CreateAndAddTextNode(pXMLDom, L"Ansi", pEntryTypeNode));
						break;
					case PSTTYPE_UNICODE:
						CHK_HR(CreateAndAddTextNode(pXMLDom, L"Unicode", pEntryTypeNode));
						break;
					}
					SAFE_RELEASE(pEntryTypeNode);

					SAFE_RELEASE(pServiceNode);
				}
			}
			SAFE_RELEASE(pServicesRootNode);

		}
		SAFE_RELEASE(pProfileNode);
	}

	// Add NEWLINE for identation before </root>.
	CHK_HR(CreateAndAddTextNode(pXMLDom, L"\n", pRoot));
	// add <root> to document
	CHK_HR(AppendChildToParent(pRoot, pXMLDom));

	CHK_HR(pXMLDom->get_xml(&bstrXML));
	Logger::Write(logLevelSuccess, L"Wrote info to xml :" + std::wstring(bstrXML));
	if (szExportPath != L"")
	{
		std::wstring szComputerName = _wgetenv(L"COMPUTERNAME");
		std::wstring szUserName = _wgetenv(L"USERNAME");
		std::wstring szFullExportPath = szExportPath + L"\\" + szComputerName + L"_" + szUserName + L".xml";
		CHK_HR(VariantFromString(szFullExportPath.c_str(), varFileName));
		CHK_HR(pXMLDom->save(varFileName));
		Logger::Write(logLevelSuccess, L"Profile information saved to " + szFullExportPath);
	}
	else
	{
		std::wstring szComputerName = _wgetenv(L"COMPUTERNAME");
		std::wstring szUserName = _wgetenv(L"USERNAME");
		std::wstring szFullExportPath = szComputerName + L"_" + szUserName + L".xml";
		CHK_HR(VariantFromString(szFullExportPath.c_str(), varFileName));
		CHK_HR(pXMLDom->save(varFileName));
		Logger::Write(logLevelSuccess, L"Profile information saved to " + szFullExportPath);
	}

CleanUp:
	SAFE_RELEASE(pXMLDom);
	SAFE_RELEASE(pRoot);
	SysFreeString(bstrXML);
	VariantClear(&varFileName);
}


