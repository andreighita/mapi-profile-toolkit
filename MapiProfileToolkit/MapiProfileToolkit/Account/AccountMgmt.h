#pragma once
#pragma once

#include <mapix.h>
#include <mapidefs.h>
#include <guiddef.h>
#include <MAPIUTIL.H>
#include "AccountObjects.h"

typedef unsigned long PT_DWORD;

#define ACCT_INIT_NOSYNCH_MAPI_ACCTS	0x00000001 
#define ACCT_INIT_NO_STORES_CHECK		0x00000002

#define ACCTUI_NO_WARNING				0x0100
#define	ACCTUI_SHOW_DATA_TAB			0x0200
#define ACCTUI_SHOW_ACCTWIZARD			0x0400

#define E_ACCT_NOT_FOUND				0x800C8101 
#define E_ACCT_UI_BUSY					0x800C8102
#define E_ACCT_WRONG_SORT_ORDER			0x800C8105 

#define E_OLK_ALREADY_INITIALIZED		0x800C8002 
#define E_OLK_NOT_INITIALIZED			0x800C8005 
#define E_OLK_PARAM_NOT_SUPPORTED		0x800C8003 
#define E_OLK_PROP_READ_ONLY			0x800C800D 
#define E_OLK_REGISTRY					0x800C8002 

#define NOTIFY_ACCT_CHANGED				1 
#define NOTIFY_ACCT_CREATED				2 
#define NOTIFY_ACCT_DELETED				3 
#define NOTIFY_ACCT_ORDER_CHANGED		4 
#define NOTIFY_ACCT_PREDELETED			5 

#define OLK_ACCOUNT_NO_FLAGS			0 

#define SECURE_FLAG						0x8000	

// Capone profile section {00020D0A-0000-0000-C000-000000000046}

DEFINE_OLEGUID(IID_CAPONE_PROF, 0x00020d0a, 0, 0);

//{ed475410-b0d6-11d2-8c3b-00104b2a6676}

DEFINE_GUID(CLSID_OlkAccountManager, 0xed475410, 0xb0d6, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{ed475411-b0d6-11d2-8c3b-00104b2a6676}

DEFINE_GUID(CLSID_OlkPOP3Account, 0xed475411, 0xb0d6, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{ed475412-b0d6-11d2-8c3b-00104b2a6676}

DEFINE_GUID(CLSID_OlkIMAP4Account, 0xed475412, 0xb0d6, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{ed475414-b0d6-11d2-8c3b-00104b2a6676}

DEFINE_GUID(CLSID_OlkMAPIAccount, 0xed475414, 0xb0d6, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{ed475418-b0d6-11d2-8c3b-00104b2a6676}

DEFINE_GUID(CLSID_OlkMail, 0xed475418, 0xb0d6, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{ed475419-b0d6-11d2-8c3b-00104b2a6676}

DEFINE_GUID(CLSID_OlkAddressBook, 0xed475419, 0xb0d6, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{ed475420-b0d6-11d2-8c3b-00104b2a6676}

DEFINE_GUID(CLSID_OlkStore, 0xed475420, 0xb0d6, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{4db5cbf0-3b77-4852-bc8e-bb81908861f3}

DEFINE_GUID(CLSID_OlkHotmailAccount, 0x4db5cbf0, 0x3b77, 0x4852, 0xbc, 0x8e, 0xbb, 0x81, 0x90, 0x88, 0x61, 0xf3);

//{4db5cbf2-3b77-4852-bc8e-bb81908861f3}

DEFINE_GUID(CLSID_OlkLDAPAccount, 0x4db5cbf2, 0x3b77, 0x4852, 0xbc, 0x8e, 0xbb, 0x81, 0x90, 0x88, 0x61, 0xf3);


//Interface Identifiers

//{9240A6C0-AF41-11d2-8C3B-00104B2A6676}

DEFINE_GUID(IID_IOlkErrorUnknown, 0x9240a6c0, 0xaf41, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{9240A6C1-AF41-11d2-8C3B-00104B2A6676}

DEFINE_GUID(IID_IOlkEnum, 0x9240a6c1, 0xaf41, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{9240a6c3-af41-11d2-8c3b-00104b2a6676}

DEFINE_GUID(IID_IOlkAccountNotify, 0x9240a6c3, 0xaf41, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{9240a6cb-af41-11d2-8c3b-00104b2a6676}

DEFINE_GUID(IID_IOlkAccountHelper, 0x9240a6cb, 0xaf41, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{9240a6cd-af41-11d2-8c3b-00104b2a6676}

DEFINE_GUID(IID_IOlkAccountManager, 0x9240a6cd, 0xaf41, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);

//{9240a6d2-af41-11d2-8c3b-00104b2a6676}

DEFINE_GUID(IID_IOlkAccount, 0x9240a6d2, 0xaf41, 0x11d2, 0x8c, 0x3b, 0x0, 0x10, 0x4b, 0x2a, 0x66, 0x76);


typedef struct {
	DWORD	cb;
	BYTE	* pb;
} ACCT_BIN;

typedef struct
{
	DWORD dwType;
	union
	{
		DWORD dw;
		WCHAR *pwsz;
		ACCT_BIN bin;
	} Val;
} ACCT_VARIANT;

typedef struct {
	CLSID clsidType;
	DWORD dwFlags;
	WCHAR wszName[128];
	WCHAR wszFlavor[128];
} ACCT_TYPE;

// Common Properties - 0x0001 : 0x00FF
// ===================================
#define PROP_ACCT_ID					PROP_TAG(PT_LONG,   0x0001)
#define PROP_ACCT_NAME					PROP_TAG(PT_UNICODE, 0x0002)
#define PROP_ACCT_MINI_UID              PROP_TAG(PT_LONG,   0x0003)
#define PROP_ACCT_TYPE					PROP_TAG(PT_UNICODE, 0x0004)
#define PROP_ACCT_IDENTITY				PROP_TAG(PT_UNICODE, 0x0007)
#define PROP_ACCT_FLAVOR				PROP_TAG(PT_UNICODE, 0x0008)
#define PROP_ACCT_IS_DEFAULT_MAIL		PROP_TAG(PT_LONG,   0x000A)
#define PROP_ACCT_USER_DISPLAY_NAME		PROP_TAG(PT_UNICODE, 0x000B)
#define PROP_ACCT_USER_EMAIL_ADDR		PROP_TAG(PT_UNICODE, 0x000C)
#define PROP_ACCT_STAMP					PROP_TAG(PT_UNICODE, 0x000D)
#define PROP_ACCT_SEND_STAMP			PROP_TAG(PT_UNICODE, 0x000E)
#define PROP_ACCT_IS_EXCH				PROP_TAG(PT_LONG,  0x0014)
#define PROP_ACCT_DISABLED				PROP_TAG(PT_LONG,  0x0015)
#define PROP_ACCT_PREFERENCES_UID		PROP_TAG(PT_BINARY, 0x0022)

// Specific MAPI Properties - 0x2000 : 0x20FF
// ==========================================
#define PROP_MAPI_SERVICE_NAME			PROP_ACCT_FLAVOR
#define PROP_MAPI_SERVICE_UID			PROP_TAG(PT_BINARY,  0x2000)
#define PROP_MAPI_PROVIDER_TYPE			PROP_TAG(PT_LONG,   0x2001)
#define PROP_MAPI_IDENTITY_ENTRYID		PROP_TAG(PT_BINARY,  0x2002)

#define PR_PRIMARY_SEND_ACCT				PROP_TAG(PT_UNICODE, 0x0e28)
#define PR_NEXT_SEND_ACCT					PROP_TAG(PT_UNICODE, 0x0e29)

//Properties related to profile having exchange account only
#define PROP_EXCHANGE_EMAILID			PROP_TAG(PT_UNICODE, 0x663D) //email ID
#define PROP_EXCHANGE_EMAILID2			PROP_TAG(PT_UNICODE, 0x6641) //email ID
#define PR_ROH_PROXY_SERVER				PROP_TAG(PT_UNICODE, 0x6622) //RPC server name
#define PR_INTERNET_CONTENT_ID			PROP_TAG(PT_UNICODE, 0x662A) //server name

//start::IOlkAccountManager::DisplayAccountList flags
#define E_ACCT_UI_BUSY 0x800C8102
#define ACCTUI_NO_WARNING      0x0100
#define ACCTUI_SHOW_DATA_TAB   0x0200
#define ACCTUI_SHOW_ACCTWIZARD 0x0400
//end::IOlkAccountManager::DisplayAccountList flags

interface IOlkErrorUnknown : IUnknown
{
	//GetLastError Gets a message string for the specified error.  
	virtual STDMETHODIMP GetLastError(HRESULT hr, LPWSTR* ppwszError);
};

typedef IOlkErrorUnknown FAR * LPOLKERRORUNKNOWN;

interface IOlkAccount : IOlkErrorUnknown
{
public:
	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder1();
	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder2();
	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder3();
	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder4();
	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder5();
	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder6();

	//GetAccountInfo Gets the type and categories of the specified account. 
	virtual STDMETHODIMP GetAccountInfo(CLSID* pclsidType, DWORD* pcCategories, CLSID** prgclsidCategory);
	//GetProp Gets the value of the specified account property. See the Properties table below. 
	virtual STDMETHODIMP GetProp(DWORD dwProp, ACCT_VARIANT* pVar);
	//SetProp Sets the value of the specified account property. See the Properties table below. 
	virtual STDMETHODIMP SetProp(DWORD dwProp, ACCT_VARIANT* pVar);

	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder7();
	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder8();
	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder9();

	//FreeMemory Frees memory allocated by the IOlkAccount interface. 
	virtual STDMETHODIMP FreeMemory(BYTE* pv);

	//Placeholder member Not supported or documented. 
	virtual STDMETHODIMP PlaceHolder10();

	//SaveChanges Saves changes to the specified account. 
	virtual STDMETHODIMP SaveChanges(DWORD dwFlags);
};

typedef IOlkAccount FAR * LPOLKACCOUNT;

interface IOlkAccountHelper : IUnknown
{
public:
	//Placeholder1 This member is a placeholder and is not supported.
	virtual STDMETHODIMP PlaceHolder1(LPVOID) = 0;

	//GetIdentity Gets the profile name of an account. 
	virtual STDMETHODIMP GetIdentity(LPWSTR pwszIdentity, DWORD * pcch) = 0;
	//GetMapiSession Gets the current MAPI session. 
	virtual STDMETHODIMP GetMapiSession(LPUNKNOWN * ppmsess) = 0;
	//HandsOffSession Releases the current MAPI session that has been created by 
	//IOlkAccountHelper::GetMapiSession. 
	virtual STDMETHODIMP HandsOffSession() = 0;
};

typedef IOlkAccountHelper FAR * LPOLKACCOUNTHELPER;

interface IOlkAccountNotify : IOlkErrorUnknown
{
public:
	//Notify Notifies the client of changes to the specified account. 
	STDMETHODIMP Notify(DWORD dwNotify, DWORD dwAcctID, DWORD dwFlags);
};

typedef IOlkAccountNotify FAR * LPOLKACCOUNTNOTIFY;

interface IOlkEnum : IUnknown
{
public:
	//GetCount  Gets the number of accounts in the enumerator. 
	virtual STDMETHODIMP GetCount(DWORD *pulCount);
	//Reset Resets the enumerator to the beginning. 
	virtual STDMETHODIMP Reset();
	//GetNext Gets the next account in the enumerator. 
	virtual STDMETHODIMP GetNext(LPUNKNOWN* ppunk);
	//Skip Skips a specified number of accounts in the enumerator. 
	virtual STDMETHODIMP Skip(DWORD cSkip);
};

typedef IOlkEnum FAR * LPOLKENUM;

interface IOlkAccountManager : IOlkErrorUnknown
{
public:
	//Init Initializes the account manager for use. 
	virtual STDMETHODIMP Init(IOlkAccountHelper* pAcctHelper, DWORD dwFlags);

	//Placeholder member Not supported or documented 
	//virtual STDMETHODIMP PlaceHolder1();
	//DisplayAccountList Displays the account list wizard
	virtual STDMETHODIMP DisplayAccountList(
		HWND hwnd,
		DWORD dwFlags,
		LPCWSTR lpwszReserved, // Not used
		DWORD dwReserved, // Not used
		const CLSID * pclsidReserved1, // Not used
		const CLSID * pclsidReserved2); // Not used


										//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder2();
	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder3();
	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder4();
	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder5();
	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder6();

	//FindAccount Finds an account by property value. 
	virtual STDMETHODIMP FindAccount(DWORD dwProp, ACCT_VARIANT* pVar, IOlkAccount** ppAccount);

	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder7();
	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder8();
	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder9();

	//DeleteAccount Deletes the specified account. 
	virtual STDMETHODIMP DeleteAccount(DWORD dwAcctID);

	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder10();

	//SaveChanges Saves changes to the specified account. 
	virtual STDMETHODIMP SaveChanges(DWORD dwAcctID, DWORD dwFlags);
	//GetOrder Gets the ordering of the specified category of accounts. 
	virtual STDMETHODIMP GetOrder(const CLSID* pclsidCategory, DWORD* pcAccts, DWORD* prgAccts[]);
	//SetOrder Modifies the ordering of the specified category of accounts. 
	virtual STDMETHODIMP SetOrder(const CLSID* pclsidCategory, DWORD* pcAccts, DWORD* prgAccts[]);
	//EnumerateAccounts Gets an enumerator for the accounts of the specific category and type. 
	virtual STDMETHODIMP EnumerateAccounts(const CLSID* pclsidCategory, const CLSID* pclsidType, DWORD dwFlags, IOlkEnum** ppEnum);

	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder11();
	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder12();

	//FreeMemory Frees memory allocated by the IOlkAccountManager interface. 
	virtual STDMETHODIMP FreeMemory(BYTE* pv);
	//Advise Registers an account for notifications sent by the account manager. 
	virtual STDMETHODIMP Advise(IOlkAccountNotify* pNotify, DWORD* pdwCookie);
	//Unadvise Unregisters an account for notifications sent by the account manager. 
	virtual STDMETHODIMP Unadvise(DWORD* pdwCookie);

	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder13();
	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder14();
	//Placeholder member Not supported or documented 
	virtual STDMETHODIMP PlaceHolder15();
};

typedef IOlkAccountManager FAR * LPOLKACCOUNTMANAGER;
