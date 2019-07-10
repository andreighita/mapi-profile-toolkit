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

#pragma once
#include "..\\stdafx.h"
#include <MAPIX.h>
#define AB_PROVIDER_BASE_ID						0x6600  // Look at the comments in MAPITAGS.H
#define PROP_AB_PROVIDER_DISPLAY_NAME			PR_DISPLAY_NAME
#define PROP_AB_PROVIDER_SERVER_NAME			PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0000))	// "example.contoso.com"
#define PROP_AB_PROVIDER_SERVER_PORT			PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0001)) // "389"
#define PROP_AB_PROVIDER_USER_NAME				PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0002)) // contoso\administrator
#define PROP_AB_PROVIDER_SEARCH_BASE			PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0003)) // SEARCH_FILTER_VALUE
#define PROP_AB_PROVIDER_SEARCH_FILTER			PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0004)) // "(&(mail=*)(|(mail=%s*)(|(cn=%s*)(|(sn=%s*)(givenName=%s*)))))"
#define PROP_AB_PROVIDER_ADDRTYPE				PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0005))	// "SMTP"
#define PROP_AB_PROVIDER_SOURCE					PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0006))	// "mail"
#define PROP_AB_PROVIDER_SEARCH_TIMEOUT			PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0007))	// "60"
#define PROP_AB_PROVIDER_MAX_ENTRIES			PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0008)) // "100"
#define PROP_AB_PROVIDER_SEARCH_TIMEOUT_LBW		PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0009))	// "120"
#define PROP_AB_PROVIDER_MAX_ENTRIES_LBW		PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x000a)) // "15"
#define PROP_AB_PROVIDER_LOGFILE				PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x000b)) // ""
#define PROP_AB_PROVIDER_ERRLOGGING				PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x000c)) // "OFF"
#define PROP_AB_PROVIDER_DIAGTRACING			PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x000d)) // "OFF"
#define PROP_AB_PROVIDER_TRACELEVEL				PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x000e)) // "NONE"
#define PROP_AB_PROVIDER_DEBUGWIN				PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x000f)) // "OFF"
#define PROP_AB_PROVIDER_ADDITIONAL_INFO_SOURCE	PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0010)) // "postalAddress"
#define PROP_AB_PROVIDER_DISPLAY_NAME_SOURCE	PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0011)) // "cn"
#define PROP_AB_PROVIDER_LDAP_UI				PROP_TAG (PT_TSTRING,	(AB_PROVIDER_BASE_ID + 0x0012)) // "1"
#define PROP_AB_PROVIDER_USE_SSL				PROP_TAG (PT_BOOLEAN,	(AB_PROVIDER_BASE_ID + 0x0013)) // False
#define PROP_AB_PROVIDER_SERVER_SPA				PROP_TAG (PT_BOOLEAN,	(AB_PROVIDER_BASE_ID + 0x0015)) // False
#define PROP_AB_PROVIDER_USER_PASSWORD_ENCODED	PROP_TAG (PT_BINARY,	(AB_PROVIDER_BASE_ID + 0x0017)) // ENCODED_PWD
#define PROP_AB_PROVIDER_ENABLE_BROWSING		PROP_TAG(PT_BOOLEAN,	(AB_PROVIDER_BASE_ID + 0x0022)) // False
#define PROP_AB_PROVIDER_SEARCH_BASE_DEFAULT	PROP_TAG(PT_LONG,		(AB_PROVIDER_BASE_ID + 0x0023)) // 0 or 1

struct ABProvider
{
	LPTSTR lpszDisplayName;  // LPTSTR = LPWSTR; LPSTR  
	LPTSTR lpszServerName;
	LPTSTR lpszServerPort;
	BOOL bUseSSL;
	LPTSTR lpszUsername;
	LPTSTR lpszPassword;
	BOOL bRequireSPA;
	LPTSTR lpszTimeout;
	LPTSTR lpszMaxResults;
	ULONG ulDefaultSearchBase;
	LPTSTR lpszCustomSearchBase;
	BOOL bEnableBrowsing;
	LPTSTR lpszServiceName;
};

