#pragma once
#include "../stdafx.h"

struct OlkAccount
{
public:

	typedef struct {
		DWORD	cb;
		BYTE * pb;
	} BIN;

	// Common Properties - 0x0001 : 0x00FF
	long	lAcctId;
	LPWSTR szAcctName;
	long	lAcctMiniUid;
	LPWSTR szAcctType;
	LPWSTR szConnInfo;
	LPWSTR szSchedInfo;
	LPWSTR szAcctIdentity;
	LPWSTR szAcctFlavor;
	long	lAcctInclude;
	long	lAcctIsDefaultMail;
	LPWSTR	szAcctUserDisplayName;
	LPWSTR szAcctUserEmailAddr;
	LPWSTR szAcctStamp;
	LPWSTR szAcctSendStamp;
	long	lAcctConnectionType;
	LPWSTR szAcctConnectoid;
	long	lAcctForceId;
	long	lAcctUnicodeSupport;
	long	lAcctMigrationFlags;
	long	lAcctIsExch;
	long	lAcctDisabled;
	LPWSTR szAcctNewSignature;
	LPWSTR szAcctReplyFwdSignature;
	long	lAcctTcidIcon;
	long	lAcctPreferencesUid;

	//General Internet Mail Properties - 0x0100 : 0x01FF
	LPWSTR szInetServer;
	LPWSTR szInetUser;
	LPWSTR szInetPassword;
	LPWSTR szInetReplyEmail;
	long	lInetPort;
	long	lInetSSL;
	long	lInetRememberPassword;
	LPWSTR szInetOrganization;
	long	lInetUseSpa;


	// Specific SMTP Properties - 0x0200 : 0x02FF
	LPWSTR szSmtpServer;
	long	lSmtpPort;
	long	lSmtpSSL;
	long	lSmtpUseAuth;
	LPWSTR szSmtpUser;
	LPWSTR szSmtpPassword;
	long	lSmtpRememberPassword;
	long	lSmtpUseSpa;
	long	lSmtpAuthMethod;
	long	lSmtpTimeout;

	// Specific POP Properties - 0x1000 : 0x10FF
	long	lPopLeaveOnServer;
	LPWSTR szPopMigrateRhc;
	BIN		bPopUserEntryId;
	BIN		bPopUserSearchKey;

	// Specific IMAP Properties - 0x1100 : 0x11FF
	long	lImapPollAllFolders;
	LPWSTR lImapRootFolder;
	long	lImapUseLsub;
	long	lImapFullList;
	long	lImapNoopInterval;

	// Specific MAPI Properties - 0x2000 : 0x20FF
	LPWSTR szMapiServiceName;
	BIN		bMapiServiceUid;
	long	lMapiProviderType;
	BIN		bMapiIdentityEntryId;
	long	lMapiXpStatusCode;
	long	lMapiXpCapabilities;

	LPWSTR szPrimarySendAcct;
	LPWSTR szNextSendAcct;

};
