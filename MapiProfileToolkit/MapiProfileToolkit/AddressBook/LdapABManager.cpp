//// LdapABManager.cpp : Defines the entry point for the console application.
////
///*
//* © 2016 Microsoft Corporation
//*
//* written by Andrei Ghita
//*
//* Microsoft provides programming examples for illustration only, without warranty either expressed or implied.
//* This includes, but is not limited to, the implied warranties of merchantability or fitness for a particular purpose.
//* This sample assumes that you are familiar with the programming language that is being demonstrated and with
//* the tools that are used to create and to debug procedures. Microsoft support engineers can help explain the
//* functionality of a particular procedure, but they will not modify these examples to provide added functionality
//* or construct procedures to meet your specific requirements.
//*/
//
//#include "..\stdafx.h"
//#include "LdapABManager.h"
//#include "ABProviderObjects.h"
//#include "ConfigXmlParser.h"
//#include "..\Profile.h"
//#include "Shlwapi.h"
//
//BOOL ParseArgs(int argc, char * argv[], ABManagerOptions * pRunOpts)
//{
//	if (!pRunOpts) return FALSE;
//
//	ZeroMemory(pRunOpts, sizeof(ABManagerOptions));
//
//	// Setting running mode to Read as a default
//	pRunOpts->ulRunningMode = RUNNINGMODE_LIST;
//
//	for (int i = 1; i < argc; i++)
//	{
//		switch (argv[i][0])
//		{
//		case '-':
//		case '/':
//		case '\\':
//			if (0 == argv[i][1])
//			{
//				// Bad argument - get out of here
//				return false;
//			}
//			switch (tolower(argv[i][1]))
//			{
//			case 'p':
//				if (tolower(argv[i][2]) == 'm')
//				{
//					if (i + 1 < argc)
//					{
//						std::string profileMode = argv[i + 1];
//						std::transform(profileMode.begin(), profileMode.end(), profileMode.begin(), ::tolower);
//						if (profileMode == "one")
//						{
//							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_SPECIFIC;
//							i++;
//						}
//						else if (profileMode == "default")
//						{
//							pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_DEFAULT;
//							i++;
//						}
//						else
//						{
//							return false;
//						}
//					}
//				}
//				else if (tolower(argv[i][2]) == 'n')
//				{
//					if (i + 1 < argc)
//					{
//						pRunOpts->szProfileName = argv[i + 1];
//						pRunOpts->ulProfileMode = (ULONG)PROFILEMODE_SPECIFIC;
//						i++;
//					}
//					else return false;
//				}
//				else return false;
//				break;
//			case 'd':
//				if (tolower(argv[i][2]) == 'n')
//				{
//					if (i + 1 < argc)
//					{
//						pRunOpts->szABDisplayName = argv[i + 1];
//						i++;
//					}
//					else return false;
//				}
//				else return false;
//				break;
//			case 's':
//				if (tolower(argv[i][2]) == 'n')
//				{
//					if (i + 1 < argc)
//					{
//						pRunOpts->szABDisplayName = argv[i + 1];
//						i++;
//					}
//					else return false;
//				}
//				else return false;
//				break;
//			case 'f':
//				if (i + 1 < argc)
//				{
//					pRunOpts->szConfigFilePath = argv[i + 1];
//					i++;
//				}
//				else return false;
//				break;
//			case 'm':
//				if (i + 1 < argc)
//				{
//					std::string runningMode = argv[i + 1];
//					std::transform(runningMode.begin(), runningMode.end(), runningMode.begin(), ::tolower);
//					if (runningMode == "list")
//					{
//						pRunOpts->ulRunningMode = (ULONG)RUNNINGMODE_LIST;
//						i++;
//					}
//					else if (runningMode == "create")
//					{
//						pRunOpts->ulRunningMode = (ULONG)RUNNINGMODE_CREATE;
//						i++;
//					}
//					else if (runningMode == "update")
//					{
//						pRunOpts->ulRunningMode = (ULONG)RUNNINGMODE_UPDATE;
//						i++;
//					}
//					else if (runningMode == "remove")
//					{
//						pRunOpts->ulRunningMode = (ULONG)RUNNINGMODE_REMOVE;
//						i++;
//					}
//					else
//					{
//						return false;
//					}
//				}
//				else return false;
//				break;
//			case '?':
//			default:
//				// display help
//				return false;
//				break;
//			}
//		}
//	}
//
//
//
//	// If no profile mode or index or name specified then use default
//	if (pRunOpts->szProfileName.empty())
//	{
//		if (pRunOpts->ulProfileMode == 0)
//		{
//			pRunOpts->ulProfileMode = PROFILEMODE_DEFAULT;
//		}
//
//	}
//
//	// If running mode is RUNNINGMODE_WRITE then expect a profile section name or a service index or a service type
//	if (pRunOpts->ulRunningMode == RUNNINGMODE_REMOVE)
//	{
//		if (!pRunOpts->szABDisplayName.empty())
//		{
//			return true;
//		}
//		else return false;
//	}
//	else if (pRunOpts->ulRunningMode == RUNNINGMODE_UPDATE)
//	{
//		if ((!pRunOpts->szConfigFilePath.empty()) && (!pRunOpts->szABDisplayName.empty()))
//		{
//			return true;
//		}
//		else return false;
//	}
//	else if (pRunOpts->ulRunningMode == RUNNINGMODE_CREATE)
//	{
//		if (!pRunOpts->szConfigFilePath.empty())
//		{
//			return true;
//		}
//		else return false;
//	}
//	return true;
//}
//
//void DisplayUsage()
//{
//	wprintf(L"DISCLAIMER:\n");
//
//	wprintf(L"Microsoft provides programming examples for illustration only, without \n");
//	wprintf(L"warranty either expressed or implied.This includes, but is not limited\n");
//	wprintf(L"to, the implied warranties of merchantability or fitness for a particular\n");
//	wprintf(L"purpose.This sample assumes that you are familiar with the programming\n");
//	wprintf(L"language that is being demonstrated and with the tools that are used to\n");
//	wprintf(L"create and to debug procedures.Microsoft support engineers can help\n");
//	wprintf(L"explain the functionality of a particular procedure, but they will not\n");
//	wprintf(L"modify these examples to provide added functionality or construct\n");
//	wprintf(L"procedures to meet your specific requirements.\n");
//	wprintf(L"\n");
//	wprintf(L"\n");
//	wprintf(L"LdapABManager - Ldap Address Book Manager\n");
//	wprintf(L"    Sample utility for listing, creating, updating or removing Ldap address books. \n");
//	wprintf(L"\n");
//	wprintf(L"Usage: LdapABManager [-?] [-pm <one, default>] [-pn profilename] \n");
//	wprintf(L"       [-dn displayname] [-sn servername] [-f configurationfilepath] \n");
//	wprintf(L"       [-m <listall, listone, create, update, remove>] \n");
//	wprintf(L"\n");
//	wprintf(L"Options:\n");
//	wprintf(L"       -pm : Sets the profile mode.\n");
//	wprintf(L"              \"default\" to process the default profile.\n");
//	wprintf(L"              \"one\" to process a specific profile. The profile Name needs to be \n");
//	wprintf(L"              specified using -pn. Default profile will be used if -pm is not \n");
//	wprintf(L"              used.\n");
//	wprintf(L"       -pn : Name of the profile to process.\n");
//	wprintf(L"             Default profile will be used if -pn is not used.\n");
//	wprintf(L"\n");
//	wprintf(L"       -dn : Display name of the Ldap Addressbook to list.\n");
//	wprintf(L"\n");
//	wprintf(L"       -f  : Full path to the XML configuration file. For example: \n");
//	wprintf(L"             \"C:\\LdapABManager\\ABConfiguration.xml\".\n");
//	wprintf(L"\n");
//	wprintf(L"       -m  : Sets the running mode.\n");
//	wprintf(L"              \"list\" to list all Ldap Addressbooks \n");
//	wprintf(L"              \"create\" to create a new Ldap Addressbook \n");
//	wprintf(L"       	      Must be used in conjunction with -f .\n");
//	wprintf(L"              \"update\" to update a specifc Ldap Addressbook \n");
//	wprintf(L"       	      Must be used in conjunction with -dn, optionally -sn and -f.\n");
//	wprintf(L"              \"remove\" to remove a specifc Ldap Addressbook \n");
//	wprintf(L"       	      Must be used in conjunction with -dn and optionally -sn.\n");
//	wprintf(L"\n");
//	wprintf(L"       -?  : Displays this usage information.\n");
//}
//
//
//void main(int argc, char* argv[])
//{
//	HRESULT hRes = S_OK;
//	try
//	{
//		// Create a new instance of ABManagerOptions to store the runtime options
//		ABManagerOptions runtimeOptions = { 0 };
//		// Parse the input parameters and populate the runtime options.
//		if (!ParseArgs(argc, argv, &runtimeOptions))
//		{
//			DisplayUsage();
//			return;
//		}
//		else
//		{
//			if (!runtimeOptions.szConfigFilePath.empty())
//			{
//				// If a path was specified 
//				if (!PathFileExists(LPCSTR(runtimeOptions.szConfigFilePath.c_str())))
//				{
//					wprintf(L"WARNING: The specified file \"%s\" does not exsits.\n", LPTSTR(runtimeOptions.szConfigFilePath.c_str()));
//					return;
//				}
//			}
//		}
//
//		MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
//
//		// Increments the MAPI subsystem reference count and initializes global data for the MAPI DLL. 
//		// This call is require prior to any mapi threads in the current call.
//		hRes = MAPIInitialize(&MAPIINIT);
//
//		if (SUCCEEDED(hRes))
//		{
//			LPPROFADMIN lpProfAdmin = NULL;		// profile administration object pointer
//			LPSERVICEADMIN lpSvcAdmin = NULL;	// service administration object pointer
//			MAPIUID mapiUid = { 0 };			// MAPIUID structure
//			LPMAPIUID lpMapiUid = &mapiUid;		// pointer to a MAPIUID structure
//			BOOL fValidPath = false;
//			BOOL fServiceExists = false;
//			// Create a new ABProvider instance and set the service name to EMABLT (Address Book service)
//			ABProvider pABProvider = { 0 };
//			pABProvider.lpszServiceName = "EMABLT";
//
//			// Make sure the file path is valid and parse the XML to populate the ABProvider parameters
//			if (!runtimeOptions.szConfigFilePath.empty())
//			{
//				fValidPath = true;
//				EC_HRES(ParseXml(LPTSTR(runtimeOptions.szConfigFilePath.c_str()), &pABProvider));
//			}
//
//			// If we're processing the default profile then fetch the name of it and populate that in the runtime options.
//			if (runtimeOptions.ulProfileMode == PROFILEMODE_DEFAULT)
//			{
//				runtimeOptions.szProfileName = GetDefaultProfileName();
//				if (runtimeOptions.szProfileName.empty())
//				{
//					wprintf(L"ERROR: No default profile found, please specify a valid profile name.");
//					return;
//				}
//
//			}
//
//			// Create a profile administration object.
//			EC_HRES(MAPIAdminProfiles(0,		// Bitmask of flags indicating options for the service entry function. 
//				&lpProfAdmin));					// Pointer to a pointer to the new profile administration object.
//			wprintf(L"Retrieved IProfAdmin interface pointer.\n");
//
//			// Get access to a message service administration object for making changes to the message services in a profile. 
//			EC_HRES(lpProfAdmin->AdminServices(LPTSTR(runtimeOptions.szProfileName.c_str()),	// A pointer to the name of the profile to be modified. The lpszProfileName parameter must not be NULL.
//				NULL,																			// Always NULL. 
//				NULL,																			// A handle of the parent window for any dialog boxes or windows that this method displays.
//				0,																				// A bitmask of flags that controls the retrieval of the message service administration object. The following flags can be set:
//				&lpSvcAdmin));																	// A pointer to a pointer to a message service administration object.
//			wprintf(L"Retrieved IMsgServiceAdmin interface pointer.\n");
//
//			// Branching based on the selected running mode
//			switch (runtimeOptions.ulRunningMode)
//			{
//			case RUNNINGMODE_LIST:
//				wprintf(L"Running in List mode.\n");
//				// Calling ListAllABServices to list all the existing Ldap AB Servies in the selected profile
//				EC_HRES(ListAllABServices(lpSvcAdmin));
//				break;
//			case RUNNINGMODE_CREATE:
//				wprintf(L"Running in Create mode.\n");
//				if (fValidPath)
//				{
//					// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
//					// display name (and optionally, the ldap server name) supplied
//					EC_HRES(CheckABServiceExists(lpSvcAdmin, pABProvider.lpszDisplayName, pABProvider.lpszServerName, &mapiUid, &fServiceExists));
//					if (!fServiceExists)
//						// If no existing service is found then call CreateAbService to create the new service
//						EC_HRES(CreateABService(lpSvcAdmin, &pABProvider));
//					else
//						wprintf(L"The specified AB already exists.\n");
//				}
//				else
//					wprintf(L"ERROR: Invalid input file or invalid file path.");
//				break;
//			case RUNNINGMODE_UPDATE:
//				wprintf(L"Running in Update mode.\n");
//				if (fValidPath)
//				{
//					if (!runtimeOptions.szABServerName.empty())
//					{
//						// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
//						// display name (and optionally, the ldap server name) supplied
//						EC_HRES(CheckABServiceExists(lpSvcAdmin, LPTSTR(runtimeOptions.szABDisplayName.c_str()), LPTSTR(runtimeOptions.szABServerName.c_str()), &mapiUid, &fServiceExists));
//					}
//					else
//						// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
//						// display name (and optionally, the ldap server name) supplied
//						EC_HRES(CheckABServiceExists(lpSvcAdmin, LPTSTR(runtimeOptions.szABDisplayName.c_str()), &mapiUid, &fServiceExists));
//					if (fServiceExists)
//						// If the searched for service is found then call UpdateABService to update the service properties
//						EC_HRES(UpdateABService(lpSvcAdmin, &pABProvider, lpMapiUid));
//					else
//						wprintf(L"The specified AB doesn't exist.\n");
//				}
//				else
//					wprintf(L"ERROR: Invalid input file or invalid file path.");
//				break;
//			case RUNNINGMODE_REMOVE:
//				wprintf(L"Running in Remove mode.\n");
//				if (!runtimeOptions.szABServerName.empty())
//				{
//					// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
//					// display name (and optionally, the ldap server name) supplied
//					EC_HRES(CheckABServiceExists(lpSvcAdmin, LPTSTR(runtimeOptions.szABDisplayName.c_str()), LPTSTR(runtimeOptions.szABServerName.c_str()), &mapiUid, &fServiceExists));
//				}
//				else
//					// Calling CheckABServiceExists to retrieve a pointer to a MAPIUID for an existing AB service that matches the
//					// display name (and optionally, the ldap server name) supplied
//					EC_HRES(CheckABServiceExists(lpSvcAdmin, LPTSTR(runtimeOptions.szABDisplayName.c_str()), &mapiUid, &fServiceExists));
//				if (fServiceExists)
//					// If the searched for service is found then call RemoveABService to update the service properties
//					EC_HRES(RemoveABService(lpSvcAdmin, lpMapiUid));
//				else
//					wprintf(L"The specified AB doesn't exist.\n");
//				break;
//			}
//			if (lpSvcAdmin) lpSvcAdmin->Release();
//			if (lpProfAdmin) lpProfAdmin->Release();
//			MAPIUninitialize();
//		}
//		else
//			EC_HRES(hRes);
//	}
//	catch (const exception &ex)
//	{
//		std::cout << ex.what() << endl;
//	}
//
//Error:
//	return;
//}
//
