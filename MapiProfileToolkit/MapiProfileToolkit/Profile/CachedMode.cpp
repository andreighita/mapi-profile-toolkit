#include "CachedMode.h"
#include <MAPIUtil.h>

HRESULT HrSetCachedMode(LPWSTR lpwszProfileName, BOOL bDefaultProfile, BOOL bAllProfiles, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths, int iOutlookVersion)
{
	HRESULT hRes = S_OK;

	if (bDefaultProfile)
	{
		ProfileInfo* profileInfo = new ProfileInfo();
		EC_HRES_MSG(HrGetProfile((LPWSTR)GetDefaultProfileName().c_str(), profileInfo), L"Calling GetProfile");
		EC_HRES_MSG(HrSetCachedModeOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultService, bAllServices, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths, iOutlookVersion), L"Calling HrSetCachedModeOneProfile");

	}
	else if (bAllProfiles)
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo* profileInfo = new ProfileInfo[ulProfileCount];
		hRes = HrGetProfiles(ulProfileCount, profileInfo);
		for (ULONG i = 0; i <= ulProfileCount; i++)
		{
			EC_HRES_MSG(HrSetCachedModeOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultService, bAllServices, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths, iOutlookVersion), L"Calling HrSetCachedModeOneProfile");
		}
	}
	else
	{
		if (lpwszProfileName)
		{
			ProfileInfo* profileInfo = new ProfileInfo();
			hRes = HrGetProfile(lpwszProfileName, profileInfo);
			EC_HRES_MSG(HrSetCachedModeOneProfile((LPWSTR)profileInfo->wszProfileName.c_str(), profileInfo, iServiceIndex, bDefaultService, bAllServices, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths, iOutlookVersion), L"Calling HrSetCachedModeOneProfile");
		}
		else
			wprintf(L"The specified profile name is invalid or no profile name was specified.\n");
	}

Error:
	return hRes;
}

HRESULT HrSetCachedModeOneProfile(LPWSTR lpwszProfileName, ProfileInfo* pProfileInfo, int iServiceIndex, BOOL bDefaultService, BOOL bAllServices, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths, int iOutlookVersion)
{
	HRESULT hRes = S_OK;

	for (ULONG i = 0; i <= pProfileInfo->ulServiceCount; i++)
	{
		if (bDefaultService)
		{
			if (pProfileInfo->profileServices[i].bDefaultStore)
			{
				if (pProfileInfo->profileServices[i].serviceType == ServiceType::ServiceType_Mailbox)
				{
					EC_HRES_MSG(HrSetCachedModeOneService(ConvertWideCharToMultiByte(lpwszProfileName), &pProfileInfo->profileServices[i].muidServiceUid, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths, iOutlookVersion), L"Calling HrSetCachedModeOneService on service");
				}
			}
		}
		else if (iServiceIndex != -1)
		{
			if (pProfileInfo->profileServices[iServiceIndex].serviceType == ServiceType::ServiceType_Mailbox)
			{
				EC_HRES_MSG(HrSetCachedModeOneService(ConvertWideCharToMultiByte(lpwszProfileName), &pProfileInfo->profileServices[iServiceIndex].muidServiceUid, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths, iOutlookVersion), L"Calling HrSetCachedModeOneService on service");

			}
		}
		else if (bAllServices)
		{
			if (pProfileInfo->profileServices[i].serviceType == ServiceType::ServiceType_Mailbox)
			{
				EC_HRES_MSG(HrSetCachedModeOneService(ConvertWideCharToMultiByte(lpwszProfileName), &pProfileInfo->profileServices[i].muidServiceUid, bCachedModeOwner, bCachedModeShared, bCachedModePublicFolders, iCachedModeMonths, iOutlookVersion), L"Calling HrSetCachedModeOneService on service");
			}
		}
	}
Error:
	return hRes;
}

HRESULT HrSetCachedModeOneService(LPSTR lpszProfileName, LPMAPIUID lpServiceUid, bool bCachedModeOwner, bool bCachedModeShared, bool bCachedModePublicFolders, int iCachedModeMonths, int iOutlookVersion)
{
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer

	EC_HRES_MSG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin
									 // Get an IProfAdmin interface.

									 // Begin process services
	LPSERVICEADMIN lpServiceAdmin = NULL;
	EC_HRES_MSG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling AdminServices");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		LPPROVIDERADMIN lpProvAdmin = NULL;
		LPPROFSECT lpEmsMdbProfSect, lpStoreProviderProfSect = NULL;
		if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpServiceUid,
			0,
			&lpProvAdmin)))
		{
			EC_HRES_MSG(HrGetSections(lpServiceAdmin, lpServiceUid, &lpEmsMdbProfSect, &lpStoreProviderProfSect), L"Calling HrGetSections");
			// Access the EMSMDB section
			if (lpEmsMdbProfSect)
			{
				LPMAPIPROP pMAPIProp = NULL;
				if (SUCCEEDED(lpEmsMdbProfSect->QueryInterface(IID_IMAPIProp, (void**)& pMAPIProp)))
				{
					// bind to the PR_PROFILE_CONFIG_FLAGS property
					LPSPropValue profileConfigFlags = NULL;
					if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_CONFIG_FLAGS, &profileConfigFlags)))
					{
						if (profileConfigFlags)
						{
							if (bCachedModeOwner)
							{
								if (!(profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PRIVATE))
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PRIVATE;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									wprintf(L"Cached mode owner enabled.\n");
								}
								else
								{
									wprintf(L"Cached mode owner already enabled on service.\n");
								}

							}
							else
							{
								if (profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PRIVATE)
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PRIVATE;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									wprintf(L"Cached mode owner disabled.\n");
								}
								else
								{
									wprintf(L"Cached mode owner already disabled on service.\n");
								}

							}


							if (bCachedModeShared)
							{
								if (!(profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_DELEGATE_PIM))
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_DELEGATE_PIM;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									wprintf(L"Cached mode shared enabled.\n");
								}
								else
								{
									wprintf(L"Cached mode shared already enabled on service.\n");
								}
							}
							else
							{
								if (profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_DELEGATE_PIM)
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_DELEGATE_PIM;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									wprintf(L"Cached mode shared disabled.\n");
								}
								else
								{
									wprintf(L"Cached mode shared already disabled on service.\n");
								}
							}


							if (bCachedModePublicFolders)
							{
								if (!(profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PUBLIC))
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PUBLIC;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									wprintf(L"Cached mode public folders enabled.\n");
								}
								else
								{
									wprintf(L"Cached mode public folders already enabled on service.\n");
								}
							}
							else
							{
								if (profileConfigFlags[0].Value.l & CONFIG_OST_CACHE_PUBLIC)
								{
									profileConfigFlags[0].Value.l = profileConfigFlags[0].Value.l ^ CONFIG_OST_CACHE_PUBLIC;
									EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileConfigFlags, NULL), L"Calling SetProps");
									wprintf(L"Cached mode public folders disabled.\n");
								}
								else
								{
									wprintf(L"Cached mode public folders already disabled on service.\n");
								}
							}

							EC_HRES_MSG(lpEmsMdbProfSect->SaveChanges(0), L"Calling #");
							if (profileConfigFlags) MAPIFreeBuffer(profileConfigFlags);
						}
					}

					// We require split logic for 2010 or older, where all e-mail is cached, vs 2013 and newer. 

					switch (iOutlookVersion)
					{
					case 2013:
					case 2016:
						// bind to the PR_RULE_ACTION_TYPE property for setting the amout of mail to cache
						LPSPropValue profileRuleActionType = NULL;
						if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_RULE_ACTION_TYPE, &profileRuleActionType)))
						{
							if (profileRuleActionType)
							{

								profileRuleActionType[0].Value.i = iCachedModeMonths;
								EC_HRES_MSG(lpEmsMdbProfSect->SetProps(1, profileRuleActionType, NULL), L"Calling SetProps");
								wprintf(L"Cached mode amount to sync set.\n");

								EC_HRES_MSG(lpEmsMdbProfSect->SaveChanges(0), L"Calling SaveChanges");
								if (profileRuleActionType) MAPIFreeBuffer(profileRuleActionType);
							}
						}
						break;
					}
				}
				if (lpEmsMdbProfSect) lpEmsMdbProfSect->Release();
			}

			if (lpProvAdmin) lpProvAdmin->Release();

		}

		if (lpServiceAdmin) lpServiceAdmin->Release();

	}
	// End process services

Error:
	goto Cleanup;
Cleanup:
	// Free up memory
	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;
	return S_OK;
}

