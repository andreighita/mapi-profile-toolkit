#pragma once
#include "ServiceWorker.h"
namespace MAPIToolkit
{
	class ProfileWorker
	{
	public:
		std::wstring profileName;			// pn
		std::wstring delegateName;
	public:

		// ACTION_PROFILE_ADD
		void AddProfile();

		// ACTION_PROFILE_CLONE
		void CloneProfile();

		// ACTION_PROFILE_RENAME
		void RenameProfile();

		// ACTION_PROFILE_LIST
		void ListProfile();
		void ListDefaultProfile();

		// ACTION_PROFILE_LISTALL
		void ListAllProfiles();

		// ACTION_PROFILE_REMOVE
		void RemoveProfile();

		// ACTION_PROFILE_REMOVEALL
		void RemoveAllProfiles();

		// ACTION_PROFILE_SETDEFAULT
		void SetDefaultProfile();

		// ACTION_PROFILE_PROMOTEDELEGATES
		void PromoteDelegates();

		// ACTION_PROFILE_PROMOTEONEDELEGATE
		void PromoteOneDelegate();
	};
}
