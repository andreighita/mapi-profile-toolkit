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
//#pragma once
//#include "..\stdafx.h"
//#include <iostream>
//#include <string>
//#include <utility>
//#include <algorithm>  
//
//enum { PROFILEMODE_DEFAULT = 1, PROFILEMODE_SPECIFIC };
//enum { RUNNINGMODE_LIST = 1, RUNNINGMODE_UPDATE, RUNNINGMODE_CREATE, RUNNINGMODE_REMOVE };
//
//struct ABManagerOptions
//{
//	ULONG ulRunningMode; // 1 = List one; 2 = List all; 2 = Update; 3 = Create 
//	ULONG ulProfileMode; // 1 = default; 2 = specific;
//	std::wstring szProfileName;
//	std::wstring szABDisplayName;
//	std::wstring	szConfigFilePath;
//	std::wstring	szABServerName;
//};