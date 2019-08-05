// MAPIToolkitConsole.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <tchar.h>
#include <map>
#pragma comment (lib, "MAPIToolkit.lib")
//#include "C:\Users\anixi\source\repos\mapi-toolkit\MapiProfileToolkit\MAPIToolkit\MAPIToolkit.h"

int wmain(int argc, wchar_t* argv[])
{
    std::cout << "Hello World!\n";
	//MAPIToolkit::Run(argc, argv);
	std::map<std::string, int> mymap = {
					{ "alpha", 10 },
					{ "beta", 20 },
					{ "gamma", 30 } };
	try
	{
		int value1 = mymap.at("alpha");
		int value2 = mymap.at("something");
	}
	catch (const std::exception& e)
	{

	}
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
