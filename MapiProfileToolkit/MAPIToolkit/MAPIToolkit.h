#pragma once
#include <Windows.h>
#include <tchar.h>
#include "Toolkit.h"
namespace MAPIToolkit
{
	// Is64BitProcess
// Returns true if 64 bit process or false if 32 bit.
	BOOL Is64BitProcess(void);

	// GetOutlookVersion
	int GetOutlookVersion();

	// IsCorrectBitness
	// Matches the App bitness against Outlook's bitness to avoid MAPI version errors at startup
	// The execution will only continue if the bitness is matched.
	BOOL _cdecl IsCorrectBitness();

	static Toolkit* m_toolkit;

	void Run(int argc, wchar_t* argv[]);
	BOOL ParseParams(int argc, wchar_t* argv[]);
}