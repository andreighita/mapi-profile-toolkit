/*
* © 2015 Microsoft Corporation
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

#pragma once
#include "stdafx.h"
#include "ToolkitObjects.h"

BOOL Is64BitProcess(void);
BOOL IsCorrectBitness();
BOOL ValidateScenario(int argc, _TCHAR* argv[], RuntimeOptions * pRunOpts);
BOOL ParseArgsProfile(int argc, _TCHAR* argv[], ProfileOptions * profileOptions);
BOOL ParseArgsService(int argc, _TCHAR* argv[], ServiceOptions * serviceOptions);
BOOL ParseArgsMailbox(int argc, _TCHAR* argv[], MailboxOptions * mailboxOptions);