#pragma once
#pragma once

#include <iostream>
#include <fstream>
#include <string>
#include <time.h>
#include "ToolkitTypeDefs.h"
// Log, version 0.1: a simple logging class
namespace MAPIToolkit
{
	class Logger
	{
	public:
		static void Initialise(std::wstring wszPath);
		static LogCallStatus Write(LogLevel llLevel, std::wstring szMessage, LoggingMode loggingMode);
		static LogCallStatus Write(LogLevel llLevel, std::wstring szMessage);
		static void SetLoggingMode(LoggingMode loggingMode);
	private:
		~Logger();

		static std::wofstream m_ofsLogFile;
		static std::wstring m_szLogFilePath;
		static bool m_bIsLogFileOpen;
		static LoggingMode m_loggingMode;
		LogLevel m_logLevel;

	};
}
