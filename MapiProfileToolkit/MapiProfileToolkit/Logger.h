#pragma once

#include <iostream>
#include <fstream>
#include <string>
#include <time.h>

// Log, version 0.1: a simple logging class
enum LogLevel { logLevelInfo, logLevelWarning, logLevelError, logLevelSuccess, logLevelFailed, logLevelDebug };
enum LogCallStatus { logCallStatusSuccess, logCallStatusError, logCallStatusNoFile, logCallStatusLoggingDisabled };
enum LoggingMode { loggingModeNone, loggingModeConsole, loggingModeFile, loggingModeConsoleandFile };
class Logger
{
public:
	static void Initialise(std::wstring wszPath);
	static LogCallStatus Write(LogLevel llLevel, std::wstring szMessage, LoggingMode loggingMode);
	static LogCallStatus Write(LogLevel llLevel, std::wstring szMessage);
	static void SetLoggingMode(LoggingMode loggingMode);
private:
	~Logger()
	{
		if (m_bIsLogFileOpen)
		{
			m_ofsLogFile.close();
		}
	}
	
	static std::wofstream m_ofsLogFile;
	static std::wstring m_szLogFilePath;
	static bool m_bIsLogFileOpen;
	static LoggingMode m_loggingMode;
	LogLevel m_logLevel;
	
};

