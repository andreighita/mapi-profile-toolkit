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
private:
	~Logger()
	{
		if (bIsLogFileOpen)
		{
			ofsLogFile.close();
		}
	}

	static std::wofstream ofsLogFile;
	static std::wstring szLogFilePath;
	static bool bIsLogFileOpen;
	LogLevel logLevel;
};

