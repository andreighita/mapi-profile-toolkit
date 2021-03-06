#include "stdafx.h"
#include "Logger.h"

std::wstring Logger::m_szLogFilePath;
std::wofstream Logger::m_ofsLogFile;
bool Logger::m_bIsLogFileOpen;
LoggingMode Logger::m_loggingMode = loggingModeConsole;

void Logger::Initialise(std::wstring wszPath)
{
	Logger::m_szLogFilePath = wszPath;
	Logger::m_ofsLogFile = std::wofstream(Logger::m_szLogFilePath, std::ios_base::app);
	if (!Logger::m_ofsLogFile.is_open()) {
		std::cerr << "Couldn't open 'output.txt'" << std::endl;
		Logger::m_bIsLogFileOpen = false;
	}
	else
		Logger::m_bIsLogFileOpen = true;
}

void Logger::SetLoggingMode(LoggingMode loggingMode)
{
	Logger::m_loggingMode = loggingMode;
}

LogCallStatus Logger::Write(LogLevel llLevel, std::wstring szMessage, LoggingMode loggingMode)
{
	if (loggingMode == loggingModeNone)
	{
		return logCallStatusLoggingDisabled;
	}

	if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
	{
		try
		{
			Logger::Initialise(Logger::m_szLogFilePath);
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered";
			Logger::Write(logLevelError, oss.str(), loggingMode);
			return logCallStatusError;
		}
	}
	try
	{
		time_t  timev = time(0);

		struct tm * now = localtime(&timev);
		std::wstring wszTimeStampNow = std::to_wstring(now->tm_year + 1900) + L"." + std::to_wstring(now->tm_mon + 1) + L"." + std::to_wstring(now->tm_mday) + L" " + std::to_wstring(now->tm_hour) + L":" + std::to_wstring(now->tm_min) + L":" + std::to_wstring(now->tm_sec);
		if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
			Logger::m_ofsLogFile << wszTimeStampNow;
		else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
		{
			std::wcout << wszTimeStampNow;
		}

		switch (llLevel)
		{
		case LogLevel::logLevelInfo:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" INFO ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" INFO ";
			}
			break;
		case LogLevel::logLevelWarning:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" WARNING ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" WARNING ";
			}
			break;
		case LogLevel::logLevelError:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" ERROR ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" ERROR ";
			}
			break;
		case LogLevel::logLevelSuccess:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" SUCCESS ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" SUCCESS ";
			}
			break;
		case LogLevel::logLevelFailed:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" FAILED ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" FAILED ";
			}
			break;
		case LogLevel::logLevelDebug:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" DEBUG ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" DEBUG ";
			}
			break;
		default:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" UNKN ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" UNKN ";
			}
			break;
		}

		if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
			Logger::m_ofsLogFile << szMessage << std::endl;
		else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
		{
			std::wcout << szMessage << std::endl;
		}
		return logCallStatusSuccess;
	}
	catch (int exception)
	{
		std::wostringstream oss; \
			oss << L"Error " << std::dec << exception << L" encountered";
		Logger::Write(logLevelError, oss.str(), loggingMode);
		return logCallStatusError;
	}
	return logCallStatusError;

}

LogCallStatus Logger::Write(LogLevel llLevel, std::wstring szMessage)
{
	LoggingMode loggingMode = Logger::m_loggingMode;

	if (loggingMode == loggingModeNone)
	{
		return logCallStatusLoggingDisabled;
	}

	if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
	{
		try
		{
			Logger::Initialise(Logger::m_szLogFilePath);
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered";
			Logger::Write(logLevelError, oss.str(), loggingMode);
			return logCallStatusError;
		}
	}
	try
	{
		time_t  timev = time(0);

		struct tm * now = localtime(&timev);
		std::wstring wszTimeStampNow = std::to_wstring(now->tm_year + 1900) + L"." + std::to_wstring(now->tm_mon + 1) + L"." + std::to_wstring(now->tm_mday) + L" " + std::to_wstring(now->tm_hour) + L":" + std::to_wstring(now->tm_min) + L":" + std::to_wstring(now->tm_sec);
		if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
			Logger::m_ofsLogFile << wszTimeStampNow;
		else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
		{
			std::wcout << wszTimeStampNow;
		}

		switch (llLevel)
		{
		case LogLevel::logLevelInfo:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" INFO ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" INFO ";
			}
			break;
		case LogLevel::logLevelWarning:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" WARNING ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" WARNING ";
			}
			break;
		case LogLevel::logLevelError:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" ERROR ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" ERROR ";
			}
			break;
		case LogLevel::logLevelSuccess:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" SUCCESS ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" SUCCESS ";
			}
			break;
		case LogLevel::logLevelFailed:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" FAILED ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" FAILED ";
			}
			break;
		case LogLevel::logLevelDebug:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" DEBUG ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" DEBUG ";
			}
			break;
		default:
			if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << L" UNKN ";
			else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
			{
				std::wcout << L" UNKN ";
			}
			break;
		}

		if ((loggingMode == LoggingMode::loggingModeConsoleandFile || loggingMode == LoggingMode::loggingModeFile) && Logger::m_bIsLogFileOpen)
			Logger::m_ofsLogFile << szMessage << std::endl;
		else if ((loggingMode == LoggingMode::loggingModeConsole) || (loggingMode == LoggingMode::loggingModeConsoleandFile))
		{
			std::wcout << szMessage << std::endl;
		}
		return logCallStatusSuccess;
	}
	catch (int exception)
	{
		std::wostringstream oss; \
			oss << L"Error " << std::dec << exception << L" encountered";
		Logger::Write(logLevelError, oss.str(), loggingMode);
		return logCallStatusError;
	}
	return logCallStatusError;

}
