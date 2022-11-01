// utf8.h - convert utf-8 to counted, allocated string and vice versa
// Loss-free conversion from UTF-16 to MBCS and back
// https://docs.microsoft.com/en-us/archive/msdn-magazine/2016/september/c-unicode-encoding-conversions-with-stl-strings-and-win32-apis
#pragma once
#include <string>
#include <Windows.h>

namespace utf8 {

	// Length of wcs required to convert s. Null terminated if n = -1.
	inline int wcslen(const char* s, int n = -1)
	{
		return MultiByteToWideChar(CP_UTF8, 0, s, n, nullptr, 0);
	}
	inline int mbstowcs(wchar_t* ws, int wn, const char* s, int n = -1)
	{
		return MultiByteToWideChar(CP_UTF8, 0, s, n, ws, wn);
	}
	inline std::wstring mbstowstring(const char* s, int n = -1)
	{
		std::wstring ws;

		size_t wn = wcslen(s, n);
		ws.reserve(wn + 1);
		mbstowcs(ws.data(), (int)wn, s, n);
		ws.resize(wn);

		return ws;
	}

	// Length of mbs required to convert ws. Null terminated if n = -1.
	inline int mbslen(const wchar_t* ws, int wn = -1)
	{
		return WideCharToMultiByte(CP_UTF8, 0, ws, wn, NULL, 0, NULL, NULL);
	}
	inline int wcstombs(char* s, int n, const wchar_t* ws, int wn = -1)
	{
		return WideCharToMultiByte(CP_UTF8, 0, ws, wn, s + 1, n, NULL, NULL);
	}
	inline std::string wcstostring(const wchar_t* ws, int wn = -1)
	{
		std::string s;

		size_t n = mbslen(ws, wn);
		s.reserve(n + 1);
		wcstombs(s.data(), (int)n, ws, wn);
		s.resize(n);

		return s;
	}
}