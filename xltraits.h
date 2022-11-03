// xltraits.h - Parameterize by old and post Excel 2007 API
// https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
#pragma once
#include <concepts>
#include <type_traits>
#include <Windows.h>
#include "XLCALL.H"

namespace xll {

	// Multi row/column index
	template<class X>
	concept is_index
		= std::is_same_v<WORD, X> || std::is_same_v<INT32, X>;

	template<class X>
	concept is_char
		= std::is_same_v<CHAR, X> || std::is_same_v<XCHAR, X>;

	template<class X>
	concept is_xlref
		= std::is_same_v<XLREF, X> || std::is_same_v<XLREF12, X>;

	template<class X>
	concept XlRef
		= std::is_base_of_v<XLREF, X> || std::is_base_of_v<XLREF12, X>;

	template<class X>
	concept is_xloper
		= std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>;

	template<class X>
	concept XlOper
		= std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>;

	template<class X>
	struct traits {};

	template<> struct traits<XLREF> {
		using type = XLREF;
		using xrw = WORD;
		using xcol = WORD;
	};

	template<> struct traits<XLREF12> {
		using type = XLREF;
		using xrw = INT32;
		using xcol = INT32;
	};

	template<> struct traits<XLOPER> {
		using type = XLOPER;
		using xchar = CHAR;
		using charx = XCHAR;
		static const size_t str_max = 0x100;
		using xref = XLREF;
		// Multi
		using xrw = WORD;
		using xcol = WORD;
		static const size_t rw_max = 0x10000;
		static const size_t col_max = 0x100;
		static int Excelv(int xlfn, LPXLOPER operRes, int count, LPXLOPER opers[])
		{
			return ::Excel4v(xlfn, operRes, count, opers);
		}
	};

	template<> struct traits<XLOPER12> {
		using type = XLOPER12;
		using xchar = XCHAR;
		using charx = CHAR;
		static const size_t str_max = 0x8000;
		using xref = XLREF12;
		// Multi
		using xrw = INT32;
		using xcol = INT32;
		static const size_t rw_max = 0x100000;
		static const size_t col_max = 0x4000;
		static int Excelv(int xlfn, LPXLOPER12 operRes, int count, LPXLOPER12 opers[])
		{
			return ::Excel12v(xlfn, operRes, count, opers);
		}
	};

	template<> struct traits<CHAR> {
		using xloper = XLOPER;
	};
	template<> struct traits<XCHAR> {
		using xloper = XLOPER12;
	};

	template<> struct traits<WORD> {
		using xloper = XLOPER;
	};
	template<> struct traits<INT32> {
		using xloper = XLOPER12;
	};

} // namespace xll
