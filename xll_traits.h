// xll_traits.h - Parameterize by old and post Excel 2007 API
#pragma once
#include <concepts>
#include <type_traits>
#include <Windows.h>
#include "XLCALL.H"

namespace xll {

	template<class X>
	concept is_ref
		= std::is_same_v<XLREF, X> || std::is_same_v<XLREF12, X>;

	template<class X>
	concept is_xloper
		= std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>;

	template<class X>
	struct type_traits {};

	template<>
	struct type_traits<XLOPER> {
		using xchar = CHAR;
		using charx = XCHAR;
		static const size_t str_max = 0xFF;
		// XLREF
		using xref = XLREF;
		using xrw = WORD;
		using xcol = BYTE;
	};

	template<>
	struct type_traits<XLOPER12> {
		using xchar = XCHAR;
		using charx = CHAR;
		static const size_t str_max = 0x7FFF;
		// XLREF12
		using xref = XLREF12;
		using xrw = RW;
		using xcol = COL;
	};

} // namespace xll
