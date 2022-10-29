// xloper.h - XLOPER functions
#pragma once
#pragma once
#include <concepts>
#include <memory>
#include <type_traits>
#include <Windows.h>
#include "XLCALL.H"

namespace xll {

	template<class X>
	concept is_xloper
		= std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>;
	template<class X>
	inline auto type(const X& x)
	{
		return x.xltype & ~(xlbitDLLFree | xlbitXLFree);
	}

	template<class X> requires is_xloper<X>
	inline auto rows(const X& x)
	{
		switch (type(x)) {
		case xltypeMulti: 
			return x.val.array.rows;
		case xltypeSRef:
			return x.val.sref
		case xltypeNil:
		case xltypeMissing:
			return 0;
		}

		return 1;
	}

	template<class X>
	inline auto columns(const X& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.columns;
		case xltypeNil:
		case xltypeMissing:
			return 0;
		}

		return 1;
	}

		template<class X>
	inline auto size(const X& x)
}