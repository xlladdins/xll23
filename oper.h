// xloper.h
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
	struct type_traits {};

	template<>
	struct type_traits<XLOPER> {
		using xchar = CHAR;
		using charx = XCHAR;
		static const size_t str_max = 0xFF;
	};
	
	template<>
	struct type_traits<XLOPER12> {
		using xchar = XCHAR;
		using charx = CHAR;
		static const size_t str_max = 0x7FFF;
	};

	// null terminated length
	template<class T>
	size_t len(const T* t)
	{
		size_t n = 0;

		while (*t++) ++n;

		return n;
	}

	template<class X> requires is_xloper<X>
	class OPER : public X {
		using xchar = type_traits<X>::xchar;
		using charx = type_traits<X>::charx;

		OPER()
			: X{.xltype = xltypeNil}
		{ }
		OPER(const OPER& o)
		{
			switch (o.type()) {
			case xltypeStr:

			case xltypeNum:
				*this = o;
			}
		}
		~OPER()
		{
			if (xltype == xltypeStr) {
				free(val.str);
			}

			xltype = xltypeNil;
		}

		int type() const
		{
			return xltype & ~(xlbitDLLFree | xlbitXLFree);
		}

		size_t size() const
		{
			return size(*this);
		}

		operator bool() const
		{
			switch (type()) {
			case xltypeNum: return val.num;
			case xltypeStr: return val.str[0];
			case xltypeBool: return val.xbool;
			case xltypeSRef: return size(val.sref.ref);
			case xltypeErr: return false;
			case xltypeInt: return val.w;
			case xltypeMulti: return val.array.rows > 1 || val.array.columns > 1 || val.
			}

			return false;
		}

		// Num
		explicit OPER(double num)
			: X{Num(num)}
		{ }

		// Str
		explicit OPER(size_t len, const xchar* str)
		{
			xltype = xltypeStr;
			val.str = (xchar*)malloc(sizeof(xchar) * (len + 1)];
			val.str[0] = const_cast<xchar>(n);
			memcpy_s(val.str + 1, n, str, n);
		}
		explicit OPER(const xchar* str)
			: OPER(len(str), str)
		{ }
		template<size_t N>
		Str(const xchar(&str)[N])
			: OPER(N, str)
		{ }

		// Bool
		explicit OPER(bool xbool)
			: X{Bool(xbool)}
		{ }
	};

#ifdef _DEBUG
	//inline constexpr int i = 1;
#endif // _DEBUG

} // namespace xll
