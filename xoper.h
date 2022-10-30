// xloper.h
#pragma once
#include <concepts>
#include <memory>
#include <type_traits>
#include "xloper.h"

namespace xll {


	// null terminated length
	template<class T>
	size_t len(const T* t)
	{
		size_t n = 0;

		while (*t++) ++n;

		return n;
	}

	template<is_xloper X>
	class XOPER : public X {
		using X::xltype;
		using X::val;
		using xchar = type_traits<X>::xchar;
		using charx = type_traits<X>::charx;

		constexpr XOPER()
			: X{.xltype = xltypeNil}
		{ }
		XOPER(const XOPER& o)
		{
			switch (o.type()) {
			case xltypeStr:
				*this = XOPER(o.val.str[0], o.val.str + 1);
				break;
			default:
				*this = o;
			}
		}
		XOPER& operator=(const XOPER& o)
		{
			return *this;
		}
		~XOPER()
		{
			if (X::xltype == xltypeStr) {
				free(val.str);
			}

			xltype = xltypeNil;
		}

		constexpr int type() const
		{
			return xltype & ~(xlbitDLLFree | xlbitXLFree);
		}

		constexpr size_t size() const
		{
			return size(*this);
		}

		constexpr operator bool() const
		{
			switch (type()) {
			case xltypeNum: return val.num;
			case xltypeStr: return val.str[0];
			case xltypeBool: return val.xbool;
			case xltypeErr: return false;
			case xltypeMulti: return val.array.rows > 1 || val.array.columns > 1 || val.array.lparray[0];
			case xltypeSRef: return size(val.sref.ref);
			case xltypeInt: return val.w;
			}

			return false;
		}

		// Num
		explicit constexpr XOPER(double num)
			: X{Num<X>(num)}
		{ }

		// Str
		explicit XOPER(size_t len, const xchar* str)
		{
			xltype = xltypeStr;
			val.str = (xchar*)malloc(sizeof(xchar) * (len + 1));
			val.str[0] = const_cast<xchar>(len);
			// stop if str is null terminated
			for (size_t i = 0; i < len && str[i]; ++i) {
				val.str[i + 1] = str[i];
			}
		}
		explicit XOPER(const xchar* str)
			: XOPER(len(str), str)
		{ }
		template<size_t N>
		XOPER(const xchar(&str)[N])
			: XOPER(N, str)
		{ }

		// Bool
		explicit constexpr XOPER(bool xbool)
			: X{Bool<X>(xbool)}
		{ }

		// Err
		explicit constexpr XOPER(enum XlErr err)
			: X{Err<X>(err)}
		{ }

		XOPER& operator[](int i)
		{
			return index(*this, i);
		}
		const XOPER& operator[](int i) const
		{
			return index(*this, i);
		}
		XOPER& operator()(int i, int j)
		{
			return index(*this, i, j);
		}
		const XOPER& operator()(int i, int j) const
		{
			return index(*this, i, j);
		}
	};

	using OPER4 = XOPER<XLOPER>;
	using OPER12 = XOPER<XLOPER12>;
	using OPER = XOPER<XLOPERX>;

#ifdef _DEBUG
	//inline constexpr int i = 1;
#endif // _DEBUG

} // namespace xll
