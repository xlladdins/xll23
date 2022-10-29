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

		operator bool() const
		{
			switch (type()) {
			case xltypeNum: return val.num;
			case xltypeStr: return val.str[0];
			case xltypeBool: return val.xbool;
			case xltypeInt: return val.w;
			case xltypeErr: return false;
			case xltypeMulti: return val.array.rows > 1 || val.array.columns > 1 || val.
			}
		}

		// Num
		explicit OPER(double num)
			: X{.val.num = num, .xltype = xltypeNum}
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
			: X{.val.xbool = xbool, .xltype = xltypeBool}
		{ }
	};

	template<class X> requires is_xloper<X>
	class Str : OPER<X> {
		// length of null terminated string
		size_t len(const xchar* s)
		{
			size_t n = 0;

			while (*s++) ++n;

			return n;
		}

		// allocate and count
		void allocate(size_t n)
		{
			xltype = xltypeStr;
			val.str = new xchar[n + 1];
			val.str[0] = static_cast<xchar>(n);
		}
		// copy counted string
		void copy(const xchar* str)
		{
			size_t n = val.str[0];
			memcpy_s(val.str + 1, n, s.val.str + 1, n);

		}
		void deallocate()
		{
			delete[] val.str;
		}
	public:
		explicit Str(const xchar* str)
		{
			allocate(len(str));
			copy(str);
		}
		Str(size_t n, const xchar* str)
		{
			allocate(n);
			copy(str);
		}
		template<size_t N>
		constexpr Str(const xchar(&_str)[N])
		{
			allocate(N);
			copy(_str);
		}
		Str(const Str& s)
		{
			allocate(s.val.str[0]);
			copy(s.val.str + 1);
		}
		Str& operator=(const Str& s)
		{
			deallocate();
			allocate(s.val.str[0]);
			copy(s.val.str + 1);

			return *this;
		}
		~Str()
		{
			delete[] val.str;
		}

		bool operator==(const Str& str) const
		{
			return val.str[0] == str.val.str[0] && operator==(str.val.str + 1, str.val.str[0]);

		}
		bool operator==(const xchar* str, xchar len = 0) const
		{
			xchar i = 1;

			while (i <= val.str[0] && *str) {
				if (val.str[i] != *str) {
					return false;
				}
				++i;
				++str;
			}

			return i == val.str[0] && !*str;
		}
	};


#ifdef _DEBUG
	//inline constexpr int i = 1;
#endif // _DEBUG

} // namespace xll
