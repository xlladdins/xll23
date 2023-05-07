// oper.h - Excel cell or 2-d range of cells
#pragma once
#include <concepts>
#include <scoped_allocator>
#include <type_traits>
#include "utf8.h"
#include "xloper.h"

namespace xll {

	template<class T>
	constexpr std::size_t len(const T* s)
	{
		return std::char_traits<T>::length(s);
	}
	/*
	// null terminated length
	template<class T>
	inline consteval size_t len(const T* t)
	{
		size_t n = 0;

		while (*t++) ++n;

		return n;
	}
	*/
	template<is_xloper X>
	struct XOPER : X {
		using xchar = traits<X>::xchar;
		using charx = traits<X>::charx;
		using X::xltype;
		using X::val;
	private:
		void swap(XOPER& x)
		{
			std::swap(xltype, x.xltype);
			std::swap(val, x.val);
		}
		void alloc_str(size_t len, const xchar* str)
		{
			ensure(len <= traits<X>::str_max);

			xltype = xltypeStr;
			val.str = new xchar[len + 1];
			//new xchar[len + 1];
			val.str[0] = static_cast<xchar>(len);
			// stop if str is null terminated
			for (size_t i = 0; i < len && str[i]; ++i) {
				val.str[i + 1] = str[i];
			}
		}
		void _XOPER()
		{
			if (X::xltype == xltypeStr) {
				delete[] val.str;
			}
			else if (X::xltype & xlbitXLFree) {
				X* this_[1] = { this };
				traits<X>::Excelv(xlFree, 0, 1, (X**)this_);
			}

			xltype = xltypeNil;
		}
	public:
		XOPER()
			: X{ .xltype = xltypeNil }
		{ }
		XOPER(const X& x)
			: X{.xltype = xll::type(x)}
		{
			switch (xltype) {
			case xltypeNum:
				val.num = x.val.num;
				break;
			case xltypeStr:
				alloc_str(x.val.str[0], x.val.str + 1);
				break;
			case xltypeBool:
				val.xbool = x.val.xbool;
				break;
			case xltypeErr:
				val.err = x.val.err;
				break;
			//case xltypeMulti:
			case xltypeSRef:
				val.sref.ref = x.val.sref.ref;
				break;
			case xltypeInt:
				val.w = x.val.w;
				break;
			//case xltypeMulti:
			}
		}
		XOPER(const XOPER& o)
			: XOPER(static_cast<const X&>(o))
		{
		}
		XOPER& operator=(const XOPER& o)
		{
			_XOPER();

			XOPER o_(o);
			swap(o_);

			return *this;
		}
		XOPER(XOPER&& o)
			: XOPER{}
		{
			swap(o);
		}
		XOPER& operator=(XOPER&& o)
		{
			swap(o);

			return *this;
		}
		~XOPER()
		{
			_XOPER();
		}
		constexpr bool operator==(const X& o) const
		{
			return operator==(*this, o);
		}

		size_t size() const
		{
			return xll::size(*this);
		}

		// Num
		explicit XOPER(double num)
			: X{XNum<X>(num)}
		{ }
		bool operator==(double num) const
		{
			return as_num(*this) == num;
		}
		operator double& ()
		{
			return val.num;
		}
		operator const double& () const
		{
			return as_num(*this);
		}

#ifdef _DEBUG
		inline void test_oper_num()
		{
			{
				constexpr XOPER<XLOPER> o(1.23);
				static_assert(o.xltype == xltypeNum);
				static_assert(o.val.num == 1.23);
				static_assert(o == 1.23);
			}
		}
#endif // _DEBUG

		// Str
		XOPER(size_t len, const xchar* str)
		{
			str_alloc(len, str);
		}
		explicit XOPER(const xchar* str)
			: XOPER(len(str), str)
		{ }
		template<size_t N>
		XOPER(const xchar(&str)[N])
			: XOPER(N - 1, str)
		{ }
#if XLL_VERSION == 12
		XOPER(const char* str, int len = -1)
		{
			xltype = xltypeStr;
			size_t wlen = utf8::wcslen(str, len);
			X::val.str = new wchar_t[wlen + 1];
			utf8::mbstowcs(X::val.str + 1, (int)wlen, str, len);
			X::val.str[0] = (wchar_t)wlen - 1;
		}
#endif
		template<is_char T>
		constexpr bool operator==(const T* str)
		{
			if (xll::type(*this) != xltypeStr) {
				return false;
			}

			xchar len = 0;

			for (len = 0; len < val.str[0] && str[len]; ++len) {
				if (val.str[len + 1] != str[len]) {
					return false;
				}
			}

			return val.str[0] == len;
		}
#ifdef _DEBUG
		inline void test_oper_str()
		{
			{
				constexpr XOPER o("foo");
				static_assert(o.xltype == xltypeStr);
				static_assert(o == "foo");
				static_assert(o == L"foo");
			}
			{
				constexpr XOPER o(L"foo");
				static_assert(o.xltype == xltypeStr);
				static_assert(o == L"foo");
				static_assert(o == "foo");
			}
			{
				constexpr XOPER<XLOPER12> o("abc");
				static_assert(o.xltype == xltypeStr);
				static_assert(o.val.str[0] == 3);
				static_assert(o.val.str[3] == 'c');
				static_assert(o.val.str[3] == 0);
			}
		}
#endif // _DEBUG

		// Bool
		explicit XOPER(bool xbool) noexcept
			: X{XBool<X>(xbool)}
		{ }
		bool operator==(bool b) const
		{
			return static_cast<bool>(as_num(*this)) == b;
		}
		constexpr operator bool() const
		{
			return as_bool(*this);
		}

		// Err
		explicit XOPER(enum XlErr err)
			: X{XErr<X>(err)}
		{ }
		bool operator==(const XErr<X>& err)
		{
			return val.err == err.val.err;
		}

		// Multi
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

		// Int
		explicit XOPER(int w)
			: X{ XInt<X>(w) }
		{ }
		bool operator==(int w) const
		{
			return as_num(*this) == w;
		}
	};

	using OPER4 = XOPER<XLOPER>;
	using OPER12 = XOPER<XLOPER12>;
	using OPER = XOPER<XLOPERX>;

#ifdef _DEBUG
	//inline constexpr int i = 1;
#endif // _DEBUG


} // namespace xll
