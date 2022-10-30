// xloper.h - XLOPER functions
// Use native C SDK XLOPER/12 structs and compile time string literals for xltypeStr with fixed length
#pragma once
#include <map>
#include "xlref.h"

namespace xll {


	template<is_xloper X>
	inline constexpr auto type(const X& x)
	{
		return x.xltype & ~(xlbitDLLFree | xlbitXLFree);
	}

	template<is_xloper X>
	inline constexpr auto rows(const X& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.rows;
		case xltypeSRef:
			return rows(x.val.sref);
		case xltypeNil:
		case xltypeMissing:
			return 0;
		}

		return 1;
	}

	template<is_xloper X>
	inline constexpr auto columns(const X& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.columns;
		case xltypeSRef:
			return columns(x.val.sref);
		case xltypeNil:
		case xltypeMissing:
			return 0;
		}

		return 1;
	}

	template<is_xloper X>
	inline constexpr auto size(const X& x)
	{
		return type(x) == xltypeStr ? x.val.str[0] : rows(x) * columns(x);
	}

	template<is_xloper X>
	inline constexpr auto begin(const X& x)
	{
		switch(type(x)) {
		case xltypeMulti:
			return x.val.array.lparray;
		case xltypeStr:
			return x.val.str + 1;
		default:
			return &x;
		}
	}
	template<is_xloper X>
	inline constexpr auto end(const X& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.lparray + size(x);
		case xltypeStr:
			return x.val.str + x.val.str[0] + 1;
		default:
			return &x + 1;
		}
	}

	template<is_xloper X>
	X& index(X& x, int i)
	{
		return x.val.array[i];
	}
	template<is_xloper X>
	const X& index(X& x, int i)
	{
		return x.val.array[i];
	}
	template<is_xloper X>
	X& index(X& x, int i, int j)
	{
		return x.val.array[i * columns(x) + j];
	}
	template<is_xloper X>
	const X& index(X& x, int i, int j)
	{
		return x.val.array[i*columns(x) + j];
	}

	// xltypeNum = 1
	template<is_xloper X>
	struct Num : X {
		constexpr Num(double num)
			: X{ .val = {.num = num}, .xltype = xltypeNum }
		{ }
		operator double&()
		{
			return X::val.num;
		}
	};

	// counted Pascal string
	template<size_t N>
	struct PStr {
		char str[N]{};
		constexpr PStr(const char(&_str)[N])
		{
			str[0] = static_cast<char>(N - 1);
			for (size_t i = 0; i < N - 1; ++i) {
				str[i + 1] = _str[i];
			}
		}
	};
	template<PStr str>
	constexpr auto operator"" _pstr()
	{
		return str.str;
	}
	// xltypeStr = 2
	template<size_t N>
	struct Str : public XLOPER {
		PStr<N> str;
		constexpr Str(const char(&_str)[N])
			: XLOPER{ .xltype = xltypeStr }, str{ PStr(_str) }
		{
			XLOPER::val.str = str.str;
		}
	};

	// counted Pascal string
	template<size_t N>
	struct WPStr {
		wchar_t str[N]{};
		constexpr WPStr(const wchar_t(&_str)[N])
		{
			str[0] = static_cast<wchar_t>(N - 1);
			for (size_t i = 0; i < N - 1; ++i) {
				str[i + 1] = _str[i];
			}
		}
	};
	template<WPStr str>
	constexpr auto operator"" _wpstr()
	{
		return str.str;
	}
	// xltypeStr = 2
	template<size_t N>
	struct Str12 : public XLOPER12 {
		WPStr<N> str;
		constexpr Str12(const wchar_t(&_str)[N])
			: XLOPER12{ .xltype = xltypeStr }, str{ WPStr(_str) }
		{
			XLOPER12::val.str = str.str;
		}
	};

	// xltypeBool = 4
	template<is_xloper X>
	inline constexpr X Bool(bool xbool)
	{
		return X{ .val = {.xbool = xbool}, .xltype = xltypeBool };
	}

	// xlerrX, Excel error string, error message
#define XLL_ERR(X)	                                                        \
	X(Null,  "#NULL!",  "intersection of two ranges that do not intersect") \
	X(Div0,  "#DIV/0!", "formula divides by zero")                          \
	X(Value, "#VALUE!", "variable in formula has wrong type")               \
	X(Ref,   "#REF!",   "formula contains an invalid cell reference")       \
	X(Name,  "#NAME?",  "unrecognised formula name or text")                \
	X(Num,   "#NUM!",   "invalid number")                                   \
	X(NA,    "#N/A",    "value not available to a formula.")                \

#define XLL_ERR_ENUM(a,b,c) a = xlerr##a,
	enum class XlErr {
		XLL_ERR(XLL_ERR_ENUM)
	};
#undef XLL_ERR_ENUM
#define XLL_ERR_ENUM(a,b,c) {xlerr##a, b},
	inline const std::map<int, const char*> xll_err_str = {
		XLL_ERR(XLL_ERR_ENUM)
	};
#undef XLL_ERR_ENUM
#define XLL_ERR_ENUM(a,b,c) {xlerr##a, c},
	inline const std::map<int, const char*> xll_err_msg = {
		XLL_ERR(XLL_ERR_ENUM)
	};
#undef XLL_ERR_ENUM

	// xltypeErr = 16
	template<is_xloper X>
	inline constexpr X Err(XlErr err)
	{
		return X{ .val = {.err = err}, .xltype = xltypeErr };
	}

#ifdef _DEBUG
	inline void test_xloper()
	{
		{
			constexpr auto s = "abc"_pstr;
			static_assert(3 == s[0]);
			static_assert('c' == s[3]);
		}
		{
			constexpr auto s = L"abc"_wpstr;
			static_assert(3 == s[0]);
			static_assert('c' == s[3]);
		}
		{
			constexpr auto s = Str("abc");
			static_assert(3 == s.val.str[0]);
			static_assert('c' == s.val.str[3]);
		}
		{
			constexpr auto s = Str12(L"abc");
			static_assert(3 == s.val.str[0]);
			static_assert('c' == s.val.str[3]);
		}
	}
#endif // _DEBUG


} // namespace xll