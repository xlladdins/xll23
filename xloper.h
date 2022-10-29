// xloper.h - XLOPER functions
#pragma once
#include "ref.h"

namespace xll {

	template<class X>
	concept is_xloper
		= std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>;

	template<class X> requires is_xloper<X>
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
			return rows(x.val.sref);
		case xltypeNil:
		case xltypeMissing:
			return 0;
		}

		return 1;
	}

	template<class X> requires is_xloper<X>
	inline auto columns(const X& x)
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

	template<class X> requires is_xloper<X>
	inline auto size(const X& x)
	{
		return rows(x) * columns(x);
	}

	template<class X> requires is_xloper<X>
	inline auto begin(const X& x)
	{
		return type(x) == xltypeMulti ? x.val.array.lparray : &x;
	}
	template<class X> requires is_xloper<X>
	inline auto end(const X& x)
	{
		return type(x) == xltypeMulti ? x.val.array.lparray + size(x) : &x + 1;
	}

	template<class X> requires is_xloper<X>
	inline X Num(double num)
	{
		return X{ .val.num = num, .xltype = xltypeNum };
	}
	template<class X> requires is_xloper<X>
	inline X Bool(bool xbool)
	{
		return X{ .val.xbool = xbool, .xltype = xltypeBool };
	}
	// xlerrX, Excel error string, error description
#define XLL_ERR(X)	                                                        \
	X(Null,  "#NULL!",  "intersection of two ranges that do not intersect") \
	X(Div0,  "#DIV/0!", "formula divides by zero")                          \
	X(Value, "#VALUE!", "variable in formula has wrong type")               \
	X(Ref,   "#REF!",   "formula contains an invalid cell reference")       \
	X(Name,  "#NAME?",  "unrecognised formula name or text")                \
	X(Num,   "#NUM!",   "invalid number")                                   \
	X(NA,    "#N/A",    "value not available to a formula.")                \

#define XLL_ERR_ENUM(a,b,c) a = xlerr##a,
	enum class Err {
		XLL_ERR(XLL_ERR_ENUM)
	};
#undef XLL_ERR_ENUM

	template<class X> requires is_xloper<X>
	inline X Err(Err err)
	{
		return X{ .val.err = err, .xltype = xltypeErr };
	}

} // namespace xll