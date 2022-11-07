// xloper.h - XLOPER functions
// Use native C SDK XLOPER/12 structs
#pragma once
#include <algorithm>
#include <map>
#include "defines.h"
#include "xlref.h"

namespace xll {

	template<XlOper X>
	inline constexpr auto type(const X& x)
	{
		return x.xltype & ~(xlbitDLLFree | xlbitXLFree);
	}

	template<XlOper X>
	constexpr bool as_bool(const X& x)
	{
		switch (type(x)) {
			case xltypeNum: return x.val.num;
			case xltypeStr: return x.val.str[0];
			case xltypeBool: return x.val.xbool;
			case xltypeErr: return false;
			case xltypeMulti: return std::any_of(begin(x), end(x), as_bool);
			case xltypeSRef: return area(x.val.sref.ref);
			case xltypeInt: return x.val.w;
		}

		return false;
	}

	template<XlOper X>
	constexpr double as_num(const X& x)
	{
		switch (type(x)) {
			case xltypeNum: return x.val.num;
			case xltypeBool: return x.val.xbool;
			case xltypeInt: return x.val.w;
		}

		return std::numeric_limits<double>::quiet_NaN();
	}

	template<XlOper X>
	inline constexpr traits<typename X::type>::xrw rows(const X& x)
	{
		switch (type(x)) {
			case xltypeMulti:
				return x.val.array.rows;
			case xltypeSRef:
				return height(x.val.sref.ref);
			case xltypeNil:
			case xltypeMissing:
				return 0;
		}

		return 1;
	}

	template<XlOper X>
	inline constexpr traits<typename X::type>::xcol columns(const X& x)
	{
		switch (type(x)) {
			case xltypeMulti:
				return x.val.array.columns;
			case xltypeSRef:
				return width(x.val.sref.ref);
			case xltypeNil:
			case xltypeMissing:
				return 0;
		}

		return 1;
	}

	template<XlOper X>
	inline constexpr traits<typename X::type>::xrw size(const X& x)
	{
		return rows(x) * columns(x);
	}

	template<XlOper X>
	inline constexpr X* begin(X& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return (X*)x.val.array.lparray;
		default:
			return &x;
		}
	}
	template<XlOper X>
	inline constexpr const X* begin(const X& x)
	{
		switch(type(x)) {
		case xltypeMulti:
			return (const X*)x.val.array.lparray;
		default:
			return &x;
		}
	}
	template<XlOper X>
	inline constexpr X* end(X& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return (X*)x.val.array.lparray + size(x);
		default:
			return &x + 1;
		}
	}
	template<XlOper X>
	inline constexpr const X* end(const X& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return (const X*)x.val.array.lparray + size(x);
		default:
			return &x + 1;
		}
	}

	template<XlOper X, XlOper Y>
	inline constexpr bool operator==(const X& x, const Y& y)
	{
		if (type(x) != type(y)) {
			return false;
		}

		switch (type(x)) {
		case xltypeNum:
			return x.val.num == y.val.num;
		case xltypeStr:
			if (x.val.str[0] != y.val.str[0]) {
				return false;
			}
			for (int i = 1; i <= x.val.str[0]; ++i) {
				if (x.val.str[i] != y.val.str[i]) {
					return false;
				}
			}
			return true;
		case xltypeBool:
			return x.val.xbool == y.val.xbool;
		case xltypeErr:
			return x.val.err == y.val.err;
		case xltypeMulti:
			if (rows(x) != rows(y) || columns(x) != columns(y)) {
				return false;
			}
			return std::equal(begin(x), end(x), begin(y), end(y));
		case xltypeSRef:
			return x.val.sref.ref == y.val.sref.ref;
		case xltypeInt:
			return x.val.w == y.val.w;
		}

		return true; // Missing, Nil
	}

	// 1-d index
	template<XlOper X>
	constexpr X& index(X& x, int i)
	{
		return x.val.array[i];
	}
	template<XlOper X>
	constexpr const X& index(X& x, int i)
	{
		return x.val.array[i];
	}

	// 2-d index
	template<XlOper X>
	constexpr X& index(X& x, int i, int j)
	{
		return x.val.array[i * columns(x) + j];
	}
	template<XlOper X>
	constexpr const X& index(X& x, int i, int j)
	{
		return x.val.array[i*columns(x) + j];
	}

	// xltypeNum = 1
	template<is_xloper X>
	struct XNum : X {
		using type = X;
		explicit constexpr XNum(double num = 0)
			: X{ .val = {.num = num}, .xltype = xltypeNum }
		{ }
		constexpr operator double&()
		{
			return X::val.num;
		}
		constexpr operator const double& () const
		{
			return X::val.num;
		}

		constexpr XNum& operator++()
		{
			++X::val.num;

			return *this;
		}
		XNum operator++(int)
		{
			auto num(*this);

			++X::val.num;

			return num;
		}
		constexpr XNum& operator--()
		{
			--X::val.num;

			return *this;
		}
		XNum operator--(int)
		{
			auto num(*this);

			--X::val.num;

			return num;
		}
	};
	using Num4 = XNum<XLOPER>;
	using Num12 = XNum<XLOPER12>;
	using Num = XNum<XLOPERX>;

#ifdef _DEBUG
	inline void test_xloper_num()
	{
		{
			constexpr Num zero;
			static_assert(xltypeNum == type(zero));
			static_assert(0 == zero);
			static_assert(1.23 == Num(1.23));
		}
		{
			constexpr auto num = Num(1);
			static_assert(1 == num);
			static_assert(2 == 1 + num);
			static_assert(2 == ++Num(1));
			static_assert(1 == --Num(2));
			static_assert(2 == (Num(1) += 1));
			static_assert(1 == (Num(2) -= 1));
		}
	}
#endif // _DEBUG

	// xltypeStr = 2
	template<size_t N, is_char T, class X = typename traits<T>::xloper>
	struct Str : X {
		using type = X;
		T str[N]{};
		constexpr Str(const T(&_str)[N])
			: X{ .val = { .str = str }, .xltype = xltypeStr}
		{
			static_assert(N <= traits<X>::str_max);

			str[0] = static_cast<T>(N - 1);
			for (size_t i = 0; i < N - 1; ++i) {
				str[i + 1] = _str[i];
			}
		}
	};

#ifdef _DEBUG
	inline void test_xloper_str()
	{
		{
			constexpr Str s("abc");
			static_assert(3 == s.val.str[0]);
			static_assert('c' == s.val.str[3]);
		}
		{
			constexpr auto s = Str(L"abc");
			static_assert(3 == s.val.str[0]);
			static_assert('c' == s.val.str[3]);
		}
		{
			constexpr auto s = Str("abc");
			static_assert(sizeof(s) == sizeof(XLOPER) + 8);
		}
		{
			constexpr auto s = Str(L"abc");
			static_assert(sizeof(s) == sizeof(XLOPER12) + 8);
		}
	}
#endif // _DEBUG

	// xltypeBool = 4
	template<is_xloper X>
	struct XBool : X {
		using type = X;
		explicit constexpr XBool(bool xbool = false)
			: X{ .val = {.xbool = xbool}, .xltype = xltypeBool }
		{ }

		constexpr operator bool() const
		{
			return  X::val.xbool != 0;
		}
	};
	using Bool4 = XBool<XLOPER>;
	using Bool12 = XBool<XLOPER12>;
	using Bool = XBool<XLOPERX>;

#ifdef _DEBUG
	inline void test_xloper_bool()
	{
		{
			constexpr Bool s;
			static_assert(!s);
		}
		{
			constexpr Bool s(true);
			static_assert(s);
		}
	}
#endif // _DEBUG

	// xltypeErr = 0x10
	template<is_xloper X>
	struct XErr : X {
		using type = X;
		constexpr XErr(XlErr err)
			: X{ .val = {.err = (int)err}, .xltype = xltypeErr }
		{ }
	};
	using Err4 = XErr<XLOPER>;
	using Err12 = XErr<XLOPER12>;
	using Err = XErr<XLOPERX>;
	inline constexpr Err ErrNA(XlErr::NA);

#ifdef _DEBUG
	inline void test_xloper_err()
	{
		{
			constexpr Err s(XlErr::NA);
			static_assert(s == ErrNA);
		}
	}
#endif // _DEBUG

	// xltypeMulti = 0x40
	template<is_index R, is_index C, class X = typename traits<R>::xloper>
	struct Multi : X {
		using type = X;

		X arr[R * C];
		constexpr Multi()
			: X{ .val = {.array = {.rows = R, .columns = C}, .xltpe = xltypeMulti } }
		{ 
			static_assert(R <= traits<X>::rw_max);
			static_assert(C <= traits<X>::col_max);

			type::val.array.lparray = arr;
			for (auto& a : arr) {
				a.xltype = xltypeNil;
			}
		}
	};
	//template<WORD R, WORD C>
	//using Multi4 = Multi<R, C, XLOPER>;
/*
template<int R = 1, int C = 1>
	struct Multi12 : XLOPER12 {
		using type = XLOPER12;

		XLOPER12 arr[R * C];
		constexpr Multi12()
			: type{ .val = {.array = {.rows = R, .columns = C}}, .xltype = xltypeMulti}
		{
			static_assert(R <= traits<XLOPER12>::rw_max);
			static_assert(C <= traits<XLOPER12>::col_max);

			type::val.array.lparray = arr;
			for (auto& a : arr) {
				a.xltype = xltypeNil;
			}
		}
	};
	template<unsigned int R = 1, unsigned int C = 1>
	struct Multi4 : XLOPER {
		using type = XLOPER;

		XLOPER arr[R * C] = { 0 };
		constexpr Multi4()
			: type{ .val = {.array = {.lparray = arr, .rows = R, .columns = C }}, .xltype = xltypeMulti }
		{
			static_assert(R <= traits<XLOPER>::rw_max);
			static_assert(C <= traits<XLOPER>::col_max);

			for (auto& a : arr) {
				a.xltype = xltypeNil;
			}
		}
	};
	*/
#ifdef _DEBUG
	inline void test_xloper_multi()
	{
		{
//			constexpr Multi<1, 2> m;
//			static_assert(1 == rows(m));
		}
		/*
		{
			constexpr Multi<1u, 2u> m;
			static_assert(2 == columns(m));
		}
		{
			constexpr Multi<2u, 3u> m;
			static_assert(6 == size(m));
		}
		*/
	}
#endif // _DEBUG

	template<is_xloper X>
	class XMissing : X { 
		using type = X;
		constexpr XMissing()
			: X{ .xltype = xltypeMissing } 
		{ } 
	};
	using Missing4 = XMissing<XLOPER>;
	using Missing12 = XMissing<XLOPER12>;
	using Missing = XMissing<XLOPERX>;

	template<is_xloper X>
	class XNil : X { 
		using type = X;
		constexpr XNil()
			: X{ .xltype = xltypeNil } 
		{ } 
	};
	using Nil4 = XNil<XLOPER>;
	using Nil12 = XNil<XLOPER12>;
	using Nil = XNil<XLOPERX>;

	// sref

	// xltypeInt = 0x0800
	template<is_xloper X>
	struct XInt : X {
		using type = X;
		explicit constexpr XInt(int w)
			: X{ .val = {.w = w}, .xltype = xltypeInt }
		{ }

		constexpr operator int() const
		{
			return  X::val.w;
		}
	};
	using Int4 = XInt<XLOPER>;
	using Int12 = XInt<XLOPER12>;
	using Int = XInt<XLOPERX>;


} // namespace xll