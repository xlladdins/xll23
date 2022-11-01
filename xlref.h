// xlref.h - XLREF functions
#pragma once
#include "xltraits.h"

namespace xll {

	constexpr bool operator==(const XLREF12& r, const XLREF12& s)
	{
		return r.rwFirst == s.rwFirst
			&& r.rwLast == s.rwLast
			&& r.colFirst == s.colFirst
			&& r.colLast == s.colLast;
	}
	constexpr bool operator==(const XLREF& r, const XLREF& s)
	{
		return r.rwFirst == s.rwFirst
			&& r.rwLast == s.rwLast
			&& r.colFirst == s.colFirst
			&& r.colLast == s.colLast;
	}

	inline constexpr INT32 height(const XLREF12& x)
	{
		return x.rwLast - x.rwFirst + 1;
	}
	inline constexpr INT32 width(const XLREF12& x)
	{
		return x.colLast - x.colFirst + 1;
	}
	inline constexpr INT32 area(const XLREF12& x)
	{
		return height(x) * width(x);
	}

	inline constexpr WORD height(const XLREF& x)
	{
		return x.rwLast - x.rwFirst + 1;
	}
	inline constexpr WORD width(const XLREF& x)
	{
		return x.colLast - x.colFirst + 1;
	}
	inline constexpr WORD area(const XLREF& x)
	{
		return height(x) * width(x);
	}

	// translate by r rows and c columns
	inline constexpr XLREF12 move(XLREF12 x, RW r, COL c)
	{
		x.rwFirst += r;
		x.rwLast += r;
		x.colFirst += c;
		x.colLast += c;

		return x;
	}
	inline constexpr XLREF move(XLREF x, WORD r, BYTE c)
	{
		x.rwFirst += r;
		x.rwLast += r;
		x.colFirst += c;
		x.colLast += c;

		return x;
	}

	// Reference to a single range of cells.
	struct REF : XLREF12 {

		// Default to A1:A1
		constexpr REF(RW r = 0, COL c = 0, RW h = 1, COL w = 1)
			: XLREF12{.rwFirst = r, .rwLast = r + h - 1,
				.colFirst = c, .colLast = c + w - 1}
		{ }
		constexpr REF(const XLREF12& r)
			: XLREF12{r}
		{ }

		// set height
		constexpr REF& height(RW h)
		{
			XLREF12::rwLast = XLREF12::rwFirst + h - 1;

			return *this;
		}
		// set width
		constexpr REF& width(COL w)
		{
			XLREF12::colLast = XLREF12::colFirst + w - 1;

			return *this;
		}
	};

	struct REF4 : XLREF {

		// Default to A1:A1
		constexpr REF4(WORD r = 0, BYTE c = 0, WORD h = 1, BYTE w = 1)
			: XLREF{ .rwFirst = r, .rwLast = static_cast<WORD>(r + h - 1),
				.colFirst = c, .colLast = static_cast<BYTE>(c + w - 1) }
		{ }
		constexpr REF4(const XLREF& r)
			: XLREF{ r }
		{ }

		// set height
		constexpr REF4& height(WORD h)
		{
			XLREF::rwLast = XLREF::rwFirst + h - 1;

			return *this;
		}
		// set width
		constexpr REF4& width(BYTE w)
		{
			XLREF::colLast = XLREF::colFirst + w - 1;

			return *this;
		}
	};

#ifdef _DEBUG
	inline void test_xlref()
	{
		{
			constexpr REF r;
			constexpr REF r2(1, 2);
			static_assert(r == r);
			static_assert(1 == height(r));
			static_assert(1 == height(r2));
			static_assert(3 == height(REF(1, 2, 3, 4)));
			static_assert(4 == width(REF(1, 2, 3, 4)));
			static_assert(12 == area(REF(1, 2, 3, 4)));
		}
		{
			static_assert(2 == move(REF{}, 2, 3).rwFirst);
			static_assert(3 == move(REF{}, 2, 3).colFirst);
			constexpr REF r = move(REF{}, 2, 3);
			constexpr REF r2(2, 3);
			static_assert(r == r2);
		}
		{
			constexpr XLREF12 r12{ .rwFirst = 1, .rwLast = 2, .colFirst = 3, .colLast = 4 };
			constexpr REF r{ r12 };
			static_assert(r == r12);
			static_assert(r12 == r);
		}
		{
			constexpr XLREF r4{ .rwFirst = 1, .rwLast = 2, .colFirst = 3, .colLast = 4 };
			constexpr REF4 r{ r4 };
			static_assert(r == r4);
			static_assert(r4 == r);
		}
	}
#endif // _DEBUG

} // namespace xll