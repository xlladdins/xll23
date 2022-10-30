// xlref.h - XLREF functions
#pragma once
#include "xll_traits.h"

namespace xll {

	template<is_ref X>
	inline constexpr auto height(const X& x)
	{
		return x.rwLast - x.rwFirst + 1;
	}
	template<is_ref X>
	inline constexpr auto width(const X& x)
	{
		return x.colLast - x.colFirst + 1;
	}
	template<is_ref X>
	inline constexpr auto size(const X& x)
	{
		return height(x) * width(x);
	}

	// translate by r rows and c columns
	template<is_ref X>
	inline constexpr X move(X x, int r, int c)
	{
		x.rwFirst += r;
		x.rwLast += r;
		x.colFirst += c;
		x.colLast += c;

		return x;
	}

	// Reference to a single range of cells.
	template<is_xloper X>
	struct REF : public type_traits<X>::xref {
		using xrw = type_traits<X>::xrw;
		using xcol = type_traits<X>::xcol;
		// Default to A1:A1
		constexpr REF(xrw r = 0, xcol c = 0, xrw h = 1, xcol w = 1)
			: type_traits<X>::xref{
				.rwFirst = r, .rwLast = static_cast<xrw>(r + h - 1), 
				.colFirst = c, .colLast = static_cast<xcol>(c + w - 1)}
		{ }

		constexpr bool operator==(const REF& r) const = default;

		// set height
		constexpr REF& height(xrw h)
		{
			X::rwLast = X::rwFirst + h - 1;

			return *this;
		}
		// set width
		constexpr REF& width(xcol w)
		{
			X::colLast = X::colFirst + w - 1;

			return *this;
		}
	};

#ifdef _DEBUG
	// hidden namespace
	//namespace {
	//	constexpr REF<XLOPERX> r;
	//	constexpr REF<XLOPERX> r2{ r };
	//	static_assert(r == r2);
	//	//r = r2;
	//	//static_assert(!(r2 != r));

	//	//static_assert(1 == rows(r));
	//	//static_assert(1 == columns(r));
	//	//static_assert(1 == size(r));
	//}
	//{
	//		REF<XLOPERX> r(1,2,3,4);
	//		REF<XLOPERX> r2{ r };
	//		static_assert(r == r2);
	//		r = r2;
	//		static_assert(!(r2 != r));

	//		static_assert(1 == r.rwFirst);
	//		static_assert(2 == r.colFirst);

	//		static_assert(3 == height(r));
	//		static_assert(4 == width(r));
	//		static_assert(12 == size(r));

	//		r = move(r, 5, 6);
	//		static_assert(6 == r.rwFirst);
	//		static_assert(8 == r.colFirst);

	//		static_assert(3 == height(r));
	//		static_assert(4 == width(r));
	//		static_assert(12 == size(r));

	//		r.height(20);
	//		static_assert(2 == height(r));
	//	}
	//}
#endif // _DEBUG

} // namespace xll