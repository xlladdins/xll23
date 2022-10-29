// ref.h - REF functions
#pragma once
#include <concepts>
#include <Windows.h>
#include "XLCALL.H"

namespace xll {

	template<class X>
	concept is_ref
		= std::is_same_v<XLREF, X> || std::is_same_v<XLREF12, X>;

	template<class X> requires is_ref<X>
	inline auto height(const X& x)
	{
		return x.rwLast - x.rwFirst + 1;
	}
	template<class X> requires is_ref<X>
	inline auto width(const X& x)
	{
		return x.colLast - x.colFirst + 1;
	}
	template<class X> requires is_ref<X>
	inline auto size(const X& x)
	{
		return height(x) * width(x);
	}

	template<class X> requires is_ref<X>
	inline X move(X x, int r, int c)
	{
		x.rwFirst += r;
		x.rwLast += r;
		x.colFIrst += c;
		x.colLasest += c;

		return x;
	}

	template<class X> requires is_ref<X>
	struct REF : public X {
		// Default to A1:A1
		REF(int r = 0, int c = 0, int h = 1, int w = 1)
			: X{.rwFirst = r, .rwLast = r + h - 1, .colFirst = c, .colLast = c + w - 1}
		{ }

		bool operator==(const REF& r) const = default;
	};

#ifdef _DEBUG
	template<class X> requires is_ref<X>
	inline void test_ref()
	{
		{
			REF<X> r;
			REF<X> r2{ r };
			assert(r == r2);
			r = r2;
			assert(!(r2 != r));

			assert(1 == rows(r));
			assert(1 == columns(r));
			assert(1 == size(r));
		}
		{
			REF<X> r(1,2,3,4);
			REF<X> r2{ r };
			assert(r == r2);
			r = r2;
			assert(!(r2 != r));

			assert(1 == r.rwFirst);
			assert(2 == r.colFirst);

			assert(3 == rows(r));
			assert(4 == columns(r));
			assert(12 == size(r));

			r = move(r, 5, 6);
			assert(6 == r.rwFirst);
			assert(8 == r.colFirst);

			assert(3 == rows(r));
			assert(4 == columns(r));
			assert(12 == size(r));
		}
	}
#endif // _DEBUG

} // namespace xll