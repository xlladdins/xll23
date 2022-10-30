// xll_register.h - register an add-in
#pragma once
#include <vector>
#include "xoper.h"

namespace xll {

	template<is_xloper X>
	inline int Register(const X& xs)
	{
		// array of pointers
		std::vector<X*> args(size(xs));
		for (const auto x : xs) {
			args[i] = &index(xs, i);
		}
		return 0;
	}

} // namespace xll
