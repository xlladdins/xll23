// counted_array_view.h
#pragma once
#include <memory>
#include <stdexcept>
#include "Win.h"

namespace xll {

	// Use proxy to compute size of object.
	// XLOPER type multi.
	// |XLOPER|rows*columns*sizeof(XLOPER)|
	// size = 1 + rows*columns

	// array of counted strings
	// |n|p0...|p1...
	template<class T>
	class counted_array_view {
		HANDLE h;
		T* pt;
	public:
		// Open with FILE_MAP_ALL_ACCESS to create
		counted_array_view(const char* name, DWORD access = FILE_MAP_READ, DWORD hi = 0, DWORD lo = 1 << 20)
		{
			if (access == FILE_MAP_READ) {
				h = OpenFileMappingA(access, 0, name);
				if (!h) {
					throw std::runtime_error(Win::GetFormatMessage());
				}
			}
			else {
				h = CreateFileMappingA(INVALID_HANDLE_VALUE, 0, access, hi, lo, name);
				if (!h) {
					throw std::runtime_error(Win::GetFormatMessage());
				}
			}

			pt = static_cast<T*>(MapViewOfFile(h, access, hi, lo, lo));
		}
		counted_array_view(const counted_array_view&) = delete;
		counted_array_view& operator=(const counted_array_view&) = delete;
		~counted_array_view()
		{
			UnmapViewOfFile(pt);
		}

		T size() const
		{
			return *pt;
		}

		T* operator[](T i)
		{
			if (i < 0 || i >= size()) {
				return nullptr;
			}

			T* s = pt + 1;
			while (i--) {
				s += *s + 1;
			}

			return s;
		}

		counted_array_view& append(const T* t)
		{
			++*pt;
			memcpy(operator[](*pt), t, (t[0] + 1) * sizeof(T));

			return *this;
		}

		counted_array_view& pop_back(T n = 1)
		{
			*pt -= n;

			return *this;
		}
	};

} // namespace xll

