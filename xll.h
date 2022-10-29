// xll.h - Excel add-ins
#pragma once
#include <memory>
#include <winbase.h>

namespace xll {

	// array of counted strings
	// |n|p0...|p1...
	template<class T>
	class counted_array_view {
		HANDLE h;
		T* pt;
	public:
		counted_array_view(const char* name, DWORD hi = 0, DWORD lo = 1<<20)
			: h(CreateFileMappingA(INVALID_HANDLE_VALUE, 0, PAGE_READWRITE, hi, lo, name))
		{
			pt = MapViewOfFile(h, FILE_MAP_ALL_ACCESS, hi, lo, lo);
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
			if (i >= size()) {
				return nullptr;
			}

			T* s = pt + 1;
			while (i--) {
				s += *s + 1;
			}

			return s;
		}

		T* append(const T* t)
		{
			++*pt;
			memcpy(operator[](*pt), t, (t[0] + 1)*sizeof(T));
		}

		counted_array_view& pop_back(T n = 1)
		{
			*pt -= n;

			return *this;
		}
	};

} // namespace xll
