#pragma once
#include <Windows.h>

namespace Win {

	inline char* GetFormatMessage(DWORD id = GetLastError())
	{
		// not thread safe
		static constexpr DWORD size = 1024;
		static char buf[size];

		FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, 0, id, 0, buf, size, 0);

		return buf;
	}

	// Parameterized handle class
	template<class H, BOOL(WINAPI* D)(H)>
	class Hnd {
		H h;
	public:
		Hnd(H h = INVALID_HANDLE_VALUE)
			: h(h)
		{ }
		Hnd(const Hnd&) = delete;
		Hnd& operator=(const Hnd&) = delete;
		Hnd(Hnd&& h) noexcept
			: h(std::exchange(h.h, INVALID_HANDLE_VALUE))
		{ }
		Hnd& operator=(Hnd&& h_) noexcept
		{
			if (this != &h_) {
				h = std::exchange(h_.h, INVALID_HANDLE_VALUE);
			}

			return *this;
		}
		~Hnd()
		{
			D(h);
		}
		operator H()
		{
			return h;
		}
	};

	using Handle = Hnd<HANDLE, CloseHandle>;
}