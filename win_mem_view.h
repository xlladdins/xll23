// xll_win_mem.cpp - in memory data
#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <memoryapi.h>

namespace Win {

	template<class T>
	class mem_view {
		HANDLE h;
	public:
		T* buf;
		DWORD len;

		/// <summary>
		/// Map file of temporary anonymous memory.
		/// </summary>
		/// <param name="h">optional handle to file</param>
		/// <param name="len">maximum size of buffer</param>
		mem_view(HANDLE h_ = INVALID_HANDLE_VALUE, DWORD len = 1 << 20)
			: h(CreateFileMapping(h_, 0, PAGE_READWRITE, 0, len * sizeof(T), nullptr)), buf(nullptr), len(0)
		{
			if (h != NULL) {
				len = 0;
				buf = (T*)MapViewOfFile(h, FILE_MAP_ALL_ACCESS, 0, 0, len * sizeof(T));
			}
		}
		mem_view(const mem_view&) = delete;
		mem_view(mem_view&& mv) noexcept
		{
			*this = std::move(mv);
		}
		mem_view& operator=(const mem_view&) = delete;
		mem_view& operator=(mem_view&& mv) noexcept
		{
			if (this != &mv) {
				h = std::exchange(mv.h, INVALID_HANDLE_VALUE);
				buf = std::exchange(mv.buf, nullptr);
				len = std::exchange(mv.len, 0);
			}

			return *this;
		}
		~mem_view()
		{
			if (buf) UnmapViewOfFile(buf);
			if (h) CloseHandle(h);
		}

		mem_view& reset(DWORD _len = 0)
		{
			len = _len;

			return *this;
		}

		operator T* ()
		{
			return buf;
		}
		operator const T* () const
		{
			return buf;
		}
		T* begin()
		{
			return buf;
		}
		const T* begin() const
		{
			return buf;
		}
		T* end()
		{
			return buf + len;
		}
		const T* end() const
		{
			return buf + len;
		}

		explicit operator bool() const
		{
			return h != NULL && buf != nullptr;
		}

		// Write to buffered memory.
		mem_view& append(const T* s, DWORD n)
		{
			if (n) {
				std::copy(s, s + n, buf + len);
				len += n;
			}

			return *this;
		}
		mem_view& append(const T* b, const T* e)
		{
			return append(b, (DWORD)(e - b));
		}
		mem_view& append(T t)
		{
			DWORD one = 1;

			return append(&t, one);
		}

		T* back(const T& t)
		{
			return buf + len - 1;
		}
		const T* back(const T& t) const
		{
			return buf + len - 1;
		}
	};

} // namespace Win