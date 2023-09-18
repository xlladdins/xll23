// xlauto.cpp - xlAuto* functions
#include "xll.h"
#include "win_mem_view.h"

extern "C" int __declspec(dllexport) xlAutoOpen()
{
	try {
	}
	catch (const std::exception& ex) {
		const char* s;
		s = ex.what();
	}

	// use counted_array_view to register addins
	return TRUE;
}