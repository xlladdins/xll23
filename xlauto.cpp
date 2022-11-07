// xlauto.cpp - xlAuto* functions
#include "xll.h"
#include "win_mem_view.h"

extern "C" int __declspec(dllexport) xlAutoOpen()
{
	try {
		//xll::counted_array_view<wchar_t> cav("Local\\xll", SECTION_MAP_WRITE| SECTION_MAP_READ, 0, 1024);
		//xll::OPER o("abc");
		//cav.append(o.val.str);
		//xll::OPER o2("def");
		//cav.append(o2.val.str);
		//int n;
		//n = cav.size();
		//const wchar_t* p = cav[0];
		//p = cav[1];
	}
	catch (const std::exception& ex) {
		const char* s;
		s = ex.what();
	}

	// use counted_array_view to register addins
	return TRUE;
}