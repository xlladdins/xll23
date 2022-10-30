// auto.cpp - xlAuto* functions
#include "xll.h"

extern "C" int xlAutoOpen()
{
	// use counted_array_view to register addins
	return TRUE;
}