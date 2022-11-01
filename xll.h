// xll.h - Excel add-ins
#pragma once
#include <concepts>
#ifndef _INC_WINDOWS
#include <Windows.h>
#endif
#include "XLCALL.H"
#include "ensure.h"

#ifndef XLL_VERSION
#define XLL_VERSION 12
#endif

#if XLL_VERSION == 12
using XLREFX = XLREF12;
using XLOPERX = XLOPER12;
using xchar = XCHAR;
#else
using XLREFX = XLREF;
using XLOPERX = XLOPER;
using xchar = CHAR;
#endif

#include "excel.h"
//#include "counted_array_view.h"

