// xll.h - Excel add-ins
#pragma once
#include <concepts>
#ifndef _INC_WINDOWS
#include <Windows.h>
#endif
#include "XLCALL.H"

#ifndef XLL_VERSION
#define XLL_VERSION 12
#endif

#if XLL_VERSION == 12
using XLOPERX = XLOPER12;
#else
using XLOPERX = XLOPER;
#endif

#include "xoper.h"
//#include "counted_array_view.h"

