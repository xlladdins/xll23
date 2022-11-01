// defines.h - 
#pragma once
#include <map>

// Used to export undecorated function name from a dll.
// Put '#pragma XLLEXPORT' in every add-in function body.
#define XLLEXPORT comment(linker, "/export:" __FUNCDNAME__ "=" __FUNCTION__)

// xlerrX, Excel error string, error message
#define XLL_ERR(X)	                                                        \
	X(Null,  "#NULL!",  "intersection of two ranges that do not intersect") \
	X(Div0,  "#DIV/0!", "formula divides by zero")                          \
	X(Value, "#VALUE!", "variable in formula has wrong type")               \
	X(Ref,   "#REF!",   "formula contains an invalid cell reference")       \
	X(Name,  "#NAME?",  "unrecognised formula name or text")                \
	X(Num,   "#NUM!",   "invalid number")                                   \
	X(NA,    "#N/A",    "value not available to a formula.")                \

namespace xll {

#define XLL_ERR_ENUM(a,b,c) a = xlerr##a,
	enum class XlErr {
		XLL_ERR(XLL_ERR_ENUM)
	};
#undef XLL_ERR_ENUM
#define XLL_ERR_ENUM(a,b,c) {xlerr##a, b},
	inline const std::map<int, const char*> xll_err_str = {
		XLL_ERR(XLL_ERR_ENUM)
	};
#undef XLL_ERR_ENUM
#define XLL_ERR_ENUM(a,b,c) {xlerr##a, c},
	inline const std::map<int, const char*> xll_err_msg = {
		XLL_ERR(XLL_ERR_ENUM)
	};
#undef XLL_ERR_ENUM

} // namespace xll