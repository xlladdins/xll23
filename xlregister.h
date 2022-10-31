// xll_register.h - register an add-in
/*
    class AddIn xai_foo(Function(XLL_DOUBLE, "xll_foo", "XLL.FOO"));

	OPER args;
	args = Function(args, XLL_DOUBLE, "xll_foo", "XLL.FOO");

	Register(LPXLOPER12 pxProcedure,
    LPXLOPER12 pxTypeText,     LPXLOPER12 pxFunctionText,
    LPXLOPER12 pxArgumentText, LPXLOPER12 pxMacroType,
    LPXLOPER12 pxCategory,     LPXLOPER12 pxShortcutText,
    LPXLOPER12 pxHelpTopic,    LPXLOPER12 pxFunctionHelp,
    LPXLOPER12 pxArgumentHelp1, LPXLOPER12 pxArgumentHelp2,...)

	Register("xll_foo", XLL_DOUBLE XLL_CSTRING, "XLL.FOO", "arg,...",
		MACRO, "Cat", void, "https://foo.org", "call foo", "first arg");

*/
#pragma once
#include <vector>
#include "oper.h"

namespace xll {

	template<class Procedure, class TypeText, class FunctionText, class ArgumentText,
		class MacroType, class Category, class ShortcutText,
		class HelpTopic, class FunctionHelp, class... ArgumentHelp>

	inline int Register(const Procedure& procedure)
	{
		OPER arg(11 + sizeof...(ArgumentHelp), 1);
		arg[0] = Excel(xlGetName);
		arg[1] = procedure;
		// array of pointers
		std::vector<LPXLOPER> args(args.size());
		for (const auto& a : args) {
			args[i] = &arg[i];
		}
		return 0;
	}

} // namespace xll
