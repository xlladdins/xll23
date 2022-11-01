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

	template<
		class Procedure,
		class TypeText,
		class FunctionText,
		class ArgumentText,
		class MacroType,
		class Category,
		class ShortcutText,
		class HelpTopic,
		class FunctionHelp,
		class... ArgumentHelp>
	inline XLOPERX Register(
		const Procedure& procedure,
		const TypeText& typeText,
		const FunctionText& functionText,
		const ArgumentText& argumentText,
		const MacroType& macroType,
		const Category& category,
		const ShortcutText& shortcutText,
		const HelpTopic& helpTopic,
		const FunctionHelp* functionHelp
		//const ArgumentHelp...& argumentHelp
		)
	{
		OPER arg(11 + sizeof...(ArgumentHelp), 1);
		int i = 0;
		arg[0] = Excel(xlGetName);
		arg[1] = procedure;
		arg[2] = typeText;
		arg[3] = typeText;
		arg[4] = functionText;
		arg[5] = argumentText;
		arg[6] = macroType;
		arg[7] = category;
		arg[8] = shortcutText;
		arg[9] = helpTopic;
		arg[10] = functionHelp;
//				class... ArgumentHelp
				// array of pointers
		int count = 10;
		std::vector<XLOPERX*> args(args.size());
		for (const auto& a : args) {
			args[i] = &arg[i];
		}

		XLOPERX registerId = { .xltype = xltypeNil };
		int ret = traits<XLOPERX>::Excelv(xlfRegister, &registerId, count, (XLOPERX**)&args[0]);
		if (ret != xlretSuccess || registerId.xltype != xltypeNum) {
			OPER xMsg("Failed to register: ");
			//Excel(xlcAlert, xMsg /* & args.FunctionText()*/, OPER(2));
		}

		return registerId;
		return 0;
	}

} // namespace xll
