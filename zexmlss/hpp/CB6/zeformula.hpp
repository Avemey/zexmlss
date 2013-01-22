// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zeformula.pas' rev: 6.00

#ifndef zeformulaHPP
#define zeformulaHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zeformula
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
extern PACKAGE int __fastcall ZEGetColByA1(AnsiString AA, bool StartZero = true);
extern PACKAGE AnsiString __fastcall ZEGetA1byCol(int ColNum, bool StartZero = true);

}	/* namespace Zeformula */
using namespace Zeformula;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zeformula
