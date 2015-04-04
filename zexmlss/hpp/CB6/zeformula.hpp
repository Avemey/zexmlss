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
static const Shortint ZE_RTA_ODF = 0x1;
static const Shortint ZE_RTA_ODF_PREFIX = 0x2;
static const Shortint ZE_RTA_NO_ABSOLUTE = 0x4;
static const Shortint ZE_RTA_ONLY_ABSOLUTE = 0x8;
static const Shortint ZE_RTA_ODF_NO_BRACKET = 0x10;
static const Shortint ZE_ATR_DEL_PREFIX = 0x1;
extern PACKAGE bool __fastcall ZEGetCellCoords(const AnsiString cell, /* out */ int &column, /* out */ int &row, bool StartZero = true);
extern PACKAGE AnsiString __fastcall ZER1C1ToA1(const AnsiString formula, int CurCol, int CurRow, int options, bool StartZero = true);
extern PACKAGE AnsiString __fastcall ZEA1ToR1C1(const AnsiString formula, int CurCol, int CurRow, int options, bool StartZero = true);
extern PACKAGE int __fastcall ZEGetColByA1(AnsiString AA, bool StartZero = true);
extern PACKAGE AnsiString __fastcall ZEGetA1byCol(int ColNum, bool StartZero = true);

}	/* namespace Zeformula */
using namespace Zeformula;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zeformula
