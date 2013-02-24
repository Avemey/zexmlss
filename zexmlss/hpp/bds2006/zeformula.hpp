// Borland C++ Builder
// Copyright (c) 1995, 2005 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'Zeformula.pas' rev: 10.00

#ifndef ZeformulaHPP
#define ZeformulaHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <Sysinit.hpp>	// Pascal unit
#include <Sysutils.hpp>	// Pascal unit

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
extern PACKAGE bool __fastcall ZEGetCellCoords(const AnsiString cell, int &column, int &row, bool StartZero = true);
extern PACKAGE AnsiString __fastcall ZER1C1ToA1(const AnsiString formula, int CurCol, int CurRow, int options, bool StartZero = true);
extern PACKAGE AnsiString __fastcall ZEA1ToR1C1(const AnsiString formula, int CurCol, int CurRow, int options, bool StartZero = true);
extern PACKAGE int __fastcall ZEGetColByA1(AnsiString AA, bool StartZero = true);
extern PACKAGE AnsiString __fastcall ZEGetA1byCol(int ColNum, bool StartZero = true);

}	/* namespace Zeformula */
using namespace Zeformula;
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// Zeformula
