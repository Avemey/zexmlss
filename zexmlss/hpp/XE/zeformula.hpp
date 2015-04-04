// CodeGear C++Builder
// Copyright (c) 1995, 2010 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zeformula.pas' rev: 22.00

#ifndef ZeformulaHPP
#define ZeformulaHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zeformula
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
static const System::ShortInt ZE_RTA_ODF = 0x1;
static const System::ShortInt ZE_RTA_ODF_PREFIX = 0x2;
static const System::ShortInt ZE_RTA_NO_ABSOLUTE = 0x4;
static const System::ShortInt ZE_RTA_ONLY_ABSOLUTE = 0x8;
static const System::ShortInt ZE_RTA_ODF_NO_BRACKET = 0x10;
static const System::ShortInt ZE_ATR_DEL_PREFIX = 0x1;
extern PACKAGE bool __fastcall ZEGetCellCoords(const System::UnicodeString cell, /* out */ int &column, /* out */ int &row, bool StartZero = true);
extern PACKAGE System::UnicodeString __fastcall ZER1C1ToA1(const System::UnicodeString formula, int CurCol, int CurRow, int options, bool StartZero = true);
extern PACKAGE System::UnicodeString __fastcall ZEA1ToR1C1(const System::UnicodeString formula, int CurCol, int CurRow, int options, bool StartZero = true);
extern PACKAGE int __fastcall ZEGetColByA1(System::UnicodeString AA, bool StartZero = true);
extern PACKAGE System::UnicodeString __fastcall ZEGetA1byCol(int ColNum, bool StartZero = true);

}	/* namespace Zeformula */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE)
using namespace Zeformula;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZeformulaHPP
