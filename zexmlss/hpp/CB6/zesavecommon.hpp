// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zesavecommon.pas' rev: 6.00

#ifndef zesavecommonHPP
#define zesavecommonHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <zsspxml.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zesavecommon
{
//-- type declarations -------------------------------------------------------
typedef DynamicArray<int >  TZESaveIntArray;

typedef DynamicArray<AnsiString >  TZESaveStrArray;

//-- var, const, procedure ---------------------------------------------------
extern PACKAGE bool __fastcall ZETryStrToBoolean(const AnsiString st, bool valueIfError = false);
extern PACKAGE double __fastcall ZETryStrToFloat(const AnsiString st, double valueIfError = 0.000000E+00);
extern PACKAGE AnsiString __fastcall ZEFloatSeparator(AnsiString st);
extern PACKAGE AnsiString __fastcall IntToStrN(int value, int NullCount);
extern PACKAGE AnsiString __fastcall ZEDateToStr(System::TDateTime ATime);
extern PACKAGE void __fastcall ZEWriteHeaderCommon(Zsspxml::TZsspXMLWriter* xml, const AnsiString CodePageName, const AnsiString BOM);
extern PACKAGE void __fastcall ZESClearArrays(TZESaveIntArray &_pages, TZESaveStrArray &_names);
extern PACKAGE bool __fastcall ZECheckTablesTitle(Zexmlss::TZEXMLSS* &XMLSS, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, TZESaveIntArray &_pages, TZESaveStrArray &_names, int &retCount);

}	/* namespace Zesavecommon */
using namespace Zesavecommon;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zesavecommon
