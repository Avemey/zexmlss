// Borland C++ Builder
// Copyright (c) 1995, 2005 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'Zesavecommon.pas' rev: 10.00

#ifndef ZesavecommonHPP
#define ZesavecommonHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <Sysinit.hpp>	// Pascal unit
#include <Sysutils.hpp>	// Pascal unit
#include <Zexmlss.hpp>	// Pascal unit
#include <Zsspxml.hpp>	// Pascal unit

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
extern PACKAGE bool __fastcall ZECheckTablesTitle(Zexmlss::TZEXMLSS* &XMLSS, int const * SheetsNumbers, const int SheetsNumbers_Size, AnsiString const * SheetsNames, const int SheetsNames_Size, TZESaveIntArray &_pages, TZESaveStrArray &_names, int &retCount);

}	/* namespace Zesavecommon */
using namespace Zesavecommon;
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// Zesavecommon
