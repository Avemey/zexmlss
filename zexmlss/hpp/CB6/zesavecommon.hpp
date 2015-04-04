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
#include <Types.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zesavecommon
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
extern PACKAGE double ZE_MMinInch;
#define ZELibraryVersion "0.0.7"
#define ZELibraryFork ""
extern PACKAGE AnsiString __fastcall ZELibraryName();
extern PACKAGE int __fastcall ZENormalizeAngle90(const Zexmlss::TZCellTextRotate value);
extern PACKAGE int __fastcall ZENormalizeAngle180(const Zexmlss::TZCellTextRotate value);
extern PACKAGE AnsiString __fastcall ZEReplaceEntity(const AnsiString st);
extern PACKAGE bool __fastcall ZE_CheckDirExist(AnsiString &DirName);
extern PACKAGE bool __fastcall ZEStrToBoolean(const AnsiString val);
extern PACKAGE bool __fastcall ZETryStrToBoolean(const AnsiString st, bool valueIfError = false);
extern PACKAGE bool __fastcall ZEIsTryStrToFloat(const AnsiString st, /* out */ double &retValue);
extern PACKAGE double __fastcall ZETryStrToFloat(const AnsiString st, /* out */ bool &isOk, double valueIfError = 0.000000E+00)/* overload */;
extern PACKAGE double __fastcall ZETryStrToFloat(const AnsiString st, double valueIfError = 0.000000E+00)/* overload */;
extern PACKAGE AnsiString __fastcall ZEFloatSeparator(AnsiString st);
extern PACKAGE AnsiString __fastcall IntToStrN(int value, int NullCount);
extern PACKAGE AnsiString __fastcall ZEDateToStr(System::TDateTime ATime);
extern PACKAGE void __fastcall ZEWriteHeaderCommon(Zsspxml::TZsspXMLWriter* xml, const AnsiString CodePageName, const AnsiString BOM);
extern PACKAGE void __fastcall ZESClearArrays(TIntegerDynArray &_pages, TStringDynArray &_names);
extern PACKAGE bool __fastcall ZECheckTablesTitle(Zexmlss::TZEXMLSS* &XMLSS, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, /* out */ TIntegerDynArray &_pages, /* out */ TStringDynArray &_names, /* out */ int &retCount);
extern PACKAGE bool __fastcall TryStrToIntPercent(AnsiString s, /* out */ int &Value);

}	/* namespace Zesavecommon */
using namespace Zesavecommon;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zesavecommon
