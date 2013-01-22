// CodeGear C++Builder
// Copyright (c) 1995, 2010 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zesavecommon.pas' rev: 22.00

#ifndef ZesavecommonHPP
#define ZesavecommonHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zesavecommon
{
//-- type declarations -------------------------------------------------------
typedef System::DynamicArray<int> TZESaveIntArray;

typedef System::DynamicArray<System::UnicodeString> TZESaveStrArray;

//-- var, const, procedure ---------------------------------------------------
extern PACKAGE bool __fastcall ZETryStrToBoolean(const System::UnicodeString st, bool valueIfError = false);
extern PACKAGE double __fastcall ZETryStrToFloat(const System::UnicodeString st, double valueIfError = 0.000000E+00);
extern PACKAGE System::UnicodeString __fastcall ZEFloatSeparator(System::UnicodeString st);
extern PACKAGE System::UnicodeString __fastcall IntToStrN(int value, int NullCount);
extern PACKAGE System::UnicodeString __fastcall ZEDateToStr(System::TDateTime ATime);
extern PACKAGE void __fastcall ZEWriteHeaderCommon(Zsspxml::TZsspXMLWriterH* xml, const System::UnicodeString CodePageName, const System::AnsiString BOM);
extern PACKAGE void __fastcall ZESClearArrays(TZESaveIntArray &_pages, TZESaveStrArray &_names);
extern PACKAGE bool __fastcall ZECheckTablesTitle(Zexmlss::TZEXMLSS* &XMLSS, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size, TZESaveIntArray &_pages, TZESaveStrArray &_names, int &retCount);

}	/* namespace Zesavecommon */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE)
using namespace Zesavecommon;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZesavecommonHPP
