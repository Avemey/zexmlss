// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zeodfs.pas' rev: 6.00

#ifndef zeodfsHPP
#define zeodfsHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <Graphics.hpp>	// Pascal unit
#include <zesavecommon.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zeodfs
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
extern PACKAGE int __fastcall ODFCreateStyles(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const Zesavecommon::TZESaveIntArray _pages, const Zesavecommon::TZESaveStrArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ODFCreateSettings(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const Zesavecommon::TZESaveIntArray _pages, const Zesavecommon::TZESaveStrArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ODFCreateContent(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const Zesavecommon::TZESaveIntArray _pages, const Zesavecommon::TZESaveStrArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ODFCreateMeta(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall SaveXmlssToODFSPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE bool __fastcall ReadODFContent(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* stream);
extern PACKAGE int __fastcall ReadODFSPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString DirName);

}	/* namespace Zeodfs */
using namespace Zeodfs;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zeodfs
