// Borland C++ Builder
// Copyright (c) 1995, 2005 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'Zeodfs.pas' rev: 10.00

#ifndef ZeodfsHPP
#define ZeodfsHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <Sysinit.hpp>	// Pascal unit
#include <Sysutils.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <Types.hpp>	// Pascal unit
#include <Zsspxml.hpp>	// Pascal unit
#include <Zexmlss.hpp>	// Pascal unit
#include <Zesavecommon.hpp>	// Pascal unit
#include <Zezippy.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zeodfs
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
extern PACKAGE int __fastcall ODFCreateStyles(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ODFCreateSettings(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ODFCreateContent(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ODFCreateMeta(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall SaveXmlssToODFSPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName, int const * SheetsNumbers, const int SheetsNumbers_Size, AnsiString const * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall ExportXmlssToODFS(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName, int const * SheetsNumbers, const int SheetsNumbers_Size, AnsiString const * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "", bool AllowUnzippedFolder = false, TMetaClass* ZipGenerator = 0x0)/* overload */;
extern PACKAGE bool __fastcall ReadODFContent(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* stream);
extern PACKAGE bool __fastcall ReadODFSettings(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* stream);
extern PACKAGE int __fastcall ReadODFSPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString DirName);

}	/* namespace Zeodfs */
using namespace Zeodfs;
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// Zeodfs
