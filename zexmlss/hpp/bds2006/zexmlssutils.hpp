// Borland C++ Builder
// Copyright (c) 1995, 2005 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'Zexmlssutils.pas' rev: 10.00

#ifndef ZexmlssutilsHPP
#define ZexmlssutilsHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <Sysinit.hpp>	// Pascal unit
#include <Windows.hpp>	// Pascal unit
#include <Sysutils.hpp>	// Pascal unit
#include <Types.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <Grids.hpp>	// Pascal unit
#include <Zexmlss.hpp>	// Pascal unit
#include <Zsspxml.hpp>	// Pascal unit
#include <Zesavecommon.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zexmlssutils
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
extern PACKAGE TStringDynArray __fastcall SplitString(const AnsiString buffer, const char delimeter);
extern PACKAGE bool __fastcall GridToXmlSS(Zexmlss::TZEXMLSS* &XMLSS, const int PageNum, Grids::TStringGrid* &Grid, int ToCol, int ToRow, int BCol, int BRow, int ECol, int ERow, bool ignorebgcolor, Byte _border)/* overload */;
extern PACKAGE bool __fastcall XmlSSToGrid(Grids::TStringGrid* &Grid, Zexmlss::TZEXMLSS* &XMLSS, const int PageNum, int ToCol, int ToRow, int BCol, int BRow, int ECol, int ERow, Byte InsertMode, int StyleCopy = 0x1ff)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToHtml(Zexmlss::TZEXMLSS* &XMLSS, const int PageNum, AnsiString Title, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToHtml(Zexmlss::TZEXMLSS* &XMLSS, const int PageNum, AnsiString Title, AnsiString FileName, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, int const * SheetsNumbers, const int SheetsNumbers_Size, AnsiString const * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName, int const * SheetsNumbers, const int SheetsNumbers_Size, AnsiString const * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall ReadEXMLSS(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream)/* overload */;
extern PACKAGE int __fastcall ReadEXMLSS(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName)/* overload */;

}	/* namespace Zexmlssutils */
using namespace Zexmlssutils;
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// Zexmlssutils
