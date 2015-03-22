// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zexmlssutils.pas' rev: 6.00

#ifndef zexmlssutilsHPP
#define zexmlssutilsHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <Graphics.hpp>	// Pascal unit
#include <zesavecommon.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <Grids.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <Types.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <Windows.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

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
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName)/* overload */;
extern PACKAGE int __fastcall ReadEXMLSS(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream)/* overload */;
extern PACKAGE int __fastcall ReadEXMLSS(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName)/* overload */;

}	/* namespace Zexmlssutils */
using namespace Zexmlssutils;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zexmlssutils
