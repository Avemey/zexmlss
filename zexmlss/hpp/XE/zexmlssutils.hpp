// CodeGear C++Builder
// Copyright (c) 1995, 2010 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zexmlssutils.pas' rev: 22.00

#ifndef ZexmlssutilsHPP
#define ZexmlssutilsHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <Windows.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <Grids.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <zesavecommon.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zexmlssutils
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
extern PACKAGE double _MMinInch;
extern PACKAGE bool __fastcall GridToXmlSS(Zexmlss::TZEXMLSS* &XMLSS, const int PageNum, Grids::TStringGrid* &Grid, int ToCol, int ToRow, int BCol, int BRow, int ECol, int ERow, bool ignorebgcolor, System::Byte _border)/* overload */;
extern PACKAGE bool __fastcall XmlSSToGrid(Grids::TStringGrid* &Grid, Zexmlss::TZEXMLSS* &XMLSS, const int PageNum, int ToCol, int ToRow, int BCol, int BRow, int ECol, int ERow, System::Byte InsertMode, int StyleCopy = 0x1ff)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToHtml(Zexmlss::TZEXMLSS* &XMLSS, const int PageNum, System::UnicodeString Title, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToHtml(Zexmlss::TZEXMLSS* &XMLSS, const int PageNum, System::UnicodeString Title, System::UnicodeString FileName, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString FileName, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall ReadEXMLSS(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream)/* overload */;
extern PACKAGE int __fastcall ReadEXMLSS(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString FileName)/* overload */;

}	/* namespace Zexmlssutils */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE)
using namespace Zexmlssutils;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZexmlssutilsHPP
