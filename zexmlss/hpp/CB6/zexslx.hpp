// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zexslx.pas' rev: 6.00

#ifndef zexslxHPP
#define zexslxHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <Windows.hpp>	// Pascal unit
#include <zeformula.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <zesavecommon.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zexslx
{
//-- type declarations -------------------------------------------------------
#pragma pack(push, 4)
struct TZXLSXFileItem
{
	AnsiString name;
	AnsiString original;
	int ftype;
} ;
#pragma pack(pop)

typedef DynamicArray<TZXLSXFileItem >  TZXLSXFileArray;

#pragma pack(push, 4)
struct TZXLSXRelations
{
	AnsiString id;
	int ftype;
	AnsiString target;
	int fileid;
	AnsiString name;
	Byte state;
	int sheetid;
} ;
#pragma pack(pop)

typedef DynamicArray<TZXLSXRelations >  TZXLSXRelationsArray;

//-- var, const, procedure ---------------------------------------------------
extern PACKAGE bool __fastcall ZEXSLXReadContentTypes(Classes::TStream* &Stream, TZXLSXFileArray &FileArray, int &FilesCount);
extern PACKAGE bool __fastcall ZEXSLXReadSharedStrings(Classes::TStream* &Stream, Zesavecommon::TZESaveStrArray &StrArray, int &StrCount);
extern PACKAGE bool __fastcall ZEXSLXReadSheet(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* &Stream, const AnsiString SheetName, Zesavecommon::TZESaveStrArray &StrArray, int StrCount);
extern PACKAGE bool __fastcall ZEXSLXReadStyles(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* &Stream);
extern PACKAGE bool __fastcall ZEXSLXReadWorkBook(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* &Stream, TZXLSXRelationsArray &Relations, int &RelationsCount);
extern PACKAGE bool __fastcall ZE_XSLXReadRelationships(Classes::TStream* &Stream, TZXLSXRelationsArray &Relations, int &RelationsCount, bool &isWorkSheet, bool needReplaceDelimiter);
extern PACKAGE int __fastcall ReadXSLXPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString DirName);
extern PACKAGE int __fastcall ZEXLSXCreateContentTypes(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, int PageCount, int CommentCount, const Zesavecommon::TZESaveIntArray PagesComments, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ZEXLSXCreateSheet(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, int SheetNum, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM, bool &isHaveComments);
extern PACKAGE int __fastcall ZEXLSXCreateWorkBook(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const Zesavecommon::TZESaveIntArray _pages, const Zesavecommon::TZESaveStrArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ZEXLSXCreateStyles(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall SaveXmlssToXLSXPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "");

}	/* namespace Zexslx */
using namespace Zexslx;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zexslx
