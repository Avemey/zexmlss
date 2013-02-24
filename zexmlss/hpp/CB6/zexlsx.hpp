// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zexlsx.pas' rev: 6.00

#ifndef zexlsxHPP
#define zexlsxHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <Windows.hpp>	// Pascal unit
#include <zeZippy.hpp>	// Pascal unit
#include <zesavecommon.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <zeformula.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <Types.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zexlsx
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
extern PACKAGE bool __fastcall ZEXSLXReadTheme(Classes::TStream* &Stream, TIntegerDynArray &ThemaFillsColors, int &ThemaColorCount);
extern PACKAGE bool __fastcall ZEXSLXReadContentTypes(Classes::TStream* &Stream, TZXLSXFileArray &FileArray, int &FilesCount);
extern PACKAGE bool __fastcall ZEXSLXReadSharedStrings(Classes::TStream* &Stream, /* out */ TStringDynArray &StrArray, /* out */ int &StrCount);
extern PACKAGE bool __fastcall ZEXSLXReadSheet(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* &Stream, const AnsiString SheetName, TStringDynArray &StrArray, int StrCount, TZXLSXRelationsArray &Relations, int RelationsCount);
extern PACKAGE bool __fastcall ZEXSLXReadStyles(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* &Stream, TIntegerDynArray &ThemaFillsColors, int &ThemaColorCount);
extern PACKAGE bool __fastcall ZEXSLXReadWorkBook(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* &Stream, TZXLSXRelationsArray &Relations, int &RelationsCount);
extern PACKAGE bool __fastcall ZE_XSLXReadRelationships(Classes::TStream* &Stream, TZXLSXRelationsArray &Relations, int &RelationsCount, bool &isWorkSheet, bool needReplaceDelimiter);
extern PACKAGE bool __fastcall ZEXSLXReadComments(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* &Stream);
extern PACKAGE int __fastcall ReadXLSXPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString DirName);
extern PACKAGE int __fastcall ZEXLSXCreateContentTypes(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, int PageCount, int CommentCount, const TIntegerDynArray PagesComments, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ZEXLSXCreateSheet(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, int SheetNum, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM, /* out */ bool &isHaveComments);
extern PACKAGE int __fastcall ZEXLSXCreateWorkBook(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ZEXLSXCreateStyles(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ZEXLSXCreateRelsMain(Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ZEXLSXCreateSharedStrings(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ZEXLSXCreateDocPropsApp(Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ZEXLSXCreateDocPropsCore(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ExportXmlssToXLSX(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "", bool AllowUnzippedFolder = false, TMetaClass* ZipGenerator = 0x0);
extern PACKAGE int __fastcall SaveXmlssToXLSXPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "");
extern PACKAGE int __fastcall ReadXSLXPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString DirName);

}	/* namespace Zexlsx */
using namespace Zexlsx;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zexlsx
