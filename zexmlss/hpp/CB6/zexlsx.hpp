// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zexlsx.pas' rev: 6.00

#ifndef zexlsxHPP
#define zexlsxHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <zipusejcl7z.hpp>	// Pascal unit
#include <zearchhelper.hpp>	// Pascal unit
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

class DELPHICLASS TZXLSXDiffBorderItemStyle;
class PASCALIMPLEMENTATION TZXLSXDiffBorderItemStyle : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	bool FUseStyle;
	bool FUseColor;
	Graphics::TColor FColor;
	Zexmlss::TZBorderType FLineStyle;
	Byte FWeight;
	
public:
	__fastcall TZXLSXDiffBorderItemStyle(void);
	void __fastcall Clear(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	__property bool UseStyle = {read=FUseStyle, write=FUseStyle, nodefault};
	__property bool UseColor = {read=FUseColor, write=FUseColor, nodefault};
	__property Graphics::TColor Color = {read=FColor, write=FColor, nodefault};
	__property Zexmlss::TZBorderType LineStyle = {read=FLineStyle, write=FLineStyle, nodefault};
	__property Byte Weight = {read=FWeight, write=FWeight, nodefault};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZXLSXDiffBorderItemStyle(void) { }
	#pragma option pop
	
};


class DELPHICLASS TZXLSXDiffBorder;
class PASCALIMPLEMENTATION TZXLSXDiffBorder : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
public:
	TZXLSXDiffBorderItemStyle* operator[](int Num) { return Border[Num]; }
	
private:
	TZXLSXDiffBorderItemStyle* FBorder[6];
	void __fastcall SetBorder(int Num, const TZXLSXDiffBorderItemStyle* Value);
	TZXLSXDiffBorderItemStyle* __fastcall GetBorder(int Num);
	
public:
	__fastcall virtual TZXLSXDiffBorder(void);
	__fastcall virtual ~TZXLSXDiffBorder(void);
	void __fastcall Clear(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	__property TZXLSXDiffBorderItemStyle* Border[int Num] = {read=GetBorder, write=SetBorder/*, default*/};
};


class DELPHICLASS TZXLSXDiffFormattingItem;
class PASCALIMPLEMENTATION TZXLSXDiffFormattingItem : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	bool FUseFont;
	bool FUseFontColor;
	bool FUseFontStyles;
	Graphics::TColor FFontColor;
	Graphics::TFontStyles FFontStyles;
	bool FUseBorder;
	TZXLSXDiffBorder* FBorders;
	bool FUseFill;
	bool FUseCellPattern;
	Zexmlss::TZCellPattern FCellPattern;
	bool FUseBGColor;
	Graphics::TColor FBGColor;
	bool FUsePatternColor;
	Graphics::TColor FPatternColor;
	
public:
	__fastcall TZXLSXDiffFormattingItem(void);
	__fastcall virtual ~TZXLSXDiffFormattingItem(void);
	void __fastcall Clear(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	__property bool UseFont = {read=FUseFont, write=FUseFont, nodefault};
	__property bool UseFontColor = {read=FUseFontColor, write=FUseFontColor, nodefault};
	__property bool UseFontStyles = {read=FUseFontStyles, write=FUseFontStyles, nodefault};
	__property Graphics::TColor FontColor = {read=FFontColor, write=FFontColor, nodefault};
	__property Graphics::TFontStyles FontStyles = {read=FFontStyles, write=FFontStyles, nodefault};
	__property bool UseBorder = {read=FUseBorder, write=FUseBorder, nodefault};
	__property TZXLSXDiffBorder* Borders = {read=FBorders, write=FBorders};
	__property bool UseFill = {read=FUseFill, write=FUseFill, nodefault};
	__property bool UseCellPattern = {read=FUseCellPattern, write=FUseCellPattern, nodefault};
	__property Zexmlss::TZCellPattern CellPattern = {read=FCellPattern, write=FCellPattern, nodefault};
	__property bool UseBGColor = {read=FUseBGColor, write=FUseBGColor, nodefault};
	__property Graphics::TColor BGColor = {read=FBGColor, write=FBGColor, nodefault};
	__property bool UsePatternColor = {read=FUsePatternColor, write=FUsePatternColor, nodefault};
	__property Graphics::TColor PatternColor = {read=FPatternColor, write=FPatternColor, nodefault};
};


typedef DynamicArray<TZXLSXDiffFormattingItem* >  zexlsx__5;

class DELPHICLASS TZXLSXDiffFormatting;
class PASCALIMPLEMENTATION TZXLSXDiffFormatting : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
public:
	TZXLSXDiffFormattingItem* operator[](int num) { return Items[num]; }
	
private:
	int FCount;
	int FMaxCount;
	DynamicArray<TZXLSXDiffFormattingItem* >  FItems;
	
protected:
	TZXLSXDiffFormattingItem* __fastcall GetItem(int num);
	void __fastcall SetItem(int num, const TZXLSXDiffFormattingItem* Value);
	void __fastcall SetCount(int ACount);
	
public:
	__fastcall TZXLSXDiffFormatting(void);
	__fastcall virtual ~TZXLSXDiffFormatting(void);
	void __fastcall Add(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	void __fastcall Clear(void);
	__property int Count = {read=FCount, nodefault};
	__property TZXLSXDiffFormattingItem* Items[int num] = {read=GetItem, write=SetItem/*, default*/};
};


class DELPHICLASS TZEXLSXReadHelper;
class PASCALIMPLEMENTATION TZEXLSXReadHelper : public System::TObject 
{
	typedef System::TObject inherited;
	
private:
	TZXLSXDiffFormatting* FDiffFormatting;
	
protected:
	void __fastcall SetDiffFormatting(const TZXLSXDiffFormatting* Value);
	
public:
	__fastcall TZEXLSXReadHelper(void);
	__fastcall virtual ~TZEXLSXReadHelper(void);
	__property TZXLSXDiffFormatting* DiffFormatting = {read=FDiffFormatting, write=SetDiffFormatting};
};


//-- var, const, procedure ---------------------------------------------------
extern PACKAGE bool __fastcall ZEXSLXReadTheme(Classes::TStream* &Stream, TIntegerDynArray &ThemaFillsColors, int &ThemaColorCount);
extern PACKAGE bool __fastcall ZEXSLXReadContentTypes(Classes::TStream* &Stream, TZXLSXFileArray &FileArray, int &FilesCount);
extern PACKAGE bool __fastcall ZEXSLXReadSharedStrings(Classes::TStream* &Stream, /* out */ TStringDynArray &StrArray, /* out */ int &StrCount);
extern PACKAGE bool __fastcall ZEXSLXReadSheet(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* &Stream, const AnsiString SheetName, TStringDynArray &StrArray, int StrCount, TZXLSXRelationsArray &Relations, int RelationsCount, TZEXLSXReadHelper* ReadHelper);
extern PACKAGE bool __fastcall ZEXSLXReadStyles(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* &Stream, TIntegerDynArray &ThemaFillsColors, int &ThemaColorCount, TZEXLSXReadHelper* ReadHelper);
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
extern PACKAGE int __fastcall SaveXmlssToXLSXPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToXLSXPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToXLSXPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName)/* overload */;
extern PACKAGE int __fastcall ReadXSLXPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString DirName);
extern PACKAGE int __fastcall SaveXmlssToXLSX(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToXLSX(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToXLSX(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName)/* overload */;
extern PACKAGE int __fastcall ReadXLSX(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName);

}	/* namespace Zexlsx */
using namespace Zexlsx;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zexlsx
