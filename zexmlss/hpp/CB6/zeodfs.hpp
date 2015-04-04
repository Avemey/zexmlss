// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zeodfs.pas' rev: 6.00

#ifndef zeodfsHPP
#define zeodfsHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <zipusejcl7z.hpp>	// Pascal unit
#include <zearchhelper.hpp>	// Pascal unit
#include <zeZippy.hpp>	// Pascal unit
#include <zesavecommon.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <Types.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zeodfs
{
//-- type declarations -------------------------------------------------------
#pragma pack(push, 4)
struct TZEConditionMap
{
	AnsiString ConditionValue;
	AnsiString ApplyStyleName;
	int ApplyStyleIDX;
	AnsiString ApplyBaseCellAddres;
} ;
#pragma pack(pop)

typedef DynamicArray<TZEConditionMap >  zeodfs__1;

#pragma pack(push, 4)
struct TZEODFStyleProperties
{
	AnsiString name;
	int index;
	AnsiString ParentName;
	bool isHaveParent;
	int ConditionsCount;
	DynamicArray<TZEConditionMap >  Conditions;
} ;
#pragma pack(pop)

typedef DynamicArray<TZEODFStyleProperties >  TZODFStyleArray;

#pragma pack(push, 4)
struct TZEODFCFLine
{
	int CellNum;
	int StyleNumber;
	int Count;
} ;
#pragma pack(pop)

#pragma pack(push, 4)
struct TZEODFCFAreaItem
{
	int RowNum;
	int ColNum;
	int Width;
	int Height;
	int CFStyleNumber;
} ;
#pragma pack(pop)

typedef DynamicArray<TZEODFCFLine >  zeodfs__3;

typedef DynamicArray<int >  zeodfs__4;

typedef DynamicArray<TZEODFCFAreaItem >  zeodfs__5;

class DELPHICLASS TZODFConditionalReadHelper;
class DELPHICLASS TZEODFReadHelper;
class DELPHICLASS TZEODFReadWriteHelperParent;
class PASCALIMPLEMENTATION TZEODFReadWriteHelperParent : public System::TObject 
{
	typedef System::TObject inherited;
	
private:
	Zexmlss::TZEXMLSS* FXMLSS;
	
protected:
	__property Zexmlss::TZEXMLSS* XMLSS = {read=FXMLSS};
	
public:
	__fastcall virtual TZEODFReadWriteHelperParent(Zexmlss::TZEXMLSS* AXMLSS);
public:
	#pragma option push -w-inl
	/* TObject.Destroy */ inline __fastcall virtual ~TZEODFReadWriteHelperParent(void) { }
	#pragma option pop
	
};


typedef DynamicArray<Zexmlss::TZStyle* >  zeodfs__31;

typedef DynamicArray<Zexmlss::TZSheetOptions* >  zeodfs__41;

typedef DynamicArray<AnsiString >  zeodfs__51;

typedef DynamicArray<Zexmlss::TZSheetOptions* >  zeodfs__61;

typedef DynamicArray<AnsiString >  zeodfs__71;

class PASCALIMPLEMENTATION TZEODFReadHelper : public TZEODFReadWriteHelperParent 
{
	typedef TZEODFReadWriteHelperParent inherited;
	
private:
	int FStylesCount;
	DynamicArray<Zexmlss::TZStyle* >  FStyles;
	int FMasterPagesCount;
	DynamicArray<Zexmlss::TZSheetOptions* >  FMasterPages;
	DynamicArray<AnsiString >  FMasterPagesNames;
	int FPageLayoutsCount;
	DynamicArray<Zexmlss::TZSheetOptions* >  FPageLayouts;
	DynamicArray<AnsiString >  FPageLayoutsNames;
	TZODFConditionalReadHelper* FConditionReader;
	Zexmlss::TZStyle* __fastcall GetStyle(int num);
	
public:
	DynamicArray<TZEODFStyleProperties >  StylesProperties;
	__fastcall virtual TZEODFReadHelper(Zexmlss::TZEXMLSS* AXMLSS);
	__fastcall virtual ~TZEODFReadHelper(void);
	void __fastcall ReadAutomaticStyles(Zsspxml::TZsspXMLReader* xml);
	void __fastcall ReadMasterStyles(Zsspxml::TZsspXMLReader* xml);
	void __fastcall AddStyle(void);
	void __fastcall ApplyMasterPageStyle(Zexmlss::TZSheetOptions* SheetOptions, const AnsiString MasterPageName);
	__property int StylesCount = {read=FStylesCount, nodefault};
	__property Zexmlss::TZStyle* Style[int num] = {read=GetStyle};
	__property TZODFConditionalReadHelper* ConditionReader = {read=FConditionReader};
};


class PASCALIMPLEMENTATION TZODFConditionalReadHelper : public System::TObject 
{
	typedef System::TObject inherited;
	
private:
	Zexmlss::TZEXMLSS* FXMLSS;
	int FCountInLine;
	int FMaxCountInLine;
	DynamicArray<TZEODFCFLine >  FCurrentLine;
	DynamicArray<int >  FColumnSCFNumbers;
	int FColumnsCount;
	int FMaxColumnsCount;
	DynamicArray<TZEODFCFAreaItem >  FAreas;
	int FAreasCount;
	int FMaxAreasCount;
	int FLineItemWidth;
	int FLineItemStartCell;
	int FLineItemStyleCFNumber;
	TZEODFReadHelper* FReadHelper;
	
protected:
	void __fastcall AddToLine(const int CellNum, const int AStyleCFNumber, const int ACount);
	bool __fastcall ODFReadGetConditional(const AnsiString ConditionalValue, /* out */ Zexmlss::TZCondition &Condition, /* out */ Zexmlss::TZConditionalOperator &ConditionOperator, /* out */ AnsiString &Value1, /* out */ AnsiString &Value2);
	
public:
	__fastcall TZODFConditionalReadHelper(Zexmlss::TZEXMLSS* XMLSS);
	__fastcall virtual ~TZODFConditionalReadHelper(void);
	void __fastcall CheckCell(int CellNum, int AStyleCFNumber, int RepeatCount = 0x1);
	void __fastcall ApplyConditionStylesToSheet(int SheetNumber, TZODFStyleArray &DefStylesArray, int DefStylesCount, TZODFStyleArray &StylesArray, int StylesCount);
	void __fastcall AddColumnCF(int ColumnNumber, int StyleCFNumber);
	int __fastcall GetColumnCF(int ColumnNumber);
	void __fastcall ApplyBaseCellAddr(const AnsiString BaseCellTxt, const Zexmlss::TZConditionalStyleItem* ACFStyle, int PageNum);
	void __fastcall Clear(void);
	void __fastcall ClearLine(void);
	void __fastcall ProgressLine(int RowNumber, int RepeatCount = 0x1);
	void __fastcall ReadCalcextTag(Zsspxml::TZsspXMLReader* &xml, int SheetNum);
	__property int ColumnsCount = {read=FColumnsCount, nodefault};
	__property int LineItemWidth = {read=FLineItemWidth, write=FLineItemWidth, nodefault};
	__property int LineItemStartCell = {read=FLineItemStartCell, write=FLineItemStartCell, nodefault};
	__property int LineItemStyleID = {read=FLineItemStyleCFNumber, write=FLineItemStyleCFNumber, nodefault};
	__property TZEODFReadHelper* ReadHelper = {read=FReadHelper, write=FReadHelper};
};


typedef DynamicArray<Zexmlss::TZConditionalAreas* >  TODFCFAreas;

#pragma pack(push, 4)
struct TODFCFmatch
{
	int StyleID;
	int StyleCFID;
} ;
#pragma pack(pop)

typedef DynamicArray<TODFCFmatch >  zeodfs__6;

#pragma pack(push, 4)
struct TODFStyleCFID
{
	int Count;
	DynamicArray<TODFCFmatch >  ID;
} ;
#pragma pack(pop)

typedef DynamicArray<TODFStyleCFID >  zeodfs__7;

#pragma pack(push, 4)
struct TODFCFWriterArray
{
	int CountCF;
	DynamicArray<TODFStyleCFID >  StyleCFID;
	DynamicArray<Zexmlss::TZConditionalAreas* >  Areas;
} ;
#pragma pack(pop)

typedef DynamicArray<TODFCFWriterArray >  zeodfs__9;

typedef DynamicArray<int >  zeodfs__01;

class DELPHICLASS TZODFConditionalWriteHelper;
class PASCALIMPLEMENTATION TZODFConditionalWriteHelper : public System::TObject 
{
	typedef System::TObject inherited;
	
private:
	int FPagesCount;
	DynamicArray<int >  FPageIndex;
	DynamicArray<AnsiString >  FPageNames;
	DynamicArray<TODFCFWriterArray >  FPageCF;
	DynamicArray<int >  FFirstCFIdInPage;
	Zexmlss::TZEXMLSS* FXMLSS;
	int FStylesCount;
	int FApplyCFStylesCount;
	int FMaxApplyCFStylesCount;
	DynamicArray<int >  FApplyCFStyles;
	
protected:
	AnsiString __fastcall GetBaseCellAddr(const Zexmlss::TZConditionalStyleItem* StCondition, const AnsiString CurrPageName);
	bool __fastcall AddBetweenCond(const AnsiString ConditName, const AnsiString Value1, const AnsiString Value2, /* out */ AnsiString &retCondition);
	
public:
	__fastcall TZODFConditionalWriteHelper(Zexmlss::TZEXMLSS* ZEXMLSS, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount);
	bool __fastcall TryAddApplyCFStyle(int AStyleIndex, /* out */ int &retCFIndex);
	int __fastcall GetApplyCFStyle(int AStyleIndex);
	void __fastcall WriteCFStyles(Zsspxml::TZsspXMLWriter* xml);
	void __fastcall WriteCalcextCF(Zsspxml::TZsspXMLWriter* xml, int PageIndex);
	AnsiString __fastcall ODSGetOperatorStr(Zexmlss::TZConditionalOperator AOperator);
	int __fastcall GetStyleNum(const int PageIndex, const int Col, const int Row);
	__fastcall virtual ~TZODFConditionalWriteHelper(void);
	__property int StylesCount = {read=FStylesCount, nodefault};
};


typedef DynamicArray<int >  zeodfs__91;

typedef DynamicArray<int >  zeodfs__02;

typedef DynamicArray<int >  zeodfs__12;

typedef DynamicArray<int >  zeodfs__22;

typedef DynamicArray<AnsiString >  zeodfs__32;

class DELPHICLASS TZEODFWriteHelper;
class PASCALIMPLEMENTATION TZEODFWriteHelper : public TZEODFReadWriteHelperParent 
{
	typedef TZEODFReadWriteHelperParent inherited;
	
private:
	TZODFConditionalWriteHelper* FConditionWriter;
	DynamicArray<int >  FUniquePageLayouts;
	int FUniquePageLayoutsCount;
	DynamicArray<int >  FPageLayoutsIndexes;
	DynamicArray<int >  FMasterPagesIndexes;
	int FMasterPagesCount;
	DynamicArray<int >  FMasterPages;
	DynamicArray<AnsiString >  FMasterPagesNames;
	
public:
	__fastcall virtual TZEODFWriteHelper(Zexmlss::TZEXMLSS* AXMLSS, const TIntegerDynArray _pages, const TStringDynArray _names, int PagesCount)/* overload */;
	__fastcall virtual ~TZEODFWriteHelper(void);
	void __fastcall WriteStylesPageLayouts(Zsspxml::TZsspXMLWriter* xml, const TIntegerDynArray _pages);
	void __fastcall WriteStylesMasterPages(Zsspxml::TZsspXMLWriter* xml, const TIntegerDynArray _pages);
	AnsiString __fastcall GetMasterPageName(int PageNum);
	__property TZODFConditionalWriteHelper* ConditionWriter = {read=FConditionWriter};
};


//-- var, const, procedure ---------------------------------------------------
extern PACKAGE int __fastcall ODFCreateStyles(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM, const TZEODFWriteHelper* WriteHelper);
extern PACKAGE int __fastcall ODFCreateSettings(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall ODFCreateContent(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM, const TZEODFWriteHelper* WriteHelper);
extern PACKAGE int __fastcall ODFCreateMeta(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM);
extern PACKAGE int __fastcall SaveXmlssToODFSPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToODFSPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToODFSPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString PathName)/* overload */;
extern PACKAGE int __fastcall ExportXmlssToODFS(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "", bool AllowUnzippedFolder = false, TMetaClass* ZipGenerator = 0x0)/* overload */;
extern PACKAGE bool __fastcall ReadODFContent(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* stream, TZEODFReadHelper* &ReadHelper);
extern PACKAGE bool __fastcall ReadODFSettings(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* stream);
extern PACKAGE int __fastcall ReadODFSPath(Zexmlss::TZEXMLSS* &XMLSS, AnsiString DirName);
extern PACKAGE int __fastcall SaveXmlssToODFS(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, AnsiString CodePageName, AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToODFS(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName, const int * SheetsNumbers, const int SheetsNumbers_Size, const AnsiString * SheetsNames, const int SheetsNames_Size)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToODFS(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName)/* overload */;
extern PACKAGE int __fastcall ReadODFS(Zexmlss::TZEXMLSS* &XMLSS, AnsiString FileName);

}	/* namespace Zeodfs */
using namespace Zeodfs;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zeodfs
