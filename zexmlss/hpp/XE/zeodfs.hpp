// CodeGear C++Builder
// Copyright (c) 1995, 2010 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zeodfs.pas' rev: 22.00

#ifndef ZeodfsHPP
#define ZeodfsHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <Types.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <zesavecommon.hpp>	// Pascal unit
#include <zeZippy.hpp>	// Pascal unit
#include <zearchhelper.hpp>	// Pascal unit
#include <zipuseab.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zeodfs
{
//-- type declarations -------------------------------------------------------
struct DECLSPEC_DRECORD TZEConditionMap
{
	
public:
	System::UnicodeString ConditionValue;
	System::UnicodeString ApplyStyleName;
	int ApplyStyleIDX;
	System::UnicodeString ApplyBaseCellAddres;
};


struct DECLSPEC_DRECORD TZEODFStyleProperties
{
	
private:
	typedef System::DynamicArray<TZEConditionMap> _TZEODFStyleProperties__1;
	
	
public:
	System::UnicodeString name;
	int index;
	System::UnicodeString ParentName;
	bool isHaveParent;
	int ConditionsCount;
	_TZEODFStyleProperties__1 Conditions;
};


typedef System::DynamicArray<TZEODFStyleProperties> TZODFStyleArray;

struct DECLSPEC_DRECORD TZEODFCFLine
{
	
public:
	int CellNum;
	int StyleNumber;
	int Count;
};


struct DECLSPEC_DRECORD TZEODFCFAreaItem
{
	
public:
	int RowNum;
	int ColNum;
	int Width;
	int Height;
	int CFStyleNumber;
};


class DELPHICLASS TZODFConditionalReadHelper;
class DELPHICLASS TZEODFReadHelper;
class PASCALIMPLEMENTATION TZODFConditionalReadHelper : public System::TObject
{
	typedef System::TObject inherited;
	
private:
	typedef System::DynamicArray<TZEODFCFLine> _TZODFConditionalReadHelper__1;
	
	typedef System::DynamicArray<System::StaticArray<int, 2> > _TZODFConditionalReadHelper__2;
	
	typedef System::DynamicArray<TZEODFCFAreaItem> _TZODFConditionalReadHelper__3;
	
	
private:
	Zexmlss::TZEXMLSS* FXMLSS;
	int FCountInLine;
	int FMaxCountInLine;
	_TZODFConditionalReadHelper__1 FCurrentLine;
	_TZODFConditionalReadHelper__2 FColumnSCFNumbers;
	int FColumnsCount;
	int FMaxColumnsCount;
	_TZODFConditionalReadHelper__3 FAreas;
	int FAreasCount;
	int FMaxAreasCount;
	int FLineItemWidth;
	int FLineItemStartCell;
	int FLineItemStyleCFNumber;
	TZEODFReadHelper* FReadHelper;
	
protected:
	void __fastcall AddToLine(const int CellNum, const int AStyleCFNumber, const int ACount);
	bool __fastcall ODFReadGetConditional(const System::UnicodeString ConditionalValue, /* out */ Zexmlss::TZCondition &Condition, /* out */ Zexmlss::TZConditionalOperator &ConditionOperator, /* out */ System::UnicodeString &Value1, /* out */ System::UnicodeString &Value2);
	
public:
	__fastcall TZODFConditionalReadHelper(Zexmlss::TZEXMLSS* XMLSS);
	__fastcall virtual ~TZODFConditionalReadHelper(void);
	void __fastcall CheckCell(int CellNum, int AStyleCFNumber, int RepeatCount = 0x1);
	void __fastcall ApplyConditionStylesToSheet(int SheetNumber, TZODFStyleArray &DefStylesArray, int DefStylesCount, TZODFStyleArray &StylesArray, int StylesCount);
	void __fastcall AddColumnCF(int ColumnNumber, int StyleCFNumber);
	int __fastcall GetColumnCF(int ColumnNumber);
	void __fastcall ApplyBaseCellAddr(const System::UnicodeString BaseCellTxt, const Zexmlss::TZConditionalStyleItem* ACFStyle, int PageNum);
	void __fastcall Clear(void);
	void __fastcall ClearLine(void);
	void __fastcall ProgressLine(int RowNumber, int RepeatCount = 0x1);
	void __fastcall ReadCalcextTag(Zsspxml::TZsspXMLReaderH* &xml, int SheetNum);
	__property int ColumnsCount = {read=FColumnsCount, nodefault};
	__property int LineItemWidth = {read=FLineItemWidth, write=FLineItemWidth, nodefault};
	__property int LineItemStartCell = {read=FLineItemStartCell, write=FLineItemStartCell, nodefault};
	__property int LineItemStyleID = {read=FLineItemStyleCFNumber, write=FLineItemStyleCFNumber, nodefault};
	__property TZEODFReadHelper* ReadHelper = {read=FReadHelper, write=FReadHelper};
};


typedef System::DynamicArray<Zexmlss::TZConditionalAreas*> TODFCFAreas;

struct DECLSPEC_DRECORD TODFCFmatch
{
	
public:
	int StyleID;
	int StyleCFID;
};


struct DECLSPEC_DRECORD TODFStyleCFID
{
	
private:
	typedef System::DynamicArray<TODFCFmatch> _TODFStyleCFID__1;
	
	
public:
	int Count;
	_TODFStyleCFID__1 ID;
};


struct DECLSPEC_DRECORD TODFCFWriterArray
{
	
private:
	typedef System::DynamicArray<TODFStyleCFID> _TODFCFWriterArray__1;
	
	
public:
	int CountCF;
	_TODFCFWriterArray__1 StyleCFID;
	TODFCFAreas Areas;
};


class DELPHICLASS TZODFConditionalWriteHelper;
class PASCALIMPLEMENTATION TZODFConditionalWriteHelper : public System::TObject
{
	typedef System::TObject inherited;
	
private:
	typedef System::DynamicArray<TODFCFWriterArray> _TZODFConditionalWriteHelper__1;
	
	typedef System::DynamicArray<int> _TZODFConditionalWriteHelper__2;
	
	
private:
	int FPagesCount;
	TIntegerDynArray FPageIndex;
	TStringDynArray FPageNames;
	_TZODFConditionalWriteHelper__1 FPageCF;
	TIntegerDynArray FFirstCFIdInPage;
	Zexmlss::TZEXMLSS* FXMLSS;
	int FStylesCount;
	int FApplyCFStylesCount;
	int FMaxApplyCFStylesCount;
	_TZODFConditionalWriteHelper__2 FApplyCFStyles;
	
protected:
	System::UnicodeString __fastcall GetBaseCellAddr(const Zexmlss::TZConditionalStyleItem* StCondition, const System::UnicodeString CurrPageName);
	bool __fastcall AddBetweenCond(const System::UnicodeString ConditName, const System::UnicodeString Value1, const System::UnicodeString Value2, /* out */ System::UnicodeString &retCondition);
	
public:
	__fastcall TZODFConditionalWriteHelper(Zexmlss::TZEXMLSS* ZEXMLSS, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount);
	bool __fastcall TryAddApplyCFStyle(int AStyleIndex, /* out */ int &retCFIndex);
	int __fastcall GetApplyCFStyle(int AStyleIndex);
	void __fastcall WriteCFStyles(Zsspxml::TZsspXMLWriterH* xml);
	void __fastcall WriteCalcextCF(Zsspxml::TZsspXMLWriterH* xml, int PageIndex);
	System::UnicodeString __fastcall ODSGetOperatorStr(Zexmlss::TZConditionalOperator AOperator);
	int __fastcall GetStyleNum(const int PageIndex, const int Col, const int Row);
	__fastcall virtual ~TZODFConditionalWriteHelper(void);
	__property int StylesCount = {read=FStylesCount, nodefault};
};


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
	/* TObject.Destroy */ inline __fastcall virtual ~TZEODFReadWriteHelperParent(void) { }
	
};


class PASCALIMPLEMENTATION TZEODFReadHelper : public TZEODFReadWriteHelperParent
{
	typedef TZEODFReadWriteHelperParent inherited;
	
private:
	typedef System::DynamicArray<Zexmlss::TZStyle*> _TZEODFReadHelper__1;
	
	typedef System::DynamicArray<Zexmlss::TZSheetOptions*> _TZEODFReadHelper__2;
	
	typedef System::DynamicArray<System::UnicodeString> _TZEODFReadHelper__3;
	
	typedef System::DynamicArray<Zexmlss::TZSheetOptions*> _TZEODFReadHelper__4;
	
	typedef System::DynamicArray<System::UnicodeString> _TZEODFReadHelper__5;
	
	
private:
	int FStylesCount;
	_TZEODFReadHelper__1 FStyles;
	int FMasterPagesCount;
	_TZEODFReadHelper__2 FMasterPages;
	_TZEODFReadHelper__3 FMasterPagesNames;
	int FPageLayoutsCount;
	_TZEODFReadHelper__4 FPageLayouts;
	_TZEODFReadHelper__5 FPageLayoutsNames;
	TZODFConditionalReadHelper* FConditionReader;
	Zexmlss::TZStyle* __fastcall GetStyle(int num);
	
public:
	TZODFStyleArray StylesProperties;
	__fastcall virtual TZEODFReadHelper(Zexmlss::TZEXMLSS* AXMLSS);
	__fastcall virtual ~TZEODFReadHelper(void);
	void __fastcall ReadAutomaticStyles(Zsspxml::TZsspXMLReaderH* xml);
	void __fastcall ReadMasterStyles(Zsspxml::TZsspXMLReaderH* xml);
	void __fastcall AddStyle(void);
	void __fastcall ApplyMasterPageStyle(Zexmlss::TZSheetOptions* SheetOptions, const System::UnicodeString MasterPageName);
	__property int StylesCount = {read=FStylesCount, nodefault};
	__property Zexmlss::TZStyle* Style[int num] = {read=GetStyle};
	__property TZODFConditionalReadHelper* ConditionReader = {read=FConditionReader};
};


class DELPHICLASS TZEODFWriteHelper;
class PASCALIMPLEMENTATION TZEODFWriteHelper : public TZEODFReadWriteHelperParent
{
	typedef TZEODFReadWriteHelperParent inherited;
	
private:
	typedef System::DynamicArray<int> _TZEODFWriteHelper__1;
	
	typedef System::DynamicArray<int> _TZEODFWriteHelper__2;
	
	typedef System::DynamicArray<int> _TZEODFWriteHelper__3;
	
	typedef System::DynamicArray<int> _TZEODFWriteHelper__4;
	
	typedef System::DynamicArray<System::UnicodeString> _TZEODFWriteHelper__5;
	
	
private:
	TZODFConditionalWriteHelper* FConditionWriter;
	_TZEODFWriteHelper__1 FUniquePageLayouts;
	int FUniquePageLayoutsCount;
	_TZEODFWriteHelper__2 FPageLayoutsIndexes;
	_TZEODFWriteHelper__3 FMasterPagesIndexes;
	int FMasterPagesCount;
	_TZEODFWriteHelper__4 FMasterPages;
	_TZEODFWriteHelper__5 FMasterPagesNames;
	
public:
	__fastcall virtual TZEODFWriteHelper(Zexmlss::TZEXMLSS* AXMLSS, const TIntegerDynArray _pages, const TStringDynArray _names, int PagesCount)/* overload */;
	__fastcall virtual ~TZEODFWriteHelper(void);
	void __fastcall WriteStylesPageLayouts(Zsspxml::TZsspXMLWriterH* xml, const TIntegerDynArray _pages);
	void __fastcall WriteStylesMasterPages(Zsspxml::TZsspXMLWriterH* xml, const TIntegerDynArray _pages);
	System::UnicodeString __fastcall GetMasterPageName(int PageNum);
	__property TZODFConditionalWriteHelper* ConditionWriter = {read=FConditionWriter};
};


//-- var, const, procedure ---------------------------------------------------
extern PACKAGE int __fastcall ODFCreateStyles(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM, const TZEODFWriteHelper* WriteHelper);
extern PACKAGE int __fastcall ODFCreateSettings(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM);
extern PACKAGE int __fastcall ODFCreateContent(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, const TIntegerDynArray _pages, const TStringDynArray _names, int PageCount, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM, const TZEODFWriteHelper* WriteHelper);
extern PACKAGE int __fastcall ODFCreateMeta(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* Stream, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM);
extern PACKAGE int __fastcall SaveXmlssToODFSPath(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString PathName, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToODFSPath(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString PathName, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToODFSPath(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString PathName)/* overload */;
extern PACKAGE int __fastcall ExportXmlssToODFS(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString FileName, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM = "", bool AllowUnzippedFolder = false, Zezippy::CZxZipGens ZipGenerator = 0x0)/* overload */;
extern PACKAGE bool __fastcall ReadODFContent(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* stream, TZEODFReadHelper* &ReadHelper);
extern PACKAGE bool __fastcall ReadODFSettings(Zexmlss::TZEXMLSS* &XMLSS, Classes::TStream* stream);
extern PACKAGE int __fastcall ReadODFSPath(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString DirName);
extern PACKAGE int __fastcall SaveXmlssToODFS(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString FileName, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM = "")/* overload */;
extern PACKAGE int __fastcall SaveXmlssToODFS(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString FileName, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size)/* overload */;
extern PACKAGE int __fastcall SaveXmlssToODFS(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString FileName)/* overload */;
extern PACKAGE int __fastcall ReadODFS(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString FileName);

}	/* namespace Zeodfs */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE)
using namespace Zeodfs;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZeodfsHPP
