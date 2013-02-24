// Borland C++ Builder
// Copyright (c) 1995, 2005 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'Zexmlss.pas' rev: 10.00

#ifndef ZexmlssHPP
#define ZexmlssHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <Sysinit.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <Sysutils.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <Windows.hpp>	// Pascal unit
#include <Zsspxml.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zexmlss
{
//-- type declarations -------------------------------------------------------
#pragma option push -b-
enum TZCellType { ZENumber, ZEDateTime, ZEBoolean, ZEString, ZEError };
#pragma option pop

#pragma option push -b-
enum TZBorderType { ZENone, ZEContinuous, ZEDot, ZEDash, ZEDashDot, ZEDashDotDot, ZESlantDashDot, ZEDouble };
#pragma option pop

#pragma option push -b-
enum TZHorizontalAlignment { ZHAutomatic, ZHLeft, ZHCenter, ZHRight, ZHFill, ZHJustify, ZHCenterAcrossSelection, ZHDistributed, ZHJustifyDistributed };
#pragma option pop

#pragma option push -b-
enum TZVerticalAlignment { ZVAutomatic, ZVTop, ZVBottom, ZVCenter, ZVJustify, ZVDistributed, ZVJustifyDistributed };
#pragma option pop

#pragma option push -b-
enum TZCellPattern { ZPNone, ZPSolid, ZPGray75, ZPGray50, ZPGray25, ZPGray125, ZPGray0625, ZPHorzStripe, ZPVertStripe, ZPReverseDiagStripe, ZPDiagStripe, ZPDiagCross, ZPThickDiagCross, ZPThinHorzStripe, ZPThinVertStripe, ZPThinReverseDiagStripe, ZPThinDiagStripe, ZPThinHorzCross, ZPThinDiagCross };
#pragma option pop

#pragma option push -b-
enum TZSplitMode { ZSplitNone, ZSplitFrozen, ZSplitSplit };
#pragma option pop

class DELPHICLASS TZCell;
class PASCALIMPLEMENTATION TZCell : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	AnsiString FFormula;
	AnsiString FData;
	AnsiString FHref;
	AnsiString FHRefScreenTip;
	AnsiString FComment;
	AnsiString FCommentAuthor;
	bool FAlwaysShowComment;
	bool FShowComment;
	TZCellType FCellType;
	int FCellStyle;
	double __fastcall GetDataAsDouble(void);
	void __fastcall SetDataAsDouble(const double Value);
	void __fastcall SetDataAsInteger(const int Value);
	int __fastcall GetDataAsInteger(void);
	
public:
	__fastcall virtual TZCell(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	void __fastcall Clear(void);
	__property bool AlwaysShowComment = {read=FAlwaysShowComment, write=FAlwaysShowComment, default=0};
	__property AnsiString Comment = {read=FComment, write=FComment};
	__property AnsiString CommentAuthor = {read=FCommentAuthor, write=FCommentAuthor};
	__property int CellStyle = {read=FCellStyle, write=FCellStyle, default=-1};
	__property TZCellType CellType = {read=FCellType, write=FCellType, default=3};
	__property AnsiString Data = {read=FData, write=FData};
	__property AnsiString Formula = {read=FFormula, write=FFormula};
	__property AnsiString HRef = {read=FHref, write=FHref};
	__property AnsiString HRefScreenTip = {read=FHRefScreenTip, write=FHRefScreenTip};
	__property bool ShowComment = {read=FShowComment, write=FShowComment, default=0};
	__property double AsDouble = {read=GetDataAsDouble, write=SetDataAsDouble};
	__property int AsInteger = {read=GetDataAsInteger, write=SetDataAsInteger, nodefault};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZCell(void) { }
	#pragma option pop
	
};


class DELPHICLASS TZBorderStyle;
class PASCALIMPLEMENTATION TZBorderStyle : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZBorderType FLineStyle;
	Byte FWeight;
	Graphics::TColor FColor;
	void __fastcall SetLineStyle(const TZBorderType Value);
	void __fastcall SetWeight(const Byte Value);
	void __fastcall SetColor(const Graphics::TColor Value);
	
public:
	__fastcall virtual TZBorderStyle(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	
__published:
	__property TZBorderType LineStyle = {read=FLineStyle, write=SetLineStyle, default=0};
	__property Byte Weight = {read=FWeight, write=SetWeight, default=0};
	__property Graphics::TColor Color = {read=FColor, write=SetColor, default=0};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZBorderStyle(void) { }
	#pragma option pop
	
};


class DELPHICLASS TZBorder;
class PASCALIMPLEMENTATION TZBorder : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
public:
	TZBorderStyle* operator[](int Num) { return Border[Num]; }
	
private:
	TZBorderStyle* FBorder[6];
	void __fastcall SetBorder(int Num, const TZBorderStyle* Value);
	TZBorderStyle* __fastcall GetBorder(int Num);
	
public:
	__fastcall virtual TZBorder(void);
	__fastcall virtual ~TZBorder(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	__property TZBorderStyle* Border[int Num] = {read=GetBorder, write=SetBorder/*, default*/};
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	
__published:
	__property TZBorderStyle* Left = {read=GetBorder, write=SetBorder, index=0};
	__property TZBorderStyle* Top = {read=GetBorder, write=SetBorder, index=1};
	__property TZBorderStyle* Right = {read=GetBorder, write=SetBorder, index=2};
	__property TZBorderStyle* Bottom = {read=GetBorder, write=SetBorder, index=3};
	__property TZBorderStyle* DiagonalLeft = {read=GetBorder, write=SetBorder, index=4};
	__property TZBorderStyle* DiagonalRight = {read=GetBorder, write=SetBorder, index=5};
};


typedef short TZCellTextRotate;

class DELPHICLASS TZAlignment;
class PASCALIMPLEMENTATION TZAlignment : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZHorizontalAlignment FHorizontal;
	int FIndent;
	TZCellTextRotate FRotate;
	bool FShrinkToFit;
	TZVerticalAlignment FVertical;
	bool FVerticalText;
	bool FWrapText;
	void __fastcall SetHorizontal(const TZHorizontalAlignment Value);
	void __fastcall SetIndent(const int Value);
	void __fastcall SetRotate(const TZCellTextRotate Value);
	void __fastcall SetShrinkToFit(const bool Value);
	void __fastcall SetVertical(const TZVerticalAlignment Value);
	void __fastcall SetVerticalText(const bool Value);
	void __fastcall SetWrapText(const bool Value);
	
public:
	__fastcall virtual TZAlignment(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	
__published:
	__property TZHorizontalAlignment Horizontal = {read=FHorizontal, write=SetHorizontal, default=0};
	__property int Indent = {read=FIndent, write=SetIndent, default=0};
	__property TZCellTextRotate Rotate = {read=FRotate, write=SetRotate, nodefault};
	__property bool ShrinkToFit = {read=FShrinkToFit, write=SetShrinkToFit, default=0};
	__property TZVerticalAlignment Vertical = {read=FVertical, write=SetVertical, default=0};
	__property bool VerticalText = {read=FVerticalText, write=SetVerticalText, default=0};
	__property bool WrapText = {read=FWrapText, write=SetWrapText, default=0};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZAlignment(void) { }
	#pragma option pop
	
};


class DELPHICLASS TZStyle;
class PASCALIMPLEMENTATION TZStyle : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZBorder* FBorder;
	TZAlignment* FAlignment;
	Graphics::TFont* FFont;
	Graphics::TColor FBGColor;
	Graphics::TColor FPatternColor;
	TZCellPattern FCellPattern;
	AnsiString FNumberFormat;
	bool FProtect;
	bool FHideFormula;
	void __fastcall SetFont(const Graphics::TFont* Value);
	void __fastcall SetBorder(const TZBorder* Value);
	void __fastcall SetAlignment(const TZAlignment* Value);
	void __fastcall SetBGColor(const Graphics::TColor Value);
	void __fastcall SetPatternColor(const Graphics::TColor Value);
	void __fastcall SetCellPattern(const TZCellPattern Value);
	
protected:
	virtual void __fastcall SetNumberFormat(const AnsiString Value);
	
public:
	__fastcall virtual TZStyle(void);
	__fastcall virtual ~TZStyle(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	
__published:
	__property Graphics::TFont* Font = {read=FFont, write=SetFont};
	__property TZBorder* Border = {read=FBorder, write=SetBorder};
	__property TZAlignment* Alignment = {read=FAlignment, write=SetAlignment};
	__property Graphics::TColor BGColor = {read=FBGColor, write=SetBGColor, default=-16777211};
	__property Graphics::TColor PatternColor = {read=FPatternColor, write=SetPatternColor, default=-16777211};
	__property bool Protect = {read=FProtect, write=FProtect, default=1};
	__property bool HideFormula = {read=FHideFormula, write=FHideFormula, default=0};
	__property TZCellPattern CellPattern = {read=FCellPattern, write=SetCellPattern, default=0};
	__property AnsiString NumberFormat = {read=FNumberFormat, write=SetNumberFormat};
};


typedef DynamicArray<TZStyle* >  zexmlss__7;

class DELPHICLASS TZStyles;
class PASCALIMPLEMENTATION TZStyles : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
public:
	TZStyle* operator[](int num) { return Items[num]; }
	
private:
	TZStyle* FDefaultStyle;
	DynamicArray<TZStyle* >  FStyles;
	int FCount;
	void __fastcall SetDefaultStyle(const TZStyle* Value);
	TZStyle* __fastcall GetStyle(int num);
	void __fastcall SetStyle(int num, const TZStyle* Value);
	void __fastcall SetCount(const int Value);
	
public:
	__fastcall virtual TZStyles(void);
	__fastcall virtual ~TZStyles(void);
	int __fastcall Add(const TZStyle* Style, bool CheckMatch = false);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual void __fastcall Clear(void);
	virtual int __fastcall DeleteStyle(int num);
	int __fastcall Find(const TZStyle* Style);
	__property TZStyle* Items[int num] = {read=GetStyle, write=SetStyle/*, default*/};
	__property int Count = {read=FCount, write=SetCount, nodefault};
	
__published:
	__property TZStyle* DefaultStyle = {read=FDefaultStyle, write=SetDefaultStyle};
};


typedef DynamicArray<Types::TRect >  zexmlss__9;

class DELPHICLASS TZMergeCells;
class DELPHICLASS TZSheet;
class DELPHICLASS TZEXMLSS;
class DELPHICLASS TZSheets;
class PASCALIMPLEMENTATION TZSheets : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
public:
	TZSheet* operator[](int num) { return Sheet[num]; }
	
private:
	TZEXMLSS* FStore;
	DynamicArray<TZSheet* >  FSheets;
	int FCount;
	void __fastcall SetSheetCount(const int Value);
	void __fastcall SetSheet(int num, const TZSheet* Value);
	TZSheet* __fastcall GetSheet(int num);
	
public:
	__fastcall virtual TZSheets(TZEXMLSS* AStore);
	__fastcall virtual ~TZSheets(void);
	__property int Count = {read=FCount, write=SetSheetCount, nodefault};
	__property TZSheet* Sheet[int num] = {read=GetSheet, write=SetSheet/*, default*/};
};


class DELPHICLASS TZEXMLDocumentProperties;
class PASCALIMPLEMENTATION TZEXMLDocumentProperties : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	AnsiString FAuthor;
	AnsiString FLastAuthor;
	System::TDateTime FCreated;
	AnsiString FCompany;
	AnsiString FVersion;
	Word FWindowHeight;
	Word FWindowWidth;
	int FWindowTopX;
	int FWindowTopY;
	bool FModeR1C1;
	
protected:
	void __fastcall SetAuthor(const AnsiString Value);
	void __fastcall SetLastAuthor(const AnsiString Value);
	void __fastcall SetCompany(const AnsiString Value);
	void __fastcall SetVersion(const AnsiString Value);
	
public:
	__fastcall virtual TZEXMLDocumentProperties(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	
__published:
	__property AnsiString Author = {read=FAuthor, write=SetAuthor};
	__property AnsiString LastAuthor = {read=FLastAuthor, write=SetLastAuthor};
	__property System::TDateTime Created = {read=FCreated, write=FCreated};
	__property AnsiString Company = {read=FCompany, write=SetCompany};
	__property AnsiString Version = {read=FVersion, write=SetVersion};
	__property bool ModeR1C1 = {read=FModeR1C1, write=FModeR1C1, default=0};
	__property Word WindowHeight = {read=FWindowHeight, write=FWindowHeight, default=20000};
	__property Word WindowWidth = {read=FWindowWidth, write=FWindowWidth, default=20000};
	__property int WindowTopX = {read=FWindowTopX, write=FWindowTopX, default=150};
	__property int WindowTopY = {read=FWindowTopY, write=FWindowTopY, default=150};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZEXMLDocumentProperties(void) { }
	#pragma option pop
	
};


class DELPHICLASS TZSheetOptions;
class PASCALIMPLEMENTATION TZSheetOptions : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	Word FActiveCol;
	Word FActiveRow;
	Word FMarginBottom;
	Word FMarginLeft;
	Word FMarginTop;
	Word FMarginRight;
	bool FPortraitOrientation;
	bool FCenterHorizontal;
	bool FCenterVertical;
	int FStartPageNumber;
	Word FHeaderMargin;
	Word FFooterMargin;
	AnsiString FHeaderData;
	AnsiString FFooterData;
	Byte FPaperSize;
	TZSplitMode FSplitVerticalMode;
	TZSplitMode FSplitHorizontalMode;
	int FSplitVerticalValue;
	int FSplitHorizontalValue;
	
public:
	__fastcall virtual TZSheetOptions(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	
__published:
	__property Word ActiveCol = {read=FActiveCol, write=FActiveCol, default=0};
	__property Word ActiveRow = {read=FActiveRow, write=FActiveRow, default=0};
	__property Word MarginBottom = {read=FMarginBottom, write=FMarginBottom, default=25};
	__property Word MarginLeft = {read=FMarginLeft, write=FMarginLeft, default=20};
	__property Word MarginTop = {read=FMarginTop, write=FMarginTop, default=25};
	__property Word MarginRight = {read=FMarginRight, write=FMarginRight, default=20};
	__property Byte PaperSize = {read=FPaperSize, write=FPaperSize, default=9};
	__property bool PortraitOrientation = {read=FPortraitOrientation, write=FPortraitOrientation, default=1};
	__property bool CenterHorizontal = {read=FCenterHorizontal, write=FCenterHorizontal, default=0};
	__property bool CenterVertical = {read=FCenterVertical, write=FCenterVertical, default=0};
	__property int StartPageNumber = {read=FStartPageNumber, write=FStartPageNumber, default=1};
	__property Word HeaderMargin = {read=FHeaderMargin, write=FHeaderMargin, default=13};
	__property Word FooterMargin = {read=FFooterMargin, write=FFooterMargin, default=13};
	__property AnsiString HeaderData = {read=FHeaderData, write=FHeaderData};
	__property AnsiString FooterData = {read=FFooterData, write=FFooterData};
	__property TZSplitMode SplitVerticalMode = {read=FSplitVerticalMode, write=FSplitVerticalMode, default=0};
	__property TZSplitMode SplitHorizontalMode = {read=FSplitHorizontalMode, write=FSplitHorizontalMode, default=0};
	__property int SplitVerticalValue = {read=FSplitVerticalValue, write=FSplitVerticalValue, nodefault};
	__property int SplitHorizontalValue = {read=FSplitHorizontalValue, write=FSplitHorizontalValue, nodefault};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZSheetOptions(void) { }
	#pragma option pop
	
};


class PASCALIMPLEMENTATION TZEXMLSS : public Classes::TComponent 
{
	typedef Classes::TComponent inherited;
	
private:
	TZSheets* FSheets;
	TZEXMLDocumentProperties* FDocumentProperties;
	TZStyles* FStyles;
	double FHorPixelSize;
	double FVertPixelSize;
	TZSheetOptions* FDefaultSheetOptions;
	void __fastcall SetHorPixelSize(double Value);
	void __fastcall SetVertPixelSize(double Value);
	TZSheetOptions* __fastcall GetDefaultSheetOptions(void);
	void __fastcall SetDefaultSheetOptions(TZSheetOptions* Value);
	
public:
	__fastcall virtual TZEXMLSS(Classes::TComponent* AOwner);
	__fastcall virtual ~TZEXMLSS(void);
	void __fastcall GetPixelSize(HWND hdc);
	__property TZSheets* Sheets = {read=FSheets, write=FSheets};
	
__published:
	__property TZStyles* Styles = {read=FStyles, write=FStyles};
	__property TZSheetOptions* DefaultSheetOptions = {read=GetDefaultSheetOptions, write=SetDefaultSheetOptions};
	__property TZEXMLDocumentProperties* DocumentProperties = {read=FDocumentProperties, write=FDocumentProperties};
	__property double HorPixelSize = {read=FHorPixelSize, write=SetHorPixelSize};
	__property double VertPixelSize = {read=FVertPixelSize, write=SetVertPixelSize};
};


class DELPHICLASS TZSheetPrintTitles;
class PASCALIMPLEMENTATION TZSheetPrintTitles : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZSheet* FOwner;
	bool FColumns;
	bool FActive;
	Word FTill;
	Word FFrom;
	void __fastcall SetActive(const bool Value);
	void __fastcall SetFrom(const Word Value);
	void __fastcall SetTill(const Word Value);
	bool __fastcall Valid(const Word AFrom, const Word ATill);
	void __fastcall RequireValid(const Word AFrom, const Word ATill);
	
public:
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	__fastcall TZSheetPrintTitles(const TZSheet* owner, const bool ForColumns);
	AnsiString __fastcall ToString();
	
__published:
	__property Word From = {read=FFrom, write=SetFrom, nodefault};
	__property Word Till = {read=FTill, write=SetTill, nodefault};
	__property bool Active = {read=FActive, write=SetActive, nodefault};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZSheetPrintTitles(void) { }
	#pragma option pop
	
};


class DELPHICLASS TZColOptions;
class DELPHICLASS TZRowOptions;
class PASCALIMPLEMENTATION TZSheet : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZEXMLSS* FStore;
	DynamicArray<DynamicArray<TZCell* > >  FCells;
	DynamicArray<TZRowOptions* >  FRows;
	DynamicArray<TZColOptions* >  FColumns;
	AnsiString FTitle;
	int FRowCount;
	int FColCount;
	Graphics::TColor FTabColor;
	double FDefaultRowHeight;
	double FDefaultColWidth;
	TZMergeCells* FMergeCells;
	bool FProtect;
	bool FRightToLeft;
	TZSheetOptions* FSheetOptions;
	bool FSelected;
	TZSheetPrintTitles* FPrintRows;
	TZSheetPrintTitles* FPrintCols;
	void __fastcall SetColumn(int num, const TZColOptions* Value);
	TZColOptions* __fastcall GetColumn(int num);
	void __fastcall SetRow(int num, const TZRowOptions* Value);
	TZRowOptions* __fastcall GetRow(int num);
	TZSheetOptions* __fastcall GetSheetOptions(void);
	void __fastcall SetSheetOptions(TZSheetOptions* Value);
	void __fastcall SetPrintCols(const TZSheetPrintTitles* Value);
	void __fastcall SetPrintRows(const TZSheetPrintTitles* Value);
	
protected:
	virtual void __fastcall SetColWidth(int num, const double Value);
	virtual double __fastcall GetColWidth(int num);
	virtual void __fastcall SetRowHeight(int num, const double Value);
	virtual double __fastcall GetRowHeight(int num);
	virtual void __fastcall SetDefaultColWidth(const double Value);
	virtual void __fastcall SetDefaultRowHeight(const double Value);
	virtual TZCell* __fastcall GetCell(int ACol, int ARow);
	virtual void __fastcall SetCell(int ACol, int ARow, const TZCell* Value);
	virtual int __fastcall GetRowCount(void);
	virtual void __fastcall SetRowCount(const int Value);
	virtual int __fastcall GetColCount(void);
	virtual void __fastcall SetColCount(const int Value);
	
public:
	__fastcall virtual TZSheet(TZEXMLSS* AStore);
	__fastcall virtual ~TZSheet(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual void __fastcall Clear(void);
	__property double ColWidths[int num] = {read=GetColWidth, write=SetColWidth};
	__property TZColOptions* Columns[int num] = {read=GetColumn, write=SetColumn};
	__property TZRowOptions* Rows[int num] = {read=GetRow, write=SetRow};
	__property double RowHeights[int num] = {read=GetRowHeight, write=SetRowHeight};
	__property double DefaultColWidth = {read=FDefaultColWidth, write=SetDefaultColWidth};
	__property double DefaultRowHeight = {read=FDefaultRowHeight, write=SetDefaultRowHeight};
	__property TZCell* Cell[int ACol][int ARow] = {read=GetCell, write=SetCell};
	__property bool Protect = {read=FProtect, write=FProtect, default=0};
	__property Graphics::TColor TabColor = {read=FTabColor, write=FTabColor, default=-16777211};
	__property AnsiString Title = {read=FTitle, write=FTitle};
	__property int RowCount = {read=GetRowCount, write=SetRowCount, nodefault};
	__property bool RightToLeft = {read=FRightToLeft, write=FRightToLeft, default=0};
	__property int ColCount = {read=GetColCount, write=SetColCount, nodefault};
	__property TZMergeCells* MergeCells = {read=FMergeCells, write=FMergeCells};
	__property TZSheetOptions* SheetOptions = {read=GetSheetOptions, write=SetSheetOptions};
	__property bool Selected = {read=FSelected, write=FSelected, nodefault};
	__property TZSheetPrintTitles* RowsToRepeat = {read=FPrintRows, write=SetPrintRows};
	__property TZSheetPrintTitles* ColsToRepeat = {read=FPrintCols, write=SetPrintCols};
};


class PASCALIMPLEMENTATION TZMergeCells : public System::TObject 
{
	typedef System::TObject inherited;
	
public:
	Types::TRect operator[](int Num) { return Items[Num]; }
	
private:
	TZSheet* FSheet;
	int FCount;
	DynamicArray<Types::TRect >  FMergeArea;
	Types::TRect __fastcall GetItem(int Num);
	
public:
	__fastcall virtual TZMergeCells(TZSheet* ASheet);
	__fastcall virtual ~TZMergeCells(void);
	Byte __fastcall AddRect(const Types::TRect &Rct);
	Byte __fastcall AddRectXY(int x1, int y1, int x2, int y2);
	bool __fastcall DeleteItem(int num);
	int __fastcall InLeftTopCorner(int ACol, int ARow);
	int __fastcall InMergeRange(int ACol, int ARow);
	void __fastcall Clear(void);
	__property int Count = {read=FCount, nodefault};
	__property Types::TRect Items[int Num] = {read=GetItem/*, default*/};
};


typedef DynamicArray<TZCell* >  TZCellColumn;

class DELPHICLASS TZRowColOptions;
class PASCALIMPLEMENTATION TZRowColOptions : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZSheet* FSheet;
	bool FHidden;
	int FStyleID;
	double FSize;
	bool FAuto;
	bool FBreaked;
	
protected:
	bool __fastcall GetAuto(void);
	void __fastcall SetAuto(bool Value);
	double __fastcall GetSizePoint(void);
	void __fastcall SetSizePoint(double Value);
	double __fastcall GetSizeMM(void);
	void __fastcall SetSizeMM(double Value);
	virtual int __fastcall GetSizePix(void) = 0 ;
	virtual void __fastcall SetSizePix(int Value) = 0 ;
	
public:
	__fastcall virtual TZRowColOptions(TZSheet* ASheet);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	__property bool Hidden = {read=FHidden, write=FHidden, default=0};
	__property int StyleID = {read=FStyleID, write=FStyleID, default=-1};
	__property bool Breaked = {read=FBreaked, write=FBreaked, default=0};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZRowColOptions(void) { }
	#pragma option pop
	
};


class PASCALIMPLEMENTATION TZColOptions : public TZRowColOptions 
{
	typedef TZRowColOptions inherited;
	
protected:
	virtual int __fastcall GetSizePix(void);
	virtual void __fastcall SetSizePix(int Value);
	
public:
	__fastcall virtual TZColOptions(TZSheet* ASheet);
	__property bool AutoFitWidth = {read=GetAuto, write=SetAuto, nodefault};
	__property double Width = {read=GetSizePoint, write=SetSizePoint};
	__property double WidthMM = {read=GetSizeMM, write=SetSizeMM};
	__property int WidthPix = {read=GetSizePix, write=SetSizePix, nodefault};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZColOptions(void) { }
	#pragma option pop
	
};


class PASCALIMPLEMENTATION TZRowOptions : public TZRowColOptions 
{
	typedef TZRowColOptions inherited;
	
protected:
	virtual int __fastcall GetSizePix(void);
	virtual void __fastcall SetSizePix(int Value);
	
public:
	__fastcall virtual TZRowOptions(TZSheet* ASheet);
	__property bool AutoFitHeight = {read=GetAuto, write=SetAuto, nodefault};
	__property double Height = {read=GetSizePoint, write=SetSizePoint};
	__property double HeightMM = {read=GetSizeMM, write=SetSizeMM};
	__property int HeightPix = {read=GetSizePix, write=SetSizePix, nodefault};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZRowOptions(void) { }
	#pragma option pop
	
};


typedef DynamicArray<DynamicArray<TZCell* > >  zexmlss__61;

typedef DynamicArray<TZRowOptions* >  zexmlss__71;

typedef DynamicArray<TZColOptions* >  zexmlss__81;

typedef DynamicArray<TZSheet* >  zexmlss__02;

//-- var, const, procedure ---------------------------------------------------
extern PACKAGE double _PointToMM;
#define ZEAnsiString (TZCellType)(3)
extern PACKAGE void __fastcall CorrectStrForXML(const AnsiString St, AnsiString &Corrected, bool &UseXMLNS);
extern PACKAGE AnsiString __fastcall ColorToHTMLHex(Graphics::TColor Color);
extern PACKAGE Graphics::TColor __fastcall HTMLHexToColor(AnsiString value);
extern PACKAGE double __fastcall PixelToPoint(int inPixel, double PixelSizeMM = 2.650000E-01);
extern PACKAGE int __fastcall PointToPixel(double inPoint, double PixelSizeMM = 2.650000E-01);
extern PACKAGE double __fastcall PointToMM(double inPoint);
extern PACKAGE double __fastcall MMToPoint(double inMM);
extern PACKAGE AnsiString __fastcall HAlToStr(TZHorizontalAlignment HA);
extern PACKAGE AnsiString __fastcall VAlToStr(TZVerticalAlignment VA);
extern PACKAGE AnsiString __fastcall ZBorderTypeToStr(TZBorderType ZB);
extern PACKAGE AnsiString __fastcall ZCellPatternToStr(TZCellPattern pp);
extern PACKAGE AnsiString __fastcall ZCellTypeToStr(TZCellType pp);
extern PACKAGE TZHorizontalAlignment __fastcall StrToHal(AnsiString Value);
extern PACKAGE TZVerticalAlignment __fastcall StrToVAl(AnsiString Value);
extern PACKAGE TZBorderType __fastcall StrToZBorderType(AnsiString Value);
extern PACKAGE TZCellPattern __fastcall StrToZCellPattern(AnsiString Value);
extern PACKAGE TZCellType __fastcall StrToZCellType(AnsiString Value);
extern PACKAGE void __fastcall Register(void);

}	/* namespace Zexmlss */
using namespace Zexmlss;
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// Zexmlss
