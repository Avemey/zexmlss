// CodeGear C++Builder
// Copyright (c) 1995, 2010 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zexmlss.pas' rev: 22.00

#ifndef ZexmlssHPP
#define ZexmlssHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <Windows.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <Types.hpp>	// Pascal unit

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
	System::UnicodeString FFormula;
	System::UnicodeString FData;
	System::UnicodeString FHref;
	System::UnicodeString FHRefScreenTip;
	System::UnicodeString FComment;
	System::UnicodeString FCommentAuthor;
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
	__property System::UnicodeString Comment = {read=FComment, write=FComment};
	__property System::UnicodeString CommentAuthor = {read=FCommentAuthor, write=FCommentAuthor};
	__property int CellStyle = {read=FCellStyle, write=FCellStyle, default=-1};
	__property TZCellType CellType = {read=FCellType, write=FCellType, default=3};
	__property System::UnicodeString Data = {read=FData, write=FData};
	__property System::UnicodeString Formula = {read=FFormula, write=FFormula};
	__property System::UnicodeString HRef = {read=FHref, write=FHref};
	__property System::UnicodeString HRefScreenTip = {read=FHRefScreenTip, write=FHRefScreenTip};
	__property bool ShowComment = {read=FShowComment, write=FShowComment, default=0};
	__property double AsDouble = {read=GetDataAsDouble, write=SetDataAsDouble};
	__property int AsInteger = {read=GetDataAsInteger, write=SetDataAsInteger, nodefault};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZCell(void) { }
	
};


class DELPHICLASS TZBorderStyle;
class PASCALIMPLEMENTATION TZBorderStyle : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	TZBorderType FLineStyle;
	System::Byte FWeight;
	Graphics::TColor FColor;
	void __fastcall SetLineStyle(const TZBorderType Value);
	void __fastcall SetWeight(const System::Byte Value);
	void __fastcall SetColor(const Graphics::TColor Value);
	
public:
	__fastcall virtual TZBorderStyle(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	
__published:
	__property TZBorderType LineStyle = {read=FLineStyle, write=SetLineStyle, default=0};
	__property System::Byte Weight = {read=FWeight, write=SetWeight, default=0};
	__property Graphics::TColor Color = {read=FColor, write=SetColor, default=0};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZBorderStyle(void) { }
	
};


class DELPHICLASS TZBorder;
class PASCALIMPLEMENTATION TZBorder : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
public:
	TZBorderStyle* operator[](int Num) { return Border[Num]; }
	
private:
	System::StaticArray<TZBorderStyle*, 6> FBorder;
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
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZAlignment(void) { }
	
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
	System::UnicodeString FNumberFormat;
	bool FProtect;
	bool FHideFormula;
	void __fastcall SetFont(const Graphics::TFont* Value);
	void __fastcall SetBorder(const TZBorder* Value);
	void __fastcall SetAlignment(const TZAlignment* Value);
	void __fastcall SetBGColor(const Graphics::TColor Value);
	void __fastcall SetPatternColor(const Graphics::TColor Value);
	void __fastcall SetCellPattern(const TZCellPattern Value);
	
protected:
	virtual void __fastcall SetNumberFormat(const System::UnicodeString Value);
	
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
	__property System::UnicodeString NumberFormat = {read=FNumberFormat, write=SetNumberFormat};
};


class DELPHICLASS TZStyles;
class PASCALIMPLEMENTATION TZStyles : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	typedef System::DynamicArray<TZStyle*> _TZStyles__1;
	
	
public:
	TZStyle* operator[](int num) { return Items[num]; }
	
private:
	TZStyle* FDefaultStyle;
	_TZStyles__1 FStyles;
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


class DELPHICLASS TZMergeCells;
class DELPHICLASS TZSheet;
class PASCALIMPLEMENTATION TZMergeCells : public System::TObject
{
	typedef System::TObject inherited;
	
private:
	typedef System::DynamicArray<Types::TRect> _TZMergeCells__1;
	
	
public:
	Types::TRect operator[](int Num) { return Items[Num]; }
	
private:
	TZSheet* FSheet;
	int FCount;
	_TZMergeCells__1 FMergeArea;
	Types::TRect __fastcall GetItem(int Num);
	
public:
	__fastcall virtual TZMergeCells(TZSheet* ASheet);
	__fastcall virtual ~TZMergeCells(void);
	System::Byte __fastcall AddRect(const Types::TRect &Rct);
	System::Byte __fastcall AddRectXY(int x1, int y1, int x2, int y2);
	bool __fastcall DeleteItem(int num);
	int __fastcall InLeftTopCorner(int ACol, int ARow);
	int __fastcall InMergeRange(int ACol, int ARow);
	void __fastcall Clear(void);
	__property int Count = {read=FCount, nodefault};
	__property Types::TRect Items[int Num] = {read=GetItem/*, default*/};
};


typedef System::DynamicArray<TZCell*> TZCellColumn;

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
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZRowColOptions(void) { }
	
};


class DELPHICLASS TZColOptions;
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
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZColOptions(void) { }
	
};


class DELPHICLASS TZRowOptions;
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
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZRowOptions(void) { }
	
};


class DELPHICLASS TZSheetPrintTitles;
class PASCALIMPLEMENTATION TZSheetPrintTitles : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	TZSheet* FOwner;
	bool FColumns;
	bool FActive;
	System::Word FTill;
	System::Word FFrom;
	void __fastcall SetActive(const bool Value);
	void __fastcall SetFrom(const System::Word Value);
	void __fastcall SetTill(const System::Word Value);
	bool __fastcall Valid(const System::Word AFrom, const System::Word ATill);
	void __fastcall RequireValid(const System::Word AFrom, const System::Word ATill);
	
public:
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	__fastcall TZSheetPrintTitles(const TZSheet* owner, const bool ForColumns);
	virtual System::UnicodeString __fastcall ToString(void);
	
__published:
	__property System::Word From = {read=FFrom, write=SetFrom, nodefault};
	__property System::Word Till = {read=FTill, write=SetTill, nodefault};
	__property bool Active = {read=FActive, write=SetActive, nodefault};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZSheetPrintTitles(void) { }
	
};


class DELPHICLASS TZSheetFooterHeader;
class PASCALIMPLEMENTATION TZSheetFooterHeader : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	System::UnicodeString FDataLeft;
	System::UnicodeString FData;
	System::UnicodeString FDataRight;
	bool FIsDisplay;
	
public:
	__fastcall TZSheetFooterHeader(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	
__published:
	__property System::UnicodeString DataLeft = {read=FDataLeft, write=FDataLeft};
	__property System::UnicodeString Data = {read=FData, write=FData};
	__property System::UnicodeString DataRight = {read=FDataRight, write=FDataRight};
	__property bool IsDisplay = {read=FIsDisplay, write=FIsDisplay, nodefault};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZSheetFooterHeader(void) { }
	
};


class DELPHICLASS TZHeaderFooterMargins;
class PASCALIMPLEMENTATION TZHeaderFooterMargins : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	System::Word FMarginTopBottom;
	System::Word FMarginLeft;
	System::Word FMarginRight;
	System::Word FHeight;
	bool FUseAutoFitHeight;
	
public:
	__fastcall TZHeaderFooterMargins(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	
__published:
	__property System::Word MarginTopBottom = {read=FMarginTopBottom, write=FMarginTopBottom, default=13};
	__property System::Word MarginLeft = {read=FMarginLeft, write=FMarginLeft, default=0};
	__property System::Word MarginRight = {read=FMarginRight, write=FMarginRight, default=0};
	__property System::Word Height = {read=FHeight, write=FHeight, default=7};
	__property bool UseAutoFitHeight = {read=FUseAutoFitHeight, write=FUseAutoFitHeight, default=1};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZHeaderFooterMargins(void) { }
	
};


class DELPHICLASS TZSheetOptions;
class PASCALIMPLEMENTATION TZSheetOptions : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	System::Word FActiveCol;
	System::Word FActiveRow;
	System::Word FMarginBottom;
	System::Word FMarginLeft;
	System::Word FMarginTop;
	System::Word FMarginRight;
	TZHeaderFooterMargins* FHeaderMargins;
	TZHeaderFooterMargins* FFooterMargins;
	bool FPortraitOrientation;
	bool FCenterHorizontal;
	bool FCenterVertical;
	int FStartPageNumber;
	bool FIsEvenFooterEqual;
	bool FIsEvenHeaderEqual;
	TZSheetFooterHeader* FHeader;
	TZSheetFooterHeader* FFooter;
	TZSheetFooterHeader* FEvenHeader;
	TZSheetFooterHeader* FEvenFooter;
	Graphics::TColor FHeaderBGColor;
	Graphics::TColor FFooterBGColor;
	int FScaleToPercent;
	int FScaleToPages;
	System::Byte FPaperSize;
	int FPaperWidth;
	int FPaperHeight;
	TZSplitMode FSplitVerticalMode;
	TZSplitMode FSplitHorizontalMode;
	int FSplitVerticalValue;
	int FSplitHorizontalValue;
	System::UnicodeString __fastcall GetHeaderData(void);
	void __fastcall SetHeaderData(System::UnicodeString Value);
	System::UnicodeString __fastcall GetFooterData(void);
	void __fastcall SetfooterData(System::UnicodeString Value);
	System::Word __fastcall GetHeaderMargin(void);
	void __fastcall SetHeaderMargin(System::Word Value);
	System::Word __fastcall GetFooterMargin(void);
	void __fastcall SetFooterMargin(System::Word Value);
	
public:
	__fastcall virtual TZSheetOptions(void);
	__fastcall virtual ~TZSheetOptions(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	
__published:
	__property System::Word ActiveCol = {read=FActiveCol, write=FActiveCol, default=0};
	__property System::Word ActiveRow = {read=FActiveRow, write=FActiveRow, default=0};
	__property System::Word MarginBottom = {read=FMarginBottom, write=FMarginBottom, default=25};
	__property System::Word MarginLeft = {read=FMarginLeft, write=FMarginLeft, default=20};
	__property System::Word MarginTop = {read=FMarginTop, write=FMarginTop, default=25};
	__property System::Word MarginRight = {read=FMarginRight, write=FMarginRight, default=20};
	__property System::Byte PaperSize = {read=FPaperSize, write=FPaperSize, default=9};
	__property int PaperWidth = {read=FPaperWidth, write=FPaperWidth, default=0};
	__property int PaperHeight = {read=FPaperHeight, write=FPaperHeight, default=0};
	__property bool PortraitOrientation = {read=FPortraitOrientation, write=FPortraitOrientation, default=1};
	__property bool CenterHorizontal = {read=FCenterHorizontal, write=FCenterHorizontal, default=0};
	__property bool CenterVertical = {read=FCenterVertical, write=FCenterVertical, default=0};
	__property int StartPageNumber = {read=FStartPageNumber, write=FStartPageNumber, default=1};
	__property System::Word HeaderMargin = {read=GetHeaderMargin, write=SetHeaderMargin, default=13};
	__property System::Word FooterMargin = {read=GetFooterMargin, write=SetFooterMargin, default=13};
	__property TZHeaderFooterMargins* HeaderMargins = {read=FHeaderMargins};
	__property TZHeaderFooterMargins* FooterMargins = {read=FFooterMargins};
	__property bool IsEvenFooterEqual = {read=FIsEvenFooterEqual, write=FIsEvenFooterEqual, default=1};
	__property bool IsEvenHeaderEqual = {read=FIsEvenHeaderEqual, write=FIsEvenHeaderEqual, default=1};
	__property System::UnicodeString HeaderData = {read=GetHeaderData, write=SetHeaderData};
	__property System::UnicodeString FooterData = {read=GetFooterData, write=SetfooterData};
	__property TZSheetFooterHeader* Header = {read=FHeader};
	__property TZSheetFooterHeader* Footer = {read=FFooter};
	__property TZSheetFooterHeader* EvenHeader = {read=FEvenHeader};
	__property TZSheetFooterHeader* EvenFooter = {read=FEvenFooter};
	__property Graphics::TColor HeaderBGColor = {read=FHeaderBGColor, write=FHeaderBGColor, default=-16777211};
	__property Graphics::TColor FooterBGColor = {read=FFooterBGColor, write=FFooterBGColor, default=-16777211};
	__property int ScaleToPercent = {read=FScaleToPercent, write=FScaleToPercent, default=100};
	__property int ScaleToPages = {read=FScaleToPages, write=FScaleToPages, default=1};
	__property TZSplitMode SplitVerticalMode = {read=FSplitVerticalMode, write=FSplitVerticalMode, default=0};
	__property TZSplitMode SplitHorizontalMode = {read=FSplitHorizontalMode, write=FSplitHorizontalMode, default=0};
	__property int SplitVerticalValue = {read=FSplitVerticalValue, write=FSplitVerticalValue, nodefault};
	__property int SplitHorizontalValue = {read=FSplitHorizontalValue, write=FSplitHorizontalValue, nodefault};
};


#pragma option push -b-
enum TZCondition { ZCFIsTrueFormula, ZCFCellContentIsBetween, ZCFCellContentIsNotBetween, ZCFCellContentOperator, ZCFNumberValue, ZCFString, ZCFBoolTrue, ZCFBoolFalse, ZCFFormula, ZCFContainsText, ZCFNotContainsText, ZCFBeginsWithText, ZCFEndsWithText, ZCFCellIsEmpty, ZCFDuplicate, ZCFUnique, ZCFAboveAverage, ZCFBellowAverage, ZCFAboveEqualAverage, ZCFBelowEqualAverage, ZCFTopElements, ZCFBottomElements, ZCFTopPercent, ZCFBottomPercent, ZCFIsError, ZCFIsNoError };
#pragma option pop

#pragma option push -b-
enum TZConditionalOperator { ZCFOpGT, ZCFOpLT, ZCFOpGTE, ZCFOpLTE, ZCFOpEqual, ZCFOpNotEqual };
#pragma option pop

class DELPHICLASS TZConditionalStyleItem;
class PASCALIMPLEMENTATION TZConditionalStyleItem : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	TZCondition FCondition;
	TZConditionalOperator FConditionOperator;
	System::UnicodeString FValue1;
	System::UnicodeString FValue2;
	int FApplyStyleID;
	int FBaseCellPageIndex;
	int FBaseCellRowIndex;
	int FBaseCellColumnIndex;
	
public:
	__fastcall virtual TZConditionalStyleItem(void);
	void __fastcall Clear(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	__property int ApplyStyleID = {read=FApplyStyleID, write=FApplyStyleID, nodefault};
	__property int BaseCellColumnIndex = {read=FBaseCellColumnIndex, write=FBaseCellColumnIndex, nodefault};
	__property int BaseCellPageIndex = {read=FBaseCellPageIndex, write=FBaseCellPageIndex, nodefault};
	__property int BaseCellRowIndex = {read=FBaseCellRowIndex, write=FBaseCellRowIndex, nodefault};
	__property TZCondition Condition = {read=FCondition, write=FCondition, nodefault};
	__property TZConditionalOperator ConditionOperator = {read=FConditionOperator, write=FConditionOperator, nodefault};
	__property System::UnicodeString Value1 = {read=FValue1, write=FValue1};
	__property System::UnicodeString Value2 = {read=FValue2, write=FValue2};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZConditionalStyleItem(void) { }
	
};


class DELPHICLASS TZConditionalAreaItem;
class PASCALIMPLEMENTATION TZConditionalAreaItem : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	int FRow;
	int FColumn;
	int FWidth;
	int FHeight;
	void __fastcall SetRow(int Value);
	void __fastcall SetColumn(int Value);
	void __fastcall SetWidth(int Value);
	void __fastcall SetHeight(int Value);
	
public:
	__fastcall virtual TZConditionalAreaItem(void)/* overload */;
	__fastcall virtual TZConditionalAreaItem(int ColumnNum, int RowNum, int AreaWidth, int AreaHeight)/* overload */;
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	__property int Row = {read=FRow, write=SetRow, nodefault};
	__property int Column = {read=FColumn, write=SetColumn, nodefault};
	__property int Width = {read=FWidth, write=SetWidth, nodefault};
	__property int Height = {read=FHeight, write=SetHeight, nodefault};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZConditionalAreaItem(void) { }
	
};


class DELPHICLASS TZConditionalAreas;
class PASCALIMPLEMENTATION TZConditionalAreas : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	typedef System::DynamicArray<TZConditionalAreaItem*> _TZConditionalAreas__1;
	
	
public:
	TZConditionalAreaItem* operator[](int num) { return Items[num]; }
	
private:
	int FCount;
	_TZConditionalAreas__1 FItems;
	void __fastcall SetCount(int Value);
	TZConditionalAreaItem* __fastcall GetItem(int num);
	void __fastcall SetItem(int num, TZConditionalAreaItem* Value);
	
public:
	__fastcall virtual TZConditionalAreas(void);
	__fastcall virtual ~TZConditionalAreas(void);
	TZConditionalAreaItem* __fastcall Add(void)/* overload */;
	TZConditionalAreaItem* __fastcall Add(int ColumnNum, int RowNum, int AreaWidth, int AreaHeight)/* overload */;
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	void __fastcall Delete(int num);
	bool __fastcall IsCellInArea(int ColumnNum, int RowNum);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	__property int Count = {read=FCount, write=SetCount, nodefault};
	__property TZConditionalAreaItem* Items[int num] = {read=GetItem, write=SetItem/*, default*/};
};


class DELPHICLASS TZConditionalStyle;
class PASCALIMPLEMENTATION TZConditionalStyle : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	typedef System::DynamicArray<TZConditionalStyleItem*> _TZConditionalStyle__1;
	
	
public:
	TZConditionalStyleItem* operator[](int num) { return Items[num]; }
	
private:
	int FCount;
	int FMaxCount;
	TZConditionalAreas* FAreas;
	_TZConditionalStyle__1 FConditions;
	TZConditionalStyleItem* __fastcall GetItem(int num);
	void __fastcall SetItem(int num, TZConditionalStyleItem* Value);
	void __fastcall SetCount(int value);
	void __fastcall SetAreas(TZConditionalAreas* Value);
	
public:
	__fastcall virtual TZConditionalStyle(void);
	__fastcall virtual ~TZConditionalStyle(void);
	TZConditionalStyleItem* __fastcall Add(void)/* overload */;
	TZConditionalStyleItem* __fastcall Add(TZConditionalStyleItem* StyleItem)/* overload */;
	void __fastcall Delete(int num);
	void __fastcall Insert(int num)/* overload */;
	void __fastcall Insert(int num, TZConditionalStyleItem* StyleItem)/* overload */;
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	__property TZConditionalAreas* Areas = {read=FAreas, write=SetAreas};
	__property int Count = {read=FCount, write=SetCount, nodefault};
	__property TZConditionalStyleItem* Items[int num] = {read=GetItem, write=SetItem/*, default*/};
};


class DELPHICLASS TZConditionalFormatting;
class PASCALIMPLEMENTATION TZConditionalFormatting : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	typedef System::DynamicArray<TZConditionalStyle*> _TZConditionalFormatting__1;
	
	
public:
	TZConditionalStyle* operator[](int num) { return Items[num]; }
	
private:
	_TZConditionalFormatting__1 FStyles;
	int FCount;
	void __fastcall SetCount(int Value);
	TZConditionalStyle* __fastcall GetItem(int num);
	void __fastcall SetItem(int num, TZConditionalStyle* Value);
	
public:
	__fastcall virtual TZConditionalFormatting(void);
	__fastcall virtual ~TZConditionalFormatting(void);
	TZConditionalStyle* __fastcall Add(void)/* overload */;
	TZConditionalStyle* __fastcall Add(TZConditionalStyle* Style)/* overload */;
	TZConditionalStyle* __fastcall Add(int ColumnNum, int RowNum, int AreaWidth, int AreaHeight)/* overload */;
	void __fastcall Clear(void);
	bool __fastcall Delete(int num);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual bool __fastcall IsEqual(Classes::TPersistent* Source);
	__property int Count = {read=FCount, write=SetCount, nodefault};
	__property TZConditionalStyle* Items[int num] = {read=GetItem, write=SetItem/*, default*/};
};


class DELPHICLASS TZEXMLSS;
class PASCALIMPLEMENTATION TZSheet : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	typedef System::DynamicArray<TZCellColumn> _TZSheet__1;
	
	typedef System::DynamicArray<TZRowOptions*> _TZSheet__2;
	
	typedef System::DynamicArray<TZColOptions*> _TZSheet__3;
	
	
private:
	TZEXMLSS* FStore;
	_TZSheet__1 FCells;
	_TZSheet__2 FRows;
	_TZSheet__3 FColumns;
	System::UnicodeString FTitle;
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
	TZConditionalFormatting* FConditionalFormatting;
	void __fastcall SetConditionalFormatting(TZConditionalFormatting* Value);
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
	__property TZCell* Cell[int ACol][int ARow] = {read=GetCell, write=SetCell/*, default*/};
	__property bool Protect = {read=FProtect, write=FProtect, default=0};
	__property Graphics::TColor TabColor = {read=FTabColor, write=FTabColor, default=-16777211};
	__property System::UnicodeString Title = {read=FTitle, write=FTitle};
	__property int RowCount = {read=GetRowCount, write=SetRowCount, nodefault};
	__property bool RightToLeft = {read=FRightToLeft, write=FRightToLeft, default=0};
	__property int ColCount = {read=GetColCount, write=SetColCount, nodefault};
	__property TZMergeCells* MergeCells = {read=FMergeCells, write=FMergeCells};
	__property TZSheetOptions* SheetOptions = {read=GetSheetOptions, write=SetSheetOptions};
	__property bool Selected = {read=FSelected, write=FSelected, nodefault};
	__property TZSheetPrintTitles* RowsToRepeat = {read=FPrintRows, write=SetPrintRows};
	__property TZSheetPrintTitles* ColsToRepeat = {read=FPrintCols, write=SetPrintCols};
	__property TZConditionalFormatting* ConditionalFormatting = {read=FConditionalFormatting, write=SetConditionalFormatting};
};


class DELPHICLASS TZSheets;
class PASCALIMPLEMENTATION TZSheets : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	typedef System::DynamicArray<TZSheet*> _TZSheets__1;
	
	
public:
	TZSheet* operator[](int num) { return Sheet[num]; }
	
private:
	TZEXMLSS* FStore;
	_TZSheets__1 FSheets;
	int FCount;
	void __fastcall SetSheetCount(const int Value);
	void __fastcall SetSheet(int num, const TZSheet* Value);
	TZSheet* __fastcall GetSheet(int num);
	
public:
	__fastcall virtual TZSheets(TZEXMLSS* AStore);
	__fastcall virtual ~TZSheets(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	__property int Count = {read=FCount, write=SetSheetCount, nodefault};
	__property TZSheet* Sheet[int num] = {read=GetSheet, write=SetSheet/*, default*/};
};


class DELPHICLASS TZEXMLDocumentProperties;
class PASCALIMPLEMENTATION TZEXMLDocumentProperties : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	System::UnicodeString FAuthor;
	System::UnicodeString FLastAuthor;
	System::TDateTime FCreated;
	System::UnicodeString FCompany;
	System::UnicodeString FVersion;
	System::Word FWindowHeight;
	System::Word FWindowWidth;
	int FWindowTopX;
	int FWindowTopY;
	bool FModeR1C1;
	
protected:
	void __fastcall SetAuthor(const System::UnicodeString Value);
	void __fastcall SetLastAuthor(const System::UnicodeString Value);
	void __fastcall SetCompany(const System::UnicodeString Value);
	void __fastcall SetVersion(const System::UnicodeString Value);
	
public:
	__fastcall virtual TZEXMLDocumentProperties(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	
__published:
	__property System::UnicodeString Author = {read=FAuthor, write=SetAuthor};
	__property System::UnicodeString LastAuthor = {read=FLastAuthor, write=SetLastAuthor};
	__property System::TDateTime Created = {read=FCreated, write=FCreated};
	__property System::UnicodeString Company = {read=FCompany, write=SetCompany};
	__property System::UnicodeString Version = {read=FVersion, write=SetVersion};
	__property bool ModeR1C1 = {read=FModeR1C1, write=FModeR1C1, default=0};
	__property System::Word WindowHeight = {read=FWindowHeight, write=FWindowHeight, default=20000};
	__property System::Word WindowWidth = {read=FWindowWidth, write=FWindowWidth, default=20000};
	__property int WindowTopX = {read=FWindowTopX, write=FWindowTopX, default=150};
	__property int WindowTopY = {read=FWindowTopY, write=FWindowTopY, default=150};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TZEXMLDocumentProperties(void) { }
	
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
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	void __fastcall GetPixelSize(HWND hdc);
	__property TZSheets* Sheets = {read=FSheets, write=FSheets};
	
__published:
	__property TZStyles* Styles = {read=FStyles, write=FStyles};
	__property TZSheetOptions* DefaultSheetOptions = {read=GetDefaultSheetOptions, write=SetDefaultSheetOptions};
	__property TZEXMLDocumentProperties* DocumentProperties = {read=FDocumentProperties, write=FDocumentProperties};
	__property double HorPixelSize = {read=FHorPixelSize, write=SetHorPixelSize};
	__property double VertPixelSize = {read=FVertPixelSize, write=SetVertPixelSize};
};


//-- var, const, procedure ---------------------------------------------------
extern PACKAGE double _PointToMM;
static const TZCellType ZEAnsiString [[deprecated("use ZEString")]] = (TZCellType)(3);
extern PACKAGE void __fastcall CorrectStrForXML(const System::UnicodeString St, System::UnicodeString &Corrected, bool &UseXMLNS);
extern PACKAGE System::UnicodeString __fastcall ColorToHTMLHex(Graphics::TColor Color);
extern PACKAGE Graphics::TColor __fastcall HTMLHexToColor(System::UnicodeString value);
extern PACKAGE double __fastcall PixelToPoint(int inPixel, double PixelSizeMM = 2.650000E-01);
extern PACKAGE int __fastcall PointToPixel(double inPoint, double PixelSizeMM = 2.650000E-01);
extern PACKAGE double __fastcall PointToMM(double inPoint);
extern PACKAGE double __fastcall MMToPoint(double inMM);
extern PACKAGE System::UnicodeString __fastcall HAlToStr(TZHorizontalAlignment HA);
extern PACKAGE System::UnicodeString __fastcall VAlToStr(TZVerticalAlignment VA);
extern PACKAGE System::UnicodeString __fastcall ZBorderTypeToStr(TZBorderType ZB);
extern PACKAGE System::UnicodeString __fastcall ZCellPatternToStr(TZCellPattern pp);
extern PACKAGE System::UnicodeString __fastcall ZCellTypeToStr(TZCellType pp);
extern PACKAGE TZHorizontalAlignment __fastcall StrToHal(System::UnicodeString Value);
extern PACKAGE TZVerticalAlignment __fastcall StrToVAl(System::UnicodeString Value);
extern PACKAGE TZBorderType __fastcall StrToZBorderType(System::UnicodeString Value);
extern PACKAGE TZCellPattern __fastcall StrToZCellPattern(System::UnicodeString Value);
extern PACKAGE TZCellType __fastcall StrToZCellType(System::UnicodeString Value);
extern PACKAGE void __fastcall Register(void);

}	/* namespace Zexmlss */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE)
using namespace Zexmlss;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZexmlssHPP
