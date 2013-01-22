// Borland C++ Builder
// Copyright (c) 1995, 1999 by Borland International
// All rights reserved

// (DO NOT EDIT: machine generated header) 'ZColorStringGrid.pas' rev: 5.00

#ifndef ZColorStringGridHPP
#define ZColorStringGridHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <zcftext.hpp>	// Pascal unit
#include <StdCtrls.hpp>	// Pascal unit
#include <Grids.hpp>	// Pascal unit
#include <Dialogs.hpp>	// Pascal unit
#include <Forms.hpp>	// Pascal unit
#include <Controls.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <Messages.hpp>	// Pascal unit
#include <Windows.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zcolorstringgrid
{
//-- type declarations -------------------------------------------------------
#pragma option push -b-
enum TVerticalAlignment { vaTop, vaCenter, vaBottom };
#pragma option pop

struct TPenBrush
{
	Graphics::TColor PColor;
	int Width;
	Graphics::TPenMode Mode;
	Graphics::TPenStyle PStyle;
	Graphics::TColor BColor;
	Graphics::TBrushStyle BStyle;
} ;

#pragma option push -b-
enum TBorderCellStyle { sgLowered, sgRaised, sgNone };
#pragma option pop

typedef void __fastcall (__closure *TDrawMergeCellEvent)(System::TObject* Sender, int ACol, int ARow
	, const Windows::TRect &Rect, Grids::TGridDrawState State, Graphics::TCanvas* &CellCanvas);

class DELPHICLASS TSelectColor;
class DELPHICLASS TZColorStringGrid;
typedef DynamicArray<DynamicArray<TCellStyle* > >  ZColorStringGrid__9;

class DELPHICLASS TMergeCells;
typedef DynamicArray<Windows::TRect >  ZColorStringGrid__5;

class PASCALIMPLEMENTATION TMergeCells : public System::TObject 
{
	typedef System::TObject inherited;
	
private:
	TZColorStringGrid* FGrid;
	int FCount;
	DynamicArray<Windows::TRect >  FMergeArea;
	Windows::TRect __fastcall Get_Item(int Num);
	
public:
	__fastcall virtual TMergeCells(TZColorStringGrid* AGrid);
	__fastcall virtual ~TMergeCells(void);
	void __fastcall Clear(void);
	Byte __fastcall AddRect(const Windows::TRect &Rct);
	Byte __fastcall AddRectXY(int x1, int y1, int x2, int y2);
	bool __fastcall DeleteItem(int num);
	int __fastcall InLeftTopCorner(int ACol, int ARow);
	int __fastcall InMergeRange(int ACol, int ARow);
	int __fastcall GetWidthArea(int num);
	int __fastcall GetHeightArea(int num);
	Grids::TGridRect __fastcall GetSelectedArea(bool SetSelected);
	Windows::TPoint __fastcall GetMergeYY(int ARow);
	__property int Count = {read=FCount, nodefault};
	__property Windows::TRect Items[int Num] = {read=Get_Item};
};


class DELPHICLASS TCellStyle;
class PASCALIMPLEMENTATION TCellStyle : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZColorStringGrid* FGrid;
	Graphics::TFont* FFont;
	Graphics::TColor FBGColor;
	Classes::TAlignment FHorizontalAlignment;
	TVerticalAlignment FVerticalAlignment;
	bool FWordWrap;
	bool FSizingHeight;
	bool FSizingWidth;
	TBorderCellStyle FBorderCellStyle;
	Byte FIndentH;
	Byte FIndentV;
	int FRotate;
	void __fastcall SetFont(const Graphics::TFont* Value);
	void __fastcall SetBGColor(const Graphics::TColor Value);
	void __fastcall SetHAlignment(const Classes::TAlignment Value);
	void __fastcall SetVAlignment(const TVerticalAlignment Value);
	void __fastcall SetBorderCellStyle(const TBorderCellStyle Value);
	void __fastcall SetIndentV(const Byte Value);
	void __fastcall SetindentH(const Byte Value);
	void __fastcall SetRotate(const int Value);
	
public:
	__fastcall virtual TCellStyle(TZColorStringGrid* AGrid);
	__fastcall virtual ~TCellStyle(void);
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	virtual void __fastcall AssignTo(Classes::TPersistent* Source);
	
__published:
	__property Graphics::TFont* Font = {read=FFont, write=SetFont};
	__property Graphics::TColor BGColor = {read=FBGColor, write=SetBGColor, nodefault};
	__property TBorderCellStyle BorderCellStyle = {read=FBorderCellStyle, write=SetBorderCellStyle, default=2
		};
	__property Byte IndentH = {read=FIndentH, write=SetindentH, default=2};
	__property Byte IndentV = {read=FIndentV, write=SetIndentV, default=0};
	__property TVerticalAlignment VerticalAlignment = {read=FVerticalAlignment, write=SetVAlignment, default=1
		};
	__property Classes::TAlignment HorizontalAlignment = {read=FHorizontalAlignment, write=SetHAlignment
		, default=0};
	__property int Rotate = {read=FRotate, write=SetRotate, default=0};
	__property bool SizingHeight = {read=FSizingHeight, write=FSizingHeight, default=0};
	__property bool SizingWidth = {read=FSizingWidth, write=FSizingWidth, default=0};
	__property bool WordWrap = {read=FWordWrap, write=FWordWrap, default=0};
};


class DELPHICLASS TLineDesign;
class PASCALIMPLEMENTATION TLineDesign : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZColorStringGrid* FGrid;
	Graphics::TColor FLineColor;
	Graphics::TColor FLineUpColor;
	Graphics::TColor FLineDownColor;
	void __fastcall SetLineColor(const Graphics::TColor Value);
	void __fastcall SetLineUpColor(const Graphics::TColor Value);
	void __fastcall SetLineDownColor(const Graphics::TColor Value);
	
public:
	__fastcall virtual TLineDesign(TZColorStringGrid* AGrid);
	
__published:
	__property Graphics::TColor LineColor = {read=FLineColor, write=SetLineColor, default=8421504};
	__property Graphics::TColor LineUpColor = {read=FLineUpColor, write=SetLineUpColor, default=8421504
		};
	__property Graphics::TColor LineDownColor = {read=FLineDownColor, write=SetLineDownColor, default=0
		};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TLineDesign(void) { }
	#pragma option pop
	
};


class DELPHICLASS TZInplaceEditor;
class PASCALIMPLEMENTATION TZInplaceEditor : public Stdctrls::TMemo 
{
	typedef Stdctrls::TMemo inherited;
	
private:
	TZColorStringGrid* FGrid;
	int FExEn;
	
protected:
	DYNAMIC void __fastcall DoEnter(void);
	DYNAMIC void __fastcall DoExit(void);
	DYNAMIC void __fastcall Change(void);
	DYNAMIC void __fastcall DblClick(void);
	DYNAMIC void __fastcall KeyDown(Word &Key, Classes::TShiftState Shift);
	DYNAMIC void __fastcall KeyPress(char &Key);
	DYNAMIC void __fastcall KeyUp(Word &Key, Classes::TShiftState Shift);
	
public:
	__fastcall virtual TZInplaceEditor(Classes::TComponent* AOwner);
	DYNAMIC bool __fastcall DoMouseWheelDown(Classes::TShiftState Shift, const Windows::TPoint &MousePos
		);
	DYNAMIC bool __fastcall DoMouseWheelUp(Classes::TShiftState Shift, const Windows::TPoint &MousePos)
		;
public:
	#pragma option push -w-inl
	/* TCustomMemo.Destroy */ inline __fastcall virtual ~TZInplaceEditor(void) { }
	#pragma option pop
	
public:
	#pragma option push -w-inl
	/* TWinControl.CreateParented */ inline __fastcall TZInplaceEditor(HWND ParentWindow) : Stdctrls::TMemo(
		ParentWindow) { }
	#pragma option pop
	
};


class DELPHICLASS TInplaceEditorOptions;
class PASCALIMPLEMENTATION TInplaceEditorOptions : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZColorStringGrid* FGrid;
	Graphics::TColor FFontColor;
	Graphics::TColor FBGColor;
	Forms::TFormBorderStyle FBorderStyle;
	Classes::TAlignment FAlignment;
	bool FWordWrap;
	bool FUseCellStyle;
	void __fastcall SetFontColor(const Graphics::TColor Value);
	void __fastcall SetBGColor(const Graphics::TColor Value);
	void __fastcall SetBorderStyle(const Forms::TBorderStyle Value);
	void __fastcall SetAlignment(const Classes::TAlignment Value);
	void __fastcall SetWordWrap(const bool Value);
	void __fastcall SetUseCellStyle(const bool Value);
	
public:
	__fastcall virtual TInplaceEditorOptions(TZColorStringGrid* AGrid);
	
__published:
	__property Graphics::TColor FontColor = {read=FFontColor, write=SetFontColor, default=0};
	__property Graphics::TColor BGColor = {read=FBGColor, write=SetBGColor, default=16777215};
	__property Forms::TBorderStyle BorderStyle = {read=FBorderStyle, write=SetBorderStyle, default=0};
	__property Classes::TAlignment Alignment = {read=FAlignment, write=SetAlignment, default=0};
	__property bool WordWrap = {read=FWordWrap, write=SetWordWrap, default=1};
	__property bool UseCellStyle = {read=FUseCellStyle, write=SetUseCellStyle, default=1};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TInplaceEditorOptions(void) { }
	#pragma option pop
	
};


class PASCALIMPLEMENTATION TZColorStringGrid : public Grids::TStringGrid 
{
	typedef Grids::TStringGrid inherited;
	
private:
	Grids::TDrawCellEvent FOnBeforeTextDrawCell;
	TDrawMergeCellEvent FOnBeforeTextDrawCellMerge;
	TDrawMergeCellEvent FOnDrawCellMerge;
	DynamicArray<DynamicArray<TCellStyle* > >  FCellStyle;
	TMergeCells* FMergeCells;
	TCellStyle* FDefaultFixedCellParam;
	TCellStyle* FDefaultCellParam;
	TLineDesign* FLineDesign;
	bool FWordWrap;
	bool FUseCellWordWrap;
	Grids::TGridOptions FOptions;
	bool FgoEditing;
	bool FSizingHeight;
	bool FSizingWidth;
	bool FUseCellSizingHeight;
	bool FUseCellSizingWidth;
	TSelectColor* FSelectedColors;
	bool FUseBitMap;
	Windows::TRect FRctBitmap;
	Graphics::TBitmap* FBmp;
	Windows::TRect FRctRectangle;
	TPenBrush FPenBrush;
	TZInplaceEditor* FZInplaceEditor;
	TInplaceEditorOptions* FInplaceEditorOptions;
	bool FDontDraw;
	Grids::TGridCoord FMouseXY;
	int FLastCell[2][7];
	void __fastcall SetCellStyle(int ACol, int ARow, const TCellStyle* value);
	TCellStyle* __fastcall GetCellStyle(int ACol, int ARow);
	void __fastcall SetCellStyleRow(int ARow, bool fixedCol, const TCellStyle* value);
	void __fastcall SetCellStyleCol(int ACol, bool fixedRow, const TCellStyle* value);
	void __fastcall SetDefaultCellParam(const TCellStyle* Value);
	TCellStyle* __fastcall GetDefaultCellParam(void);
	void __fastcall SetDefaultFixedCellParam(const TCellStyle* Value);
	TCellStyle* __fastcall GetDefaultFixedCellParam(void);
	HIDESBASE void __fastcall SetRowCount(const int Value);
	int __fastcall GetRowCount(void);
	HIDESBASE void __fastcall SetColCount(const int Value);
	int __fastcall GetColCount(void);
	HIDESBASE void __fastcall SetFixedColor(const Graphics::TColor Value);
	Graphics::TColor __fastcall GetFixedColor(void);
	HIDESBASE void __fastcall SetFixedCols(const int value);
	int __fastcall GetFixedCols(void);
	HIDESBASE void __fastcall SetFixedRows(const int value);
	int __fastcall GetFixedRows(void);
	HIDESBASE void __fastcall SetCells(int ACol, int ARow, const AnsiString Value);
	HIDESBASE AnsiString __fastcall GetCells(int ACol, int ARow);
	HIDESBASE void __fastcall Initialize(void);
	void __fastcall StyleRowMove(int FromIndex, int ToIndex);
	void __fastcall GetCalcRect(Windows::TRect &ARect, int &frmparams, int ACol, int ARow, int &NumMergeArea
		);
	void __fastcall SetCellCount(int oldColCount, int newColCount, int oldRowCount, int newRowCount);
	HIDESBASE void __fastcall SetEditorMode(bool Value);
	bool __fastcall GetEditorMode(void);
	void __fastcall SavePenBrush(Graphics::TCanvas* &CellCanvas);
	void __fastcall RestorePenBrush(Graphics::TCanvas* &CellCanvas);
	Grids::TGridOptions __fastcall GetOptions(void);
	HIDESBASE void __fastcall SetOptions(Grids::TGridOptions Value);
	void __fastcall GetParamsForText(int ACol, int ARow, int &frmparams, int &AlignmentHorizontal, int 
		&AlignmentVertical, bool &ZWordWrap);
	
protected:
	virtual void __fastcall CalcCellSize(int ACol, int ARow);
	DYNAMIC void __fastcall ColumnMoved(int FromIndex, int ToIndex);
	virtual void __fastcall DrawCell(int ACol, int ARow, const Windows::TRect &ARect, Grids::TGridDrawState 
		AState);
	DYNAMIC void __fastcall DblClick(void);
	virtual void __fastcall DeleteColumn(int ACol);
	virtual void __fastcall DeleteRow(int ARow);
	DYNAMIC void __fastcall DoEnter(void);
	DYNAMIC void __fastcall DoExit(void);
	DYNAMIC bool __fastcall DoMouseWheelDown(Classes::TShiftState Shift, const Windows::TPoint &MousePos
		);
	DYNAMIC bool __fastcall DoMouseWheelUp(Classes::TShiftState Shift, const Windows::TPoint &MousePos)
		;
	DYNAMIC void __fastcall KeyDown(Word &Key, Classes::TShiftState Shift);
	DYNAMIC void __fastcall KeyPress(char &Key);
	DYNAMIC void __fastcall MouseDown(Controls::TMouseButton Button, Classes::TShiftState Shift, int X, 
		int Y);
	DYNAMIC void __fastcall MouseMove(Classes::TShiftState Shift, int X, int Y);
	virtual void __fastcall Paint(void);
	DYNAMIC void __fastcall RowMoved(int FromIndex, int ToIndex);
	virtual void __fastcall RowSelectYY(Word key);
	DYNAMIC void __fastcall SetEditText(int ACol, int ARow, const AnsiString Value);
	virtual bool __fastcall SelectCell(int ACol, int ARow);
	HIDESBASE virtual void __fastcall ShowEditor(void);
	virtual void __fastcall ShowEditorFirst(void);
	DYNAMIC void __fastcall TopLeftChanged(void);
	HIDESBASE virtual void __fastcall HideEditor(void);
	
public:
	__fastcall virtual TZColorStringGrid(Classes::TComponent* AOwner);
	__fastcall virtual ~TZColorStringGrid(void);
	__property bool DontDraw = {read=FDontDraw, write=FDontDraw, nodefault};
	__property AnsiString Cells[int ACol][int ARow] = {read=GetCells, write=SetCells};
	__property TCellStyle* CellStyle[int ACol][int ARow] = {read=GetCellStyle, write=SetCellStyle};
	__property TCellStyle* CellStyleRow[int ARow][bool fixedCol] = {write=SetCellStyleRow};
	__property TCellStyle* CellStyleCol[int ACol][bool fixedRow] = {write=SetCellStyleCol};
	__property bool EditorMode = {read=GetEditorMode, write=SetEditorMode, nodefault};
	__property TMergeCells* MergeCells = {read=FMergeCells, write=FMergeCells};
	__property TZInplaceEditor* ZInplaceEditor = {read=FZInplaceEditor, write=FZInplaceEditor};
	
__published:
	__property int ColCount = {read=GetColCount, write=SetColCount, default=5};
	__property TCellStyle* DefaultCellStyle = {read=GetDefaultCellParam, write=SetDefaultCellParam};
	__property TCellStyle* DefaultFixedCellStyle = {read=GetDefaultFixedCellParam, write=SetDefaultFixedCellParam
		};
	__property Graphics::TColor FixedColor = {read=GetFixedColor, write=SetFixedColor, nodefault};
	__property int FixedCols = {read=GetFixedCols, write=SetFixedCols, default=1};
	__property int FixedRows = {read=GetFixedRows, write=SetFixedRows, default=1};
	__property TLineDesign* LineDesign = {read=FLineDesign, write=FLineDesign};
	__property Grids::TDrawCellEvent OnBeforeTextDrawCell = {read=FOnBeforeTextDrawCell, write=FOnBeforeTextDrawCell
		};
	__property TDrawMergeCellEvent OnBeforeTextDrawMergeCell = {read=FOnBeforeTextDrawCellMerge, write=
		FOnBeforeTextDrawCellMerge};
	__property TDrawMergeCellEvent OnDrawMergeCell = {read=FOnDrawCellMerge, write=FOnDrawCellMerge};
	__property Grids::TGridOptions Options = {read=GetOptions, write=SetOptions, default=31};
	__property int RowCount = {read=GetRowCount, write=SetRowCount, default=5};
	__property TSelectColor* SelectedColors = {read=FSelectedColors, write=FSelectedColors};
	__property bool SizingHeight = {read=FSizingHeight, write=FSizingHeight, default=0};
	__property bool SizingWidth = {read=FSizingWidth, write=FSizingWidth, default=0};
	__property bool UseCellSizingHeight = {read=FUseCellSizingHeight, write=FUseCellSizingHeight, default=0
		};
	__property bool UseCellSizingWidth = {read=FUseCellSizingWidth, write=FUseCellSizingWidth, default=0
		};
	__property bool UseCellWordWrap = {read=FUseCellWordWrap, write=FUseCellWordWrap, default=0};
	__property bool WordWrap = {read=FWordWrap, write=FWordWrap, default=0};
	__property TInplaceEditorOptions* InplaceEditorOptions = {read=FInplaceEditorOptions, write=FInplaceEditorOptions
		};
public:
	#pragma option push -w-inl
	/* TWinControl.CreateParented */ inline __fastcall TZColorStringGrid(HWND ParentWindow) : Grids::TStringGrid(
		ParentWindow) { }
	#pragma option pop
	
};


class PASCALIMPLEMENTATION TSelectColor : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
private:
	TZColorStringGrid* FGrid;
	Graphics::TColor FBGColor;
	Graphics::TColor FFontColor;
	bool FColoredSelect;
	bool FUseFocusRect;
	void __fastcall SetBGColor(const Graphics::TColor Value);
	void __fastcall SetColoredSelect(const bool Value);
	void __fastcall SetUseFocusRect(const bool Value);
	
public:
	__fastcall virtual TSelectColor(TZColorStringGrid* AGrid);
	
__published:
	__property Graphics::TColor BGColor = {read=FBGColor, write=SetBGColor, default=-2147483635};
	__property Graphics::TColor FontColor = {read=FFontColor, write=FFontColor, default=16777215};
	__property bool ColoredSelect = {read=FColoredSelect, write=SetColoredSelect, default=1};
	__property bool UseFocusRect = {read=FUseFocusRect, write=SetUseFocusRect, default=1};
public:
	#pragma option push -w-inl
	/* TPersistent.Destroy */ inline __fastcall virtual ~TSelectColor(void) { }
	#pragma option pop
	
};


typedef DynamicArray<TCellStyle* >  TCellLine;

//-- var, const, procedure ---------------------------------------------------
extern PACKAGE void __fastcall Register(void);

}	/* namespace Zcolorstringgrid */
#if !defined(NO_IMPLICIT_NAMESPACE_USE)
using namespace Zcolorstringgrid;
#endif
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZColorStringGrid
