// CodeGear C++Builder
// Copyright (c) 1995, 2011 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'ZColorStringGrid.pas' rev: 23.00 (Win32)

#ifndef ZcolorstringgridHPP
#define ZcolorstringgridHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <Winapi.Windows.hpp>	// Pascal unit
#include <Winapi.Messages.hpp>	// Pascal unit
#include <System.SysUtils.hpp>	// Pascal unit
#include <System.Classes.hpp>	// Pascal unit
#include <Vcl.Graphics.hpp>	// Pascal unit
#include <Vcl.Controls.hpp>	// Pascal unit
#include <Vcl.Forms.hpp>	// Pascal unit
#include <Vcl.Dialogs.hpp>	// Pascal unit
#include <Vcl.Grids.hpp>	// Pascal unit
#include <Vcl.StdCtrls.hpp>	// Pascal unit
#include <zcftext.hpp>	// Pascal unit
#include <System.UITypes.hpp>	// Pascal unit
#include <System.Types.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zcolorstringgrid
{
//-- type declarations -------------------------------------------------------
#pragma option push -b-
enum TVerticalAlignment : unsigned char { vaTop, vaCenter, vaBottom };
#pragma option pop

struct DECLSPEC_DRECORD TPenBrush
{
	
public:
	System::Uitypes::TColor PColor;
	int Width;
	Vcl::Graphics::TPenMode Mode;
	Vcl::Graphics::TPenStyle PStyle;
	System::Uitypes::TColor BColor;
	Vcl::Graphics::TBrushStyle BStyle;
};


#pragma option push -b-
enum TBorderCellStyle : unsigned char { sgLowered, sgRaised, sgNone };
#pragma option pop

typedef void __fastcall (__closure *TDrawMergeCellEvent)(System::TObject* Sender, int ACol, int ARow, const System::Types::TRect &Rect, Vcl::Grids::TGridDrawState State, Vcl::Graphics::TCanvas* &CellCanvas);

class DELPHICLASS TSelectColor;
class DELPHICLASS TZColorStringGrid;
#pragma pack(push,4)
class PASCALIMPLEMENTATION TSelectColor : public System::Classes::TPersistent
{
	typedef System::Classes::TPersistent inherited;
	
private:
	TZColorStringGrid* FGrid;
	System::Uitypes::TColor FBGColor;
	System::Uitypes::TColor FFontColor;
	bool FColoredSelect;
	bool FUseFocusRect;
	void __fastcall SetBGColor(const System::Uitypes::TColor Value);
	void __fastcall SetColoredSelect(const bool Value);
	void __fastcall SetUseFocusRect(const bool Value);
	
public:
	__fastcall virtual TSelectColor(TZColorStringGrid* AGrid);
	
__published:
	__property System::Uitypes::TColor BGColor = {read=FBGColor, write=SetBGColor, default=-16777203};
	__property System::Uitypes::TColor FontColor = {read=FFontColor, write=FFontColor, default=16777215};
	__property bool ColoredSelect = {read=FColoredSelect, write=SetColoredSelect, default=1};
	__property bool UseFocusRect = {read=FUseFocusRect, write=SetUseFocusRect, default=1};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TSelectColor(void) { }
	
};

#pragma pack(pop)

class DELPHICLASS TLineDesign;
#pragma pack(push,4)
class PASCALIMPLEMENTATION TLineDesign : public System::Classes::TPersistent
{
	typedef System::Classes::TPersistent inherited;
	
private:
	TZColorStringGrid* FGrid;
	System::Uitypes::TColor FLineColor;
	System::Uitypes::TColor FLineUpColor;
	System::Uitypes::TColor FLineDownColor;
	void __fastcall SetLineColor(const System::Uitypes::TColor Value);
	void __fastcall SetLineUpColor(const System::Uitypes::TColor Value);
	void __fastcall SetLineDownColor(const System::Uitypes::TColor Value);
	
public:
	__fastcall virtual TLineDesign(TZColorStringGrid* AGrid);
	
__published:
	__property System::Uitypes::TColor LineColor = {read=FLineColor, write=SetLineColor, default=8421504};
	__property System::Uitypes::TColor LineUpColor = {read=FLineUpColor, write=SetLineUpColor, default=8421504};
	__property System::Uitypes::TColor LineDownColor = {read=FLineDownColor, write=SetLineDownColor, default=0};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TLineDesign(void) { }
	
};

#pragma pack(pop)

class DELPHICLASS TCellStyle;
#pragma pack(push,4)
class PASCALIMPLEMENTATION TCellStyle : public System::Classes::TPersistent
{
	typedef System::Classes::TPersistent inherited;
	
private:
	TZColorStringGrid* FGrid;
	Vcl::Graphics::TFont* FFont;
	System::Uitypes::TColor FBGColor;
	System::Classes::TAlignment FHorizontalAlignment;
	TVerticalAlignment FVerticalAlignment;
	bool FWordWrap;
	bool FSizingHeight;
	bool FSizingWidth;
	TBorderCellStyle FBorderCellStyle;
	System::Byte FIndentH;
	System::Byte FIndentV;
	int FRotate;
	void __fastcall SetFont(const Vcl::Graphics::TFont* Value);
	void __fastcall SetBGColor(const System::Uitypes::TColor Value);
	void __fastcall SetHAlignment(const System::Classes::TAlignment Value);
	void __fastcall SetVAlignment(const TVerticalAlignment Value);
	void __fastcall SetBorderCellStyle(const TBorderCellStyle Value);
	void __fastcall SetIndentV(const System::Byte Value);
	void __fastcall SetindentH(const System::Byte Value);
	void __fastcall SetRotate(const int Value);
	
public:
	__fastcall virtual TCellStyle(TZColorStringGrid* AGrid);
	__fastcall virtual ~TCellStyle(void);
	virtual void __fastcall Assign(System::Classes::TPersistent* Source);
	virtual void __fastcall AssignTo(System::Classes::TPersistent* Source);
	
__published:
	__property Vcl::Graphics::TFont* Font = {read=FFont, write=SetFont};
	__property System::Uitypes::TColor BGColor = {read=FBGColor, write=SetBGColor, nodefault};
	__property TBorderCellStyle BorderCellStyle = {read=FBorderCellStyle, write=SetBorderCellStyle, default=2};
	__property System::Byte IndentH = {read=FIndentH, write=SetindentH, default=2};
	__property System::Byte IndentV = {read=FIndentV, write=SetIndentV, default=0};
	__property TVerticalAlignment VerticalAlignment = {read=FVerticalAlignment, write=SetVAlignment, default=1};
	__property System::Classes::TAlignment HorizontalAlignment = {read=FHorizontalAlignment, write=SetHAlignment, default=0};
	__property int Rotate = {read=FRotate, write=SetRotate, default=0};
	__property bool SizingHeight = {read=FSizingHeight, write=FSizingHeight, default=0};
	__property bool SizingWidth = {read=FSizingWidth, write=FSizingWidth, default=0};
	__property bool WordWrap = {read=FWordWrap, write=FWordWrap, default=0};
};

#pragma pack(pop)

typedef System::DynamicArray<TCellStyle*> TCellLine;

class DELPHICLASS TMergeCells;
#pragma pack(push,4)
class PASCALIMPLEMENTATION TMergeCells : public System::TObject
{
	typedef System::TObject inherited;
	
private:
	typedef System::DynamicArray<System::Types::TRect> _TMergeCells__1;
	
	
private:
	TZColorStringGrid* FGrid;
	int FCount;
	_TMergeCells__1 FMergeArea;
	System::Types::TRect __fastcall Get_Item(int Num);
	
public:
	__fastcall virtual TMergeCells(TZColorStringGrid* AGrid);
	__fastcall virtual ~TMergeCells(void);
	void __fastcall Clear(void);
	System::Byte __fastcall AddRect(const System::Types::TRect &Rct);
	System::Byte __fastcall AddRectXY(int x1, int y1, int x2, int y2);
	bool __fastcall DeleteItem(int num);
	int __fastcall InLeftTopCorner(int ACol, int ARow);
	int __fastcall InMergeRange(int ACol, int ARow);
	int __fastcall GetWidthArea(int num);
	int __fastcall GetHeightArea(int num);
	Vcl::Grids::TGridRect __fastcall GetSelectedArea(bool SetSelected);
	System::Types::TPoint __fastcall GetMergeYY(int ARow);
	__property int Count = {read=FCount, nodefault};
	__property System::Types::TRect Items[int Num] = {read=Get_Item};
};

#pragma pack(pop)

class DELPHICLASS TZInplaceEditor;
class PASCALIMPLEMENTATION TZInplaceEditor : public Vcl::Stdctrls::TMemo
{
	typedef Vcl::Stdctrls::TMemo inherited;
	
private:
	TZColorStringGrid* FGrid;
	int FExEn;
	
protected:
	DYNAMIC void __fastcall DoEnter(void);
	DYNAMIC void __fastcall DoExit(void);
	DYNAMIC void __fastcall Change(void);
	DYNAMIC void __fastcall DblClick(void);
	DYNAMIC void __fastcall KeyDown(System::Word &Key, System::Classes::TShiftState Shift);
	DYNAMIC void __fastcall KeyPress(System::WideChar &Key);
	DYNAMIC void __fastcall KeyUp(System::Word &Key, System::Classes::TShiftState Shift);
	
public:
	__fastcall virtual TZInplaceEditor(System::Classes::TComponent* AOwner);
	DYNAMIC bool __fastcall DoMouseWheelDown(System::Classes::TShiftState Shift, const System::Types::TPoint &MousePos);
	DYNAMIC bool __fastcall DoMouseWheelUp(System::Classes::TShiftState Shift, const System::Types::TPoint &MousePos);
public:
	/* TCustomMemo.Destroy */ inline __fastcall virtual ~TZInplaceEditor(void) { }
	
public:
	/* TWinControl.CreateParented */ inline __fastcall TZInplaceEditor(HWND ParentWindow) : Vcl::Stdctrls::TMemo(ParentWindow) { }
	
};


class DELPHICLASS TInplaceEditorOptions;
#pragma pack(push,4)
class PASCALIMPLEMENTATION TInplaceEditorOptions : public System::Classes::TPersistent
{
	typedef System::Classes::TPersistent inherited;
	
private:
	TZColorStringGrid* FGrid;
	System::Uitypes::TColor FFontColor;
	System::Uitypes::TColor FBGColor;
	Vcl::Forms::TFormBorderStyle FBorderStyle;
	System::Classes::TAlignment FAlignment;
	bool FWordWrap;
	bool FUseCellStyle;
	void __fastcall SetFontColor(const System::Uitypes::TColor Value);
	void __fastcall SetBGColor(const System::Uitypes::TColor Value);
	void __fastcall SetBorderStyle(const Vcl::Forms::TBorderStyle Value);
	void __fastcall SetAlignment(const System::Classes::TAlignment Value);
	void __fastcall SetWordWrap(const bool Value);
	void __fastcall SetUseCellStyle(const bool Value);
	
public:
	__fastcall virtual TInplaceEditorOptions(TZColorStringGrid* AGrid);
	
__published:
	__property System::Uitypes::TColor FontColor = {read=FFontColor, write=SetFontColor, default=0};
	__property System::Uitypes::TColor BGColor = {read=FBGColor, write=SetBGColor, default=16777215};
	__property Vcl::Forms::TBorderStyle BorderStyle = {read=FBorderStyle, write=SetBorderStyle, default=0};
	__property System::Classes::TAlignment Alignment = {read=FAlignment, write=SetAlignment, default=0};
	__property bool WordWrap = {read=FWordWrap, write=SetWordWrap, default=1};
	__property bool UseCellStyle = {read=FUseCellStyle, write=SetUseCellStyle, default=1};
public:
	/* TPersistent.Destroy */ inline __fastcall virtual ~TInplaceEditorOptions(void) { }
	
};

#pragma pack(pop)

class PASCALIMPLEMENTATION TZColorStringGrid : public Vcl::Grids::TStringGrid
{
	typedef Vcl::Grids::TStringGrid inherited;
	
private:
	typedef System::DynamicArray<TCellLine> _TZColorStringGrid__1;
	
	
private:
	Vcl::Grids::TDrawCellEvent FOnBeforeTextDrawCell;
	TDrawMergeCellEvent FOnBeforeTextDrawCellMerge;
	TDrawMergeCellEvent FOnDrawCellMerge;
	_TZColorStringGrid__1 FCellStyle;
	TMergeCells* FMergeCells;
	TCellStyle* FDefaultFixedCellParam;
	TCellStyle* FDefaultCellParam;
	TLineDesign* FLineDesign;
	bool FWordWrap;
	bool FUseCellWordWrap;
	Vcl::Grids::TGridOptions FOptions;
	bool FgoEditing;
	bool FSizingHeight;
	bool FSizingWidth;
	bool FUseCellSizingHeight;
	bool FUseCellSizingWidth;
	TSelectColor* FSelectedColors;
	bool FUseBitMap;
	System::Types::TRect FRctBitmap;
	Vcl::Graphics::TBitmap* FBmp;
	System::Types::TRect FRctRectangle;
	TPenBrush FPenBrush;
	TZInplaceEditor* FZInplaceEditor;
	TInplaceEditorOptions* FInplaceEditorOptions;
	bool FDontDraw;
	Vcl::Grids::TGridCoord FMouseXY;
	System::StaticArray<System::StaticArray<int, 7>, 2> FLastCell;
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
	HIDESBASE void __fastcall SetFixedColor(const System::Uitypes::TColor Value);
	System::Uitypes::TColor __fastcall GetFixedColor(void);
	HIDESBASE void __fastcall SetFixedCols(const int value);
	int __fastcall GetFixedCols(void);
	HIDESBASE void __fastcall SetFixedRows(const int value);
	int __fastcall GetFixedRows(void);
	HIDESBASE void __fastcall SetCells(int ACol, int ARow, const System::UnicodeString Value);
	HIDESBASE System::UnicodeString __fastcall GetCells(int ACol, int ARow);
	HIDESBASE void __fastcall Initialize(void);
	void __fastcall StyleRowMove(int FromIndex, int ToIndex);
	void __fastcall GetCalcRect(System::Types::TRect &ARect, int &frmparams, int ACol, int ARow, int &NumMergeArea);
	void __fastcall SetCellCount(int oldColCount, int newColCount, int oldRowCount, int newRowCount);
	HIDESBASE void __fastcall SetEditorMode(bool Value);
	bool __fastcall GetEditorMode(void);
	void __fastcall SavePenBrush(Vcl::Graphics::TCanvas* &CellCanvas);
	void __fastcall RestorePenBrush(Vcl::Graphics::TCanvas* &CellCanvas);
	Vcl::Grids::TGridOptions __fastcall GetOptions(void);
	HIDESBASE void __fastcall SetOptions(Vcl::Grids::TGridOptions Value);
	void __fastcall GetParamsForText(int ACol, int ARow, int &frmparams, int &AlignmentHorizontal, int &AlignmentVertical, bool &ZWordWrap);
	
protected:
	virtual void __fastcall CalcCellSize(int ACol, int ARow);
	DYNAMIC void __fastcall ColumnMoved(int FromIndex, int ToIndex);
	virtual void __fastcall DrawCell(int ACol, int ARow, const System::Types::TRect &ARect, Vcl::Grids::TGridDrawState AState);
	DYNAMIC void __fastcall DblClick(void);
	virtual void __fastcall DeleteColumn(int ACol);
	virtual void __fastcall DeleteRow(int ARow);
	DYNAMIC void __fastcall DoEnter(void);
	DYNAMIC void __fastcall DoExit(void);
	DYNAMIC bool __fastcall DoMouseWheelDown(System::Classes::TShiftState Shift, const System::Types::TPoint &MousePos);
	DYNAMIC bool __fastcall DoMouseWheelUp(System::Classes::TShiftState Shift, const System::Types::TPoint &MousePos);
	DYNAMIC void __fastcall KeyDown(System::Word &Key, System::Classes::TShiftState Shift);
	DYNAMIC void __fastcall KeyPress(System::WideChar &Key);
	DYNAMIC void __fastcall MouseDown(System::Uitypes::TMouseButton Button, System::Classes::TShiftState Shift, int X, int Y);
	DYNAMIC void __fastcall MouseMove(System::Classes::TShiftState Shift, int X, int Y);
	virtual void __fastcall Paint(void);
	DYNAMIC void __fastcall RowMoved(int FromIndex, int ToIndex);
	virtual void __fastcall RowSelectYY(System::Word key);
	DYNAMIC void __fastcall SetEditText(int ACol, int ARow, const System::UnicodeString Value);
	virtual bool __fastcall SelectCell(int ACol, int ARow);
	HIDESBASE virtual void __fastcall ShowEditor(void);
	virtual void __fastcall ShowEditorFirst(void);
	DYNAMIC void __fastcall TopLeftChanged(void);
	HIDESBASE virtual void __fastcall HideEditor(void);
	
public:
	__fastcall virtual TZColorStringGrid(System::Classes::TComponent* AOwner);
	__fastcall virtual ~TZColorStringGrid(void);
	__property bool DontDraw = {read=FDontDraw, write=FDontDraw, nodefault};
	__property System::UnicodeString Cells[int ACol][int ARow] = {read=GetCells, write=SetCells};
	__property TCellStyle* CellStyle[int ACol][int ARow] = {read=GetCellStyle, write=SetCellStyle};
	__property TCellStyle* CellStyleRow[int ARow][bool fixedCol] = {write=SetCellStyleRow};
	__property TCellStyle* CellStyleCol[int ACol][bool fixedRow] = {write=SetCellStyleCol};
	__property bool EditorMode = {read=GetEditorMode, write=SetEditorMode, nodefault};
	__property TMergeCells* MergeCells = {read=FMergeCells, write=FMergeCells};
	__property TZInplaceEditor* ZInplaceEditor = {read=FZInplaceEditor, write=FZInplaceEditor};
	
__published:
	__property int ColCount = {read=GetColCount, write=SetColCount, default=5};
	__property TCellStyle* DefaultCellStyle = {read=GetDefaultCellParam, write=SetDefaultCellParam};
	__property TCellStyle* DefaultFixedCellStyle = {read=GetDefaultFixedCellParam, write=SetDefaultFixedCellParam};
	__property System::Uitypes::TColor FixedColor = {read=GetFixedColor, write=SetFixedColor, nodefault};
	__property int FixedCols = {read=GetFixedCols, write=SetFixedCols, default=1};
	__property int FixedRows = {read=GetFixedRows, write=SetFixedRows, default=1};
	__property TLineDesign* LineDesign = {read=FLineDesign, write=FLineDesign};
	__property Vcl::Grids::TDrawCellEvent OnBeforeTextDrawCell = {read=FOnBeforeTextDrawCell, write=FOnBeforeTextDrawCell};
	__property TDrawMergeCellEvent OnBeforeTextDrawMergeCell = {read=FOnBeforeTextDrawCellMerge, write=FOnBeforeTextDrawCellMerge};
	__property TDrawMergeCellEvent OnDrawMergeCell = {read=FOnDrawCellMerge, write=FOnDrawCellMerge};
	__property Vcl::Grids::TGridOptions Options = {read=GetOptions, write=SetOptions, default=31};
	__property int RowCount = {read=GetRowCount, write=SetRowCount, default=5};
	__property TSelectColor* SelectedColors = {read=FSelectedColors, write=FSelectedColors};
	__property bool SizingHeight = {read=FSizingHeight, write=FSizingHeight, default=0};
	__property bool SizingWidth = {read=FSizingWidth, write=FSizingWidth, default=0};
	__property bool UseCellSizingHeight = {read=FUseCellSizingHeight, write=FUseCellSizingHeight, default=0};
	__property bool UseCellSizingWidth = {read=FUseCellSizingWidth, write=FUseCellSizingWidth, default=0};
	__property bool UseCellWordWrap = {read=FUseCellWordWrap, write=FUseCellWordWrap, default=0};
	__property bool WordWrap = {read=FWordWrap, write=FWordWrap, default=0};
	__property TInplaceEditorOptions* InplaceEditorOptions = {read=FInplaceEditorOptions, write=FInplaceEditorOptions};
public:
	/* TWinControl.CreateParented */ inline __fastcall TZColorStringGrid(HWND ParentWindow) : Vcl::Grids::TStringGrid(ParentWindow) { }
	
};


//-- var, const, procedure ---------------------------------------------------
extern PACKAGE void __fastcall Register(void);

}	/* namespace Zcolorstringgrid */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE) && !defined(NO_USING_NAMESPACE_ZCOLORSTRINGGRID)
using namespace Zcolorstringgrid;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZcolorstringgridHPP
