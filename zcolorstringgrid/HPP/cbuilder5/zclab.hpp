// Borland C++ Builder
// Copyright (c) 1995, 1999 by Borland International
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zclab.pas' rev: 5.00

#ifndef zclabHPP
#define zclabHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <Menus.hpp>	// Pascal unit
#include <Messages.hpp>	// Pascal unit
#include <Windows.hpp>	// Pascal unit
#include <zcftext.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <Controls.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zclab
{
//-- type declarations -------------------------------------------------------
class DELPHICLASS TZCLabel;
class PASCALIMPLEMENTATION TZCLabel : public Controls::TGraphicControl 
{
	typedef Controls::TGraphicControl inherited;
	
private:
	Byte FAlignmentVertical;
	Byte FAlignmentHorizontal;
	bool FAutoSizeHeight;
	bool FAutoSizeWidth;
	bool FAutoSizeGrowOnly;
	Byte FIndentVert;
	Byte FIndentHor;
	int FLineSpacing;
	int FRotate;
	bool FSymbolWrap;
	bool FTransparent;
	bool FWordWrap;
	Classes::TNotifyEvent FOnMouseLeave;
	Classes::TNotifyEvent FOnMouseEnter;
	void __fastcall SetAlignmentHorizontal(Byte Value);
	void __fastcall SetAlignmentVertical(Byte Value);
	void __fastcall SetAutoSizeHeight(bool Value);
	void __fastcall SetAutoSizeWidth(bool Value);
	void __fastcall SetAutoSizeGrowOnly(bool Value);
	void __fastcall SetIndentVert(Byte Value);
	void __fastcall SetIndentHor(Byte Value);
	void __fastcall SetLineSpacing(int Value);
	void __fastcall SetRotate(int Value);
	void __fastcall SetSymbolWrap(bool Value);
	void __fastcall SetTransparent(bool Value);
	void __fastcall SetWordWrap(bool Value);
	MESSAGE void __fastcall CMTextChanged(Messages::TMessage &Message);
	HIDESBASE MESSAGE void __fastcall CMMouseEnter(Messages::TMessage &Message);
	HIDESBASE MESSAGE void __fastcall CMMouseLeave(Messages::TMessage &Message);
	
protected:
	virtual void __fastcall Paint(void);
	
public:
	__fastcall virtual TZCLabel(Classes::TComponent* AOwner);
	void __fastcall DrawTextOn(Graphics::TCanvas* ACanvas, AnsiString AText, Windows::TRect &ARect, bool 
		CalcOnly, bool ClipArea)/* overload */;
	void __fastcall DrawTextOn(Graphics::TCanvas* ACanvas, AnsiString AText, Windows::TRect &ARect, Graphics::TColor 
		AColor, bool CalcOnly, bool ClipArea)/* overload */;
	__property Canvas ;
	
__published:
	__property Byte AlignmentVertical = {read=FAlignmentVertical, write=SetAlignmentVertical, default=0
		};
	__property Byte AlignmentHorizontal = {read=FAlignmentHorizontal, write=SetAlignmentHorizontal, default=0
		};
	__property bool AutoSizeHeight = {read=FAutoSizeHeight, write=SetAutoSizeHeight, default=0};
	__property bool AutoSizeWidth = {read=FAutoSizeWidth, write=SetAutoSizeWidth, default=0};
	__property bool AutoSizeGrowOnly = {read=FAutoSizeGrowOnly, write=SetAutoSizeGrowOnly, default=0};
	__property Caption ;
	__property Color ;
	__property Font ;
	__property Byte IndentVert = {read=FIndentVert, write=SetIndentVert, default=0};
	__property Byte IndentHor = {read=FIndentHor, write=SetIndentHor, default=3};
	__property int LineSpacing = {read=FLineSpacing, write=SetLineSpacing, default=0};
	__property int Rotate = {read=FRotate, write=SetRotate, default=0};
	__property bool SymbolWrap = {read=FSymbolWrap, write=SetSymbolWrap, default=0};
	__property bool Transparent = {read=FTransparent, write=SetTransparent, default=1};
	__property bool WordWrap = {read=FWordWrap, write=SetWordWrap, default=1};
	__property DragCursor ;
	__property DragKind ;
	__property DragMode ;
	__property Enabled ;
	__property ShowHint ;
	__property ParentColor ;
	__property ParentFont ;
	__property ParentShowHint ;
	__property PopupMenu ;
	__property Visible ;
	__property OnClick ;
	__property OnDblClick ;
	__property OnDragDrop ;
	__property OnDragOver ;
	__property OnEndDrag ;
	__property OnMouseDown ;
	__property OnMouseMove ;
	__property OnMouseUp ;
	__property Classes::TNotifyEvent OnMouseEnter = {read=FOnMouseEnter, write=FOnMouseEnter};
	__property Classes::TNotifyEvent OnMouseLeave = {read=FOnMouseLeave, write=FOnMouseLeave};
	__property OnContextPopup ;
	__property OnResize ;
	__property OnStartDrag ;
public:
	#pragma option push -w-inl
	/* TGraphicControl.Destroy */ inline __fastcall virtual ~TZCLabel(void) { }
	#pragma option pop
	
};


//-- var, const, procedure ---------------------------------------------------
extern PACKAGE void __fastcall Register(void);

}	/* namespace Zclab */
#if !defined(NO_IMPLICIT_NAMESPACE_USE)
using namespace Zclab;
#endif
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zclab
