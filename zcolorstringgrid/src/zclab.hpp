// CodeGear C++Builder
// Copyright (c) 1995, 2011 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zclab.pas' rev: 23.00 (Win32)

#ifndef ZclabHPP
#define ZclabHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.Classes.hpp>	// Pascal unit
#include <System.SysUtils.hpp>	// Pascal unit
#include <Vcl.Controls.hpp>	// Pascal unit
#include <Vcl.Graphics.hpp>	// Pascal unit
#include <zcftext.hpp>	// Pascal unit
#include <Winapi.Windows.hpp>	// Pascal unit
#include <Winapi.Messages.hpp>	// Pascal unit
#include <System.Types.hpp>	// Pascal unit
#include <System.UITypes.hpp>	// Pascal unit
#include <Vcl.Menus.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zclab
{
//-- type declarations -------------------------------------------------------
class DELPHICLASS TZCLabel;
class PASCALIMPLEMENTATION TZCLabel : public Vcl::Controls::TGraphicControl
{
	typedef Vcl::Controls::TGraphicControl inherited;
	
private:
	System::Byte FAlignmentVertical;
	System::Byte FAlignmentHorizontal;
	bool FAutoSizeHeight;
	bool FAutoSizeWidth;
	bool FAutoSizeGrowOnly;
	System::Byte FIndentVert;
	System::Byte FIndentHor;
	int FLineSpacing;
	int FRotate;
	bool FSymbolWrap;
	bool FTransparent;
	bool FWordWrap;
	void __fastcall SetAlignmentHorizontal(System::Byte Value);
	void __fastcall SetAlignmentVertical(System::Byte Value);
	void __fastcall SetAutoSizeHeight(bool Value);
	void __fastcall SetAutoSizeWidth(bool Value);
	void __fastcall SetAutoSizeGrowOnly(bool Value);
	void __fastcall SetIndentVert(System::Byte Value);
	void __fastcall SetIndentHor(System::Byte Value);
	void __fastcall SetLineSpacing(int Value);
	void __fastcall SetRotate(int Value);
	void __fastcall SetSymbolWrap(bool Value);
	void __fastcall SetTransparent(bool Value);
	void __fastcall SetWordWrap(bool Value);
	MESSAGE void __fastcall CMTextChanged(Winapi::Messages::TMessage &Message);
	
protected:
	virtual void __fastcall Paint(void);
	
public:
	__fastcall virtual TZCLabel(System::Classes::TComponent* AOwner);
	void __fastcall DrawTextOn(Vcl::Graphics::TCanvas* ACanvas, System::UnicodeString AText, System::Types::TRect &ARect, bool CalcOnly, bool ClipArea = false)/* overload */;
	void __fastcall DrawTextOn(Vcl::Graphics::TCanvas* ACanvas, System::UnicodeString AText, System::Types::TRect &ARect, System::Uitypes::TColor AColor, bool CalcOnly, bool ClipArea = false)/* overload */;
	__property Canvas;
	
__published:
	__property System::Byte AlignmentVertical = {read=FAlignmentVertical, write=SetAlignmentVertical, default=0};
	__property System::Byte AlignmentHorizontal = {read=FAlignmentHorizontal, write=SetAlignmentHorizontal, default=0};
	__property bool AutoSizeHeight = {read=FAutoSizeHeight, write=SetAutoSizeHeight, default=0};
	__property bool AutoSizeWidth = {read=FAutoSizeWidth, write=SetAutoSizeWidth, default=0};
	__property bool AutoSizeGrowOnly = {read=FAutoSizeGrowOnly, write=SetAutoSizeGrowOnly, default=0};
	__property Caption;
	__property Color = {default=-16777211};
	__property Font;
	__property System::Byte IndentVert = {read=FIndentVert, write=SetIndentVert, default=0};
	__property System::Byte IndentHor = {read=FIndentHor, write=SetIndentHor, default=3};
	__property int LineSpacing = {read=FLineSpacing, write=SetLineSpacing, default=0};
	__property int Rotate = {read=FRotate, write=SetRotate, default=0};
	__property bool SymbolWrap = {read=FSymbolWrap, write=SetSymbolWrap, default=0};
	__property bool Transparent = {read=FTransparent, write=SetTransparent, default=1};
	__property bool WordWrap = {read=FWordWrap, write=SetWordWrap, default=1};
	__property DragCursor = {default=-12};
	__property DragKind = {default=0};
	__property DragMode = {default=0};
	__property Enabled = {default=1};
	__property ShowHint;
	__property ParentColor = {default=1};
	__property ParentFont = {default=1};
	__property ParentShowHint = {default=1};
	__property PopupMenu;
	__property Visible = {default=1};
	__property OnClick;
	__property OnDblClick;
	__property OnDragDrop;
	__property OnDragOver;
	__property OnEndDrag;
	__property OnMouseDown;
	__property OnMouseMove;
	__property OnMouseUp;
	__property OnMouseEnter;
	__property OnMouseLeave;
	__property OnContextPopup;
	__property OnResize;
	__property OnStartDrag;
public:
	/* TGraphicControl.Destroy */ inline __fastcall virtual ~TZCLabel(void) { }
	
};


//-- var, const, procedure ---------------------------------------------------
extern PACKAGE void __fastcall Register(void);

}	/* namespace Zclab */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE) && !defined(NO_USING_NAMESPACE_ZCLAB)
using namespace Zclab;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZclabHPP
