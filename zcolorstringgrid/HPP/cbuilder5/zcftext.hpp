// Borland C++ Builder
// Copyright (c) 1995, 1999 by Borland International
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zcftext.pas' rev: 5.00

#ifndef zcftextHPP
#define zcftextHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <Controls.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit
#include <Windows.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zcftext
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
static const Shortint ZCF_SIZINGW = 0x1;
static const Shortint ZCF_SIZINGH = 0x2;
static const Shortint ZCF_CALCONLY = 0x4;
static const Shortint ZCF_SYMBOL_WRAP = 0x8;
static const Shortint ZCF_NO_TRANSPARENT = 0x10;
static const Shortint ZCF_NO_CLIP = 0x20;
static const Shortint ZCF_DOTTED = 0x40;
static const Byte ZCF_DOTTED_LINE = 0x80;
extern PACKAGE bool __fastcall IsFontTrueType(Graphics::TFont* Fnt);
extern PACKAGE bool __fastcall ZCWriteTextFormatted(Graphics::TCanvas* canvas, AnsiString text, Graphics::TFont* 
	fnt, Graphics::TColor fntColor, int HA, int VA, bool WordWrap, Windows::TRect &Rct, Byte IH, Byte IV
	, int Params, int LineSpacing, int Rotate)/* overload */;
extern PACKAGE bool __fastcall ZCWriteTextFormatted(Graphics::TCanvas* canvas, AnsiString text, Graphics::TFont* 
	fnt, int HA, int VA, bool WordWrap, Windows::TRect &Rct, Byte IH, Byte IV, int Params, int LineSpacing
	, int Rotate)/* overload */;

}	/* namespace Zcftext */
#if !defined(NO_IMPLICIT_NAMESPACE_USE)
using namespace Zcftext;
#endif
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zcftext
