// CodeGear C++Builder
// Copyright (c) 1995, 2011 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zcftext.pas' rev: 23.00 (Win32)

#ifndef ZcftextHPP
#define ZcftextHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <Winapi.Windows.hpp>	// Pascal unit
#include <Vcl.Graphics.hpp>	// Pascal unit
#include <System.SysUtils.hpp>	// Pascal unit
#include <Vcl.Controls.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zcftext
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
static const System::Int8 ZCF_SIZINGW = System::Int8(0x1);
static const System::Int8 ZCF_SIZINGH = System::Int8(0x2);
static const System::Int8 ZCF_CALCONLY = System::Int8(0x4);
static const System::Int8 ZCF_SYMBOL_WRAP = System::Int8(0x8);
static const System::Int8 ZCF_NO_TRANSPARENT = System::Int8(0x10);
static const System::Int8 ZCF_NO_CLIP = System::Int8(0x20);
static const System::Int8 ZCF_DOTTED = System::Int8(0x40);
static const System::Byte ZCF_DOTTED_LINE = System::Byte(0x80);
extern PACKAGE bool __fastcall IsFontTrueType(Vcl::Graphics::TFont* Fnt);
extern PACKAGE bool __fastcall ZCWriteTextFormatted(Vcl::Graphics::TCanvas* canvas, System::UnicodeString text, Vcl::Graphics::TFont* fnt, System::Uitypes::TColor fntColor, int HA, int VA, bool WordWrap, System::Types::TRect &Rct, System::Byte IH, System::Byte IV, int Params, int LineSpacing, int Rotate = 0x0)/* overload */;
extern PACKAGE bool __fastcall ZCWriteTextFormatted(Vcl::Graphics::TCanvas* canvas, System::UnicodeString text, Vcl::Graphics::TFont* fnt, int HA, int VA, bool WordWrap, System::Types::TRect &Rct, System::Byte IH, System::Byte IV, int Params, int LineSpacing, int Rotate = 0x0)/* overload */;

}	/* namespace Zcftext */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE) && !defined(NO_USING_NAMESPACE_ZCFTEXT)
using namespace Zcftext;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZcftextHPP
