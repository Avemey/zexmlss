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
#include <Classes.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <zesavecommon.hpp>	// Pascal unit
#include <Graphics.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zeodfs
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
extern PACKAGE int __fastcall SaveXmlssToODFSPath(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString PathName, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM = "")/* overload */;

}	/* namespace Zeodfs */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE)
using namespace Zeodfs;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZeodfsHPP
