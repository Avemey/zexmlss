// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zipusejcl7z.pas' rev: 6.00

#ifndef zipusejcl7zHPP
#define zipusejcl7zHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <JclCompression.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zipusejcl7z
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
extern PACKAGE bool __fastcall ZEZipUnpackJCL7z(AnsiString ZipName, AnsiString PathName);
extern PACKAGE bool __fastcall ZEZipPackJCL7z(AnsiString ZipName, AnsiString PathName);

}	/* namespace Zipusejcl7z */
using namespace Zipusejcl7z;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zipusejcl7z
