// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zearchhelper.pas' rev: 6.00

#ifndef zearchhelperHPP
#define zearchhelperHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zearchhelper
{
//-- type declarations -------------------------------------------------------
//-- var, const, procedure ---------------------------------------------------
extern PACKAGE AnsiString __fastcall ZEGetTempDir();
extern PACKAGE bool __fastcall ZECreateUniqueTmpDir(const AnsiString ADir, AnsiString &retTmpDir);
extern PACKAGE bool __fastcall ZEDelTree(AnsiString ADir);

}	/* namespace Zearchhelper */
using namespace Zearchhelper;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zearchhelper
