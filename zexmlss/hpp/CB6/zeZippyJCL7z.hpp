// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zeZippyJCL7z.pas' rev: 6.00

#ifndef zeZippyJCL7zHPP
#define zeZippyJCL7zHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <JclCompression.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <zeZippy.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zezippyjcl7z
{
//-- type declarations -------------------------------------------------------
class DELPHICLASS TZxZipJcl7z;
class PASCALIMPLEMENTATION TZxZipJcl7z : public Zezippy::TZxZipGen 
{
	typedef Zezippy::TZxZipGen inherited;
	
public:
	__fastcall virtual TZxZipJcl7z(const AnsiString ZipFile);
	virtual void __fastcall BeforeDestruction(void);
	
protected:
	Jclcompression::TJclCompressArchive* FZ;
	virtual void __fastcall DoSaveAndSeal(void);
	virtual bool __fastcall DoSealStream(const Classes::TStream* Data, const AnsiString RelName);
public:
	#pragma option push -w-inl
	/* TObject.Destroy */ inline __fastcall virtual ~TZxZipJcl7z(void) { }
	#pragma option pop
	
};


//-- var, const, procedure ---------------------------------------------------

}	/* namespace Zezippyjcl7z */
using namespace Zezippyjcl7z;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zeZippyJCL7z
