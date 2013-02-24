// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zeZippy.pas' rev: 6.00

#ifndef zeZippyHPP
#define zeZippyHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <SysUtils.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zezippy
{
//-- type declarations -------------------------------------------------------
#pragma option push -b-
enum TZxZipGenState { zgsAccepting, zgsFlushing, zgsSealed };
#pragma option pop

class DELPHICLASS EZxZipGen;
class PASCALIMPLEMENTATION EZxZipGen : public Sysutils::Exception 
{
	typedef Sysutils::Exception inherited;
	
public:
	#pragma option push -w-inl
	/* Exception.Create */ inline __fastcall EZxZipGen(const AnsiString Msg) : Sysutils::Exception(Msg) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateFmt */ inline __fastcall EZxZipGen(const AnsiString Msg, const System::TVarRec * Args, const int Args_Size) : Sysutils::Exception(Msg, Args, Args_Size) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateRes */ inline __fastcall EZxZipGen(int Ident)/* overload */ : Sysutils::Exception(Ident) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateResFmt */ inline __fastcall EZxZipGen(int Ident, const System::TVarRec * Args, const int Args_Size)/* overload */ : Sysutils::Exception(Ident, Args, Args_Size) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateHelp */ inline __fastcall EZxZipGen(const AnsiString Msg, int AHelpContext) : Sysutils::Exception(Msg, AHelpContext) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateFmtHelp */ inline __fastcall EZxZipGen(const AnsiString Msg, const System::TVarRec * Args, const int Args_Size, int AHelpContext) : Sysutils::Exception(Msg, Args, Args_Size, AHelpContext) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateResHelp */ inline __fastcall EZxZipGen(int Ident, int AHelpContext)/* overload */ : Sysutils::Exception(Ident, AHelpContext) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateResFmtHelp */ inline __fastcall EZxZipGen(System::PResStringRec ResStringRec, const System::TVarRec * Args, const int Args_Size, int AHelpContext)/* overload */ : Sysutils::Exception(ResStringRec, Args, Args_Size, AHelpContext) { }
	#pragma option pop
	
public:
	#pragma option push -w-inl
	/* TObject.Destroy */ inline __fastcall virtual ~EZxZipGen(void) { }
	#pragma option pop
	
};


typedef TMetaClass*CZxZipGens;

class DELPHICLASS TZxZipGen;
class DELPHICLASS TStringList;
class PASCALIMPLEMENTATION TStringList : public Classes::TStringList 
{
	typedef Classes::TStringList inherited;
	
public:
	bool OwnsObjects;
	__fastcall TStringList(void)/* overload */;
	__fastcall TStringList(bool OwnsObjects)/* overload */;
	__fastcall virtual ~TStringList(void);
	virtual void __fastcall Delete(int Index);
	virtual void __fastcall Clear(void);
};


class PASCALIMPLEMENTATION TZxZipGen : public System::TObject 
{
	typedef System::TObject inherited;
	
public:
	__fastcall virtual TZxZipGen(const AnsiString ZipFile);
	Classes::TStream* __fastcall NewStream(const AnsiString RelativeName);
	void __fastcall SealStream(const Classes::TStream* Data);
	void __fastcall AbortAndDelete(void);
	void __fastcall SaveAndSeal(void);
	
protected:
	TStringList* FActiveStreams;
	TStringList* FSealedStreams;
	virtual void __fastcall DoAbortAndDelete(void);
	virtual void __fastcall DoSaveAndSeal(void) = 0 ;
	virtual Classes::TStream* __fastcall DoNewStream(const AnsiString RelativeName);
	virtual bool __fastcall DoSealStream(const Classes::TStream* Data, const AnsiString RelativeName) = 0 ;
	AnsiString FFileName;
	TZxZipGenState FState;
	void __fastcall RequireState(const TZxZipGenState st);
	void __fastcall ChangeState(const TZxZipGenState _From, const TZxZipGenState _To);
	
public:
	__property TZxZipGenState State = {read=FState, nodefault};
	__property AnsiString ZipFileName = {read=FFileName};
	virtual void __fastcall AfterConstruction(void);
	virtual void __fastcall BeforeDestruction(void);
	
protected:
	/*         class method */ static void __fastcall RegisterZipGen(TMetaClass* vmt, const TMetaClass* zgc);
	/*         class method */ static void __fastcall UnRegisterZipGen(TMetaClass* vmt, const TMetaClass* zgc);
	
public:
	/*         class method */ static void __fastcall Register(TMetaClass* vmt);
	/*         class method */ static void __fastcall UnRegister(TMetaClass* vmt);
	/*         class method */ static TMetaClass* __fastcall QueryZipGen(TMetaClass* vmt, const int idx = 0x0);
	/*         class method */ static TMetaClass* __fastcall QueryDummyZipGen(TMetaClass* vmt);
public:
	#pragma option push -w-inl
	/* TObject.Destroy */ inline __fastcall virtual ~TZxZipGen(void) { }
	#pragma option pop
	
};


//-- var, const, procedure ---------------------------------------------------

}	/* namespace Zezippy */
using namespace Zezippy;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zeZippy
