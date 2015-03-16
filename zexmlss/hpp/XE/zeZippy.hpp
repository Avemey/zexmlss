// CodeGear C++Builder
// Copyright (c) 1995, 2010 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zeZippy.pas' rev: 22.00

#ifndef ZezippyHPP
#define ZezippyHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit

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
	/* Exception.Create */ inline __fastcall EZxZipGen(const System::UnicodeString Msg) : Sysutils::Exception(Msg) { }
	/* Exception.CreateFmt */ inline __fastcall EZxZipGen(const System::UnicodeString Msg, System::TVarRec const *Args, const int Args_Size) : Sysutils::Exception(Msg, Args, Args_Size) { }
	/* Exception.CreateRes */ inline __fastcall EZxZipGen(int Ident)/* overload */ : Sysutils::Exception(Ident) { }
	/* Exception.CreateResFmt */ inline __fastcall EZxZipGen(int Ident, System::TVarRec const *Args, const int Args_Size)/* overload */ : Sysutils::Exception(Ident, Args, Args_Size) { }
	/* Exception.CreateHelp */ inline __fastcall EZxZipGen(const System::UnicodeString Msg, int AHelpContext) : Sysutils::Exception(Msg, AHelpContext) { }
	/* Exception.CreateFmtHelp */ inline __fastcall EZxZipGen(const System::UnicodeString Msg, System::TVarRec const *Args, const int Args_Size, int AHelpContext) : Sysutils::Exception(Msg, Args, Args_Size, AHelpContext) { }
	/* Exception.CreateResHelp */ inline __fastcall EZxZipGen(int Ident, int AHelpContext)/* overload */ : Sysutils::Exception(Ident, AHelpContext) { }
	/* Exception.CreateResFmtHelp */ inline __fastcall EZxZipGen(System::PResStringRec ResStringRec, System::TVarRec const *Args, const int Args_Size, int AHelpContext)/* overload */ : Sysutils::Exception(ResStringRec, Args, Args_Size, AHelpContext) { }
	/* Exception.Destroy */ inline __fastcall virtual ~EZxZipGen(void) { }
	
};


typedef TMetaClass* CZxZipGens;

class DELPHICLASS TZxZipGen;
class PASCALIMPLEMENTATION TZxZipGen : public System::TObject
{
	typedef System::TObject inherited;
	
public:
	__fastcall virtual TZxZipGen(const Sysutils::TFileName ZipFile);
	Classes::TStream* __fastcall NewStream(const Sysutils::TFileName RelativeName);
	void __fastcall SealStream(const Classes::TStream* Data);
	void __fastcall AbortAndDelete(void);
	void __fastcall SaveAndSeal(void);
	
protected:
	Classes::TStringList* FActiveStreams;
	Classes::TStringList* FSealedStreams;
	virtual void __fastcall DoAbortAndDelete(void);
	virtual void __fastcall DoSaveAndSeal(void) = 0 ;
	virtual Classes::TStream* __fastcall DoNewStream(const Sysutils::TFileName RelativeName);
	virtual bool __fastcall DoSealStream(const Classes::TStream* Data, const Sysutils::TFileName RelativeName) = 0 ;
	Sysutils::TFileName FFileName;
	TZxZipGenState FState;
	void __fastcall RequireState(const TZxZipGenState st);
	void __fastcall ChangeState(const TZxZipGenState _From, const TZxZipGenState _To);
	
public:
	__property TZxZipGenState State = {read=FState, nodefault};
	__property Sysutils::TFileName ZipFileName = {read=FFileName};
	virtual void __fastcall AfterConstruction(void);
	virtual void __fastcall BeforeDestruction(void);
	
protected:
	__classmethod void __fastcall RegisterZipGen(const CZxZipGens zgc);
	__classmethod void __fastcall UnRegisterZipGen(const CZxZipGens zgc);
	
public:
	__classmethod void __fastcall Register();
	__classmethod void __fastcall UnRegister();
	__classmethod CZxZipGens __fastcall QueryZipGen(const int idx = 0x0);
	__classmethod CZxZipGens __fastcall QueryDummyZipGen();
public:
	/* TObject.Destroy */ inline __fastcall virtual ~TZxZipGen(void) { }
	
};


//-- var, const, procedure ---------------------------------------------------

}	/* namespace Zezippy */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE)
using namespace Zezippy;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZezippyHPP
