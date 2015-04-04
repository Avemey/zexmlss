// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zeSave.pas' rev: 6.00

#ifndef zeSaveHPP
#define zeSaveHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <Types.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <zeZippy.hpp>	// Pascal unit
#include <zsspxml.hpp>	// Pascal unit
#include <zexmlss.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zesave
{
//-- type declarations -------------------------------------------------------
#pragma pack(push, 4)
struct TZxPageInfo
{
	AnsiString name;
	int no;
} ;
#pragma pack(pop)

__interface IZXMLSSave;
typedef System::DelphiInterface<IZXMLSSave> _di_IZXMLSSave;
__interface IZXMLSSave  : public IInterface 
{
	
public:
	virtual _di_IZXMLSSave __fastcall ExportFormat(const AnsiString fmt) = 0 ;
	virtual _di_IZXMLSSave __fastcall As_(const AnsiString fmt) = 0 ;
	virtual _di_IZXMLSSave __fastcall ExportTo(const AnsiString fname) = 0 ;
	virtual _di_IZXMLSSave __fastcall To_(const AnsiString fname) = 0 ;
	virtual _di_IZXMLSSave __fastcall Pages(const TZxPageInfo * pages, const int pages_Size) = 0 /* overload */;
	virtual _di_IZXMLSSave __fastcall Pages(const int * numbers, const int numbers_Size) = 0 /* overload */;
	virtual _di_IZXMLSSave __fastcall Pages(const AnsiString * titles, const int titles_Size) = 0 /* overload */;
	virtual _di_IZXMLSSave __fastcall CharSet(const AnsiString cs) = 0 /* overload */;
	virtual _di_IZXMLSSave __fastcall CharSet(const Zsspxml::TAnsiToCPConverter converter) = 0 /* overload */;
	virtual _di_IZXMLSSave __fastcall CharSet(const AnsiString cs, const Zsspxml::TAnsiToCPConverter converter) = 0 /* overload */;
	virtual _di_IZXMLSSave __fastcall CharSet(const Word codepage) = 0 /* overload */;
	virtual _di_IZXMLSSave __fastcall BOM(const AnsiString Unicode_BOM) = 0 ;
	virtual _di_IZXMLSSave __fastcall ZipWith(const TMetaClass* ZipGenerator) = 0 ;
	virtual _di_IZXMLSSave __fastcall NoZip(void) = 0 ;
	virtual int __fastcall Save(void) = 0 /* overload */;
	virtual int __fastcall Save(const AnsiString FileName) = 0 /* overload */;
	virtual void __fastcall Discard(void) = 0 ;
};

__interface IZXMLSSaveImpl;
typedef System::DelphiInterface<IZXMLSSaveImpl> _di_IZXMLSSaveImpl;
__interface INTERFACE_UUID("{DAB9A318-ADD4-466C-AFE3-CE8623EC8977}") IZXMLSSaveImpl  : public IZXMLSSave 
{
	
public:
	virtual int __fastcall InternalSave(void) = 0 ;
};

typedef TMetaClass*CZXMLSSaveClass;

typedef DynamicArray<TZxPageInfo >  zeSave__2;

class DELPHICLASS TZXMLSSave;
class PASCALIMPLEMENTATION TZXMLSSave : public System::TInterfacedObject 
{
	typedef System::TInterfacedObject inherited;
	
protected:
	virtual int __fastcall DoSave(void);
	/* virtual class method */ virtual TStringDynArray __fastcall FormatDescriptions(TMetaClass* vmt);
	
public:
	virtual void __fastcall AfterConstruction(void);
	virtual void __fastcall BeforeDestruction(void);
	/* virtual class method */ virtual _di_IZXMLSSave __fastcall From(TMetaClass* vmt, const Zexmlss::TZEXMLSS* zxbook);
	
protected:
	__fastcall TZXMLSSave(const Zexmlss::TZEXMLSS* zxbook)/* overload */;
	__fastcall virtual TZXMLSSave(const TZXMLSSave* zxsaver)/* overload */;
	_di_IZXMLSSave __fastcall ExportFormat(const AnsiString fmt);
	_di_IZXMLSSave __fastcall As_(const AnsiString fmt);
	_di_IZXMLSSave __fastcall ExportTo(const AnsiString fname);
	_di_IZXMLSSave __fastcall To_(const AnsiString fname);
	_di_IZXMLSSave __fastcall Pages(const TZxPageInfo * APages, const int APages_Size)/* overload */;
	_di_IZXMLSSave __fastcall Pages(const int * numbers, const int numbers_Size)/* overload */;
	_di_IZXMLSSave __fastcall Pages(const AnsiString * titles, const int titles_Size)/* overload */;
	_di_IZXMLSSave __fastcall CharSet(const AnsiString cs)/* overload */;
	_di_IZXMLSSave __fastcall CharSet(const Zsspxml::TAnsiToCPConverter converter)/* overload */;
	_di_IZXMLSSave __fastcall CharSet(const AnsiString cs, const Zsspxml::TAnsiToCPConverter converter)/* overload */;
	_di_IZXMLSSave __fastcall CharSet(const Word codepage)/* overload */;
	_di_IZXMLSSave __fastcall BOM(const AnsiString Unicode_BOM);
	_di_IZXMLSSave __fastcall ZipWith(const TMetaClass* ZipGenerator);
	_di_IZXMLSSave __fastcall NoZip();
	_di_IZXMLSSave __fastcall OnErrorRaise();
	_di_IZXMLSSave __fastcall OnErrorRetCode();
	int __fastcall Save(void)/* overload */;
	int __fastcall Save(const AnsiString FileName)/* overload */;
	virtual void __fastcall Discard(void);
	int __fastcall InternalSave(void);
	Zexmlss::TZEXMLSS* fBook;
	DynamicArray<TZxPageInfo >  fPages;
	AnsiString fBOM;
	AnsiString fCharSet;
	Zsspxml::TAnsiToCPConverter fConv;
	AnsiString FFile;
	AnsiString FPath;
	TMetaClass*FZipGen;
	bool FDoNotDestroyMe;
	bool FRaiseOnError;
	TIntegerDynArray __fastcall GetPageNumbers();
	TStringDynArray __fastcall GetPageTitles();
	_di_IZXMLSSave __fastcall CreateSaverForDescription(const AnsiString desc);
	void __fastcall CheckSaveRetCode(int Result);
	HIDESBASE int __stdcall _AddRef(void);
	HIDESBASE int __stdcall _Release(void);
	/*         class method */ static void __fastcall RegisterFormat(TMetaClass* vmt, const TMetaClass* sv);
	/*         class method */ static void __fastcall UnRegisterFormat(TMetaClass* vmt, const TMetaClass* sv);
	
public:
	/*         class method */ static void __fastcall Register(TMetaClass* vmt);
	/*         class method */ static void __fastcall UnRegister(TMetaClass* vmt);
public:
	#pragma option push -w-inl
	/* TObject.Destroy */ inline __fastcall virtual ~TZXMLSSave(void) { }
	#pragma option pop
	
private:
	void *__IZXMLSSaveImpl;	/* Zesave::IZXMLSSaveImpl */
	
public:
	operator IZXMLSSaveImpl*(void) { return (IZXMLSSaveImpl*)&__IZXMLSSaveImpl; }
	operator IZXMLSSave*(void) { return (IZXMLSSave*)&__IZXMLSSaveImpl; }
	
};


class DELPHICLASS EZXSaveException;
class PASCALIMPLEMENTATION EZXSaveException : public Sysutils::Exception 
{
	typedef Sysutils::Exception inherited;
	
public:
	#pragma option push -w-inl
	/* Exception.Create */ inline __fastcall EZXSaveException(const AnsiString Msg) : Sysutils::Exception(Msg) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateFmt */ inline __fastcall EZXSaveException(const AnsiString Msg, const System::TVarRec * Args, const int Args_Size) : Sysutils::Exception(Msg, Args, Args_Size) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateRes */ inline __fastcall EZXSaveException(int Ident)/* overload */ : Sysutils::Exception(Ident) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateResFmt */ inline __fastcall EZXSaveException(int Ident, const System::TVarRec * Args, const int Args_Size)/* overload */ : Sysutils::Exception(Ident, Args, Args_Size) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateHelp */ inline __fastcall EZXSaveException(const AnsiString Msg, int AHelpContext) : Sysutils::Exception(Msg, AHelpContext) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateFmtHelp */ inline __fastcall EZXSaveException(const AnsiString Msg, const System::TVarRec * Args, const int Args_Size, int AHelpContext) : Sysutils::Exception(Msg, Args, Args_Size, AHelpContext) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateResHelp */ inline __fastcall EZXSaveException(int Ident, int AHelpContext)/* overload */ : Sysutils::Exception(Ident, AHelpContext) { }
	#pragma option pop
	#pragma option push -w-inl
	/* Exception.CreateResFmtHelp */ inline __fastcall EZXSaveException(System::PResStringRec ResStringRec, const System::TVarRec * Args, const int Args_Size, int AHelpContext)/* overload */ : Sysutils::Exception(ResStringRec, Args, Args_Size, AHelpContext) { }
	#pragma option pop
	
public:
	#pragma option push -w-inl
	/* TObject.Destroy */ inline __fastcall virtual ~EZXSaveException(void) { }
	#pragma option pop
	
};


//-- var, const, procedure ---------------------------------------------------

}	/* namespace Zesave */
using namespace Zesave;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zeSave
