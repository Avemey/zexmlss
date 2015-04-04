// Borland C++ Builder
// Copyright (c) 1995, 2002 by Borland Software Corporation
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zsspxml.pas' rev: 6.00

#ifndef zsspxmlHPP
#define zsspxmlHPP

#pragma delphiheader begin
#pragma option push -w-
#pragma option push -Vx
#include <Classes.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <System.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zsspxml
{
//-- type declarations -------------------------------------------------------
typedef AnsiString __fastcall (*TAnsiToCPConverter)(const AnsiString AnsiText);

typedef void __fastcall (__closure *TReadCPCharObj)(AnsiString &RetChar, bool &_eof);

typedef void __fastcall (*TReadCPChar)(const TReadCPCharObj ReadCPChar, AnsiString &text, bool &_eof);

typedef AnsiString __fastcall (*TCPToAnsiConverter)(const AnsiString AnsiText);

typedef AnsiString TZAttrArray[2];

#pragma pack(push, 4)
struct TagsProp
{
	AnsiString Name;
	bool CloseTagNewLine;
} ;
#pragma pack(pop)

typedef DynamicArray<AnsiString >  zsspxml__2;

class DELPHICLASS TZAttributes;
class PASCALIMPLEMENTATION TZAttributes : public Classes::TPersistent 
{
	typedef Classes::TPersistent inherited;
	
public:
	AnsiString operator[](AnsiString Att) { return ItemsByName[Att]; }
	
private:
	int FCount;
	int FMaxCount;
	DynamicArray<AnsiString >  FItems;
	AnsiString __fastcall GetAttrS(const AnsiString Att);
	void __fastcall SetAttrS(const AnsiString Att, const AnsiString Value);
	AnsiString __fastcall GetAttrI(int num);
	void __fastcall SetAttrI(int num, const AnsiString Value);
	AnsiString __fastcall GetAttr(int num);
	void __fastcall SetAttr(int num, const AnsiString * Value);
	
protected:
	void __fastcall ResizeItemsArray(int NewSize);
	
public:
	__fastcall TZAttributes(void);
	__fastcall virtual ~TZAttributes(void);
	void __fastcall Add(const AnsiString AttrName, const AnsiString Value, bool TestMatch = true)/* overload */;
	void __fastcall Add(const AnsiString * Attr, bool TestMatch = true)/* overload */;
	void __fastcall Add(const AnsiString * Att, const int Att_Size, bool TestMatch = true)/* overload */;
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	void __fastcall Clear(void);
	void __fastcall DeleteItem(int Index);
	void __fastcall Insert(int Index, const AnsiString AttrName, const AnsiString Value, bool TestMatch = true)/* overload */;
	void __fastcall Insert(int Index, const AnsiString * Attr, bool TestMatch = true)/* overload */;
	virtual AnsiString __fastcall ToString(char quote, bool CheckEntity, bool addempty)/* overload */;
	virtual AnsiString __fastcall ToString(char quote, bool CheckEntity)/* overload */;
	virtual AnsiString __fastcall ToString(char quote)/* overload */;
	virtual AnsiString __fastcall ToString(bool CheckEntity)/* overload */;
	virtual AnsiString __fastcall ToString()/* overload */;
	__property int Count = {read=FCount, nodefault};
	//__property AnsiString Items[int num] = {read=GetAttr, write=SetAttr};
	__property AnsiString ItemsByName[AnsiString Att] = {read=GetAttrS, write=SetAttrS/*, default*/};
	__property AnsiString ItemsByNum[int num] = {read=GetAttrI, write=SetAttrI};
};


typedef DynamicArray<TagsProp >  zsspxml__4;

class DELPHICLASS TZsspXMLWriter;
class PASCALIMPLEMENTATION TZsspXMLWriter : public System::TObject 
{
	typedef System::TObject inherited;
	
private:
	char FAttributeQuote;
	TZAttributes* FAttributes;
	DynamicArray<TagsProp >  FTags;
	int FTagCount;
	int FMaxTagCount;
	AnsiString FBuffer;
	int FMaxBufferLength;
	bool FInProcess;
	Classes::TStream* FStream;
	TAnsiToCPConverter FTextConverter;
	bool FNewLine;
	bool FUnixNLSeparator;
	AnsiString FNLSeparator;
	AnsiString FTab;
	AnsiString FTabSymbol;
	int FTabLength;
	Byte FDestination;
	AnsiString __fastcall GetTag(int num);
	char __fastcall GetTabSymbol(void);
	void __fastcall SetAttributeQuote(char Value);
	void __fastcall SetMaxBufferLength(int Value);
	void __fastcall SetNewLine(bool Value);
	void __fastcall SetTabLength(int Value);
	void __fastcall SetTabSymbol(char Value);
	void __fastcall SetTextConverter(TAnsiToCPConverter Value);
	void __fastcall SetUnixNLSeparator(bool Value);
	void __fastcall SetAttributes(TZAttributes* Value);
	
protected:
	void __fastcall AddText(const AnsiString Text, bool UseConverter = true);
	void __fastcall AddNode(const AnsiString TagName, bool CloseTagNewLine);
	AnsiString __fastcall GetTab(int num = 0x0);
	void __fastcall _AddTag(const AnsiString _begin, AnsiString text, const AnsiString _end, bool StartNewLine, int _tab = 0x0);
	void __fastcall ResizeTagArray(int NewSize);
	
public:
	__fastcall virtual TZsspXMLWriter(void);
	__fastcall virtual ~TZsspXMLWriter(void);
	bool __fastcall BeginSaveToStream(Classes::TStream* Stream);
	bool __fastcall BeginSaveToFile(const AnsiString FileName);
	bool __fastcall BeginSaveToString(void);
	void __fastcall EndSaveTo(void);
	void __fastcall FlushBuffer(void);
	void __fastcall WriteCDATA(AnsiString CDATA, bool CorrectCDATA, bool StartNewLine = true)/* overload */;
	void __fastcall WriteCDATA(const AnsiString CDATA)/* overload */;
	void __fastcall WriteComment(const AnsiString Comment, bool StartNewLine = true);
	void __fastcall WriteEmptyTag(const AnsiString TagName, TZAttributes* SAttributes, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteEmptyTag(const AnsiString TagName, const AnsiString * AttrArray, const int AttrArray_Size, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteEmptyTag(const AnsiString TagName, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteEmptyTag(const AnsiString TagName, TZAttributes* SAttributes)/* overload */;
	void __fastcall WriteEmptyTag(const AnsiString TagName, const AnsiString * AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteEmptyTag(const AnsiString TagName)/* overload */;
	void __fastcall WriteEndTagNode(void)/* overload */;
	void __fastcall WriteEndTagNode(bool isForce, bool CloseTagNewLine)/* overload */;
	void __fastcall WriteInstruction(const AnsiString InstructionName, TZAttributes* SAttributes, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteInstruction(const AnsiString InstructionName, const AnsiString * AttrArray, const int AttrArray_Size, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteInstruction(const AnsiString InstructionName, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteInstruction(const AnsiString InstructionName, TZAttributes* SAttributes)/* overload */;
	void __fastcall WriteInstruction(const AnsiString InstructionName, const AnsiString * AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteInstruction(const AnsiString InstructionName)/* overload */;
	void __fastcall WriteRaw(AnsiString Text, bool UseConverter, bool StartNewLine = true);
	void __fastcall WriteTag(const AnsiString TagName, const AnsiString Text, TZAttributes* SAttributes, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTag(const AnsiString TagName, const AnsiString Text, const AnsiString * AttrArray, const int AttrArray_Size, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTag(const AnsiString TagName, const AnsiString Text, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTag(const AnsiString TagName, const AnsiString Text, TZAttributes* SAttributes)/* overload */;
	void __fastcall WriteTag(const AnsiString TagName, const AnsiString Text, const AnsiString * AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteTag(const AnsiString TagName, const AnsiString Text)/* overload */;
	void __fastcall WriteTagNode(const AnsiString TagName, TZAttributes* SAttributes, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTagNode(const AnsiString TagName, const AnsiString * AttrArray, const int AttrArray_Size, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTagNode(const AnsiString TagName, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTagNode(const AnsiString TagName, TZAttributes* SAttributes)/* overload */;
	void __fastcall WriteTagNode(const AnsiString TagName, const AnsiString * AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteTagNode(const AnsiString TagName)/* overload */;
	__property TZAttributes* Attributes = {read=FAttributes, write=SetAttributes};
	__property char AttributeQuote = {read=FAttributeQuote, write=SetAttributeQuote, nodefault};
	__property AnsiString Buffer = {read=FBuffer};
	__property bool InProcess = {read=FInProcess, nodefault};
	__property int MaxBufferLength = {read=FMaxBufferLength, write=SetMaxBufferLength, nodefault};
	__property bool NewLine = {read=FNewLine, write=SetNewLine, nodefault};
	__property int TabLength = {read=FTabLength, write=SetTabLength, nodefault};
	__property char TabSymbol = {read=GetTabSymbol, write=SetTabSymbol, nodefault};
	__property int TagCount = {read=FTagCount, nodefault};
	__property AnsiString Tags[int num] = {read=GetTag};
	__property TAnsiToCPConverter TextConverter = {read=FTextConverter, write=SetTextConverter};
	__property bool UnixNLSeparator = {read=FUnixNLSeparator, write=SetUnixNLSeparator, nodefault};
};


typedef DynamicArray<AnsiString >  zsspxml__6;

class DELPHICLASS TZsspXMLReader;
class PASCALIMPLEMENTATION TZsspXMLReader : public System::TObject 
{
	typedef System::TObject inherited;
	
private:
	TZAttributes* FAttributes;
	Classes::TStream* FStream;
	DynamicArray<AnsiString >  FTags;
	int FTagCount;
	int FMaxTagCount;
	AnsiString FBuffer;
	int FMaxBufferLength;
	bool FInProcess;
	Byte FSourceType;
	int FPFirst;
	int FPLast;
	AnsiString FTextBeforeTag;
	int FErrorCode;
	AnsiString FRawTextTag;
	AnsiString FTagName;
	AnsiString FValue;
	Byte FTagType;
	TReadCPChar FCharReader;
	TAnsiToCPConverter FCharConverter;
	bool FStreamEnd;
	bool FIgnoreCase;
	bool FQuotesEqual;
	bool FAttributesMatch;
	AnsiString __fastcall GetTag(int num);
	void __fastcall SetMaxBufferLength(int Value);
	void __fastcall SetAttributes(TZAttributes* Value);
	void __fastcall AddTag(const AnsiString Value);
	void __fastcall DeleteClosedTag(void);
	void __fastcall DeleteTag(void);
	void __fastcall SetIgnoreCase(bool Value);
	void __fastcall SetQuotesEqual(bool Value);
	void __fastcall SetAttributesMatch(bool Value);
	
protected:
	void __fastcall Clear(void);
	void __fastcall ClearAll(void);
	void __fastcall RecognizeEncoding(AnsiString &txt);
	void __fastcall ReadBuffer(void);
	void __fastcall GetOneChar(AnsiString &OneChar, bool &err);
	void __fastcall ResizeTagArray(int NewSize);
	
public:
	__fastcall virtual TZsspXMLReader(void);
	__fastcall virtual ~TZsspXMLReader(void);
	int __fastcall BeginReadFile(AnsiString FileName);
	int __fastcall BeginReadStream(Classes::TStream* Stream);
	int __fastcall BeginReadString(AnsiString Source, bool IgnoreCodePage = true);
	bool __fastcall ReadTag(void);
	void __fastcall EndRead(void);
	virtual bool __fastcall Eof(void);
	__property TZAttributes* Attributes = {read=FAttributes};
	__property bool AttributesMatch = {read=FAttributesMatch, write=SetAttributesMatch, nodefault};
	__property bool InProcess = {read=FInProcess, nodefault};
	__property AnsiString RawTextTag = {read=FRawTextTag};
	__property int ErrorCode = {read=FErrorCode, nodefault};
	__property bool IgnoreCase = {read=FIgnoreCase, write=SetIgnoreCase, nodefault};
	__property AnsiString TagName = {read=FTagName};
	__property AnsiString TagValue = {read=FValue};
	__property Byte TagType = {read=FTagType, nodefault};
	__property int TagCount = {read=FTagCount, nodefault};
	__property AnsiString Tags[int num] = {read=GetTag};
	__property AnsiString TextBeforeTag = {read=FTextBeforeTag};
	__property int MaxBufferLength = {read=FMaxBufferLength, write=SetMaxBufferLength, nodefault};
	__property bool QuotesEqual = {read=FQuotesEqual, write=SetQuotesEqual, nodefault};
};


typedef AnsiString TZAttrArrayH[2];

typedef TZAttributes TZAttributesH;
;

typedef TZsspXMLWriter TZsspXMLWriterH;
;

typedef TZsspXMLReader TZsspXMLReaderH;
;

//-- var, const, procedure ---------------------------------------------------
#define BOMUTF8 "ï»¿"
#define BOMUTF16BE "þÿ"
#define BOMUTF16LE "ÿþ"
#define BOMUTF32BE ""
#define BOMUTF32LE "ÿþ"
extern PACKAGE void __fastcall ReadCharUTF8(const TReadCPCharObj ReadCPChar, AnsiString &text, bool &_eof);
extern PACKAGE void __fastcall ReadCharUTF16LE(const TReadCPCharObj ReadCPChar, AnsiString &text, bool &_eof);
extern PACKAGE void __fastcall ReadCharUTF16BE(const TReadCPCharObj ReadCPChar, AnsiString &text, bool &_eof);
extern PACKAGE void __fastcall ReadCharUTF32(const TReadCPCharObj ReadCPChar, AnsiString &text, bool &_eof);
extern PACKAGE void __fastcall ReadCharOneByte(const TReadCPCharObj ReadCPChar, AnsiString &text, bool &_eof);
extern PACKAGE AnsiString __fastcall conv_UTF8ToLocal(const AnsiString Text);
extern PACKAGE AnsiString __fastcall conv_UTF16LEToLocal(const AnsiString Text);
extern PACKAGE AnsiString __fastcall conv_UTF16BEToLocal(const AnsiString Text);
extern PACKAGE AnsiString __fastcall conv_UTF32LEToLocal(const AnsiString Text);
extern PACKAGE AnsiString __fastcall conv_UTF32BEToLocal(const AnsiString Text);
extern PACKAGE AnsiString __fastcall conv_WIN1251ToLocal(const AnsiString Text);
extern PACKAGE AnsiString __fastcall conv_CP866ToLocal(const AnsiString Text);
extern PACKAGE AnsiString __fastcall ToAttribute(const AnsiString AttrName, const AnsiString Value);
extern PACKAGE void __fastcall Correct_Entity(const AnsiString _St, int num, AnsiString &_result);
extern PACKAGE AnsiString __fastcall CheckStrEntity(const AnsiString st, bool checkamp = true);
extern PACKAGE TAnsiToCPConverter __fastcall ZEGetDefaultUTF8Converter(void);
extern PACKAGE bool __fastcall RecognizeEncodingXML(int startpos, AnsiString &txt, /* out */ int &cpfromtext, /* out */ AnsiString &cpname, /* out */ int &ftype)/* overload */;
extern PACKAGE int __fastcall RecognizeBOM(AnsiString &txt);
extern PACKAGE bool __fastcall RecognizeEncodingXML(AnsiString &txt, /* out */ int &BOM, /* out */ int &cpfromtext, /* out */ AnsiString &cpname, /* out */ int &ftype)/* overload */;

}	/* namespace Zsspxml */
using namespace Zsspxml;
#pragma option pop	// -w-
#pragma option pop	// -Vx

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// zsspxml
