// CodeGear C++Builder
// Copyright (c) 1995, 2010 by Embarcadero Technologies, Inc.
// All rights reserved

// (DO NOT EDIT: machine generated header) 'zsspxml.pas' rev: 22.00

#ifndef ZsspxmlHPP
#define ZsspxmlHPP

#pragma delphiheader begin
#pragma option push
#pragma option -w-      // All warnings off
#pragma option -Vx      // Zero-length empty class member functions
#pragma pack(push,8)
#include <System.hpp>	// Pascal unit
#include <SysInit.hpp>	// Pascal unit
#include <SysUtils.hpp>	// Pascal unit
#include <Classes.hpp>	// Pascal unit

//-- user supplied -----------------------------------------------------------

namespace Zsspxml
{
//-- type declarations -------------------------------------------------------
typedef System::AnsiString __fastcall (*TAnsiToCPConverter)(System::AnsiString AnsiText);

typedef void __fastcall (__closure *TReadCPCharObj)(System::AnsiString &RetChar, bool &_eof);

typedef void __fastcall (*TReadCPChar)(TReadCPCharObj ReadCPChar, System::AnsiString &text, bool &_eof);

typedef TAnsiToCPConverter TCPToAnsiConverter;

typedef System::StaticArray<System::AnsiString, 2> TZAttrArray;

struct DECLSPEC_DRECORD TagsProp
{
	
public:
	System::AnsiString Name;
	bool CloseTagNewLine;
};


class DELPHICLASS TZAttributes;
class PASCALIMPLEMENTATION TZAttributes : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
private:
	typedef System::DynamicArray<TZAttrArray> _TZAttributes__1;
	
	
public:
	System::AnsiString operator[](System::AnsiString Att) { return ItemsByName[Att]; }
	
private:
	int FCount;
	_TZAttributes__1 FItems;
	System::AnsiString __fastcall GetAttrS(System::AnsiString Att);
	void __fastcall SetAttrS(System::AnsiString Att, System::AnsiString Value);
	System::AnsiString __fastcall GetAttrI(int num);
	void __fastcall SetAttrI(int num, System::AnsiString Value);
	TZAttrArray __fastcall GetAttr(int num);
	void __fastcall SetAttr(int num, System::AnsiString *Value);
	
public:
	__fastcall TZAttributes(void);
	__fastcall virtual ~TZAttributes(void);
	void __fastcall Add(System::AnsiString AttrName, System::AnsiString Value, bool TestMatch = true)/* overload */;
	void __fastcall Add(System::AnsiString *Attr, bool TestMatch = true)/* overload */;
	void __fastcall Add(TZAttrArray *Att, const int Att_Size, bool TestMatch = true)/* overload */;
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	void __fastcall Clear(void);
	void __fastcall DeleteItem(int Index);
	void __fastcall Insert(int Index, System::AnsiString AttrName, System::AnsiString Value, bool TestMatch = true)/* overload */;
	void __fastcall Insert(int Index, System::AnsiString *Attr, bool TestMatch = true)/* overload */;
	HIDESBASE virtual System::AnsiString __fastcall ToString(char quote, bool CheckEntity, bool addempty)/* overload */;
	HIDESBASE virtual System::AnsiString __fastcall ToString(char quote, bool CheckEntity)/* overload */;
	HIDESBASE virtual System::AnsiString __fastcall ToString(char quote)/* overload */;
	HIDESBASE virtual System::AnsiString __fastcall ToString(bool CheckEntity)/* overload */;
	HIDESBASE virtual System::AnsiString __fastcall ToString(void)/* overload */;
	__property int Count = {read=FCount, nodefault};
//	__property TZAttrArray Items[int num] = {read=GetAttr, write=SetAttr};
	__property System::AnsiString ItemsByName[System::AnsiString Att] = {read=GetAttrS, write=SetAttrS/*, default*/};
	__property System::AnsiString ItemsByNum[int num] = {read=GetAttrI, write=SetAttrI};
};


class DELPHICLASS TZsspXMLWriter;
class PASCALIMPLEMENTATION TZsspXMLWriter : public System::TObject
{
	typedef System::TObject inherited;
	
private:
	typedef System::DynamicArray<TagsProp> _TZsspXMLWriter__1;
	
	
private:
	char FAttributeQuote;
	TZAttributes* FAttributes;
	_TZsspXMLWriter__1 FTags;
	int FTagCount;
	System::AnsiString FBuffer;
	int FMaxBufferLength;
	bool FInProcess;
	Classes::TStream* FStream;
	TAnsiToCPConverter FTextConverter;
	bool FNewLine;
	bool FUnixNLSeparator;
	System::AnsiString FNLSeparator;
	System::AnsiString FTab;
	System::AnsiString FTabSymbol;
	int FTabLength;
	System::Byte FDestination;
	System::AnsiString __fastcall GetTag(int num);
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
	void __fastcall AddText(System::AnsiString Text, bool UseConverter = true);
	void __fastcall AddNode(System::AnsiString TagName, bool CloseTagNewLine);
	System::AnsiString __fastcall GetTab(int num = 0x0);
	void __fastcall _AddTag(System::AnsiString _begin, System::AnsiString text, System::AnsiString _end, bool StartNewLine, int _tab = 0x0);
	
public:
	__fastcall virtual TZsspXMLWriter(void);
	__fastcall virtual ~TZsspXMLWriter(void);
	bool __fastcall BeginSaveToStream(Classes::TStream* Stream);
	bool __fastcall BeginSaveToFile(System::UnicodeString FileName);
	bool __fastcall BeginSaveToString(void);
	void __fastcall EndSaveTo(void);
	void __fastcall FlushBuffer(void);
	void __fastcall WriteCDATA(System::AnsiString CDATA, bool CorrectCDATA, bool StartNewLine = true)/* overload */;
	void __fastcall WriteCDATA(System::AnsiString CDATA)/* overload */;
	void __fastcall WriteComment(System::AnsiString Comment, bool StartNewLine = true);
	void __fastcall WriteEmptyTag(System::AnsiString TagName, TZAttributes* SAttributes, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteEmptyTag(System::AnsiString TagName, TZAttrArray *AttrArray, const int AttrArray_Size, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteEmptyTag(System::AnsiString TagName, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteEmptyTag(System::AnsiString TagName, TZAttributes* SAttributes)/* overload */;
	void __fastcall WriteEmptyTag(System::AnsiString TagName, TZAttrArray *AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteEmptyTag(System::AnsiString TagName)/* overload */;
	void __fastcall WriteEndTagNode(void);
	void __fastcall WriteInstruction(System::AnsiString InstructionName, TZAttributes* SAttributes, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteInstruction(System::AnsiString InstructionName, TZAttrArray *AttrArray, const int AttrArray_Size, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteInstruction(System::AnsiString InstructionName, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteInstruction(System::AnsiString InstructionName, TZAttributes* SAttributes)/* overload */;
	void __fastcall WriteInstruction(System::AnsiString InstructionName, TZAttrArray *AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteInstruction(System::AnsiString InstructionName)/* overload */;
	void __fastcall WriteRaw(System::AnsiString Text, bool UseConverter, bool StartNewLine = true);
	void __fastcall WriteTag(System::AnsiString TagName, System::AnsiString Text, TZAttributes* SAttributes, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTag(System::AnsiString TagName, System::AnsiString Text, TZAttrArray *AttrArray, const int AttrArray_Size, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTag(System::AnsiString TagName, System::AnsiString Text, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTag(System::AnsiString TagName, System::AnsiString Text, TZAttributes* SAttributes)/* overload */;
	void __fastcall WriteTag(System::AnsiString TagName, System::AnsiString Text, TZAttrArray *AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteTag(System::AnsiString TagName, System::AnsiString Text)/* overload */;
	void __fastcall WriteTagNode(System::AnsiString TagName, TZAttributes* SAttributes, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTagNode(System::AnsiString TagName, TZAttrArray *AttrArray, const int AttrArray_Size, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTagNode(System::AnsiString TagName, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTagNode(System::AnsiString TagName, TZAttributes* SAttributes)/* overload */;
	void __fastcall WriteTagNode(System::AnsiString TagName, TZAttrArray *AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteTagNode(System::AnsiString TagName)/* overload */;
	__property TZAttributes* Attributes = {read=FAttributes, write=SetAttributes};
	__property char AttributeQuote = {read=FAttributeQuote, write=SetAttributeQuote, nodefault};
	__property System::AnsiString Buffer = {read=FBuffer};
	__property bool InProcess = {read=FInProcess, nodefault};
	__property int MaxBufferLength = {read=FMaxBufferLength, write=SetMaxBufferLength, nodefault};
	__property bool NewLine = {read=FNewLine, write=SetNewLine, nodefault};
	__property int TabLength = {read=FTabLength, write=SetTabLength, nodefault};
	__property char TabSymbol = {read=GetTabSymbol, write=SetTabSymbol, nodefault};
	__property int TagCount = {read=FTagCount, nodefault};
	__property System::AnsiString Tags[int num] = {read=GetTag};
	__property TAnsiToCPConverter TextConverter = {read=FTextConverter, write=SetTextConverter};
	__property bool UnixNLSeparator = {read=FUnixNLSeparator, write=SetUnixNLSeparator, nodefault};
};


class DELPHICLASS TZsspXMLReader;
class PASCALIMPLEMENTATION TZsspXMLReader : public System::TObject
{
	typedef System::TObject inherited;
	
private:
	typedef System::DynamicArray<System::AnsiString> _TZsspXMLReader__1;
	
	
private:
	TZAttributes* FAttributes;
	Classes::TStream* FStream;
	_TZsspXMLReader__1 FTags;
	int FTagCount;
	System::AnsiString FBuffer;
	int FMaxBufferLength;
	bool FInProcess;
	System::Byte FSourceType;
	int FPFirst;
	int FPLast;
	System::AnsiString FTextBeforeTag;
	int FErrorCode;
	System::AnsiString FRawTextTag;
	System::AnsiString FTagName;
	System::AnsiString FValue;
	System::Byte FTagType;
	TReadCPChar FCharReader;
	TAnsiToCPConverter FCharConverter;
	bool FStreamEnd;
	bool FIgnoreCase;
	System::AnsiString __fastcall GetTag(int num);
	void __fastcall SetMaxBufferLength(int Value);
	void __fastcall SetAttributes(TZAttributes* Value);
	void __fastcall AddTag(System::AnsiString Value);
	void __fastcall DeleteClosedTag(void);
	void __fastcall DeleteTag(void);
	void __fastcall SetIgnoreCase(bool Value);
	
protected:
	void __fastcall Clear(void);
	void __fastcall ClearAll(void);
	void __fastcall RecognizeEncoding(System::AnsiString &txt);
	void __fastcall ReadBuffer(void);
	void __fastcall GetOneChar(System::AnsiString &OneChar, bool &err);
	
public:
	__fastcall virtual TZsspXMLReader(void);
	__fastcall virtual ~TZsspXMLReader(void);
	int __fastcall BeginReadFile(System::UnicodeString FileName);
	int __fastcall BeginReadStream(Classes::TStream* Stream);
	int __fastcall BeginReadString(System::AnsiString Source, bool IgnoreCodePage = true);
	bool __fastcall ReadTag(void);
	void __fastcall EndRead(void);
	virtual bool __fastcall Eof(void);
	__property TZAttributes* Attributes = {read=FAttributes};
	__property bool InProcess = {read=FInProcess, nodefault};
	__property System::AnsiString RawTextTag = {read=FRawTextTag};
	__property int ErrorCode = {read=FErrorCode, nodefault};
	__property bool IgnoreCase = {read=FIgnoreCase, write=SetIgnoreCase, nodefault};
	__property System::AnsiString TagName = {read=FTagName};
	__property System::AnsiString TagValue = {read=FValue};
	__property System::Byte TagType = {read=FTagType, nodefault};
	__property int TagCount = {read=FTagCount, nodefault};
	__property System::AnsiString Tags[int num] = {read=GetTag};
	__property System::AnsiString TextBeforeTag = {read=FTextBeforeTag};
	__property int MaxBufferLength = {read=FMaxBufferLength, write=SetMaxBufferLength, nodefault};
};


typedef System::StaticArray<System::UnicodeString, 2> TZAttrArrayH;

class DELPHICLASS TZAttributesH;
class PASCALIMPLEMENTATION TZAttributesH : public Classes::TPersistent
{
	typedef Classes::TPersistent inherited;
	
public:
	System::UnicodeString operator[](System::UnicodeString Att) { return ItemsByName[Att]; }
	
private:
	TZAttributes* FAttributes;
	int __fastcall GetAttrCount(void);
	System::UnicodeString __fastcall GetAttrS(System::UnicodeString Att);
	void __fastcall SetAttrS(System::UnicodeString Att, System::UnicodeString Value);
	System::UnicodeString __fastcall GetAttrI(int num);
	void __fastcall SetAttrI(int num, System::UnicodeString Value);
	TZAttrArrayH __fastcall GetAttr(int num);
	void __fastcall SetAttr(int num, System::UnicodeString *Value);
	
public:
	__fastcall TZAttributesH(void);
	__fastcall virtual ~TZAttributesH(void);
	void __fastcall Add(System::UnicodeString AttrName, System::UnicodeString Value, bool TestMatch = true)/* overload */;
	void __fastcall Add(System::UnicodeString *Attr, bool TestMatch = true)/* overload */;
	void __fastcall Add(TZAttrArrayH *Att, const int Att_Size, bool TestMatch = true)/* overload */;
	virtual void __fastcall Assign(Classes::TPersistent* Source);
	void __fastcall Clear(void);
	void __fastcall DeleteItem(int Index);
	void __fastcall Insert(int Index, System::UnicodeString AttrName, System::UnicodeString Value, bool TestMatch = true)/* overload */;
	void __fastcall Insert(int Index, System::UnicodeString *Attr, bool TestMatch = true)/* overload */;
	HIDESBASE virtual System::UnicodeString __fastcall ToString(System::WideChar quote, bool CheckEntity, bool addempty)/* overload */;
	HIDESBASE virtual System::UnicodeString __fastcall ToString(System::WideChar quote, bool CheckEntity)/* overload */;
	HIDESBASE virtual System::UnicodeString __fastcall ToString(System::WideChar quote)/* overload */;
	HIDESBASE virtual System::UnicodeString __fastcall ToString(bool CheckEntity)/* overload */;
	virtual System::UnicodeString __fastcall ToString(void)/* overload */;
	__property int Count = {read=GetAttrCount, nodefault};
//	__property TZAttrArrayH Items[int num] = {read=GetAttr, write=SetAttr};
	__property System::UnicodeString ItemsByName[System::UnicodeString Att] = {read=GetAttrS, write=SetAttrS/*, default*/};
	__property System::UnicodeString ItemsByNum[int num] = {read=GetAttrI, write=SetAttrI};
};


class DELPHICLASS TZsspXMLWriterH;
class PASCALIMPLEMENTATION TZsspXMLWriterH : public System::TObject
{
	typedef System::TObject inherited;
	
private:
	TZAttributesH* FAttributes;
	TZsspXMLWriter* FXMLWriter;
	System::UnicodeString __fastcall GetXMLBuffer(void);
	System::WideChar __fastcall GetAttributeQuote(void);
	bool __fastcall GetInProcess(void);
	int __fastcall GetMaxBufferLength(void);
	bool __fastcall GetNewLine(void);
	int __fastcall GetTabLength(void);
	int __fastcall GetTagCount(void);
	TAnsiToCPConverter __fastcall GetTextConverter(void);
	bool __fastcall GetUnixNLSeparator(void);
	System::UnicodeString __fastcall GetTag(int num);
	System::WideChar __fastcall GetTabSymbol(void);
	void __fastcall SetAttributeQuote(System::WideChar Value);
	void __fastcall SetMaxBufferLength(int Value);
	void __fastcall SetNewLine(bool Value);
	void __fastcall SetTabLength(int Value);
	void __fastcall SetTabSymbol(System::WideChar Value);
	void __fastcall SetTextConverter(TAnsiToCPConverter Value);
	void __fastcall SetUnixNLSeparator(bool Value);
	void __fastcall SetAttributes(TZAttributesH* Value);
	
public:
	__fastcall virtual TZsspXMLWriterH(void);
	__fastcall virtual ~TZsspXMLWriterH(void);
	bool __fastcall BeginSaveToStream(Classes::TStream* Stream);
	bool __fastcall BeginSaveToFile(System::UnicodeString FileName);
	bool __fastcall BeginSaveToString(void);
	void __fastcall EndSaveTo(void);
	void __fastcall FlushBuffer(void);
	void __fastcall WriteCDATA(System::UnicodeString CDATA, bool CorrectCDATA, bool StartNewLine = true)/* overload */;
	void __fastcall WriteCDATA(System::UnicodeString CDATA)/* overload */;
	void __fastcall WriteComment(System::UnicodeString Comment, bool StartNewLine = true);
	void __fastcall WriteEmptyTag(System::UnicodeString TagName, TZAttributesH* SAttributes, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteEmptyTag(System::UnicodeString TagName, TZAttrArrayH *AttrArray, const int AttrArray_Size, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteEmptyTag(System::UnicodeString TagName, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteEmptyTag(System::UnicodeString TagName, TZAttributesH* SAttributes)/* overload */;
	void __fastcall WriteEmptyTag(System::UnicodeString TagName, TZAttrArrayH *AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteEmptyTag(System::UnicodeString TagName)/* overload */;
	void __fastcall WriteEndTagNode(void);
	void __fastcall WriteInstruction(System::UnicodeString InstructionName, TZAttributesH* SAttributes, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteInstruction(System::UnicodeString InstructionName, TZAttrArrayH *AttrArray, const int AttrArray_Size, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteInstruction(System::UnicodeString InstructionName, bool StartNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteInstruction(System::UnicodeString InstructionName, TZAttributesH* SAttributes)/* overload */;
	void __fastcall WriteInstruction(System::UnicodeString InstructionName, TZAttrArrayH *AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteInstruction(System::UnicodeString InstructionName)/* overload */;
	void __fastcall WriteRaw(System::UnicodeString Text, bool UseConverter, bool StartNewLine = true)/* overload */;
	void __fastcall WriteRaw(System::AnsiString Text, bool UseConverter, bool StartNewLine = true)/* overload */;
	void __fastcall WriteTag(System::UnicodeString TagName, System::UnicodeString Text, TZAttributesH* SAttributes, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTag(System::UnicodeString TagName, System::UnicodeString Text, TZAttrArrayH *AttrArray, const int AttrArray_Size, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTag(System::UnicodeString TagName, System::UnicodeString Text, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTag(System::UnicodeString TagName, System::UnicodeString Text, TZAttributesH* SAttributes)/* overload */;
	void __fastcall WriteTag(System::UnicodeString TagName, System::UnicodeString Text, TZAttrArrayH *AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteTag(System::UnicodeString TagName, System::UnicodeString Text)/* overload */;
	void __fastcall WriteTagNode(System::UnicodeString TagName, TZAttributesH* SAttributes, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTagNode(System::UnicodeString TagName, TZAttrArrayH *AttrArray, const int AttrArray_Size, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTagNode(System::UnicodeString TagName, bool StartNewLine, bool CloseTagNewLine, bool CheckEntity = true)/* overload */;
	void __fastcall WriteTagNode(System::UnicodeString TagName, TZAttributesH* SAttributes)/* overload */;
	void __fastcall WriteTagNode(System::UnicodeString TagName, TZAttrArrayH *AttrArray, const int AttrArray_Size)/* overload */;
	void __fastcall WriteTagNode(System::UnicodeString TagName)/* overload */;
	__property TZAttributesH* Attributes = {read=FAttributes, write=SetAttributes};
	__property System::WideChar AttributeQuote = {read=GetAttributeQuote, write=SetAttributeQuote, nodefault};
	__property System::UnicodeString Buffer = {read=GetXMLBuffer};
	__property bool InProcess = {read=GetInProcess, nodefault};
	__property int MaxBufferLength = {read=GetMaxBufferLength, write=SetMaxBufferLength, nodefault};
	__property bool NewLine = {read=GetNewLine, write=SetNewLine, nodefault};
	__property int TabLength = {read=GetTabLength, write=SetTabLength, nodefault};
	__property System::WideChar TabSymbol = {read=GetTabSymbol, write=SetTabSymbol, nodefault};
	__property int TagCount = {read=GetTagCount, nodefault};
	__property System::UnicodeString Tags[int num] = {read=GetTag};
	__property TAnsiToCPConverter TextConverter = {read=GetTextConverter, write=SetTextConverter};
	__property bool UnixNLSeparator = {read=GetUnixNLSeparator, write=SetUnixNLSeparator, nodefault};
};


class DELPHICLASS TZsspXMLReaderH;
class PASCALIMPLEMENTATION TZsspXMLReaderH : public System::TObject
{
	typedef System::TObject inherited;
	
private:
	TZAttributesH* FAttributes;
	TZsspXMLReader* FXMLReader;
	TZAttributesH* __fastcall GetAttributes(void);
	bool __fastcall GetInProcess(void);
	System::UnicodeString __fastcall GetRawTextTag(void);
	int __fastcall GetErrorCode(void);
	bool __fastcall GetIgnoreCase(void);
	System::UnicodeString __fastcall GetValue(void);
	System::Byte __fastcall GetTagType(void);
	int __fastcall GetTagCount(void);
	System::UnicodeString __fastcall GetTextBeforeTag(void);
	System::UnicodeString __fastcall GetTagName(void);
	System::UnicodeString __fastcall GetTag(int num);
	void __fastcall SetMaxBufferLength(int Value);
	int __fastcall GetMaxBufferLength(void);
	void __fastcall SetAttributes(TZAttributesH* Value);
	void __fastcall SetIgnoreCase(bool Value);
	
public:
	__fastcall virtual TZsspXMLReaderH(void);
	__fastcall virtual ~TZsspXMLReaderH(void);
	int __fastcall BeginReadFile(System::UnicodeString FileName);
	int __fastcall BeginReadStream(Classes::TStream* Stream);
	int __fastcall BeginReadString(System::UnicodeString Source, bool IgnoreCodePage = true);
	bool __fastcall ReadTag(void);
	void __fastcall EndRead(void);
	virtual bool __fastcall Eof(void);
	__property TZAttributesH* Attributes = {read=GetAttributes};
	__property bool InProcess = {read=GetInProcess, nodefault};
	__property System::UnicodeString RawTextTag = {read=GetRawTextTag};
	__property int ErrorCode = {read=GetErrorCode, nodefault};
	__property bool IgnoreCase = {read=GetIgnoreCase, write=SetIgnoreCase, nodefault};
	__property System::UnicodeString TagName = {read=GetTagName};
	__property System::UnicodeString TagValue = {read=GetValue};
	__property System::Byte TagType = {read=GetTagType, nodefault};
	__property int TagCount = {read=GetTagCount, nodefault};
	__property System::UnicodeString Tags[int num] = {read=GetTag};
	__property System::UnicodeString TextBeforeTag = {read=GetTextBeforeTag};
	__property int MaxBufferLength = {read=GetMaxBufferLength, write=SetMaxBufferLength, nodefault};
};


//-- var, const, procedure ---------------------------------------------------
#define BOMUTF8 L"\u043f\u00bb\u0457"
#define BOMUTF16BE L"\u044e\u044f"
#define BOMUTF16LE L"\u044f\u044e"
#define BOMUTF32BE L"\u0000\u0000\u044e\u044f"
#define BOMUTF32LE L"\u044f\u044e\u0000\u0000"
extern PACKAGE void __fastcall ReadCharUTF8(TReadCPCharObj ReadCPChar, System::AnsiString &text, bool &_eof);
extern PACKAGE void __fastcall ReadCharUTF16LE(TReadCPCharObj ReadCPChar, System::AnsiString &text, bool &_eof);
extern PACKAGE void __fastcall ReadCharUTF16BE(TReadCPCharObj ReadCPChar, System::AnsiString &text, bool &_eof);
extern PACKAGE void __fastcall ReadCharUTF32(TReadCPCharObj ReadCPChar, System::AnsiString &text, bool &_eof);
extern PACKAGE void __fastcall ReadCharOneByte(TReadCPCharObj ReadCPChar, System::AnsiString &text, bool &_eof);
extern PACKAGE System::AnsiString __fastcall conv_UTF8ToLocal(System::AnsiString Text);
extern PACKAGE System::AnsiString __fastcall conv_UTF16LEToLocal(System::AnsiString Text);
extern PACKAGE System::AnsiString __fastcall conv_UTF16BEToLocal(System::AnsiString Text);
extern PACKAGE System::AnsiString __fastcall conv_UTF32LEToLocal(System::AnsiString Text);
extern PACKAGE System::AnsiString __fastcall conv_UTF32BEToLocal(System::AnsiString Text);
extern PACKAGE System::AnsiString __fastcall conv_WIN1251ToLocal(System::AnsiString Text);
extern PACKAGE System::AnsiString __fastcall conv_CP866ToLocal(System::AnsiString Text);
extern PACKAGE TZAttrArray __fastcall ToAttribute(System::AnsiString AttrName, System::AnsiString Value)/* overload */;
extern PACKAGE TZAttrArrayH __fastcall ToAttribute(System::UnicodeString AttrName, System::UnicodeString Value)/* overload */;
extern PACKAGE void __fastcall Correct_Entity(const System::AnsiString _St, int num, System::AnsiString &_result)/* overload */;
extern PACKAGE void __fastcall Correct_Entity(const System::UnicodeString _St, int num, System::UnicodeString &_result)/* overload */;
extern PACKAGE System::AnsiString __fastcall CheckStrEntity(System::AnsiString st, bool checkamp = true)/* overload */;
extern PACKAGE System::UnicodeString __fastcall CheckStrEntity(System::UnicodeString st, bool checkamp = true)/* overload */;
extern PACKAGE bool __fastcall RecognizeEncodingXML(int startpos, System::AnsiString &txt, int &cpfromtext, System::AnsiString &cpname, int &ftype)/* overload */;
extern PACKAGE int __fastcall RecognizeBOM(System::AnsiString &txt);
extern PACKAGE bool __fastcall RecognizeEncodingXML(System::AnsiString &txt, int &BOM, int &cpfromtext, System::AnsiString &cpname, int &ftype)/* overload */;

}	/* namespace Zsspxml */
#if !defined(DELPHIHEADER_NO_IMPLICIT_NAMESPACE_USE)
using namespace Zsspxml;
#endif
#pragma pack(pop)
#pragma option pop

#pragma delphiheader end.
//-- end unit ----------------------------------------------------------------
#endif	// ZsspxmlHPP
