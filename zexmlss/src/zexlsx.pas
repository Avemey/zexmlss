//****************************************************************
// Read/write xlsx (Office Open XML file format (Spreadsheet))
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2016.07.03
//----------------------------------------------------------------
// Modified by the_Arioch@nm.ru - uniform save API for creating
//     XLSX files in Delphi/Windows
{
 Copyright (C) 2012 Ruslan Neborak

  This software is provided 'as-is', without any express or implied
 warranty. In no event will the authors be held liable for any damages
 arising from the use of this software.

 Permission is granted to anyone to use this software for any purpose,
 including commercial applications, and to alter it and redistribute it
 freely, subject to the following restrictions:

    1. The origin of this software must not be misrepresented; you must not
    claim that you wrote the original software. If you use this software
    in a product, an acknowledgment in the product documentation would be
    appreciated but is not required.

    2. Altered source versions must be plainly marked as such, and must not be
    misrepresented as being the original software.

    3. This notice may not be removed or altered from any source
    distribution.
}
//HINT: File should have name 'zexlsx.pas', but my hands ... :)
// ... and file renamed (thanks to Arioch)
//****************************************************************
unit zexlsx;

interface

{$I zexml.inc}
{$I compver.inc}

{$IFDEF FPC}
  {$mode DELPHI}{$H+}
{$ENDIF}

uses
  SysUtils, Classes, Types, Graphics,
  zeformula, zsspxml, zexmlss, zesavecommon
  {$IFNDEF FPC}
  ,windows
  {$ELSE}
  ,LCLType,
  LResources
  {$ENDIF}
  {$IFDEF FPC},zipper{$ELSE}, zeZipper{$ENDIF};

type
  TZXLSXFileItem = record
    name: string;     //путь к файлу
    original: string; //исходная строка
    ftype: integer;   //тип контента
  end;

  TZXLSXFileArray = array of TZXLSXFileItem;

  TZXLSXRelations = record
    id: string;       //rID
    ftype: integer;   //тип ссылки
    target: string;   //ссылка на файла
    fileid: integer;  //ссылка на запись
    name: string;     //имя листа
    state: byte;      //состояние
    sheetid: integer; //номер листа
  end;

  TZXLSXRelationsArray = array of TZXLSXRelations;

///Differential Formatting

  TZXLSXDiffBorderItemStyle = class(TPersistent)
  private
    FUseStyle: boolean;             //заменять ли стиль
    FUseColor: boolean;             //заменять ли цвет
    FColor: TColor;                 //цвет линий
    FLineStyle: TZBorderType;       //стиль линий
    FWeight: byte;
  protected
  public
    constructor Create();
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    property UseStyle: boolean read FUseStyle write FUseStyle;
    property UseColor: boolean read FUseColor write FUseColor;
    property Color: TColor read FColor write FColor;
    property LineStyle: TZBorderType read FLineStyle write FLineStyle;
    property Weight: byte read FWeight write FWeight;
  end;

  TZXLSXDiffBorder = class(TPersistent)
  private
    FBorder: array [0..5] of TZXLSXDiffBorderItemStyle;
    procedure SetBorder(Num: integer; Const Value: TZXLSXDiffBorderItemStyle);
    function GetBorder(Num: integer): TZXLSXDiffBorderItemStyle;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Clear();
    procedure Assign(Source: TPersistent);override;
    property Border[Num: integer]: TZXLSXDiffBorderItemStyle read GetBorder write SetBorder; default;
  end;

  //Итем для дифференцированного форматирования
  //TODO:
  //      возможно, для excel xml тоже понадобится (перенести?)
  TZXLSXDiffFormattingItem = class(TPersistent)
  private
    FUseFont: boolean;              //заменять ли шрифт
    FUseFontColor: boolean;         //заменять ли цвет шрифта
    FUseFontStyles: boolean;        //заменять ли стиль шрифта
    FFontColor: TColor;             //цвет шрифта
    FFontStyles: TFontStyles;       //стиль шрифта
    FUseBorder: boolean;            //заменять ли рамку
    FBorders: TZXLSXDiffBorder;     //Что менять в рамке
    FUseFill: boolean;              //заменять ли заливку
    FUseCellPattern: boolean;       //Заменять ли тип заливки
    FCellPattern: TZCellPattern;    //тип заливки
    FUseBGColor: boolean;           //заменять ли цвет заливки
    FBGColor: TColor;               //цвет заливки
    FUsePatternColor: boolean;      //Заменять ли цвет шаблона заливки
    FPatternColor: TColor;          //Цвет шаблона заливки
  protected
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    property UseFont: boolean read FUseFont write FUseFont;
    property UseFontColor: boolean read FUseFontColor write FUseFontColor;
    property UseFontStyles: boolean read FUseFontStyles write FUseFontStyles;
    property FontColor: TColor read FFontColor write FFontColor;
    property FontStyles: TFontStyles read FFontStyles write FFontStyles;
    property UseBorder: boolean read FUseBorder write FUseBorder;
    property Borders: TZXLSXDiffBorder read FBorders write FBorders;
    property UseFill: boolean read FUseFill write FUseFill;
    property UseCellPattern: boolean read FUseCellPattern write FUseCellPattern;
    property CellPattern: TZCellPattern read FCellPattern write FCellPattern;
    property UseBGColor: boolean read FUseBGColor write FUseBGColor;
    property BGColor: TColor read FBGColor write FBGColor;
    property UsePatternColor: boolean read FUsePatternColor write FUsePatternColor;
    property PatternColor: TColor read FPatternColor write FPatternColor;
  end;

  //Дифференцированное форматирование
  TZXLSXDiffFormatting = class(TPersistent)
  private
    FCount: integer;
    FMaxCount: integer;
    FItems: array of TZXLSXDiffFormattingItem;
  protected
    function GetItem(num: integer): TZXLSXDiffFormattingItem;
    procedure SetItem(num: integer; const Value: TZXLSXDiffFormattingItem);
    procedure SetCount(ACount: integer);
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Add();
    procedure Assign(Source: TPersistent); override;
    procedure Clear();
    property Count: integer read FCount;
    property Items[num: integer]: TZXLSXDiffFormattingItem read GetItem write SetItem; default;
  end;

///end Differential Formatting

  //List of cell number formats (date/numbers/currencies etc formats)
  TZEXLSXNumberFormats = class
  private
    FFormatsCount: integer;
    FFormats: array of string; //numFmts (include default formats)
    FStyleFmtID: array of integer;
    FStyleFmtIDCount: integer;
  protected
    function GetFormat(num: integer): string;
    procedure SetFormat(num: integer; const value: string);
    function GetStyleFMTID(num: integer): integer;
    procedure SetStyleFMTID(num: integer; const value: integer);
    procedure SetStyleFMTCount(value: integer);
  public
    constructor Create();
    destructor Destroy(); override;

    procedure ReadNumFmts(const xml: TZsspXMLReaderH);

    function IsDateFormat(StyleNum: integer): boolean;

    function FindFormatID(const value: string): integer;

    property FormatsCount: integer read FFormatsCount;
    property Format[num: integer]: string read GetFormat write SetFormat; default;
    property StyleFMTID[num: integer]: integer read GetStyleFMTID write SetStyleFMTID;
    property StyleFMTCount: integer read FStyleFmtIDCount write SetStyleFMTCount;
  end;

  TZEXLSXReadHelper = class
  private
    FDiffFormatting: TZXLSXDiffFormatting;
    FNumberFormats: TZEXLSXNumberFormats;
  protected
    procedure SetDiffFormatting(const Value: TZXLSXDiffFormatting);
  public
    constructor Create();
    destructor Destroy(); override;
    property DiffFormatting: TZXLSXDiffFormatting read FDiffFormatting write SetDiffFormatting;
    property NumberFormats: TZEXLSXNumberFormats read FNumberFormats;
  end;

  //Store link item
  type TZEXLSXHyperLinkItem = record
    RID: integer;
    RelType: integer;
    CellRef: string;
    Target: string;
    ScreenTip: string;
    TargetMode: string;
  end;

  { TZEXLSXWriteHelper }

  TZEXLSXWriteHelper = class
  private
    FHyperLinks: array of TZEXLSXHyperLinkItem;
    FHyperLinksCount: integer;
    FMaxHyperLinksCount: integer;
    FCurrentRID: integer;                      //Current rID number (for HyperLinks/comments etc)
    FisHaveComments: boolean;                  //Is Need create comments*.xml?
    FisHaveDrawings: boolean;                  //Is Need create drawings*.xml?
    FSheetHyperlinksArray: array of integer;
    FSheetHyperlinksCount: integer;
  protected
    function GenerateRID(): integer;
  public
    constructor Create();
    destructor Destroy(); override;
    procedure AddHyperLink(const ACellRef, ATarget, AScreenTip, ATargetMode: string);
    function AddDrawing(const ATarget: string): integer;
    procedure WriteHyperLinksTag(const xml: TZsspXMLWriterH);
    function CreateSheetRels(const Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
    procedure AddSheetHyperlink(PageNum: integer);
    function IsSheetHaveHyperlinks(PageNum: integer): boolean;
    procedure Clear();
    property HyperLinksCount: integer read FHyperLinksCount;
    property isHaveComments: boolean read FisHaveComments write FisHaveComments; //Is need create comments*.xml?
    property isHaveDrawings: boolean read FisHaveDrawings write FisHaveDrawings; //Is need create drawings*.xml?
  end;


function ReadXSLXPath(var XMLSS: TZEXMLSS; DirName: string): integer; deprecated {$IFDEF FPC}'Use ReadXLSXPath!'{$ENDIF};
function ReadXLSXPath(var XMLSS: TZEXMLSS; DirName: string): integer;

function ReadXSLX(var XMLSS: TZEXMLSS; FileName: string): integer; deprecated {$IFDEF FPC}'Use ReadXLSX!'{$ENDIF};
function ReadXLSX(var XMLSS: TZEXMLSS; FileName: string): integer;

function SaveXmlssToXLSXPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers: array of integer;
                             const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
function SaveXmlssToXLSXPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers: array of integer;
                             const SheetsNames: array of string): integer; overload;
function SaveXmlssToXLSXPath(var XMLSS: TZEXMLSS; PathName: string): integer; overload;

function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string): integer; overload;
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string): integer; overload;

function ExportXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string;
  const SheetsNumbers: array of integer;
  const SheetsNames: array of string;
  TextConverter: TAnsiToCPConverter;
  CodePageName: string;
  BOM: AnsiString = '';
  AllowUnzippedFolder: boolean = false): integer;

//Дополнительные функции, на случай чтения отдельного файла
function ZEXSLXReadTheme(var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer): boolean;
function ZEXSLXReadContentTypes(var Stream: TStream; var FileArray: TZXLSXFileArray; var FilesCount: integer): boolean;
function ZEXSLXReadSharedStrings(var Stream: TStream; out StrArray: TStringDynArray; out StrCount: integer): boolean;
function ZEXSLXReadStyles(var XMLSS: TZEXMLSS; var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer; ReadHelper: TZEXLSXReadHelper): boolean;
function ZE_XSLXReadRelationships(var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer; var isWorkSheet: boolean; needReplaceDelimiter: boolean): boolean;
function ZEXSLXReadWorkBook(var XMLSS: TZEXMLSS; var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer): boolean;
function ZEXSLXReadSheet(var XMLSS: TZEXMLSS; var Stream: TStream; const SheetName: string; var StrArray: TStringDynArray; StrCount: integer; var Relations: TZXLSXRelationsArray; RelationsCount: integer; ReadHelper: TZEXLSXReadHelper): boolean;
function ZEXSLXReadComments(var XMLSS: TZEXMLSS; var Stream: TStream): boolean;

//Дополнительные функции для экспорта отдельных файлов
function ZEXLSXCreateStyles(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateWorkBook(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                              const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
function ZEXLSXCreateSheet(var XMLSS: TZEXMLSS; Stream: TStream; SheetNum: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring; const WriteHelper: TZEXLSXWriteHelper): integer;
function ZEXLSXCreateContentTypes(var XMLSS: TZEXMLSS; Stream: TStream; PageCount: integer; CommentCount: integer; const PagesComments: TIntegerDynArray;
                                  TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring;
                                  const WriteHelper: TZEXLSXWriteHelper): integer;
function ZEXLSXCreateRelsMain(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateSharedStrings(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateDocPropsApp(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateDocPropsCore(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;

{$IFDEF ZUSE_DRAWINGS}
function ZEXLSXCreateDrawing(var XMLSS: TZEXMLSS; Stream: TStream; Drawing: TZEDrawing; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
function ZEXLSXCreateDrawingRels(var XMLSS: TZEXMLSS; Stream: TStream; Drawing: TZEDrawing; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
{$ENDIF}

procedure ZEAddRelsRelation(xml: TZsspXMLWriterH; const rid: string; ridType: integer; const Target: string; const TargetMode: string = '');

implementation

uses
{$IfDef DELPHI_UNICODE}
  AnsiStrings,  // AnsiString targeted overloaded versions of Pos, Trim, etc
{$EndIf}
  StrUtils,
  Math,
  zenumberformats,
  zearchhelper;

{$IFDEF ZUSE_CONDITIONAL_FORMATTING}
const
  ZETag_conditionalFormatting   = 'conditionalFormatting';
  ZETag_cfRule                  = 'cfRule';
  ZETag_formula                 = 'formula';
{$ENDIF}

const
  SCHEMA_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006';
  SCHEMA_DOC_REL = SCHEMA_DOC + '/relationships';
  SCHEMA_PACKAGE = 'http://schemas.openxmlformats.org/package/2006';
  SCHEMA_PACKAGE_REL = SCHEMA_PACKAGE + '/relationships';
  SCHEMA_SHEET_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

  REL_TYPE_WORKSHEET = 0;
  REL_TYPE_STYLES = 1;
  REL_TYPE_SHARED_STR = 2;
  REL_TYPE_DOC = 3;
  REL_TYPE_CORE_PROP = 4;
  REL_TYPE_EXT_PROP = 5;
  REL_TYPE_HYPERLINK = 6;
  REL_TYPE_COMMENTS = 7;
  REL_TYPE_VML_DRAWING = 8;
  REL_TYPE_DRAWING = 9;

type
  //шрифт
  TZEXLSXFont = record
    name: string;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strike: boolean;
    charset: integer;
    color: TColor;
    ColorType: byte;
    LumFactor: double;
    fontsize: integer;
  end;

  TZEXLSXFontArray = array of TZEXLSXFont; //массив шрифтов

  TStreamList = class(TList)
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    function GetItem(AIndex: Integer): TStream;
  public
    property Items[AIndex: Integer]: TStream read GetItem;
  end;


  //Для распаковки в поток

  procedure XLSXSortRelationArray(var arr: TZXLSXRelationsArray; count: integer); forward;

  { TXSLXZipHelper }
type
  TXSLXZipHelper = class
  private
    FXMLSS: TZEXMLSS;
    FRetCode: integer;
    FFileArray: TZXLSXFileArray;
    FFilesCount: integer;
    FFileType: integer;
    FRelationsArray: array of TZXLSXRelationsArray;
    FRelationsCounts: array of integer;
    FRelationsCount: integer;
    FSheetRelationNumber: integer;
    FFileNumber: integer;
    FStrArray: TStringDynArray;
    FArchFiles: TStringDynArray;
    FArchFilesCount: integer;
    FMaxArchFilesCount: integer;
    FStrCount: integer;
    FListName: string;
    FThemaColor: TIntegerDynArray;
    FThemaColorCount: integer;
    FSheetRelations: TZXLSXRelationsArray;
    FSheetRelationsCount: integer;
    FNeedReadComments: boolean;
    FReadHelper: TZEXLSXReadHelper;
    function GetFileItem(num: integer): TZXLSXFileItem;
    function GetRelationsCounts(num: integer): integer;
    function GetRelationsArray(num: integer): TZXLSXRelationsArray;
    function GetArchFileItem(num: integer): string;
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure SortRelationArray();
    procedure DoCreateOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
    procedure DoDoneOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
    procedure AddArchFile(const FileName: string);
    procedure ClearCurrentSheetRelations();
    function GetCurrentPageCommentsNumber(): integer;
    procedure AddToFiles(const sname: string; ftype: integer);
    property XMLSS: TZEXMLSS read FXMLSS write FXMLSS;                  //хранилище
    property ListName: string read FListName write FListName;           //имя листа
    property RetCode: integer read FRetCode;                            //код ошибки
    property FileArray[num: integer]: TZXLSXFileItem read GetFileItem;  //файлы
    property ArchFile[num: integer]: string read GetArchFileItem;
    property ArchFilesCount: integer read FArchFilesCount;              //кол-во файлов в архиве
    property FilesCount: integer read FFilesCount;                      //кол-во файлов
    property FileType: integer read FFileType write FFileType;          //тип файла
    property FileNumber: integer read FFileNumber write FFileNumber;    //номер
    property RelationsCounts[num: integer]: integer read GetRelationsCounts; //кол-во в отношених
    property RelationsCount: integer read FRelationsCount;                   //кол-во отношений
    property RelationsArray[num: integer]: TZXLSXRelationsArray read GetRelationsArray; //отношения
    property SheetRelationNumber: integer read FSheetRelationNumber;
    property isNeedReadComments: boolean read FNeedReadComments;        //нужно ли читать примечания
  end;

constructor TXSLXZipHelper.Create();
begin
  inherited;
  FXMLSS := nil;
  FRetCode := 0;
  FFilesCount := 0;
  SetLength(FFileArray, 0);
  FFileType := -1;
  FRelationsCount := 0;
  FSheetRelationNumber := -1;
  SetLength(FRelationsArray, 0);
  SetLength(FRelationsCounts, 0);
  FFileNumber := -1;
  SetLength(FStrArray, 0);
  FStrCount := 0;
  FListName := '';
  FThemaColorCount := 0;
  FArchFilesCount := 0;
  FMaxArchFilesCount := 25;
  FSheetRelationsCount := 0;
  SetLength(FArchFiles, FMaxArchFilesCount);
  FNeedReadComments := false;
  FReadHelper := TZEXLSXReadHelper.Create();
end;

destructor TXSLXZipHelper.Destroy();
var
  i: integer;

begin
  SetLength(FFileArray, 0);
  FFileArray := nil;
  for i := 0 to FRelationsCount - 1 do
  begin
    Setlength(FRelationsArray[i], 0);
    FRelationsArray[i] := nil;
  end;
  SetLength(FRelationsCounts, 0);
  FRelationsCounts := nil;
  SetLength(FRelationsArray, 0);
  FRelationsArray := nil;
  SetLength(FStrArray, 0);
  FStrArray := nil;
  SetLength(FArchFiles, 0);
  FArchFiles := nil;
  SetLength(FThemaColor, 0);
  FThemaColor := nil;
  SetLength(FSheetRelations, 0);
  FSheetRelations := nil;
  FreeAndNil(FReadHelper);
  inherited;
end;

procedure TXSLXZipHelper.SortRelationArray();
begin
 XLSXSortRelationArray(FRelationsArray[SheetRelationNumber], RelationsCounts[SheetRelationNumber]);
end;

function TXSLXZipHelper.GetArchFileItem(num: integer): string;
begin
  result := '';
  if ((num >= 0) and (num < FArchFilesCount)) then
    result := FArchFiles[num];
end;

procedure TXSLXZipHelper.AddArchFile(const FileName: string);
begin
  inc(FArchFilesCount);
  if (FArchFilesCount >= FMaxArchFilesCount) then
  begin
    FMaxArchFilesCount := FArchFilesCount + 20;
    SetLength(FArchFiles, FMaxArchFilesCount);
  end;
  FArchFiles[FArchFilesCount - 1] := FileName;
end;

procedure TXSLXZipHelper.AddToFiles(const sname: string; ftype: integer);
begin
  inc(FFilesCount);
  SetLength(FFileArray, FFilesCount);
  FFileArray[FFilesCount - 1].name := sname;
  FFileArray[FFilesCount - 1].original := sname;
  FFileArray[FFilesCount - 1].ftype := ftype;
end;

function TXSLXZipHelper.GetFileItem(num: integer): TZXLSXFileItem;
begin
  if ((num >= 0) and (num < FilesCount)) then
    result := FFileArray[num];
end;

function TXSLXZipHelper.GetRelationsCounts(num: integer): integer;
begin
  if ((num >= 0) and (num < FRelationsCount)) then
    result := FRelationsCounts[num];
end;

function TXSLXZipHelper.GetRelationsArray(num: integer): TZXLSXRelationsArray;
begin
  if ((num >= 0) and (num < FRelationsCount)) then
    result := FRelationsArray[num];
end;

//Обнуляет SheetRelations для текущего листа (если небыло комментариев и/или ссылок)
procedure TXSLXZipHelper.ClearCurrentSheetRelations();
begin
  FSheetRelationsCount := 0;
  FNeedReadComments := false;
end;

//Возвращат номер файла с примечаниями для текущего листа
//Если возвращает отрицательное число - примечания не обноружены
function TXSLXZipHelper.GetCurrentPageCommentsNumber(): integer;
var
  i, l: integer;
  s: string;
  b: boolean;

begin
  result := -1;
  b := false;

  for i := 0 to FSheetRelationsCount - 1 do
  if (FSheetRelations[i].ftype = 7) then
  begin
    s := FSheetRelations[i].target;
    b := true;
    break;
  end;

  //Если найдены примечания
  if (b) then
  begin
    l := length(s);
    if (l >= 3) then
      if ((s[1] = '.') and (s[2] = '.')) then
        delete(s, 1, 3);
    for i := 0 to FFilesCount - 1 do
    if (FFileArray[i].ftype = 7) then
      if (pos(s, FFileArray[i].name) <> 0) then
      begin
        result := i;
        break;
      end;
  end;
end; //GetCurrentPageCommentsNumber

procedure TXSLXZipHelper.DoCreateOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
begin
  AStream := TMemoryStream.Create();
end;

procedure TXSLXZipHelper.DoDoneOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
var
  b: boolean;
  j, k: integer;

  function _FileExists(const _name: string; var _num: integer): boolean;
  var
    i: integer;
    s: string;

  begin
    s := UpperCase(_name);
    result := false;
    for i := 0 to FArchFilesCount - 1 do
      if (UpperCase(FArchFiles[i]) = s) then
      begin
        result := true;
        _num := i;
        exit;
      end;
  end; //_FileExists

  procedure _AddFile(_num, _ftype: integer);
  var
    s: string;

  begin
    SetLength(FFileArray, FFilesCount + 1);
    s := FArchFiles[_num];
    FFileArray[FFilesCount].original := s;
    FFileArray[FFilesCount].name := s;
    FFileArray[FFilesCount].ftype := _ftype;
    inc(FFilesCount);
  end; //_AddFile

  procedure _AddAdditionalFiles();
  var
    b: boolean;
    i: integer;
    s: string;

  begin
    b := false;
    for i := 0 to FFilesCount - 1 do
    if (FFileArray[i].ftype = 3) then
    begin
      b := true;
      break;
    end;

    if (not b) then
    begin
      s := '_rels/.rels';
      if (_FileExists(s, i)) then
        _AddFile(i, 3);

      s := 'xl/_rels/workbook.xml.rels';
      if (_FileExists(s, i)) then
        _AddFile(i, 3);
    end; //if

    //Styles
    b := false;
    for i := 0 to self.FFilesCount - 1 do
    if (FFileArray[i].ftype = 1) then
    begin
      b := true;
      break;
    end;

    if (not b) then
    begin
      s := 'styles.xml';
      if (_FileExists(s, i)) then
        _AddFile(i, 1);
    end; //if
  end; //_AddAdditionalFiles

begin
  if (Assigned(AStream)) then
  begin
    try
      AStream.Position := 0;
      b := false;
      //Список файлов
      if (FileType = -2) then
      begin
        if (not ZEXSLXReadContentTypes(AStream,  FFileArray, FFilesCount)) then
          FRetCode := FRetCode or 3;
        for j := 0 to FFilesCount - 1 do
        if (Length(FFileArray[j].name) > 0) then
          if (FFileArray[j].name[1] = '/') then
            delete(FFileArray[j].name, 1, 1);

        //Если в контент забыли добавить отношения, но они есть в архиве - добавляем.
        //В теории такого быть не должно, но мало ли
        _AddAdditionalFiles();

      end else
      //relationships
      if (FileType = 3) then
      begin
        SetLength(FRelationsArray, FRelationsCount + 1);
        SetLength(FRelationsCounts, FRelationsCount + 1);

        if (not ZE_XSLXReadRelationships(AStream, FRelationsArray[FRelationsCount], FRelationsCounts[FRelationsCount], b, true)) then
        begin
          FRetCode := FRetCode or 4;
          exit;
        end;
        if (b) then
        begin
          FSheetRelationNumber := FRelationsCount;
          for j := 0 to FRelationsCounts[FRelationsCount] - 1 do
          if (FRelationsArray[FRelationsCount][j].ftype = 0) then
            for k := 0 to FilesCount - 1 do
            if (FRelationsArray[FRelationsCount][j].fileid < 0) then
            begin
              if ((pos(FRelationsArray[FRelationsCount][j].target, FFileArray[k].original)) > 0) then
              begin
                FRelationsArray[RelationsCount][j].fileid := k;
                break;
              end;
            end;
        end; //if
        inc(FRelationsCount);
      end else
      //sharedstrings
      if (FileType = 4) then
      begin
        if (not ZEXSLXReadSharedStrings(AStream, FStrArray, FStrCount)) then
          FRetCode := FRetCode or 3;
      end else
      //стили
      if (FileType = 1) then
      begin
        if (not ZEXSLXReadStyles(FXMLSS, AStream, FThemaColor, FThemaColorCount, FReadHelper)) then
          FRetCode := FRetCode or 5;
      end else
      //Workbook
      if (FileType = 2) then
      begin
        if (not ZEXSLXReadWorkBook(FXMLSS, AStream, FRelationsArray[FSheetRelationNumber], FRelationsCounts[FSheetRelationNumber])) then
          FRetCode := FRetCode or 3;
      end else
      //Sheet
      if (FileType = 0) then
      begin
        if (not ZEXSLXReadSheet(FXMLSS, AStream, ListName, FStrArray, FStrCount, FSheetRelations, FSheetRelationsCount, FReadHelper)) then
          FRetCode := FRetCode or 4;
      end else
      //Тема
      if (FileType = 5) then
      begin
        if (not ZEXSLXReadTheme(AStream, FThemaColor, FThemaColorCount)) then
          FRetCode := FRetCode or 6;
      end else
      //SheetRelations для конкретного листа (чтение ссылок и примечаний)
      if (FileType = 111) then
      begin
        if (not ZE_XSLXReadRelationships(AStream, FSheetRelations, FSheetRelationsCount, b, false)) then
          FRetCode := FRetCode or 5
        else
          FNeedReadComments := true;
      end else
      //Примечания
      if (FileType = 113) then
      begin
        if (isNeedReadComments) then
          if (not ZEXSLXReadComments(FXMLSS, AStream)) then
            FRetCode := FRetCode or 6;
      end;

    finally
      FreeAndNil(AStream)
    end;
  end;
end;


////////////////////////////////////////////////////////////////////////////////
///                   Differential Formatting
////////////////////////////////////////////////////////////////////////////////

////:::::::::::::  TZXLSXDiffFormating :::::::::::::::::////

constructor TZXLSXDiffFormatting.Create();
var
  i: integer;

begin
  FCount := 0;
  FMaxCount := 20;
  SetLength(FItems, FMaxCount);
  for i := 0 to FMaxCount - 1 do
    FItems[i] := TZXLSXDiffFormattingItem.Create();
end;

destructor TZXLSXDiffFormatting.Destroy();
var
  i: integer;

begin
  for i := 0 to FMaxCount - 1 do
    if (Assigned(FItems[i])) then
      FreeAndNil(FItems[i]);
  inherited;
end;

procedure TZXLSXDiffFormatting.Add();
begin
  SetCount(FCount + 1);
  FItems[FCount - 1].Clear();
end;

procedure TZXLSXDiffFormatting.Assign(Source: TPersistent);
var
  df: TZXLSXDiffFormatting;
  i: integer;
  b: boolean;

begin
  b := true;

  if (Assigned(Source)) then
    if (Source is TZXLSXDiffFormatting) then
    begin
      b := false;
      df := Source as TZXLSXDiffFormatting;
      SetCount(df.Count);
      for i := 0 to Count - 1 do
        FItems[i].Assign(df[i]);

    end;

  if (b) then
    inherited;
end; //Assign

procedure TZXLSXDiffFormatting.SetCount(ACount: integer);
var
  i: integer;

begin
  if (ACount >= FMaxCount) then
  begin
    FMaxCount := ACount + 20;
    SetLength(FItems, FMaxCount);
    for i := FCount to FMaxCount - 1 do
      FItems[i] := TZXLSXDiffFormattingItem.Create();
  end;
  FCount := ACount;
end;

procedure TZXLSXDiffFormatting.Clear();
begin
  FCount := 0;
end;

function TZXLSXDiffFormatting.GetItem(num: integer): TZXLSXDiffFormattingItem;
begin
  if ((num >= 0) and (num < Count)) then
    result := FItems[num]
  else
    result := nil;
end;

procedure TZXLSXDiffFormatting.SetItem(num: integer; const Value: TZXLSXDiffFormattingItem);
begin
  if ((num >= 0) and (num < Count)) then
    if (Assigned(Value)) then
      FItems[num].Assign(Value);
end;

////:::::::::::::  TZXLSXDiffBorderItemStyle :::::::::::::::::////

constructor TZXLSXDiffBorderItemStyle.Create();
begin
  Clear();
end;

procedure TZXLSXDiffBorderItemStyle.Assign(Source: TPersistent);
var
  bs: TZXLSXDiffBorderItemStyle;
  b: boolean;

begin
  b := true;

  if (Assigned(Source)) then
    if (Source is TZXLSXDiffBorderItemStyle) then
    begin
      b := false;
      bs := Source as TZXLSXDiffBorderItemStyle;

      FUseStyle := bs.UseStyle;
      FUseColor := bs.UseColor;
      FColor := bs.Color;
      FWeight := bs.Weight;
      FLineStyle := bs.LineStyle;
    end;

  if (b) then
    inherited;
end; //Assign

procedure TZXLSXDiffBorderItemStyle.Clear();
begin
  FUseStyle := false;
  FUseColor := false;
  FColor := clBlack;
  FWeight := 1;
  FLineStyle := ZENone;
end;

////::::::::::::: TZXLSXDiffBorder :::::::::::::::::////

constructor TZXLSXDiffBorder.Create();
var
  i: integer;

begin
  for i := 0 to 5 do
    FBorder[i] := TZXLSXDiffBorderItemStyle.Create();
  Clear();
end;

destructor TZXLSXDiffBorder.Destroy();
var
  i: integer;

begin
  for i := 0 to 5 do
    FreeAndNil(FBorder[i]);

  inherited;
end;

procedure TZXLSXDiffBorder.Assign(Source: TPersistent);
var
  brd: TZXLSXDiffBorder;
  b: boolean;
  i: integer;

begin
  b := true;

  if (Assigned(Source)) then
    if (Source is TZXLSXDiffBorder) then
    begin
      b := false;
      brd := Source as TZXLSXDiffBorder;
      for i := 0 to 5 do
        FBorder[i].Assign(brd[i]);
    end;

  if (b) then
    inherited;
end; //Assign

procedure TZXLSXDiffBorder.Clear();
var
  i: integer;

begin
  for i := 0 to 5 do
    FBorder[i].Clear();
end;

function TZXLSXDiffBorder.GetBorder(Num: integer): TZXLSXDiffBorderItemStyle;
begin
  result := nil;
  if ((num >= 0 ) and (num < 6)) then
    result := FBorder[num];
end;

procedure TZXLSXDiffBorder.SetBorder(Num: integer; const Value: TZXLSXDiffBorderItemStyle);
begin
  if ((num >= 0) and (num < 6)) then
    if (Assigned(Value)) then
      FBorder[num].Assign(Value);
end;

////::::::::::::: TZXLSXDiffFormatingItem :::::::::::::::::////

constructor TZXLSXDiffFormattingItem.Create();
begin
  FBorders := TZXLSXDiffBorder.Create();
  Clear();
end;

destructor TZXLSXDiffFormattingItem.Destroy();
begin
  FreeAndNil(FBorders);
  inherited;
end;

procedure TZXLSXDiffFormattingItem.Assign(Source: TPersistent);
var
  dxfItem: TZXLSXDiffFormattingItem;
  b: boolean;

begin
  b := true;
  if (Assigned(Source)) then
    if (Source is TZXLSXDiffFormattingItem) then
    begin
      b := false;
      dxfItem := Source as TZXLSXDiffFormattingItem;

      FUseFont := dxfItem.UseFont;
      FUseFontColor := dxfItem.UseFontColor;
      FUseFontStyles := dxfItem.UseFontStyles;
      FFontColor := dxfItem.FontColor;
      FFontStyles := dxfItem.FontStyles;
      FUseBorder := dxfItem.UseBorder;
      FBorders.Assign(dxfItem.Borders);
      FUseFill := dxfItem.UseFill;
      FUseCellPattern := dxfItem.UseCellPattern;
      FCellPattern := dxfItem.CellPattern;
      FUseBGColor := dxfItem.UseBGColor;
      FBGColor := dxfItem.BGColor;
      FUsePatternColor := dxfItem.UsePatternColor;
      FPatternColor := dxfItem.PatternColor;
    end;

  if (b) then
    inherited;
end; //Assign

procedure TZXLSXDiffFormattingItem.Clear();
begin
  FUseFont := false;
  FUseFontColor := false;
  FUseFontStyles := false;
  FFontColor := clBlack;
  FFontStyles := [];
  FUseBorder := false;
  FBorders.Clear();
  FUseFill := false;
  FUseCellPattern := false;
  FCellPattern := ZPNone;
  FUseBGColor := false;
  FBGColor := clWindow;
  FUsePatternColor := false;
  FPatternColor := clWindow;
end; //Clear

// END Differential Formatting
////////////////////////////////////////////////////////////////////////////////

////::::::::::::: TZEXLSXNumberFormats :::::::::::::::::////

constructor TZEXLSXNumberFormats.Create();
var
  i: integer;
  //_date_nums: set of byte;

begin
  FStyleFmtIDCount := 0;
  FFormatsCount := 164;
  SetLength(FFormats, FFormatsCount);
  for i := 0 to FFormatsCount - 1 do
    FFormats[i] := '';

  //Some "Standart" formats for xlsx:
  FFormats[1] := '0';
  FFormats[2] := '0.00';
  FFormats[3] := '#,##0';
  FFormats[4] := '#,##0.00';
  FFormats[5] := '$#,##0;\-$#,##0';
  FFormats[6] := '$#,##0;[Red]\-$#,##0';
  FFormats[7] := '$#,##0.00;\-$#,##0.00';
  FFormats[8] := '$#,##0.00;[Red]\-$#,##0.00';
  FFormats[9] := '0%';
  FFormats[10] := '0.00%';
  FFormats[11] := '0.00E+00';
  FFormats[12] := '# ?/?';
  FFormats[13] := '# ??/??';
  FFormats[14] := 'mm-dd-yy';
  FFormats[15] := 'd-mmm-yy';
  FFormats[16] := 'd-mmm';
  FFormats[17] := 'mmm-yy';
  FFormats[18] := 'h:mm AM/PM';
  FFormats[19] := 'h:mm:ss AM/PM';
  FFormats[20] := 'h:mm';
  FFormats[21] := 'h:mm:ss';
  FFormats[22] := 'm/d/yy h:mm';
  FFormats[27] := '[$-404]e/m/d';
  FFormats[37] := '#,##0 ;(#,##0)';
  FFormats[38] := '#,##0 ;[Red](#,##0)';
  FFormats[39] := '#,##0.00;(#,##0.00)';
  FFormats[40] := '#,##0.00;[Red](#,##0.00)';
  FFormats[44] := '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)';
  FFormats[45] := 'mm:ss';
  FFormats[46] := '[h]:mm:ss';
  FFormats[47] := 'mmss.0';
  FFormats[48] := '##0.0E+0';
  FFormats[49] := '@';

  FFormats[30] := 'm/d/yy';
  //27..36 - date
  //50..58 - date

  FFormats[59] := 't0';
  FFormats[60] := 't0.00';
  FFormats[61] := 't#,##0';
  FFormats[62] := 't#,##0.00';
  FFormats[67] := 't0%';
  FFormats[68] := 't0.00%';
  FFormats[69] := 't# ?/?';

  FFormats[70] := 't# ??/??';

  FFormats[81] := 'd/m/bb';


  //_date_nums := [14..22, 27..36, 45..47, 50..58, 71..76, 78..81];
end;

destructor TZEXLSXNumberFormats.Destroy();
begin
  SetLength(FFormats, 0);
  SetLength(FStyleFmtID, 0);
  inherited;
end;

//Find format in formats. Return -1 if format not foud.
//INPUT
//  const value: string - format
//RETURN
//      integer - >= 0 - index number if store.
//                 -1 - not found
function TZEXLSXNumberFormats.FindFormatID(const value: string): integer;
var
  i: integer;

begin
  Result := -1;
  for i := 0 to FFormatsCount - 1 do
    if (FFormats[i] = value) then
    begin
      Result := i;
      break;
    end;
end;

function TZEXLSXNumberFormats.GetFormat(num: integer): string;
begin
  Result := '';
  if ((num >= 0) and (num < FFormatsCount)) then
    Result := FFormats[num];
end;

procedure TZEXLSXNumberFormats.SetFormat(num: integer; const value: string);
var
  i: integer;

begin
  if ((num >= 0) and (num < FFormatsCount)) then
    FFormats[num] := value
  else
    if (num >= 0) then
    begin
      SetLength(FFormats, num + 1);
      for i := FFormatsCount to num do
        FFormats[i] := '';
      FFormats[num] := value;
      FFormatsCount := num + 1;
    end;
end;

function TZEXLSXNumberFormats.GetStyleFMTID(num: integer): integer;
begin
  if ((num >= 0) and (num < FStyleFmtIDCount)) then
    Result := FStyleFmtID[num]
  else
    Result := 0;
end;

function TZEXLSXNumberFormats.IsDateFormat(StyleNum: integer): boolean;
var
  _fmtId: integer;

begin
  Result := false;

  if ((StyleNum >= 0) and (StyleNum < FStyleFmtIDCount)) then
    _fmtId := FStyleFmtID[StyleNum]
  else
    exit;

  //If default fmtID
  if ((_fmtId >= 0) and (_fmtId < 100)) then
    Result := _fmtId in [14..22, 27..36, 45..47, 50..58, 71..76, 78..81]
  else
    Result := GetXlsxNumberFormatType(FFormats[_fmtId]) = ZE_NUMFORMAT_IS_DATETIME;
end;

procedure TZEXLSXNumberFormats.ReadNumFmts(const xml: TZsspXMLReaderH);
var
  t: integer;

begin
  while (not((xml.TagName = 'numFmts') and (xml.TagType = TAG_TYPE_END))) and (not xml.Eof()) do
  begin
    xml.ReadTag();

    if (xml.TagName = 'numFmt') then
      if (TryStrToInt(xml.Attributes['numFmtId'], t)) then
        Format[t] := xml.Attributes['formatCode'];
  end;
end;

procedure TZEXLSXNumberFormats.SetStyleFMTID(num: integer; const value: integer);
begin
  if ((num >= 0) and (num < FStyleFmtIDCount)) then
    FStyleFmtID[num] := value;
end;

procedure TZEXLSXNumberFormats.SetStyleFMTCount(value: integer);
begin
  if (value >= 0) then
  begin
    if (value > FStyleFmtIDCount) then
      SetLength(FStyleFmtID, value);
    FStyleFmtIDCount := value
  end;
end;

////::::::::::::: TZEXLSXReadHelper :::::::::::::::::////

constructor TZEXLSXReadHelper.Create();
begin
  FDiffFormatting := TZXLSXDiffFormatting.Create();
  FNumberFormats := TZEXLSXNumberFormats.Create();
end;

destructor TZEXLSXReadHelper.Destroy();
begin
  FreeAndNil(FDiffFormatting);
  FreeAndNil(FNumberFormats);
  inherited;
end;

procedure TZEXLSXReadHelper.SetDiffFormatting(const Value: TZXLSXDiffFormatting);
begin
  if (Assigned(Value)) then
    FDiffFormatting.Assign(Value);
end;

////::::::::::::: TZEXLSXWriteHelper :::::::::::::::::////

//Generate next RID for references
function TZEXLSXWriteHelper.GenerateRID(): integer;
begin
  inc(FCurrentRID);
  result := FCurrentRID;
end;

constructor TZEXLSXWriteHelper.Create();
begin
  FMaxHyperLinksCount := 10;
  FSheetHyperlinksCount := 0;
  SetLength(FHyperLinks, FMaxHyperLinksCount);
  Clear();
end;

destructor TZEXLSXWriteHelper.Destroy();
begin
  SetLength(FHyperLinks, 0);
  SetLength(FSheetHyperlinksArray, 0);
  inherited Destroy;
end;

//Add hyperlink
//INPUT
//     const ACellRef: string    - cell name (like A1)
//     const ATarget: string     - hyperlink target (http://example.com)
//     const AScreenTip: string  -
//     const ATargetMode: string -
procedure TZEXLSXWriteHelper.AddHyperLink(const ACellRef, ATarget, AScreenTip,
                                          ATargetMode: string);
var
  num: integer;

begin
  num := FHyperLinksCount;
  inc(FHyperLinksCount);

  if (FHyperLinksCount >= FMaxHyperLinksCount) then
  begin
    inc(FMaxHyperLinksCount, 20);
    SetLength(FHyperLinks, FMaxHyperLinksCount);
  end;

  FHyperLinks[num].RID := GenerateRID();
  FHyperLinks[num].RelType := REL_TYPE_HYPERLINK;
  FHyperLinks[num].TargetMode := ATargetMode;
  FHyperLinks[num].CellRef := ACellRef;
  FHyperLinks[num].Target := ATarget;
  FHyperLinks[num].ScreenTip := AScreenTip;
end;

//Add hyperlink
//INPUT
//     const ATarget: string     - drawing target (../drawings/drawing2.xml)
function TZEXLSXWriteHelper.AddDrawing(const ATarget: string): integer;
var
  num: integer;
begin
  num := FHyperLinksCount;
  Inc(FHyperLinksCount);

  if (FHyperLinksCount >= FMaxHyperLinksCount) then
  begin
    Inc(FMaxHyperLinksCount, 20);
    SetLength(FHyperLinks, FMaxHyperLinksCount);
  end;

  FHyperLinks[num].RID := GenerateRID();
  FHyperLinks[num].RelType := REL_TYPE_DRAWING;
  FHyperLinks[num].TargetMode := '';
  FHyperLinks[num].CellRef := '';
  FHyperLinks[num].Target := ATarget;
  FHyperLinks[num].ScreenTip := '';
  Result := FHyperLinks[num].RID;
  FisHaveDrawings := True;
end;

//Writes tag <hyperlinks> .. </hyperlinks>
//INPUT
//     const xml: TZsspXMLWriterH
procedure TZEXLSXWriteHelper.WriteHyperLinksTag(const xml: TZsspXMLWriterH);
var
  i: integer;

begin
  if (FHyperLinksCount > 0) then
  begin
    xml.Attributes.Clear();
    xml.WriteTagNode('hyperlinks', true, true, true);
    for i := 0 to FHyperLinksCount - 1 do
    begin
      xml.Attributes.Clear();
      xml.Attributes.Add('ref', FHyperLinks[i].CellRef);
      xml.Attributes.Add('r:id', 'rId' + IntToStr(FHyperLinks[i].RID));

      if (FHyperLinks[i].ScreenTip <> '') then
        xml.Attributes.Add('tooltip', FHyperLinks[i].ScreenTip);

      xml.WriteEmptyTag('hyperlink', true);

      {
      xml.Attributes.Add('Id', 'rId' + IntToStr(FHyperLinks[i].ID));
      xml.Attributes.Add('Type', ZEXLSXGetRelationName(6));
      xml.Attributes.Add('Target', FHyperLinks[i].Target);
      if (FHyperLinks[i].TargetMode <> '')
         xml.Attributes.Add('TargetMode', FHyperLinks[i].TargetMode);
      }
    end; //for i
    xml.WriteEndTagNode(); //hyperlinks
  end; //if
end; //WriteHyperLinksTag

//Create sheet relations
//INPUT
//  const Stream: TStream
//        TextConverter: TAnsiToCPConverter
//        CodePageName: string
//        BOM: ansistring
//RETURN
function TZEXLSXWriteHelper.CreateSheetRels(const Stream: TStream;
                                            TextConverter: TAnsiToCPConverter;
                                            CodePageName: string; BOM: ansistring
                                           ): integer;
var
  _xml: TZsspXMLWriterH;
  i: integer;

begin
  result := 0;
  _xml := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    _xml.WriteTagNode('Relationships', true, true, false);

    for i := 0 to FHyperLinksCount - 1 do
      ZEAddRelsRelation(_xml,
                        'rId' + IntToStr(FHyperLinks[i].RID),
                        FHyperLinks[i].RelType,
                        FHyperLinks[i].Target,
                        FHyperLinks[i].TargetMode);

    _xml.WriteEndTagNode(); //Relationships

  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //CreateSheetRels

procedure TZEXLSXWriteHelper.AddSheetHyperlink(PageNum: integer);
begin
  SetLength(FSheetHyperlinksArray, FSheetHyperlinksCount + 1);
  FSheetHyperlinksArray[FSheetHyperlinksCount] := PageNum;
  inc(FSheetHyperlinksCount);
end;

function TZEXLSXWriteHelper.IsSheetHaveHyperlinks(PageNum: integer): boolean;
var
  i: integer;

begin
  result := false;
  for i := 0 to FSheetHyperlinksCount - 1 do
    if (FSheetHyperlinksArray[i] = PageNum) then
    begin
      result := true;
      break;
    end;
end;

procedure TZEXLSXWriteHelper.Clear();
begin
  FHyperLinksCount := 0;
  FCurrentRID := 0;
  FisHaveComments := false;
end;

//Возвращает номер Relations из rels
//INPUT
//  const name: string - текст отношения
//RETURN
//      integer - номер отношения. -1 - не определено
function ZEXLSXGetRelationNumber(const name: string): integer;
begin
  result := -1;
  if (name = SCHEMA_DOC_REL + '/worksheet') then
    result := 0
  else
  if (name = SCHEMA_DOC_REL + '/styles') then
    result := 1
  else
  if (name = SCHEMA_DOC_REL + '/sharedStrings') then
    result := 2
  else
  if (name = SCHEMA_DOC_REL + '/officeDocument') then
    result := 3
  else
  if (name = SCHEMA_PACKAGE_REL + '/metadata/core-properties') then
    result := 4
  else
  if (name = SCHEMA_DOC_REL + '/extended-properties') then
    result := 5
  else
  if (name = SCHEMA_DOC_REL + '/hyperlink') then
    result := 6
  else
  if (name = SCHEMA_DOC_REL + '/comments') then
    result := 7
  else
  if (name = SCHEMA_DOC_REL + '/vmlDrawing') then
    result := 8
  else
  if (name = SCHEMA_DOC_REL + '/drawing') then
    result := 9;
end; //ZEXLSXGetRelationNumber

//Возвращает текст Relations для rels
//INPUT
//      num: integer - номер отношения
//RETURN
//      integer - номер отношения. -1 - не определено
function ZEXLSXGetRelationName(num: integer): string;
begin
  result := '';
  case num of
    0: result := SCHEMA_DOC_REL + '/worksheet';
    1: result := SCHEMA_DOC_REL + '/styles';
    2: result := SCHEMA_DOC_REL + '/sharedStrings';
    3: result := SCHEMA_DOC_REL + '/officeDocument';
    4: result := SCHEMA_PACKAGE_REL + '/metadata/core-properties';
    5: result := SCHEMA_DOC_REL + '/extended-properties';
    6: result := SCHEMA_DOC_REL + '/hyperlink';
    7: result := SCHEMA_DOC_REL + '/comments';
    8: result := SCHEMA_DOC_REL + '/vmlDrawing';
    9: result := SCHEMA_DOC_REL + '/drawing';
  end;
end; //ZEXLSXGetRelationName

//BooleanToStr для XLSX //TODO: потом заменить
function XLSXBoolToStr(value: boolean): string;
begin
  if (value) then
    result := 'true'
  else
    result := 'false';
end;

//Читает тему (themeXXX.xml)
//INPUT
//  var Stream: TStream                   - поток чтения
//  var ThemaFillsColors: TIntegerDynArray - массив с цветами заливки
//  var ThemaColorCount: integer          - кол-во цветов заливки
//RETURN
//      boolean - true - всё прочиталось успешно
function ZEXSLXReadTheme(var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer): boolean;
var
  xml: TZsspXMLReaderH;
  maxCount: integer;
  _b: boolean;

  procedure _addFillColor(const _rgb: string);
  begin
    inc(ThemaColorCount);
    if (ThemaColorCount >= maxCount) then
    begin
      maxCount := ThemaColorCount + 20;
      SetLength(ThemaFillsColors, maxCount);
    end;
    ThemaFillsColors[ThemaColorCount - 1] := HTMLHexToColor(_rgb);
  end; //_addFillColor

begin
  result := false;
  xml := nil;
  _b := false;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;
    ThemaColorCount := 0;
    maxCount := -1;
    while (not xml.Eof()) do
    begin
      xml.ReadTag();
      if (xml.TagName = 'a:clrScheme') then
      begin
        if (xml.TagType = TAG_TYPE_START) then
          _b := true;
        if (xml.TagType = TAG_TYPE_END) then
          _b := false;
      end else
      if ((xml.TagName = 'a:sysClr') and (_b) and (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED])) then
      begin
        _addFillColor(xml.Attributes.ItemsByName['lastClr']);
      end else
      if ((xml.TagName = 'a:srgbClr') and (_b) and (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED])) then
      begin
        _addFillColor(xml.Attributes.ItemsByName['val']);
      end;
    end; //while

    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
  end;
end; //ZEXSLXReadThema

//Читает список нужных файлов из [Content_Types].xml
//INPUT
//  var Stream: TStream             - поток чтения
//  var FileArray: TZXLSXFileArray  - список файлов
//  var FilesCount: integer         - кол-во файлов
//RETURN
//      boolean - true - всё прочиталось успешно
function ZEXSLXReadContentTypes(var Stream: TStream; var FileArray: TZXLSXFileArray; var FilesCount: integer): boolean;
var
  xml: TZsspXMLReaderH;
  s: string;
  t: integer;

begin
  result := false;
  xml := nil;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;
    FilesCount := 0;
    while (not xml.Eof()) do
    begin
      xml.ReadTag();
      if ((xml.TagName = 'Override') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        s := xml.Attributes.ItemsByName['PartName'];
        if (s > '') then
        begin
          SetLength(FileArray, FilesCount + 1);
          FileArray[FilesCount].name := s;
          FileArray[FilesCount].original := s;
          s := xml.Attributes.ItemsByName['ContentType'];
          t := -1;
          if (s = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml') then
            t := 0
          else
          if (s = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml') then
            t := 1
          else
          if (s = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml') or
          	 (s = 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml') then
            t := 2
          else
          if (s = 'application/vnd.openxmlformats-package.relationships+xml') then
            t := 3
          else
          if (s = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml') then
            t := 4
          else
          if (s = 'application/vnd.openxmlformats-package.core-properties+xml') then
            t := 5
          else
          if (s = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml') then
            t := 6
          else
          if (s = 'application/vnd.openxmlformats-officedocument.vmlDrawing') then
            t := 7
          else
          if (s = 'application/vnd.openxmlformats-officedocument.theme+xml') then
            t := 8;

          FileArray[FilesCount].ftype := t;

          if (t > -1) then
            inc(FilesCount);
        end;
      end;
    end;

    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
  end;
end; //ZEXSLXReadContentTypes

//Читает строки из sharedStrings.xml
//INPUT
//  var Stream: TStream           - поток для чтения
//  var StrArray: TStringDynArray - возвращаемый массив со строками
//  var StrCount: integer         - кол-во элементов
//RETURN
//      boolean - true - всё ок
function ZEXSLXReadSharedStrings(var Stream: TStream; out StrArray: TStringDynArray; out StrCount: integer): boolean;
var
  xml: TZsspXMLReaderH;
  s: string;
  k: integer;

begin
  result := false;
  xml := nil;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;
    StrCount := 0;

    while (not xml.Eof()) do
    begin
      xml.ReadTag();
      if ((xml.TagName = 'si') and (xml.TagType = TAG_TYPE_START)) then
      begin
        s := '';
        k := 0;
        while (not((xml.TagName = 'si') and (xml.TagType = TAG_TYPE_END))) and (not xml.Eof) do
        begin
          xml.ReadTag();
          if ((xml.TagName = 't') and (xml.TagType = TAG_TYPE_END)) then
          begin
            if (k > 1) then
              s := s + {$IFDEF FPC} LineEnding {$ELSE} sLineBreak {$ENDIF};
            s := s + xml.TextBeforeTag;
          end;
          if (xml.TagName = 'r') and (xml.TagType = TAG_TYPE_END) then
            inc(k);
        end; //while
        SetLength(StrArray, StrCount + 1);
        StrArray[StrCount] := s;
        inc(StrCount);
      end; //if
    end; //while

    result := true;

  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
  end;
end; //ZEXSLXReadSharedStrings

{$IFDEF ZUSE_CONDITIONAL_FORMATTING}
//Получить условное форматирование и оператор из xlsx
//INPUT
//  const xlsxCfType: string              - xlsx тип условного форматирования
//  const xlsxCFOperator: string          - xlsx оператор
//  out CFCondition: TZCondition          - распознанное условие
//  out CFOperator: TZConditionalOperator - распознанный оператор
//RETURN
//      boolean - true - условное форматирование и оператор успешно распознаны
function ZEXLSX_getCFCondition(const xlsxCfType, xlsxCFOperator: string;
                               out CFCondition: TZCondition;
                               out CFOperator: TZConditionalOperator): boolean;
var
  isCheckOperator: boolean;

  procedure _SetCFOperator(AOperator: TZConditionalOperator);
  begin
    CFOperator := AOperator;
    CFCondition := ZCFCellContentOperator;
  end;

  //Проверить тип условного форматирования
  //  out isNeddCheckOperator: boolean - возвращает, нужно ли проверять
  //                                     оператор
  //RETURN
  //      boolean - true - всё ок, можно проверять далее
  function _CheckXLSXCfType(out isNeddCheckOperator: boolean): boolean;
  begin
    result := true;
    isNeddCheckOperator := true;
    if (xlsxCfType = 'cellIs') then
    begin
    end else
    if (xlsxCfType = 'containsText') then
      CFCondition := ZCFContainsText
    else
    if (xlsxCfType = 'notContains') then
      CFCondition := ZCFNotContainsText
    else
    if (xlsxCfType = 'beginsWith') then
      CFCondition := ZCFBeginsWithText
    else
    if (xlsxCfType = 'endsWith') then
      CFCondition := ZCFEndsWithText
    else
    if (xlsxCfType = 'containsBlanks') then
      isNeddCheckOperator := false
    else
      result := false;
  end; //_CheckXLSXCfType

  //Проверить оператор
  function _CheckCFoperator(): boolean;
  begin
    result := true;
    if (xlsxCFOperator = 'lessThan') then
      _SetCFOperator(ZCFOpLT)
    else
    if (xlsxCFOperator = 'equal') then
      _SetCFOperator(ZCFOpEqual)
    else
    if (xlsxCFOperator = 'notEqual') then
      _SetCFOperator(ZCFOpNotEqual)
    else
    if (xlsxCFOperator = 'greaterThanOrEqual') then
      _SetCFOperator(ZCFOpGTE)
    else
    if (xlsxCFOperator = 'greaterThan') then
      _SetCFOperator(ZCFOpGT)
    else
    if (xlsxCFOperator = 'lessThanOrEqual') then
      _SetCFOperator(ZCFOpLTE)
    else
    if (xlsxCFOperator = 'between') then
      CFCondition := ZCFCellContentIsBetween
    else
    if (xlsxCFOperator = 'notBetween') then
      CFCondition := ZCFCellContentIsNotBetween
    else
    if (xlsxCFOperator = 'containsText') then
      CFCondition := ZCFContainsText
    else
    if (xlsxCFOperator = 'notContains') then
      CFCondition := ZCFNotContainsText
    else
    if (xlsxCFOperator = 'beginsWith') then
      CFCondition := ZCFBeginsWithText
    else
    if (xlsxCFOperator = 'endsWith') then
      CFCondition := ZCFEndsWithText
    else
      result := false;
  end; //_CheckCFoperator

begin
  result := false;
  CFCondition := ZCFNumberValue;
  CFOperator := ZCFOpGT;

  if (_CheckXLSXCfType(isCheckOperator)) then
  begin
    if (isCheckOperator) then
      result := _CheckCFoperator()
    else
      result := true;
  end;
end; //ZEXLSX_getCFCondition
{$ENDIF}

//Читает страницу документа
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//  var Stream: TStream                 - поток для чтения
//  const SheetName: string             - название страницы
//  var StrArray: TStringDynArray       - строки для подстановки
//      StrCount: integer               - кол-во строк подстановки
//  var Relations: TZXLSXRelationsArray - отношения
//      RelationsCount: integer         - кол-во отношений
//      ReadHelper: TZEXLSXReadHelper   -
//RETURN
//      boolean - true - страница прочиталась успешно
function ZEXSLXReadSheet(var XMLSS: TZEXMLSS; var Stream: TStream; const SheetName: string; var StrArray: TStringDynArray; StrCount: integer;
                         var Relations: TZXLSXRelationsArray; RelationsCount: integer;
                         ReadHelper: TZEXLSXReadHelper): boolean;
var
  xml: TZsspXMLReaderH;
  _currPage: integer;
  _currRow: integer;
  _currCol: integer;
  _currSheet: TZSheet;
  _currCell: TZCell;
  s: string;
  _tmpr: real;
  _t: integer;
  _td: TDateTime;
  _tfloat: Double;

  //Проверить кол-во строк
  procedure CheckRow(const RowCount: integer);
  begin
    if (_currSheet.RowCount < RowCount) then
      _currSheet.RowCount := RowCount;
  end;

  //Проверить кол-во столбцов
  procedure CheckCol(const ColCount: integer);
  begin
    if (_currSheet.ColCount < ColCount) then
      _currSheet.ColCount := ColCount
  end;

  //Чтение строк/столбцов
  procedure _ReadSheetData();
  var
    t: integer;
    v: string;
    _num: integer;
    _type: string;
    _cr, _cc: integer;

  begin
    _cr := 0;
    _cc := 0;
    CheckRow(1);
    CheckCol(1);
    while (not ((xml.TagName = 'sheetData') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof) then
        break;

      //ячейка
      if (xml.TagName = 'c') then
      begin
        s := xml.Attributes.ItemsByName['r']; //номер
        if (s <> '') then
        begin
          if (ZEGetCellCoords(s, _cc, _cr)) then
          begin
            _currCol := _cc;
            CheckCol(_cc + 1);
          end;
        end;
        _type := xml.Attributes.ItemsByName['t']; //тип

        //s := xml.Attributes.ItemsByName['cm'];
        //s := xml.Attributes.ItemsByName['ph'];
        //s := xml.Attributes.ItemsByName['vm'];
        v := '';
        _num := 0;
        _currCell := _currSheet.Cell[_currCol, _currRow];
        s := xml.Attributes.ItemsByName['s']; //стиль
        if (s > '') then
          if (tryStrToInt(s, t)) then
            _currCell.CellStyle := t;
        if (xml.TagType = TAG_TYPE_START) then
        while (not ((xml.TagName = 'c') and (xml.TagType = TAG_TYPE_END))) do
        begin
          xml.ReadTag();
          if (xml.Eof) then
            break;

          //is пока игнорируем

          if (((xml.TagName = 'v') or (xml.TagName = 't')) and (xml.TagType = TAG_TYPE_END)) then
          begin
            if (_num > 0) then
              v := v + {$IFDEF FPC} LineEnding {$ELSE} sLineBreak {$ENDIF};
            v := v + xml.TextBeforeTag;  
            inc(_num);
          end else
          if ((xml.TagName = 'f') and (xml.TagType = TAG_TYPE_END)) then
            _currCell.Formula := ZEReplaceEntity(xml.TextBeforeTag);

        end; //while

        //Возможные типы:
        //  s - sharedstring
        //  b - boolean
        //  n - number
        //  e - error
        //  str - string
        //  inlineStr - inline string ??
        //  d - date
        //  тип может отсутствовать. Интерпретируем в таком случае как ZEGeneral
        if (_type = '') then _currCell.CellType := ZEGeneral
        else
        if (_type = 'n') then
        begin
          _currCell.CellType := ZENumber;
          //Trouble: if cell style is number, and number format is date, then
          // cell style is date. F****** m$!
          if (ReadHelper.NumberFormats.IsDateFormat(_currCell.CellStyle)) then
            if (ZEIsTryStrToFloat(v, _tfloat)) then
            begin
              _currCell.CellType := ZEDateTime;
              v := ZEDateTimeToStr(_tfloat);
            end;
        end
        else
        if (_type = 's') then
        begin
          _currCell.CellType := ZEString;
          if (TryStrToInt(v, t)) then
            if ((t >= 0) and (t < StrCount)) then
              v := StrArray[t];
        end
        else
        if (_type = 'd') then
        begin
          _currCell.CellType := ZEDateTime;
          if (TryZEStrToDateTime(v, _td)) then
            v := ZEDateTimeToStr(_td)
          else
          if (ZEIsTryStrToFloat(v, _tfloat)) then
            v := ZEDateTimeToStr(_tfloat)
          else
            _currCell.CellType := ZEString;
        end;

        _currCell.Data := ZEReplaceEntity(v);
        inc(_currCol);
        CheckCol(_currCol + 1);
      end else
      //строка
      if ((xml.TagName = 'row') and (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED])) then
      begin
        _currCol := 0;
        s := xml.Attributes.ItemsByName['r']; //индекс строки
        if (s > '') then
          if (TryStrToInt(s, t)) then
          begin
            _currRow := t - 1;
            CheckRow(t);
          end;
        //s := xml.Attributes.ItemsByName['collapsed'];
        //s := xml.Attributes.ItemsByName['customFormat'];
        //s := xml.Attributes.ItemsByName['customHeight'];
        s := xml.Attributes.ItemsByName['hidden'];
        if (s > '') then
          _currSheet.Rows[_currRow].Hidden := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['ht']; //в поинтах?
        if (s > '') then
        begin
          _tmpr := ZETryStrToFloat(s, 10);
          _tmpr := _tmpr / 2.835; //???
          _currSheet.Rows[_currRow].HeightMM := _tmpr;
        end;

        //s := xml.Attributes.ItemsByName['outlineLevel'];
        //s := xml.Attributes.ItemsByName['ph'];

        s := xml.Attributes.ItemsByName['s']; //номер стиля
        if (s > '') then
          if (TryStrToInt(s, t)) then
          begin
            //нужно подставить нужный стиль
          end;
        //s := xml.Attributes.ItemsByName['spans'];
        //s := xml.Attributes.ItemsByName['thickBot'];
        //s := xml.Attributes.ItemsByName['thickTop'];

        if (xml.TagType = TAG_TYPE_CLOSED) then
        begin
          inc(_currRow);
          CheckRow(_currRow + 1);
        end;
      end else
      //конец строки
      if ((xml.TagName = 'row') and (xml.TagType = TAG_TYPE_END)) then
      begin
        inc(_currRow);
        CheckRow(_currRow + 1);
      end;
    end; //while
  end; //_ReadSheetData

  //Чтение диапазона ячеек с автофильтром
  procedure _ReadAutoFilter();
  begin
    _currSheet.AutoFilter:=xml.Attributes.ItemsByName['ref'];
  end;

  //Чтение объединённых ячеек
  procedure _ReadMerge();
  var
    i, t, num: integer;
    x1, x2, y1, y2: integer;
    s1, s2: string;
    b: boolean;

    function _GetCoords(var x, y: integer): boolean;
    begin
      result := true;
      x := ZEGetColByA1(s1);
      if (x < 0) then
        result := false;
      if (not TryStrToInt(s2, y)) then
        result := false
      else
        dec(y);  
      b := result;
    end; //_GetCoords

  begin
    x1 := 0;
    y1 := 0;
    x2 := 0;
    y2 := 0;
    while (not ((xml.TagName = 'mergeCells') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof) then
        break;

      if ((xml.TagName = 'mergeCell') and (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED])) then
      begin
        s := xml.Attributes.ItemsByName['ref'];
        t := length(s);
        if (t > 0) then
        begin
          s := s + ':';
          s1 := '';
          s2 := '';
          b := true;
          num := 0;
          for i := 1 to t + 1 do
          case s[i] of
            'A'..'Z', 'a'..'z': s1 := s1 + s[i];
            '0'..'9': s2 := s2 + s[i];
            ':':
              begin
                inc(num);
                if (num > 2) then
                begin
                  b := false;
                  break;
                end;
                if (num = 1) then
                begin
                  if (not _GetCoords(x1, y1)) then
                    break;
                end else
                begin
                  if (not _GetCoords(x2, y2)) then
                    break;
                end;
                s1 := '';
                s2 := '';
              end;
            else
            begin
              b := false;
              break;
            end;
          end; //case

          if (b) then
          begin
            CheckRow(y1 + 1);
            CheckRow(y2 + 1);
            CheckCol(x1 + 1);
            CheckCol(x2 + 1);
            _currSheet.MergeCells.AddRectXY(x1, y1, x2, y2);
          end;

        end; //if
      end; //if
    end; //while
  end; //_ReadMerge

  //Столбцы
  procedure _ReadCols();
  var
    num: integer;
    t: real;
    _max, _min: integer;
    _delta, i: integer;

  begin
    num := 0;
    while (not ((xml.TagName = 'cols') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof) then
        break;

      if ((xml.TagName = 'col') and (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED])) then
      begin
        CheckCol(num + 1);
        s := xml.Attributes.ItemsByName['bestFit'];
        if (s > '') then
          _currSheet.Columns[num].AutoFitWidth := ZETryStrToBoolean(s);

        //s := xml.Attributes.ItemsByName['collapsed'];
        //s := xml.Attributes.ItemsByName['customWidth'];
        s := xml.Attributes.ItemsByName['hidden'];
        if (s > '') then
          _currSheet.Columns[num].Hidden := ZETryStrToBoolean(s);

        _min := 0;
        _max := 0;
        s := xml.Attributes.ItemsByName['max'];
        if (s > '') then
          TryStrToInt(s, _max);
        s := xml.Attributes.ItemsByName['min'];
        if (s > '') then
          TryStrToInt(s, _min);
        //s := xml.Attributes.ItemsByName['outlineLevel'];
        //s := xml.Attributes.ItemsByName['phonetic'];
        s := xml.Attributes.ItemsByName['style'];

        s := xml.Attributes.ItemsByName['width'];
        if (s > '') then
        begin
          t := ZETryStrToFloat(s, 5.14509803921569);
          t := 10 * t / 5.14509803921569;
          _currSheet.Columns[num].WidthMM := t;
        end;

        if (_max > _min) then
        begin
          _delta := _max - _min;
          CheckCol(_max);
          for i := _min to _max - 1 do
            _currSheet.Columns[i].Assign(_currSheet.Columns[num]);
          inc(num, _delta);
        end;

        inc(num);
      end; //if
    end; //while    
  end; //_ReadCols

  function _StrToMM(const st: string; var retFloat: real): boolean;
  begin
    result := false;
    if (s > '') then
    begin
      retFloat := ZETryStrToFloat(st, -1);
      if (retFloat > 0) then
      begin
        result := true;
        retFloat := retFloat * ZE_MMinInch;
      end;
    end;
  end; //_StrToMM

  procedure _GetDimension();
  var
    st, s: string;
    i, l: integer;
    _maxC, _maxR, c, r: integer;

  begin
    c := 0;
    r := 0;
    st := xml.Attributes.ItemsByName['ref'];
    l := Length(st);
    if (l > 0) then
    begin
      st := st + ':';
      inc(l);
      s := '';
      _maxC := -1;
      _maxR := -1;
      for i := 1 to l do
      if (st[i] = ':') then
      begin
        if (ZEGetCellCoords(s, c, r, true)) then
        begin;
          if (c > _maxC) then
            _maxC := c;
          if (r > _maxR) then
            _maxR := r;
        end else
          break;
        s := '';
      end else
        s := s + st[i];
      if (_maxC > 0) then
        CheckCol(_maxC);
      if (_maxR > 0) then
        CheckRow(_maxR);
    end;
  end; //_GetDimension()

  //Чтение ссылок
  procedure _ReadHyperLinks();
  var
    _c, _r: integer;
    i: integer;

  begin
    _c := 0;
    _r := 0;
    while (not ((xml.TagName = 'hyperlinks') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'hyperlink') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        s := xml.Attributes.ItemsByName['ref'];
        if (s > '') then
          if (ZEGetCellCoords(s, _c, _r, true)) then
          begin
            CheckRow(_r);
            CheckCol(_c);
            _currSheet.Cell[_c, _r].HRefScreenTip := xml.Attributes.ItemsByName['tooltip'];
            s := xml.Attributes.ItemsByName['r:id'];
            //по r:id подставить ссылку
            for i := 0 to RelationsCount - 1 do
              if ((Relations[i].id = s) and (Relations[i].ftype = 6)) then
              begin
                _currSheet.Cell[_c, _r].Href := Relations[i].target;
                break;
              end;
          end;
        //доп. атрибуты:
        //  display - ??
        //  id - id <> r:id??
        //  location - ??
      end;
    end; //while
  end; //_ReadHyperLinks();

  //<sheetViews> ... </sheetViews>
  procedure _ReadSheetViews();
  var
    vValue, hValue: integer;
    SplitMode: TZSplitMode;
    s: string;

  begin
    while (not ((xml.TagName = 'sheetViews') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'sheetView')) then
      begin
        s := xml.Attributes.ItemsByName['tabSelected'];
        _currSheet.Selected := s = '1';
      end;

      if ((xml.TagName = 'pane') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        SplitMode := ZSplitSplit;
        s := xml.Attributes.ItemsByName['state'];
        if (s = 'frozen') then
          SplitMode := ZSplitFrozen;

        s := xml.Attributes.ItemsByName['xSplit'];
        if (not TryStrToInt(s, vValue)) then
          vValue := 0;

        s := xml.Attributes.ItemsByName['ySplit'];
        if (not TryStrToInt(s, hValue)) then
          hValue := 0;

        _currSheet.SheetOptions.SplitVerticalValue := vValue;
        _currSheet.SheetOptions.SplitHorizontalValue := hValue;

        _currSheet.SheetOptions.SplitHorizontalMode := ZSplitNone;
        _currSheet.SheetOptions.SplitVerticalMode := ZSplitNone;
        if (hValue <> 0) then
          _currSheet.SheetOptions.SplitHorizontalMode := SplitMode;
        if (vValue <> 0) then
          _currSheet.SheetOptions.SplitVerticalMode := SplitMode;

        if (_currSheet.SheetOptions.SplitHorizontalMode = ZSplitSplit) then
          _currSheet.SheetOptions.SplitHorizontalValue := PointToPixel(hValue/20);
        if (_currSheet.SheetOptions.SplitVerticalMode = ZSplitSplit) then
          _currSheet.SheetOptions.SplitVerticalValue := PointToPixel(vValue/20);

      end; //if
    end; //while
  end; //_ReadSheetViews()

  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  //Чтение условного форматирования
  //<conditionalFormatting>..</conditionalFormatting>
  procedure _ReadConditionFormatting();
  var
    MaxFormulasCount: integer;
    _formulas: array of string;
    count: integer;
    _sqref: string;
    _type: string;
    _operator: string;
    _CFCondition: TZCondition;
    _CFOperator: TZConditionalOperator;
    _Style: string;
    _text: string;
    _isCFAdded: boolean;
    _isOk: boolean;
    //_priority: string;
    _CF: TZConditionalStyle;
    _tmpStyle: TZStyle;

    function _AddCF(): boolean;
    var
      s, ss: string;
      _len, i, kol: integer;
      a: array of array[0..5] of integer;
      _maxx: integer;
      ch: char;
      w, h: integer;

      function _GetOneArea(st: string): boolean;
      var
        i, j: integer;
        s: string;
        ch: char;
        _cnt: integer;
        tmpArr: array [0..1, 0..1] of integer;
        _isOk: boolean;
        t: integer;
        tmpB: boolean;

      begin
        result := false;
        if (st <> '') then
        begin
          st := st + ':';
          s := '';
          _cnt := 0;
          _isOk := true;
          for i := 1 to length(st) do
          begin
            ch := st[i];
            if (ch = ':') then
            begin
              if (_cnt < 2) then
              begin
                tmpB := ZEGetCellCoords(s, tmpArr[_cnt][0], tmpArr[_cnt][1]);
                _isOk := _isOk and tmpB;
              end;
              s := '';
              inc(_cnt);
            end else
              s := s + ch;
          end; //for

          if (_isOk) then
            if (_cnt > 0) then
            begin
              if (_cnt > 2) then
                _cnt := 2;

              a[kol][0] := _cnt;
              t := 1;
              for i := 0 to _cnt - 1 do
                for j := 0 to 1 do
                begin
                  a[kol][t] := tmpArr[i][j];
                  inc(t);
                end;
              result := true;
            end;
        end; //if
      end; //_GetOneArea

    begin
      result := false;
      if (_sqref <> '') then
      try
        _maxx := 4;
        SetLength(a, _maxx);
        ss := _sqref + ' ';
        _len := Length(ss);
        kol := 0;
        s := '';
        for i := 1 to _len do
        begin
          ch := ss[i];
          if (ch = ' ') then
          begin
            if (_GetOneArea(s)) then
            begin
              inc(kol);
              if (kol >= _maxx) then
              begin
                inc(_maxx, 4);
                SetLength(a, _maxx);
              end;
            end;
            s := '';
          end else
            s := s + ch;
        end; //for

        if (kol > 0) then
        begin
          _currSheet.ConditionalFormatting.Add();
          _CF := _currSheet.ConditionalFormatting[_currSheet.ConditionalFormatting.Count - 1];
          for i := 0 to kol - 1 do
          begin
            w := 1;
            h := 1;
            if (a[i][0] >= 2) then
            begin
              w := abs(a[i][3] - a[i][1]) + 1;
              h := abs(a[i][4] - a[i][2]) + 1;
            end;
            _CF.Areas.Add(a[i][1], a[i][2], w, h);
          end;
          result := true;
        end;
      finally
        SetLength(a, 0);
      end;
    end; //_AddCF

    //Применяем условный стиль
    procedure _TryApplyCF();
    var
      b: boolean;
      num: integer;
      _id: integer;

      procedure _CheckTextCondition();
      begin
        if (count = 1) then
          if (_formulas[0] <> '') then
            _isOk := true;
      end;

      //Найти стиль
      //  пока будем делать так: предполагаем, что все ячейки в текущей области
      //  условного форматирования имеют один стиль. Берём стиль из левой верхней
      //  ячейки, клонируем его, применяем дифф. стиль, добавляем в хранилище стилей
      //  с учётом повторов.
      //TODO: потом нужно будет переделать
      //INPUT
      //      dfNum: integer - номер дифференцированного форматирования
      //RETURN
      //      integer - номер применяемого стиля
      function _getStyleIdxForDF(dfNum: integer): integer;
      var
        _df: TZXLSXDiffFormattingItem;
        _r, _c: integer;
        _t: integer;
        i: integer;

      begin
        //_currSheet
        result := -1;
        if ((dfNum >= 0) and (dfNum < ReadHelper.DiffFormatting.Count)) then
        begin
          _df := ReadHelper.DiffFormatting[dfNum];
          _t := -1;

          if (_cf.Areas.Count > 0) then
          begin
            _r := _cf.Areas.Items[0].Row;
            _c := _cf.Areas.Items[0].Column;
            if ((_r >= 0) and (_r < _currSheet.RowCount)) then
              if ((_c >= 0) and (_c < _currSheet.ColCount)) then
                _t := _currSheet.Cell[_c, _r].CellStyle;
          end;

          _tmpStyle.Assign(XMLSS.Styles[_t]);

          if (_df.UseFont) then
          begin
            if (_df.UseFontStyles) then
              _tmpStyle.Font.Style := _df.FontStyles;
            if (_df.UseFontColor) then
              _tmpStyle.Font.Color := _df.FontColor;
          end;
          if (_df.UseFill) then
          begin
            if (_df.UseCellPattern) then
              _tmpStyle.CellPattern := _df.CellPattern;
            if (_df.UseBGColor) then
              _tmpStyle.BGColor := _df.BGColor;
            if (_df.UsePatternColor) then
              _tmpStyle.PatternColor := _df.PatternColor;
          end;
          if (_df.UseBorder) then
            for i := 0 to 5 do
            begin
              if (_df.Borders[i].UseStyle) then
              begin
                _tmpStyle.Border[i].Weight := _df.Borders[i].Weight;
                _tmpStyle.Border[i].LineStyle := _df.Borders[i].LineStyle;
              end;
              if (_df.Borders[i].UseColor) then
                _tmpStyle.Border[i].Color := _df.Borders[i].Color;
            end; //for

          result := XMLSS.Styles.Add(_tmpStyle, true);
        end; //if
      end; //_getStyleIdxForDF

    begin
      _isOk := false;
      case (_CFCondition) of
        ZCFIsTrueFormula:;
        ZCFCellContentIsBetween, ZCFCellContentIsNotBetween:
          begin
            //только числа
            if (count = 2) then
            begin
              ZETryStrToFloat(_formulas[0], b);
              if (b) then
                ZETryStrToFloat(_formulas[1], _isOk);
            end;
          end;
        ZCFCellContentOperator:
          begin
            //только числа
            if (count = 1) then
              ZETryStrToFloat(_formulas[0], _isOk);
          end;
        ZCFNumberValue:;
        ZCFString:;
        ZCFBoolTrue:;
        ZCFBoolFalse:;
        ZCFFormula:;
        ZCFContainsText: _CheckTextCondition();
        ZCFNotContainsText: _CheckTextCondition();
        ZCFBeginsWithText: _CheckTextCondition();
        ZCFEndsWithText: _CheckTextCondition();
      end; //case

      if (_isOk) then
      begin
        if (not _isCFAdded) then
          _isCFAdded := _AddCF();

        if ((_isCFAdded) and (Assigned(_CF))) then
        begin
          num := _CF.Count;
          _CF.Add();
          if (_Style <> '') then
            if (TryStrToInt(_Style, _id)) then
             _CF[num].ApplyStyleID := _getStyleIdxForDF(_id);
          _CF[num].Condition := _CFCondition;
          _CF[num].ConditionOperator := _CFOperator;

          _cf[num].Value1 := _formulas[0];
          if (count >= 2) then
            _cf[num].Value2 := _formulas[1];
        end;
      end;
    end; //_TryApplyCF

  begin
    try
      _sqref := xml.Attributes['sqref'];
      MaxFormulasCount := 2;
      SetLength(_formulas, MaxFormulasCount);
      _isCFAdded := false;
      _CF := nil;
      _tmpStyle := TZStyle.Create();
      while (not ((xml.TagType = TAG_TYPE_END) and (xml.TagName = ZETag_conditionalFormatting))) do
      begin
        xml.ReadTag();
        if (xml.Eof()) then
          break;

        // cfRule = Conditional Formatting Rule
        if ((xml.TagType = TAG_TYPE_START) and (xml.TagName = ZETag_cfRule)) then
        begin
         (*
          Атрибуты в cfRule:
          type	       	- тип
                            expression        - ??
                            cellIs            -
                            colorScale        - ??
                            dataBar           - ??
                            iconSet           - ??
                            top10             - ??
                            uniqueValues      - ??
                            duplicateValues   - ??
                            containsText      -    ?
                            notContainsText   -    ?
                            beginsWith        -    ?
                            endsWith          -    ?
                            containsBlanks    - ??
                            notContainsBlanks - ??
                            containsErrors    - ??
                            notContainsErrors - ??
                            timePeriod        - ??
                            aboveAverage      - ?
          dxfId	        - ID применяемого формата
          priority	    - приоритет
          stopIfTrue	  -  ??
          aboveAverage  -  ??
          percent	      -  ??
          bottom	      -  ??
          operator	    - оператор:
                              lessThan	          <
                              lessThanOrEqual	    <=
                              equal	              =
                              notEqual	          <>
                              greaterThanOrEqual  >=
                              greaterThan	        >
                              between	            Between
                              notBetween	        Not Between
                              containsText	      содержит текст
                              notContains	        не содержит
                              beginsWith	        начинается с
                              endsWith	          оканчивается на
          text	        -  ??
          timePeriod	  -  ??
          rank	        -  ??
          stdDev  	    -  ??
          equalAverage	-  ??
         *)
          _type := xml.Attributes['type'];
          _operator := xml.Attributes['operator'];
          _Style := xml.Attributes['dxfId'];
          _text := ZEReplaceEntity(xml.Attributes['text']);
          //_priority := xml.Attributes['priority'];

          count := 0;
          while (not ((xml.TagType = TAG_TYPE_END) and (xml.TagName = ZETag_cfRule))) do
          begin
            xml.ReadTag();
            if (xml.Eof()) then
              break;
            if (xml.TagType = TAG_TYPE_END) then
              if (xml.TagName = ZETag_formula) then
              begin
                if (count >= MaxFormulasCount) then
                begin
                  inc(MaxFormulasCount, 2);
                  SetLength(_formulas, MaxFormulasCount);
                end;
                _formulas[count] := ZEReplaceEntity(xml.TextBeforeTag);
                inc(count);
              end;
          end; //while

          if (ZEXLSX_getCFCondition(_type, _operator, _CFCondition, _CFOperator)) then
            _TryApplyCF();
        end; //if
      end; //while
    finally
      SetLength(_formulas, 0);
      FreeAndNil(_tmpStyle);
    end;
  end; //_ReadConditionFormatting
  {$ENDIF}

begin
  xml := nil;
  result := false;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    _currPage := XMLSS.Sheets.Count;
    XMLSS.Sheets.Count := XMLSS.Sheets.Count + 1;
    _currRow := 0;
    _currSheet := XMLSS.Sheets[_currPage];
    _currSheet.Title := SheetName;

    while (not xml.Eof()) do
    begin
      xml.ReadTag();
      if ((xml.TagName = 'sheetData') and (xml.TagType = TAG_TYPE_START)) then
        _ReadSheetData()
      else
      if ((xml.TagName = 'autoFilter') and (xml.TagType = TAG_TYPE_CLOSED)) then
        _ReadAutoFilter()
      else
      if ((xml.TagName = 'mergeCells') and (xml.TagType = TAG_TYPE_START)) then
        _ReadMerge()
      else
      if ((xml.TagName = 'cols') and (xml.TagType = TAG_TYPE_START)) then
        _ReadCols()
      else
      //Отступы
      if ((xml.TagName = 'pageMargins') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        //в дюймах
        s := xml.Attributes.ItemsByName['bottom'];
        if (_StrToMM(s, _tmpr)) then
          _currSheet.SheetOptions.MarginBottom := round(_tmpr);
        s := xml.Attributes.ItemsByName['footer'];
        if (_StrToMM(s, _tmpr)) then
          _currSheet.SheetOptions.FooterMargins.Height := abs(round(_tmpr));
        s := xml.Attributes.ItemsByName['header'];
        if (_StrToMM(s, _tmpr)) then
          _currSheet.SheetOptions.HeaderMargins.Height := abs(round(_tmpr));
        s := xml.Attributes.ItemsByName['left'];
        if (_StrToMM(s, _tmpr)) then
          _currSheet.SheetOptions.MarginLeft := round(_tmpr);
        s := xml.Attributes.ItemsByName['right'];
        if (_StrToMM(s, _tmpr)) then
          _currSheet.SheetOptions.MarginRight := round(_tmpr);
        s := xml.Attributes.ItemsByName['top'];
        if (_StrToMM(s, _tmpr)) then
          _currSheet.SheetOptions.MarginTop := round(_tmpr);
      end else
      //Настройки страницы
      if ((xml.TagName = 'pageSetup') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        //s := xml.Attributes.ItemsByName['blackAndWhite'];
        //s := xml.Attributes.ItemsByName['cellComments'];
        //s := xml.Attributes.ItemsByName['copies'];
        //s := xml.Attributes.ItemsByName['draft'];
        //s := xml.Attributes.ItemsByName['errors'];
        s := xml.Attributes.ItemsByName['firstPageNumber'];
        if (s > '') then
          if (TryStrToInt(s, _t)) then
            _currSheet.SheetOptions.StartPageNumber := _t;
            
        //s := xml.Attributes.ItemsByName['fitToHeight'];
        //s := xml.Attributes.ItemsByName['fitToWidth'];
        //s := xml.Attributes.ItemsByName['horizontalDpi'];
        //s := xml.Attributes.ItemsByName['id'];
        s := xml.Attributes.ItemsByName['orientation'];
        if (s > '') then
        begin
          _currSheet.SheetOptions.PortraitOrientation := false;
          if (s = 'portrait') then
            _currSheet.SheetOptions.PortraitOrientation := true;
        end;
        
        //s := xml.Attributes.ItemsByName['pageOrder'];

        s := xml.Attributes.ItemsByName['paperSize'];
        if (s > '') then
          if (TryStrToInt(s, _t)) then
            _currSheet.SheetOptions.PaperSize := _t;
        //s := xml.Attributes.ItemsByName['paperHeight']; //если утановлены paperHeight и Width, то paperSize игнорируется
        //s := xml.Attributes.ItemsByName['paperWidth'];

        //s := xml.Attributes.ItemsByName['scale'];
        //s := xml.Attributes.ItemsByName['useFirstPageNumber'];
        //s := xml.Attributes.ItemsByName['usePrinterDefaults'];
        //s := xml.Attributes.ItemsByName['verticalDpi'];
      end else
      //настройки печати
      if ((xml.TagName = 'printOptions') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        //s := xml.Attributes.ItemsByName['gridLines'];
        //s := xml.Attributes.ItemsByName['gridLinesSet'];
        //s := xml.Attributes.ItemsByName['headings'];
        s := xml.Attributes.ItemsByName['horizontalCentered'];
        if (s > '') then
          _currSheet.SheetOptions.CenterHorizontal := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['verticalCentered'];
        if (s > '') then
          _currSheet.SheetOptions.CenterVertical := ZEStrToBoolean(s);
      end else
      if ((xml.TagName = 'dimension') and (xml.TagType = TAG_TYPE_CLOSED)) then
        _GetDimension()
      else
      if ((xml.TagName = 'hyperlinks') and (xml.TagType = TAG_TYPE_START)) then
        _ReadHyperLinks()
      else
      if ((xml.TagName = 'sheetViews') and (xml.TagType = TAG_TYPE_START)) then
        _ReadSheetViews()
      {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
      else
      if ((xml.TagType = TAG_TYPE_START) and (xml.TagName = ZETag_conditionalFormatting)) then
        _ReadConditionFormatting()
      {$ENDIF}
      ;
    end; //while
    
    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
  end;
end; //ZEXSLXReadSheet

//Прочитать стили из потока (styles.xml)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//  var Stream: TStream                   - поток
//  var ThemaFillsColors: TIntegerDynArray - цвета из темы
//  var ThemaColorCount: integer          - кол-во цветов заливки в теме
//      ReadHelper: TZEXLSXReadHelper     -
//RETURN
//      boolean - true - стили прочитались без ошибок
function ZEXSLXReadStyles(var XMLSS: TZEXMLSS; var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer; ReadHelper: TZEXLSXReadHelper): boolean;
type

  TZXLSXBorderItem = record
    color: TColor;
    isColor: boolean;
    isEnabled: boolean;
    style: TZBorderType;
    Weight: byte;
  end;

  //   0 - left           левая граница
  //   1 - Top            верхняя граница
  //   2 - Right          правая граница
  //   3 - Bottom         нижняя граница
  //   4 - DiagonalLeft   диагоняль от верхнего левого угла до нижнего правого
  //   5 - DiagonalRight  диагоняль от нижнего левого угла до правого верхнего
  TZXLSXBorder = array[0..5] of TZXLSXBorderItem;
  TZXLSXBordersArray = array of TZXLSXBorder;

  TZXLSXCellAlignment = record
    horizontal: TZHorizontalAlignment;
    indent: integer;
    shrinkToFit: boolean;
    textRotation: integer;
    vertical: TZVerticalAlignment;
    wrapText: boolean;
  end;

  TZXLSXCellStyle = record
    applyAlignment: boolean;
    applyBorder: boolean;
    applyFont: boolean;
    applyProtection: boolean;
    borderId: integer;
    fillId: integer;
    fontId: integer;
    numFmtId: integer;
    xfId: integer;
    hidden: boolean;
    locked: boolean;
    alignment: TZXLSXCellAlignment;
  end;

  TZXLSXCellStylesArray = array of TZXLSXCellStyle;

  type TZXLSXStyle = record
    builtinId: integer;     //??
    customBuiltin: boolean; //??
    name: string;           //??
    xfId: integer;
  end;

  TZXLSXStyleArray = array of TZXLSXStyle;

  TZXLSXFill = record
    patternfill: TZCellPattern;
    bgColorType: byte;  //0 - rgb, 1 - indexed, 2 - theme
    bgcolor: TColor;
    patterncolor: TColor;
    patternColorType: byte;
    lumFactorBG: double;
    lumFactorPattern: double;
  end;

  TZXLSXFillArray = array of TZXLSXFill;

  TZXLSXDFFont = record
    Color: TColor;
    ColorType: byte;
    LumFactor: double;
  end;

  TZXLSXDFFontArray = array of TZXLSXDFFont;

var
  xml: TZsspXMLReaderH;
  s: string;
  FontArray: TZEXLSXFontArray;
  FontCount: integer;
  BorderArray: TZXLSXBordersArray;
  BorderCount: integer;
  CellXfsArray: TZXLSXCellStylesArray;
  CellXfsCount: integer;
  CellStyleArray: TZXLSXCellStylesArray;
  CellStyleCount: integer;
  StyleArray: TZXLSXStyleArray;
  StyleCount: integer;
  FillArray: TZXLSXFillArray;
  FillCount: integer;
  indexedColor: TIntegerDynArray;
  indexedColorCount: integer;
  indexedColorMax: integer;
  _Style: TZStyle;
  t, i, n: integer;
  h1, s1, l1: double;
  _dfFonts: TZXLSXDFFontArray;
  _dfFills: TZXLSXFillArray;

  //Приводит к шрифту по-умолчанию
  //INPUT
  //  var fnt: TZEXLSXFont - шрифт
  procedure ZEXLSXZeroFont(var fnt: TZEXLSXFont);
  begin
    fnt.name := 'Arial';
    fnt.bold := false;
    fnt.italic := false;
    fnt.underline := false;
    fnt.strike := false;
    fnt.charset := 204;
    fnt.color := clBlack;
    fnt.LumFactor := 0;
    fnt.ColorType := 0;
    fnt.fontsize := 8;
  end; //ZEXLSXZeroFont

  //Обнуляет границы
  //  var border: TZXLSXBorder - границы
  procedure ZEXLSXZeroBorder(var border: TZXLSXBorder);
  var
    i: integer;

  begin
    for i := 0 to 5 do
    begin
      border[i].isColor := false;
      border[i].isEnabled := false;
      border[i].style := ZENone;
      border[i].Weight := 0;
    end;
  end; //ZEXLSXZeroBorder

  //Меняёт местами bgColor и fgColor при несплошных заливках
  //INPUT
  //  var PattFill: TZXLSXFill - заливка
  procedure ZEXLSXSwapPatternFillColors(var PattFill: TZXLSXFill);
  var
    t: integer;
    _b: byte;

  begin
    //если не сплошная заливка - нужно поменять местами цвета (bgColor <-> fgColor)
    if (not (PattFill.patternfill in [ZPNone, ZPSolid])) then
    begin
      t := PattFill.patterncolor;
      PattFill.patterncolor := PattFill.bgcolor;
      PattFill.bgColor := t;
      l1 := PattFill.lumFactorPattern;
      PattFill.lumFactorPattern := PattFill.lumFactorBG;
      PattFill.lumFactorBG := l1;

      _b := PattFill.patternColorType;
      PattFill.patternColorType := PattFill.bgColorType;
      PattFill.bgColorType := _b;
    end; //if
  end; //ZEXLSXSwapPatternFillColors

  //Очистить заливку ячейки
  //INPUT
  //  var PattFill: TZXLSXFill - заливка
  procedure ZEXLSXClearPatternFill(var PattFill: TZXLSXFill);
  begin
    PattFill.patternfill := ZPNone;
    PattFill.bgcolor := clWindow;
    PattFill.patterncolor := clWindow;
    PattFill.bgColorType := 0;
    PattFill.patternColorType := 0;
    PattFill.lumFactorBG := 0.0;
    PattFill.lumFactorPattern := 0.0;
  end; //ZEXLSXClearPatternFill

  //Обнуляет стиль
  //INPUT
  //  var style: TZXLSXCellStyle - стиль XLSX
  procedure ZEXLSXZeroCellStyle(var style: TZXLSXCellStyle);
  begin
    style.applyAlignment := false;
    style.applyBorder := false;
    style.applyProtection := false;
    style.hidden := false;
    style.locked := false;
    style.borderId := -1;
    style.fontId := -1;
    style.fillId := -1;
    style.numFmtId := -1;
    style.xfId := -1;
    style.alignment.horizontal := ZHAutomatic;
    style.alignment.vertical := ZVAutomatic;
    style.alignment.shrinkToFit := false;
    style.alignment.wrapText := false;
    style.alignment.textRotation := 0;
    style.alignment.indent := 0;
  end; //ZEXLSXZeroCellStyle

  //TZEXLSXFont в TFont
  //INPUT
  //  var fnt: TZEXLSXFont  - XLSX шрифт
  //  var font: TFont       - стандартный шрифт
  procedure ZEXLSXFontToFont(var fnt: TZEXLSXFont; var font: TFont);
  begin
    if (Assigned(font)) then
    begin
      if (fnt.bold) then
        font.Style := font.Style + [fsBold];
      if (fnt.italic) then
        font.Style := font.Style + [fsItalic];
      if (fnt.underline) then
        font.Style := font.Style + [fsUnderline];
      if (fnt.strike) then
        font.Style := font.Style + [fsStrikeOut];
      font.Charset := fnt.charset;
      font.Name := fnt.name;
      font.Size := fnt.fontsize;
    end;
  end; //ZEXLSXFontToFont

  //Прочитать цвет
  //INPUT
  //  var retColor: TColor      - возвращаемый цвет
  //  var retColorType: byte    - тип цвета: 0 - rgb, 1 - indexed, 2 - theme
  //  var retLumfactor: double  - яркость
  procedure ZXLSXGetColor(var retColor: TColor; var retColorType: byte; var retLumfactor: double);
  var
    t: integer;

  begin
    //I hate this f****** format! m$ office >= 2007 is big piece of shit! Arrgh!
    s := xml.Attributes.ItemsByName['rgb'];
    if (length(s) > 2) then
    begin
      delete(s, 1, 2);
      if (s > '') then
        retColor := HTMLHexToColor(s);
    end;
    s := xml.Attributes.ItemsByName['theme'];
    if (s > '') then
      if (TryStrToInt(s, t)) then
      begin
        retColorType := 2;
        retColor := t;
      end;
    s := xml.Attributes.ItemsByName['indexed'];
    if (s > '') then
      if (TryStrToInt(s, t)) then
      begin
        retColorType := 1;
        retColor := t;
      end;
    s := xml.Attributes.ItemsByName['tint'];
    if (s <> '') then
      retLumfactor := ZETryStrToFloat(s, 0);
  end; //ZXLSXGetColor

  procedure _ReadFonts();
  var
    _currFont: integer;

  begin
    _currFont := -1;
    while (not ((xml.TagName = 'fonts') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof) then
        break;

      s := xml.Attributes.ItemsByName['val'];
      if ((xml.TagName = 'font') and (xml.TagType = TAG_TYPE_START)) then
      begin
        _currFont := FontCount;
        inc(FontCount);
        SetLength(FontArray, FontCount);
        ZEXLSXZeroFont(FontArray[_currFont]);
      end else
      if (_currFont >= 0) then
      begin
        if ((xml.TagName = 'name') and (xml.TagType = TAG_TYPE_CLOSED)) then
          FontArray[_currFont].name := s
        else
        if ((xml.TagName = 'b') and (xml.TagType = TAG_TYPE_CLOSED)) then
          FontArray[_currFont].bold := true
        else
        if ((xml.TagName = 'charset') and (xml.TagType = TAG_TYPE_CLOSED)) then
        begin
          if (TryStrToInt(s, t)) then
            FontArray[_currFont].charset := t;
        end else
        if ((xml.TagName = 'color') and (xml.TagType = TAG_TYPE_CLOSED)) then
        begin
          ZXLSXGetColor(FontArray[_currFont].color,
                        FontArray[_currFont].ColorType,
                        FontArray[_currFont].LumFactor);
        end else
        if ((xml.TagName = 'i') and (xml.TagType = TAG_TYPE_CLOSED)) then
          FontArray[_currFont].italic := true
        else
  //      if ((xml.TagName = 'outline') and (xml.TagType = TAG_TYPE_CLOSED)) then
  //      begin
  //      end else
        if ((xml.TagName = 'strike') and (xml.TagType = TAG_TYPE_CLOSED)) then
          FontArray[_currFont].strike := true
        else
        if ((xml.TagName = 'sz') and (xml.TagType = TAG_TYPE_CLOSED)) then
        begin
          if (TryStrToInt(s, t)) then
            FontArray[_currFont].fontsize := t;
        end else
        if ((xml.TagName = 'u') and (xml.TagType = TAG_TYPE_CLOSED)) then
        begin
          if (s > '') then
            if (s <> 'none') then
              FontArray[_currFont].underline := true;
        end
        else
        if ((xml.TagName = 'vertAlign') and (xml.TagType = TAG_TYPE_CLOSED)) then
        begin
        end;
      end; //if  
      //Тэги настройки шрифта
      //*b - bold
      //*charset
      //*color
      //?condense
      //?extend
      //?family
      //*i - italic
      //*name
      //?outline
      //?scheme
      //?shadow
      //*strike
      //*sz - size
      //*u - underline
      //*vertAlign

    end; //while
  end; //_ReadFonts

  //Получить тип заливки
  function _GetPatternFillByStr(const s: string): TZCellPattern;
  begin
    if (s = 'solid') then
      result := ZPSolid
    else
    if (s = 'none') then
      result := ZPNone
    else
    if (s = 'gray125') then
      result := ZPGray125
    else
    if (s = 'gray0625') then
      result := ZPGray0625
    else
    if (s = 'darkUp') then
      result := ZPDiagStripe
    else
    if (s = 'mediumGray') then
      result := ZPGray50
    else
    if (s = 'darkGray') then
      result := ZPGray75
    else
    if (s = 'lightGray') then
      result := ZPGray25
    else
    if (s = 'darkHorizontal') then
      result := ZPHorzStripe
    else
    if (s = 'darkVertical') then
      result := ZPVertStripe
    else
    if (s = 'darkDown') then
      result := ZPReverseDiagStripe
    else
    if (s = 'darkUpDark') then
      result := ZPDiagStripe
    else
    if (s = 'darkGrid') then
      result := ZPDiagCross
    else
    if (s = 'darkTrellis') then
      result := ZPThickDiagCross
    else
    if (s = 'lightHorizontal') then
      result := ZPThinHorzStripe
    else
    if (s = 'lightVertical') then
      result := ZPThinVertStripe
    else
    if (s = 'lightDown') then
      result := ZPThinReverseDiagStripe
    else
    if (s = 'lightUp') then
      result := ZPThinDiagStripe
    else
    if (s = 'lightGrid') then
      result := ZPThinHorzCross
    else
    if (s = 'lightTrellis') then
      result := ZPThinDiagCross
    //
    else
      result := ZPSolid; //{tut} потом подумать насчёт стилей границ
  end; //_GetPatternFillByStr

  //Определить стиль начертания границы
  //INPUT
  //  const st: string            - название стиля
  //  var retWidth: byte          - возвращаемая ширина линии
  //  var retStyle: TZBorderType  - возвращаемый стиль начертания линии
  //RETURN
  //      boolean - true - стиль определён
  function XLSXGetBorderStyle(const st: string; var retWidth: byte; var retStyle: TZBorderType): boolean;
  begin
    result := true;
    retWidth := 1;
    if (st = 'thin') then
    begin
      retStyle := ZEContinuous;
    end else
    if (st = 'hair') then
      retStyle := ZEHair
    else
    if (st = 'dashed') then
      retStyle := ZEDash
    else
    if (st = 'dotted') then
      retStyle := ZEDot
    else
    if (st = 'dashDot') then
      retStyle := ZEDashDot
    else
    if (st = 'dashDotDot') then
      retStyle := ZEDashDotDot
    else
    if (st = 'slantDashDot') then
      retStyle := ZESlantDashDot
    else
    if (st = 'double') then
      retStyle := ZEDouble
    else
    if (st = 'medium') then
    begin
      retStyle := ZEContinuous;
      retWidth := 2;
    end else
    if (st = 'thick') then
    begin
      retStyle := ZEContinuous;
      retWidth := 3;
    end else
    if (st = 'mediumDashed') then
    begin
      retStyle := ZEDash;
      retWidth := 2;
    end else
    if (st = 'mediumDashDot') then
    begin
      retStyle := ZEDashDot;
      retWidth := 2;
    end else
    if (st = 'mediumDashDotDot') then
    begin
      retStyle := ZEDashDotDot;
      retWidth := 2;
    end else
    if (st = 'none') then
      retStyle := ZENone
    else
      result := false;
  end; //XLSXGetBorderStyle

  procedure _ReadBorders();
  var
    _diagDown, _diagUP: boolean;
    _currBorder: integer; //текущий набор границ
    _currBorderItem: integer; //текущая граница (левая/правая ...)
    _color: TColor;
    _isColor: boolean;

    procedure _SetCurBorder(borderNum: integer);
    begin
      _currBorderItem := borderNum;
      s := xml.Attributes.ItemsByName['style'];
      if (s > '') then
      begin
        BorderArray[_currBorder][borderNum].isEnabled :=
          XLSXGetBorderStyle(s,
                             BorderArray[_currBorder][borderNum].Weight,
                             BorderArray[_currBorder][borderNum].style);
      end;
    end; //_SetCurBorder

  begin
    _currBorderItem := -1;
    _diagDown := false;
    _diagUP := false;
    _color := clBlack;
    while (not ((xml.TagName = 'borders') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'border') and (xml.TagType = TAG_TYPE_START)) then
      begin
        _currBorder := BorderCount;
        inc(BorderCount);
        SetLength(BorderArray, BorderCount);
        ZEXLSXZeroBorder(BorderArray[_currBorder]);
        _diagDown := false;
        _diagUP := false;
        s := xml.Attributes.ItemsByName['diagonalDown'];
        if (s > '') then
          _diagDown := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['diagonalUp'];
        if (s > '') then
          _diagUP := ZEStrToBoolean(s);
      end else
      begin
        if (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED]) then
        begin
          if (xml.TagName = 'left') then
          begin
            _SetCurBorder(0);
          end else
          if (xml.TagName = 'right') then
          begin
            _SetCurBorder(2);
          end else
          if (xml.TagName = 'top') then
          begin
            _SetCurBorder(1);
          end else
          if (xml.TagName = 'bottom') then
          begin
            _SetCurBorder(3);
          end else
          if (xml.TagName = 'diagonal') then
          begin
            if (_diagUp) then
              _SetCurBorder(5);
            if (_diagDown) then
            begin
              if (_diagUp) then
                BorderArray[_currBorder][4] := BorderArray[_currBorder][5]
              else
                _SetCurBorder(4);
            end;
          end else
          //??  {
          if (xml.TagName = 'end') then
          begin
          end else
          if (xml.TagName = 'start') then
          begin
          end else
          //??  }
          if (xml.TagName = 'color') then
          begin
            _isColor := false;
            s := xml.Attributes.ItemsByName['rgb'];
            if (length(s) > 2) then
              delete(s, 1, 2);
            if (s > '') then
            begin
              _color := HTMLHexToColor(s);
              _isColor := true;
            end;
            if (_isColor and (_currBorderItem >= 0) and (_currBorderItem < 6)) then
            begin
              BorderArray[_currBorder][_currBorderItem].color := _color;
              BorderArray[_currBorder][_currBorderItem].isColor := true;
            end;
          end;
        end; //if
      end; //else

    end; //while
  end; //_ReadBorders

  procedure _ReadFills();
  var
    _currFill: integer;

  begin
    _currFill := -1;
    while (not ((xml.TagName = 'fills') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'fill') and (xml.TagType = TAG_TYPE_START)) then
      begin
        _currFill := FillCount;
        inc(FillCount);
        SetLength(FillArray, FillCount);
        ZEXLSXClearPatternFill(FillArray[_currFill]);
      end else
      if ((xml.TagName = 'patternFill') and (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED])) then
      begin
        if (_currFill >= 0) then
        begin
          s := xml.Attributes.ItemsByName['patternType'];
          {
          *none	None
          *solid	Solid
          ?mediumGray	Medium Gray
          ?darkGray	Dary Gray
          ?lightGray	Light Gray
          ?darkHorizontal	Dark Horizontal
          ?darkVertical	Dark Vertical
          ?darkDown	Dark Down
          ?darkUpDark Up
          ?darkGrid	Dark Grid
          ?darkTrellis	Dark Trellis
          ?lightHorizontal	Light Horizontal
          ?lightVertical	Light Vertical
          ?lightDown	Light Down
          ?lightUp	Light Up
          ?lightGrid	Light Grid
          ?lightTrellis	Light Trellis
          *gray125	Gray 0.125
          *gray0625	Gray 0.0625
          }

          if (s > '') then
            FillArray[_currFill].patternfill := _GetPatternFillByStr(s);
        end;
      end else
      if ((xml.TagName = 'bgColor') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        if (_currFill >= 0) then
        begin
          ZXLSXGetColor(FillArray[_currFill].patterncolor,
                        FillArray[_currFill].patternColorType,
                        FillArray[_currFill].lumFactorPattern);

          //если не сплошная заливка - нужно поменять местами цвета (bgColor <-> fgColor)
          ZEXLSXSwapPatternFillColors(FillArray[_currFill]);
        end;
      end else
      if ((xml.TagName = 'fgColor') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        if (_currFill >= 0) then
          ZXLSXGetColor(FillArray[_currFill].bgcolor,
                        FillArray[_currFill].bgColorType,
                        FillArray[_currFill].lumFactorBG);
      end; //fgColor
    end; //while
  end; //_ReadFills

  //Читает стили (cellXfs и cellStyleXfs)
  //INPUT
  //  const TagName: string           - имя тэга
  //  var CSA: TZXLSXCellStylesArray  - массив со стилями
  //  var StyleCount: integer         - кол-во стилей
  procedure _ReadCellCommonStyles(const TagName: string; var CSA: TZXLSXCellStylesArray; var StyleCount: integer);
  var
    _currCell: integer;
    b: boolean;

  begin
    _currCell := -1;
    while (not ((xml.TagName = TagName) and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      b := false;  

      if ((xml.TagName = 'xf') and (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED])) then
      begin
        _currCell := StyleCount;
        inc(StyleCount);
        SetLength(CSA, StyleCount);
        ZEXLSXZeroCellStyle(CSA[_currCell]);
        s := xml.Attributes.ItemsByName['applyAlignment'];
        if (s > '') then
          CSA[_currCell].applyAlignment := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['applyBorder'];
        if (s > '') then
          CSA[_currCell].applyBorder := ZEStrToBoolean(s)
        else
          b := true;

        s := xml.Attributes.ItemsByName['applyFont'];
        if (s > '') then
          CSA[_currCell].applyFont := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['applyProtection'];
        if (s > '') then
          CSA[_currCell].applyProtection := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['borderId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
          begin
            CSA[_currCell].borderId := t;
            if (b and (t >= 1)) then
              CSA[_currCell].applyBorder := true;
          end;

        s := xml.Attributes.ItemsByName['fillId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].fillId := t;

        s := xml.Attributes.ItemsByName['fontId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].fontId := t;

        s := xml.Attributes.ItemsByName['numFmtId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].numFmtId := t;

        s := xml.Attributes.ItemsByName['xfId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].xfId := t;
      end else
      if ((xml.TagName = 'alignment') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        if (_currCell >= 0) then
        begin
          s := xml.Attributes.ItemsByName['horizontal'];
          if (s > '') then
          begin
            if (s = 'general') then
              CSA[_currCell].alignment.horizontal := ZHAutomatic
            else
            if (s = 'left') then
              CSA[_currCell].alignment.horizontal := ZHLeft
            else
            if (s = 'right') then
              CSA[_currCell].alignment.horizontal := ZHRight
            else
            if ((s = 'center') or (s = 'centerContinuous')) then
              CSA[_currCell].alignment.horizontal := ZHCenter
            else
            if (s = 'fill') then
              CSA[_currCell].alignment.horizontal := ZHFill
            else
            if (s = 'justify') then
              CSA[_currCell].alignment.horizontal := ZHJustify
            else
            if (s = 'distributed') then
              CSA[_currCell].alignment.horizontal := ZHDistributed;
          end;

          s := xml.Attributes.ItemsByName['indent'];
          if (s > '') then
            if (TryStrToInt(s, t)) then
              CSA[_currCell].alignment.indent := t;

          s := xml.Attributes.ItemsByName['shrinkToFit'];
          if (s > '') then
            CSA[_currCell].alignment.shrinkToFit := ZEStrToBoolean(s);

          s := xml.Attributes.ItemsByName['textRotation'];
          if (s > '') then
            if (TryStrToInt(s, t)) then
              CSA[_currCell].alignment.textRotation := t;

          s := xml.Attributes.ItemsByName['vertical'];
          if (s > '') then
          begin
            if (s = 'center') then
              CSA[_currCell].alignment.vertical := ZVCenter
            else
            if (s = 'top') then
              CSA[_currCell].alignment.vertical := ZVTop
            else
            if (s = 'bottom') then
              CSA[_currCell].alignment.vertical := ZVBottom
            else
            if (s = 'justify') then
              CSA[_currCell].alignment.vertical := ZVJustify
            else
            if (s = 'distributed') then
              CSA[_currCell].alignment.vertical := ZVDistributed;
          end;

          s := xml.Attributes.ItemsByName['wrapText'];
          if (s > '') then
            CSA[_currCell].alignment.wrapText := ZEStrToBoolean(s);
        end; //if
      end else
      if ((xml.TagName = 'protection') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        if (_currCell >= 0) then
        begin
          s := xml.Attributes.ItemsByName['hidden'];
          if (s > '') then
            CSA[_currCell].hidden := ZEStrToBoolean(s);

          s := xml.Attributes.ItemsByName['locked'];
          if (s > '') then
            CSA[_currCell].locked := ZEStrToBoolean(s);
        end;
      end;
    end; //while
  end; //_ReadCellCommonStyles

  //Сами стили ?? (или для чего они вообще?)
  procedure _ReadCellStyles();
  var
    b: boolean;

  begin
    while (not ((xml.TagName = 'cellStyles') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'cellStyle') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        b := false;
        SetLength(StyleArray, StyleCount + 1);
        s := xml.Attributes.ItemsByName['builtinId']; //?
        if (s > '') then
          if (TryStrToInt(s, t)) then
            StyleArray[StyleCount].builtinId := t;

        s := xml.Attributes.ItemsByName['customBuiltin']; //?
        if (s > '') then
          StyleArray[StyleCount].customBuiltin := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['name']; //?
          StyleArray[StyleCount].name := s;

        s := xml.Attributes.ItemsByName['xfId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
          begin
            StyleArray[StyleCount].xfId := t;
            b := true;
          end;

        if (b) then
          inc(StyleCount);
      end;
    end; //while
  end; //_ReadCellStyles

  procedure _ReadColors();
  begin
    while (not ((xml.TagName = 'colors') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'rgbColor') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        s := xml.Attributes.ItemsByName['rgb'];
        if (length(s) > 2) then
          delete(s, 1, 2);
        if (s > '') then
        begin
          inc(indexedColorCount);
          if (indexedColorCount >= indexedColorMax) then
          begin
            indexedColorMax := indexedColorCount + 80;
            SetLength(indexedColor, indexedColorMax);
          end;
          indexedColor[indexedColorCount - 1] := HTMLHexToColor(s);
        end;
      end;
    end; //while
  end; //_ReadColors

  //Конвертирует RGB в HSL
  //http://en.wikipedia.org/wiki/HSL_color_space
  //INPUT
  //      r: byte     -
  //      g: byte     -
  //      b: byte     -
  //  out h: double   - Hue - тон цвета
  //  out s: double   - Saturation - насыщенность
  //  out l: double   - Lightness (Intensity) - светлота (яркость)
  procedure ZRGBToHSL(r, g, b: byte; out h, s, l: double);
  var
    _max, _min: double;
    intMax, intMin: integer;
    _r, _g, _b: double;
    _delta: double;
    _idx: integer;

  begin
    _r := r / 255;
    _g := g / 255;
    _b := b / 255;

    intMax := Max(r, Max(g, b));
    intMin := Min(r, Min(g, b));

    _max := Max(_r, Max(_g, _b));
    _min := Min(_r, Min(_g, _b));

    h := (_max + _min) * 0.5;
    s := h;
    l := h;
    if (intMax = intMin) then
    begin
      h := 0;
      s := 0;
    end else
    begin
      _delta := _max - _min;
      if (l > 0.5) then
        s := _delta / (2 - _max - _min)
      else
        s := _delta / (_max + _min);

        if (intMax = r) then
          _idx := 1
        else
        if (intMax = g) then
          _idx := 2
        else
          _idx := 3;

        case (_idx) of
          1:
            begin
              h := (_g - _b) / _delta;
              if (g < b) then
                h := h + 6;
            end;
          2: h := (_b - _r) / _delta + 2;
          3: h := (_r - _g) / _delta + 4;
        end;
        
        h := h / 6;
    end;
  end; //ZRGBToHSL

  //Конвертирует TColor (RGB) в HSL
  //http://en.wikipedia.org/wiki/HSL_color_space
  //INPUT
  //      Color: TColor - цвет
  //  out h: double     - Hue - тон цвета
  //  out s: double     - Saturation - насыщенность
  //  out l: double     - Lightness (Intensity) - светлота (яркость)
  procedure ZColorToHSL(Color: TColor; out h, s, l: double);
  var
    _RGB: integer;

  begin
    _RGB := ColorToRGB(Color);
    ZRGBToHSL(byte(_RGB), byte(_RGB shr 8), byte(_RGB shr 16), h, s, l);
  end; //ZColorToHSL

  //Конвертирует HSL в RGB
  //http://en.wikipedia.org/wiki/HSL_color_space
  //INPUT
  //      h: double - Hue - тон цвета
  //      s: double - Saturation - насыщенность
  //      l: double - Lightness (Intensity) - светлота (яркость)
  //  out r: byte   -
  //  out g: byte   -
  //  out b: byte   -
  procedure ZHSLToRGB(h, s, l: double; out r, g, b: byte);
  var
    _r, _g, _b: double;
    q, p: double;

    function HueToRgb(p, q, t: double): double;
    begin
      result := p;
      if (t < 0) then
        t := t + 1;
      if (t > 1) then
        t := t - 1;
      if (t < 1/6) then
        result := p + (q - p) * 6 * t
      else
      if (t < 0.5) then
        result := q
      else
      if (t < 2/3) then
        result := p + (q - p) * (2/3 - t) * 6;
    end; //HueToRgb

  begin
    if (s = 0) then
    begin
      //Оттенок серого
      _r := l;
      _g := l;
      _b := l;
    end else
    begin
      if (l < 0.5) then
        q := l * (1 + s)
      else
        q := l + s - l * s;
      p := 2 * l - q;
      _r := HueToRgb(p, q, h + 1/3);
      _g := HueToRgb(p, q, h);
      _b := HueToRgb(p, q, h - 1/3);
    end;
    r := byte(round(_r * 255));
    g := byte(round(_g * 255));
    b := byte(round(_b * 255));
  end; //ZHSLToRGB

  //Конвертирует HSL в Color
  //http://en.wikipedia.org/wiki/HSL_color_space
  //INPUT
  //      h: double - Hue - тон цвета
  //      s: double - Saturation - насыщенность
  //      l: double - Lightness (Intensity) - светлота (яркость)
  //RETURN
  //      TColor - цвет
  function ZHSLToColor(h, s, l: double): TColor;
  var
    r, g, b: byte;

  begin
    ZHSLToRGB(h, s, l, r, g, b);
    result := (b shl 16) or (g shl 8) or r;
  end; //ZHSLToColor

  //Применить tint к цвету
  // Thanks Tomasz Wieckowski!
  //   http://msdn.microsoft.com/en-us/library/ff532470%28v=office.12%29.aspx
  procedure ApplyLumFactor(var Color: TColor; var lumFactor: double);
  begin
    //+delta?
    if (lumFactor <> 0.0) then
    begin
      ZColorToHSL(Color, h1, s1, l1);
      lumFactor := 1 - lumFactor;

      if (l1 = 1) then
        l1 := l1 * (1 - lumFactor)
      else
        l1 := l1 * lumFactor + (1 - lumFactor);

      Color := ZHSLtoColor(h1, s1, l1);
    end;
  end; //ApplyLumFactor

  //Differential Formatting для xlsx
  procedure _Readdxfs();
  var
    _df: TZXLSXDiffFormattingItem;
    _dfIndex: integer;
    _tmptag: string;

    procedure _addFontStyle(fnts: TFontStyle);
    begin
      _df.FontStyles := _df.FontStyles + [fnts];
      _df.UseFontStyles := true;
    end;

    procedure _ReadDFFont();
    begin
      _df.UseFont := true;
      while ((xml.TagType <> TAG_TYPE_END) and (xml.TagName <> 'font')) do
      begin
        xml.ReadTag();
        if (xml.Eof) then
          break;

        _tmptag := xml.TagName;
        if (_tmptag = 'i') then
          _addFontStyle(fsItalic);
        if (_tmptag = 'b') then
          _addFontStyle(fsBold);
        if (_tmptag = 'u') then
          _addFontStyle(fsUnderline);
        if (_tmptag = 'strike') then
          _addFontStyle(fsStrikeOut);

        if (_tmptag = 'color') then
        begin
          _df.UseFontColor := true;
          ZXLSXGetColor(_dfFonts[_dfIndex].Color,
                        _dfFonts[_dfIndex].ColorType,
                        _dfFonts[_dfIndex].LumFactor);
        end;
      end; //while
    end; //_ReadDFFont

    procedure _ReadDFFill();
    begin
      _df.UseFill := true;
      while ((xml.TagType <> TAG_TYPE_END) or (xml.TagName <> 'fill')) do
      begin
        xml.ReadTag();
        if (xml.Eof) then
          break;

        if (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED]) then
        begin
          if (xml.TagName = 'patternFill') then
          begin
            s := xml.Attributes.ItemsByName['patternType'];
            if (s <> '') then
            begin
              _df.UseCellPattern := true;
              _df.CellPattern := _GetPatternFillByStr(s);
            end;
          end else
          if (xml.TagName = 'bgColor') then
          begin
            _df.UseBGColor := true;
            ZXLSXGetColor(_dfFills[_dfIndex].bgcolor,
                          _dfFills[_dfIndex].bgColorType,
                          _dfFills[_dfIndex].lumFactorBG)
          end else
          if (xml.TagName = 'fgColor') then
          begin
            _df.UsePatternColor := true;
            ZXLSXGetColor(_dfFills[_dfIndex].patterncolor,
                          _dfFills[_dfIndex].patternColorType,
                          _dfFills[_dfIndex].lumFactorPattern);
            ZEXLSXSwapPatternFillColors(_dfFills[_dfIndex]);
          end;
        end;
      end; //while
    end; //_ReadDFFill

    procedure _ReadDFBorder();
    var
      _borderNum: integer;
      t: byte;
      _bt: TZBorderType;

      procedure _SetDFBorder(BorderNum: integer);
      begin
        _borderNum := BorderNum;
        s := xml.Attributes['style'];
        if (s <> '') then
          if (XLSXGetBorderStyle(s, t, _bt)) then
          begin
            _df.UseBorder := true;
            _df.Borders[BorderNum].Weight := t;
            _df.Borders[BorderNum].LineStyle := _bt;
            _df.Borders[BorderNum].UseStyle := true;
          end;
      end; //_SetDFBorder

    begin
      _df.UseBorder := true;
      _borderNum := 0;
      while ((xml.TagType <> TAG_TYPE_END) or (xml.TagName <> 'border')) do
      begin
        xml.ReadTag();
        if (xml.Eof) then
          break;

        if (xml.TagType in [TAG_TYPE_START, TAG_TYPE_CLOSED]) then
        begin
          _tmptag := xml.TagName;
          if (_tmptag = 'left') then
            _SetDFBorder(0)
          else
          if (_tmptag = 'right') then
            _SetDFBorder(2)
          else
          if (_tmptag = 'top') then
            _SetDFBorder(1)
          else
          if (_tmptag = 'bottom') then
            _SetDFBorder(3)
          else
          if (_tmptag = 'vertical') then
            _SetDFBorder(4)
          else
          if (_tmptag = 'horizontal') then
            _SetDFBorder(5)
          else
          if (_tmptag = 'color') then
          begin
            s := xml.Attributes['rgb'];
            if (length(s) > 2) then
              delete(s, 1, 2);
            if ((_borderNum >= 0) and (_borderNum < 6)) then
              if (s <> '') then
              begin
                _df.UseBorder := true;
                _df.Borders[_borderNum].UseColor := true;
                _df.Borders[_borderNum].Color := HTMLHexToColor(s);
              end;
          end;
        end; //if
      end; //while
    end; //_ReadDFBorder

    procedure _ReaddxfItem();
    begin
      _dfIndex := ReadHelper.DiffFormatting.Count;

      SetLength(_dfFonts, _dfIndex + 1);
      _dfFonts[_dfIndex].ColorType := 0;
      _dfFonts[_dfIndex].LumFactor := 0;

      SetLength(_dfFills, _dfIndex + 1);
      ZEXLSXClearPatternFill(_dfFills[_dfIndex]);

      ReadHelper.DiffFormatting.Add();
      _df := ReadHelper.DiffFormatting[_dfIndex];
      while ((xml.TagType <> TAG_TYPE_END) or (xml.TagName <> 'dxf')) do
      begin
        xml.ReadTag();
        if (xml.Eof) then
          break;
        if (xml.TagType = TAG_TYPE_START) then
        begin
          if (xml.TagName = 'font') then
            _ReadDFFont()
          else
          if (xml.TagName = 'fill') then
            _ReadDFFill()
          else
          if (xml.TagName = 'border') then
            _ReadDFBorder();
        end;
      end; //while
    end; //_ReaddxfItem

  begin
    while ((xml.TagType <> TAG_TYPE_END) or (xml.TagName <> 'dxfs')) do
    begin
      xml.ReadTag();
      if (xml.Eof) then
        break;
      if ((xml.TagType = TAG_TYPE_START) and (xml.TagName = 'dxf')) then
        _ReaddxfItem();
    end; //while
  end; //_Readdxfs

  procedure XLSXApplyColor(var AColor: TColor; ColorType: byte; LumFactor: double);
  begin
    //Thema color
    if (ColorType = 2) then
    begin
      t := AColor - 1;
      if ((t >= 0) and (t < ThemaColorCount)) then
        AColor := ThemaFillsColors[t];
    end;
    if (ColorType = 1) then
      if ((AColor >= 0) and (AColor < indexedColorCount))  then
        AColor := indexedColor[AColor];
    ApplyLumFactor(AColor, LumFactor);
  end; //XLSXApplyColor

  //Применить стиль
  //INPUT
  //  var XMLSSStyle: TZStyle         - стиль в хранилище
  //  var XLSXStyle: TZXLSXCellStyle  - стиль в xlsx
  procedure _ApplyStyle(var XMLSSStyle: TZStyle; var XLSXStyle: TZXLSXCellStyle);
  var
    i: integer;

  begin
    if (XLSXStyle.numFmtId >= 0) then
      XMLSSStyle.NumberFormat := ReadHelper.NumberFormats.GetFormat(XLSXStyle.numFmtId);
    
    if (XLSXStyle.applyAlignment) then
    begin
      XMLSSStyle.Alignment.Horizontal := XLSXStyle.alignment.horizontal;
      XMLSSStyle.Alignment.Vertical := XLSXStyle.alignment.vertical;
      XMLSSStyle.Alignment.Indent := XLSXStyle.alignment.indent;
      XMLSSStyle.Alignment.ShrinkToFit := XLSXStyle.alignment.shrinkToFit;
      XMLSSStyle.Alignment.WrapText := XLSXStyle.alignment.wrapText;

      XMLSSStyle.Alignment.Rotate := 0;
      i := XLSXStyle.alignment.textRotation;
      XMLSSStyle.Alignment.VerticalText := (i = 255);
      if (i >= 0) and (i <= 180) then begin
        if i > 90 then i := 90 - i;
        XMLSSStyle.Alignment.Rotate := i
      end;
    end;

    if (XLSXStyle.applyBorder) then
    begin
      n := XLSXStyle.borderId;
      if ((n >= 0) and (n < BorderCount)) then
        for i := 0 to 5 do
        if (BorderArray[n][i].isEnabled) then
        begin
          XMLSSStyle.Border[i].LineStyle := BorderArray[n][i].style;
          XMLSSStyle.Border[i].Weight := BorderArray[n][i].Weight;
          if (BorderArray[n][i].isColor) then
            XMLSSStyle.Border[i].Color := BorderArray[n][i].color;
        end;
    end;

    if (XLSXStyle.applyFont) then
    begin
      n := XLSXStyle.fontId;
      if ((n >= 0) and (n < FontCount)) then
      begin
        XLSXApplyColor(FontArray[n].color,
                       FontArray[n].ColorType,
                       FontArray[n].LumFactor);
        XMLSSStyle.Font.Name := FontArray[n].name;
        XMLSSStyle.Font.Size := FontArray[n].fontsize;
        XMLSSStyle.Font.Charset := FontArray[n].charset;
        XMLSSStyle.Font.Color := FontArray[n].color;
        if (FontArray[n].bold) then
          XMLSSStyle.Font.Style := [fsBold];
        if (FontArray[n].underline) then
          XMLSSStyle.Font.Style := XMLSSStyle.Font.Style + [fsUnderline];
        if (FontArray[n].italic) then
          XMLSSStyle.Font.Style := XMLSSStyle.Font.Style + [fsItalic];
        if (FontArray[n].strike) then
          XMLSSStyle.Font.Style := XMLSSStyle.Font.Style + [fsStrikeOut];
      end;
    end;

    if (XLSXStyle.applyProtection) then
    begin
      XMLSSStyle.Protect := XLSXStyle.locked;
      XMLSSStyle.HideFormula := XLSXStyle.hidden;
    end;

    n := XLSXStyle.fillId;
    if ((n >= 0) and (n < FillCount)) then
    begin
      XMLSSStyle.CellPattern := FillArray[n].patternfill;
      XMLSSStyle.BGColor := FillArray[n].bgcolor;
      XMLSSStyle.PatternColor := FillArray[n].patterncolor;
    end;
  end; //_ApplyStyle

  //I hate this f*****g format!
  //  Sometimes in styles.xml there are no <Colors> .. </Colors>! =_="
  //  Using standart colors.
  procedure _CheckIndexedColors();
  const
    _standart: array [0..63] of string =
    (
      '#000000',      //0
      '#FFFFFF',      //1
      '#FF0000',      //2
      '#00FF00',      //3
      '#0000FF',      //4
      '#FFFF00',      //5
      '#FF00FF',      //6
      '#00FFFF',      //7
      '#000000',      //8
      '#FFFFFF',      //9
      '#FF0000',      //10
      '#00FF00',      //11
      '#0000FF',      //12
      '#FFFF00',      //13
      '#FF00FF',      //14
      '#00FFFF',      //15
      '#800000',      //16
      '#008000',      //17
      '#000090',      //18
      '#808000',      //19
      '#800080',      //20
      '#008080',      //21
      '#C0C0C0',      //22
      '#808080',      //23
      '#9999FF',      //24
      '#993366',      //25
      '#FFFFCC',      //26
      '#CCFFFF',      //27
      '#660066',      //28
      '#FF8080',      //29
      '#0066CC',      //30
      '#CCCCFF',      //31
      '#000080',      //32
      '#FF00FF',      //33
      '#FFFF00',      //34
      '#00FFFF',      //35
      '#800080',      //36
      '#800000',      //37
      '#008080',      //38
      '#0000FF',      //39
      '#00CCFF',      //40
      '#CCFFFF',      //41
      '#CCFFCC',      //42
      '#FFFF99',      //43
      '#99CCFF',      //44
      '#FF99CC',      //45
      '#CC99FF',      //46
      '#FFCC99',      //47
      '#3366FF',      //48
      '#33CCCC',      //49
      '#99CC00',      //50
      '#FFCC00',      //51
      '#FF9900',      //52
      '#FF6600',      //53
      '#666699',      //54
      '#969696',      //55
      '#003366',      //56
      '#339966',      //57
      '#003300',      //58
      '#333300',      //59
      '#993300',      //60
      '#993366',      //61
      '#333399',      //62
      '#333333'       //63
    );
  var
    i: integer;

  begin
    if (indexedColorCount = 0) then
    begin
      indexedColorCount := 63;
      indexedColorMax := indexedColorCount + 10;
      SetLength(indexedColor, indexedColorMax);
      for i := 0 to 63 do
        indexedColor[i] := HTMLHexToColor(_standart[i]);
    end;
  end; //_CheckIndexedColors

begin
  result := false;
  xml := nil;
  CellXfsArray := nil;
  CellStyleArray := nil;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    FontCount := 0;
    BorderCount := 0;
    CellStyleCount := 0;
    StyleCount := 0;
    CellXfsCount := 0;
    FillCount := 0;
    indexedColorCount := 0;
    indexedColorMax := -1;

    while (not xml.Eof()) do
    begin
      xml.ReadTag();

      if ((xml.TagName = 'fonts') and (xml.TagType = TAG_TYPE_START)) then
        _ReadFonts()
      else
      if ((xml.TagName = 'borders') and (xml.TagType = TAG_TYPE_START)) then
        _ReadBorders()
      else
      if ((xml.TagName = 'fills') and (xml.TagType = TAG_TYPE_START)) then
        _ReadFills()
      else
      //TODO: разобраться, чем отличаются cellStyleXfs и cellXfs. Что за cellStyles?
      if ((xml.TagName = 'cellStyleXfs') and (xml.TagType = TAG_TYPE_START)) then
        _ReadCellCommonStyles('cellStyleXfs', CellStyleArray, CellStyleCount)//_ReadCellStyleXfs()
      else
      if ((xml.TagName = 'cellXfs') and (xml.TagType = TAG_TYPE_START)) then  //сами стили?
        _ReadCellCommonStyles('cellXfs', CellXfsArray, CellXfsCount) //_ReadCellXfs()
      else
      if ((xml.TagName = 'cellStyles') and (xml.TagType = TAG_TYPE_START)) then //??
        _ReadCellStyles()
      else
      if ((xml.TagName = 'colors') and (xml.TagType = TAG_TYPE_START)) then
        _ReadColors()
      else
      if ((xml.TagType = TAG_TYPE_START) and (xml.TagName = 'dxfs')) then
        _Readdxfs()
      else
      if ((xml.TagType = TAG_TYPE_START) and (xml.TagName = 'numFmts')) then
        ReadHelper.NumberFormats.ReadNumFmts(xml);
    end; //while

    //тут незабыть применить номера цветов, если были введены

    _CheckIndexedColors();

    //
    for i := 0 to FillCount - 1 do
    begin
      XLSXApplyColor(FillArray[i].bgcolor, FillArray[i].bgColorType, FillArray[i].lumFactorBG);
      XLSXApplyColor(FillArray[i].patterncolor, FillArray[i].patternColorType, FillArray[i].lumFactorPattern);
    end; //for

    //{tut}

    XMLSS.Styles.Count := CellXfsCount;
    ReadHelper.NumberFormats.StyleFMTCount := CellXfsCount;
    for i := 0 to CellXfsCount - 1 do
    begin
      t := CellXfsArray[i].xfId;
      ReadHelper.NumberFormats.StyleFMTID[i] := CellXfsArray[i].numFmtId;

      _Style := XMLSS.Styles[i];
      if ((t >= 0) and (t < CellStyleCount)) then
        _ApplyStyle(_Style, CellStyleArray[t]);
      _ApplyStyle(_Style, CellXfsArray[i]);
    end;

    //Применение цветов к DF
    for i := 0 to ReadHelper.DiffFormatting.Count - 1 do
    begin
      if (ReadHelper.DiffFormatting[i].UseFontColor) then
      begin
        XLSXApplyColor(_dfFonts[i].Color, _dfFonts[i].ColorType, _dfFonts[i].LumFactor);
        ReadHelper.DiffFormatting[i].FontColor := _dfFonts[i].Color;
      end;
      if (ReadHelper.DiffFormatting[i].UseBGColor) then
      begin
        XLSXApplyColor(_dfFills[i].bgcolor, _dfFills[i].bgColorType, _dfFills[i].lumFactorBG);
        ReadHelper.DiffFormatting[i].BGColor := _dfFills[i].bgcolor;
      end;
      if (ReadHelper.DiffFormatting[i].UsePatternColor) then
      begin
        XLSXApplyColor(_dfFills[i].patterncolor, _dfFills[i].patternColorType, _dfFills[i].lumFactorPattern);
        ReadHelper.DiffFormatting[i].PatternColor := _dfFills[i].patterncolor;
      end;
    end;

    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
    SetLength(FontArray, 0);
    FontArray := nil;
    SetLength(BorderArray, 0);
    BorderArray := nil;
    SetLength(CellStyleArray, 0);
    CellStyleArray := nil;
    SetLength(StyleArray, 0);
    StyleArray := nil;
    SetLength(CellXfsArray, 0);
    CellXfsArray := nil;
    SetLength(FillArray, 0);
    FillArray := nil;
    SetLength(indexedColor, 0);
    indexedColor := nil;
    SetLength(_dfFonts, 0);
    SetLength(_dfFills, 0);
  end;
end; //ZEXSLXReadStyles

//Читает названия листов (workbook.xml)
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//  var Stream: TStream                 - поток
//  var Relations: TZXLSXRelationsArray - связи
//  var RelationsCount: integer         - кол-во
//RETURN
//      boolean - true - названия прочитались без ошибок
function ZEXSLXReadWorkBook(var XMLSS: TZEXMLSS; var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer): boolean;
var
  xml: TZsspXMLReaderH;
  s: string;
  i: integer;
  t: integer;

begin
  result := false;
  xml := nil;
  try
    xml := TZsspXMLReaderH.Create();
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    while (not xml.Eof()) do
    begin
      xml.ReadTag();

      if ((xml.TagName = 'sheet') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        s := xml.Attributes.ItemsByName['r:id'];
        for i := 0 to RelationsCount - 1 do
          if (Relations[i].id = s) then
          begin
            Relations[i].name := ZEReplaceEntity(xml.Attributes.ItemsByName['name']);
            s := xml.Attributes.ItemsByName['sheetId'];
            relations[i].sheetid := -1;
            if (TryStrToInt(s, t)) then
              relations[i].sheetid := t;
            s := xml.Attributes.ItemsByName['state'];
            break;
          end;
      end else
      if ((xml.TagName = 'workbookView') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        s := xml.Attributes.ItemsByName['activeTab'];
        s := xml.Attributes.ItemsByName['firstSheet'];
        s := xml.Attributes.ItemsByName['showHorizontalScroll'];
        s := xml.Attributes.ItemsByName['showSheetTabs'];
        s := xml.Attributes.ItemsByName['showVerticalScroll'];
        s := xml.Attributes.ItemsByName['tabRatio'];
        s := xml.Attributes.ItemsByName['windowHeight'];
        s := xml.Attributes.ItemsByName['windowWidth'];
        s := xml.Attributes.ItemsByName['xWindow'];
        s := xml.Attributes.ItemsByName['yWindow'];
      end;
    end; //while
    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
  end;
end; //ZEXSLXReadWorkBook

//Удаляет первый символ + меняет все разделители на нужные
//INPUT
//  var FileArray: TZXLSXFileArray  - файлы
//      FilesCount: integer         - кол-во файлов
procedure ZE_XSLXReplaceDelimiter(var FileArray: TZXLSXFileArray; FilesCount: integer);
var
  i, j, k: integer;

begin
  for i := 0 to FilesCount - 1 do
  begin
    k := length(FileArray[i].name);
    if (k > 1) then
    begin
      if (FileArray[i].name[1] = '/') then
        Delete(FileArray[i].name, 1, 1);
      if (PathDelim <> '/') then
        for j := 1 to k - 1 do
          if (FileArray[i].name[j] = '/') then
            FileArray[i].name[j] := PathDelim;
    end;
  end;
end; //ZE_XSLXReplaceDelimiter

//Читает связи страниц/стилей  (*.rels: workbook.xml.rels и .rels)
//INPUT
//  var Stream: TStream                 - поток для чтения
//  var Relations: TZXLSXRelationsArray - массив с отношениями
//  var RelationsCount: integer         - кол-во
//  var isWorkSheet: boolean            - признак workbook.xml.rels
//      needReplaceDelimiter: boolean   - признак необходимости заменять разделитель
//RETURN
//      boolean - true - успешно прочитано
function ZE_XSLXReadRelationships(var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer; var isWorkSheet: boolean; needReplaceDelimiter: boolean): boolean;
var
  xml: TZsspXMLReaderH;
  s: string;
  t: integer;

begin
  result := false;
  xml := nil;
  RelationsCount := 0;
  isWorkSheet := false;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    while (not xml.Eof()) do
    begin
      xml.ReadTag();

      if ((xml.TagName = 'Relationship') and (xml.TagType = TAG_TYPE_CLOSED)) then
      begin
        SetLength(Relations, RelationsCount + 1);
        Relations[RelationsCount].id := xml.Attributes.ItemsByName['Id'];

        s := xml.Attributes.ItemsByName['Type'];
        t := ZEXLSXGetRelationNumber(s);

        if ((t >= 0) and (t < 3)) then
          isWorkSheet := true;

        Relations[RelationsCount].fileid := -1;
        Relations[RelationsCount].state := 0;
        Relations[RelationsCount].sheetid := 0;
        Relations[RelationsCount].name := '';

        Relations[RelationsCount].ftype := t;
        Relations[RelationsCount].target := xml.Attributes.ItemsByName['Target'];
        if (t >= 0) then
          inc(RelationsCount);
      end;
    end; //while
    result := true;

  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
  end;
end; //ZE_XSLXReadRelationsips

//Читает примечания (добавляет примечания на последнюю страницу)
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//  var Stream: TStream - поток для чтения
//RETURN
//      boolean - true - всё нормально
function ZEXSLXReadComments(var XMLSS: TZEXMLSS; var Stream: TStream): boolean;
var
  xml: TZsspXMLReaderH;
  s: string;
  _authorsCount: integer;
  _authors: TStringDynArray;
  _page: integer;

  procedure _ReadAuthors();
  begin
    while (not((xml.TagName = 'authors') and (xml.TagType = TAG_TYPE_END))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'author') and (xml.TagType = TAG_TYPE_END)) then
      begin
        SetLength(_authors, _authorsCount + 1);
        _authors[_authorsCount] := xml.TextBeforeTag;
        inc(_authorsCount);
      end;  
    end; //while
  end; //_ReadAuthors

  procedure _ReadComment();
  var
    _c, _r, _a: integer;
    _comment: string;
    _kol: integer;

  begin
    _c := 0;
    _r := 0;
    s := xml.Attributes.ItemsByName['ref'];
    if (s = '') then
      exit;
    if (ZEGetCellCoords(s, _c, _r, true)) then
    begin
      if (_c >= XMLSS.Sheets[_page].ColCount) then
        XMLSS.Sheets[_page].ColCount := _c + 1;
      if (_r >= XMLSS.Sheets[_page].RowCount) then
        XMLSS.Sheets[_page].RowCount := _r + 1;
        
      if (TryStrToInt(xml.Attributes.ItemsByName['authorId'], _a)) then
        if (_a >= 0) and (_a < _authorsCount) then
          XMLSS.Sheets[_page].Cell[_c, _r].CommentAuthor := _authors[_a];

      _comment := '';
      _kol := 0;
      while (not((xml.TagName = 'comment') and (xml.TagType = TAG_TYPE_END))) do
      begin
        xml.ReadTag();
        if (xml.Eof()) then
          break;
        if ((xml.TagName = 't') and (xml.TagType = TAG_TYPE_END)) then
        begin
          if (_kol > 0) then
            _comment := _comment + {$IFDEF FPC} LineEnding {$ELSE} sLineBreak {$ENDIF} + xml.TextBeforeTag
          else
            _comment := _comment + xml.TextBeforeTag;
          inc(_kol);
        end;
      end; //while
      XMLSS.Sheets[_page].Cell[_c, _r].Comment := _comment;
      XMLSS.Sheets[_page].Cell[_c, _r].ShowComment := true;
    end //if  
  end; //_ReadComment();

begin
  result := false;
  xml := nil;
  _authorsCount := 0;
  _page := XMLSS.Sheets.Count - 1;
  if (_page < 0) then
    exit;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    while (not xml.Eof()) do
    begin
      xml.ReadTag();

      if ((xml.TagName = 'authors') and (xml.TagType = TAG_TYPE_START)) then
        _ReadAuthors()
      else
      if ((xml.TagName = 'comment') and (xml.TagType = TAG_TYPE_START)) then
        _ReadComment();
        
    end; //while
    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
    SetLength(_authors, 0);
    _authors := nil;  
  end;
end; //ZEXSLXReadComments

procedure XLSXSortRelationArray(var arr: TZXLSXRelationsArray; count: integer);
var
  tmp: TZXLSXRelations;
  i, j: integer;
  _t1, _t2: integer;
  s: string;
  b: boolean;

  function _cmp(): boolean;
  begin
    b := false;
    s := arr[j].id;
    delete(s, 1, 3);
    b := TryStrToInt(s, _t1);
    if (b) then
    begin
      s := arr[j + 1].id;
      delete(s, 1, 3);
      b := TryStrToInt(s, _t2);
    end;

    if (b) then
      result := _t1 > _t2
    else
      result := arr[j].id > arr[j + 1].id;
  end;

begin
  //TODO: do not forget update sorting.
  for i := 0 to count - 2 do
    for j := 0 to count - 2 do
      if (_cmp()) then
      begin
        tmp := arr[j];
        arr[j] := arr[j + 1];
        arr[j + 1] := tmp;
      end;
end;

//Читает распакованный xlsx
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//  DirName: string     - имя папки
//RETURN
//      integer - номер ошибки (0 - всё OK)
function ReadXLSXPath(var XMLSS: TZEXMLSS; DirName: string): integer;
var
  stream: TStream;
  FileArray: TZXLSXFileArray;
  FilesCount: integer;
  StrArray: TStringDynArray;
  StrCount: integer;
  RelationsArray: array of TZXLSXRelationsArray;
  RelationsCounts: array of integer;
  SheetRelations: TZXLSXRelationsArray;
  SheetRelationsCount: integer;
  RelationsCount: integer;
  ThemaColor: TIntegerDynArray;
  ThemaColorCount: integer;
  SheetRelationNumber: integer;
  i, j, k: integer;
  s: string;
  b: boolean;
  _no_sheets: boolean;
  RH: TZEXLSXReadHelper;

  //Пытается прочитать rel для листа
  //INPUT
  //  const fname: string - имя файла листа
  //RETURN
  //      boolean - true - прочитал успешно
  function _CheckSheetRelations(const fname: string): boolean;
  var
    _rstream: TStream;
    s: string;
    i, num: integer;
    b: boolean;

  begin
    result := false;
    _rstream := nil;
    SheetRelationsCount := 0;
    num := -1;
    b := false;
    s := '';
    for i := length(fname) downto 1 do
    if (fname[i] = PathDelim) then
    begin
      num := i;
      s := fname;
      insert('_rels' + PathDelim, s, num + 1);
      s := DirName + s + '.rels';
      if (not FileExists(s)) then
        num := -1;
      break;
    end;

    if (num > 0) then
    try
      _rstream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
      result := ZE_XSLXReadRelationships(_rstream, SheetRelations, SheetRelationsCount, b, false);
    finally
      if (Assigned(_rstream)) then
        FreeAndNil(_rstream);
    end;
  end; //_CheckSheetRelations

  //Прочитать примечания
  procedure _ReadComments();
  var
    i, l: integer;
    s: string;
    b: boolean;
    _stream: TStream;

  begin
    b := false;
    s := '';
    _stream := nil;
    for i := 0 to SheetRelationsCount - 1 do
    if (SheetRelations[i].ftype = 7) then
    begin
      s := SheetRelations[i].target;
      b := true;
      break;
    end;

    //Если найдены примечания
    if (b) then
    begin
      l := length(s);
      if (l >= 3) then
        if ((s[1] = '.') and (s[2] = '.')) then
          delete(s, 1, 3);
      b := false;
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = 7) then
          if (pos(s, FileArray[i].name) <> 0) then
            if (FileExists(DirName + FileArray[i].name)) then
            begin
              s := DirName + FileArray[i].name;
              b := true;
              break;
            end;
      //Если файл не найден
      if (not b) then
      begin
        s := DirName + 'xl' + PathDelim + s;
        if (FileExists(s)) then
          b := true;
      end;

      //Файл с примечаниями таки присутствует!
      if (b) then
      try
        _stream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
        ZEXSLXReadComments(XMLSS, _stream);
      finally
        if (Assigned(_stream)) then
          FreeAndNil(_stream);
      end;
    end;
  end; //_ReadComments

begin
  result := 0;
  FilesCount := 0;
  FileArray := nil;

  if (not ZE_CheckDirExist(DirName)) then
  begin
    result := -1;
    exit;
  end;

  XMLSS.Styles.Clear();
  XMLSS.Sheets.Count := 0;
  stream := nil;
  RelationsCount := 0;
  ThemaColorCount := 0;
  SheetRelationsCount := 0;
  ThemaColor := nil;
  RH := nil;

  try
    try
      stream := TFileStream.Create(DirName + '[Content_Types].xml', fmOpenRead or fmShareDenyNone);
      if (not ZEXSLXReadContentTypes(stream,  FileArray, FilesCount)) then
      begin
        result := 3;
        exit;
      end;

      FreeAndNil(stream);

      ZE_XSLXReplaceDelimiter(FileArray, FilesCount);
      SheetRelationNumber := -1;

      b := false;
      for i := 0 to FilesCount - 1 do
      if (FileArray[i].ftype = 3) then
      begin
        b := true;
        break;
      end;

      RH := TZEXLSXReadHelper.Create();

      if (not b) then
      begin
        s := DirName + '_rels' + PathDelim + '.rels';
        if (FileExists(s)) then
        begin
          SetLength(FileArray, FilesCount + 1);
          s := '/_rels/.rels';
          FileArray[FilesCount].original := s;
          FileArray[FilesCount].name := s;
          FileArray[FilesCount].ftype := 3;
          inc(FilesCount);
        end;

        s := DirName + 'xl' + PathDelim + '_rels' + PathDelim + 'workbook.xml.rels';
        if (FileExists(s)) then
        begin
          SetLength(FileArray, FilesCount + 1);
          s := '/xl/_rels/workbook.xml.rels';
          FileArray[FilesCount].original := s;
          FileArray[FilesCount].name := s;
          FileArray[FilesCount].ftype := 3;
          inc(FilesCount);
        end;

        ZE_XSLXReplaceDelimiter(FileArray, FilesCount);
      end;

      for i := 0 to FilesCount - 1 do
      if (FileArray[i].ftype = 3) then
      begin
        SetLength(RelationsArray, RelationsCount + 1);
        SetLength(RelationsCounts, RelationsCount + 1);

        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZE_XSLXReadRelationships(stream, RelationsArray[RelationsCount], RelationsCounts[RelationsCount], b, true)) then
        begin
          result := 4;
          exit;
        end;
        if (b) then
        begin
          SheetRelationNumber := RelationsCount;
          for j := 0 to RelationsCounts[RelationsCount] - 1 do
          if (RelationsArray[RelationsCount][j].ftype = 0) then
            for k := 0 to FilesCount - 1 do
            if (RelationsArray[RelationsCount][j].fileid < 0) then
              if ((pos(RelationsArray[RelationsCount][j].target, FileArray[k].original)) > 0) then
              begin
                RelationsArray[RelationsCount][j].fileid := k;
                break;
              end;
        end; //if
        FreeAndNil(stream);
        inc(RelationsCount);
      end;

      //sharedStrings.xml
      for i:= 0 to FilesCount - 1 do
      if (FileArray[i].ftype = 4) then
      begin
        FreeAndNil(stream);
        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZEXSLXReadSharedStrings(stream, StrArray, StrCount)) then
        begin
          result := 3;
          exit;
        end;
        break;
      end;

      //тема (если есть)
      for i := 0 to FilesCount - 1 do
      if (FileArray[i].ftype = 8) then
      begin
        FreeAndNil(stream);
        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZEXSLXReadTheme(stream, ThemaColor, ThemaColorCount)) then
        begin
          result := 6;
          exit;
        end;
        break;
      end;

      //стили (styles.xml)
      for i := 0 to FilesCount - 1 do
      if (FileArray[i].ftype = 1) then
      begin
        FreeAndNil(stream);
        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZEXSLXReadStyles(XMLSS, stream, ThemaColor, ThemaColorCount, RH)) then
        begin
          result := 5;
          exit;
        end else
          b := true;
      end;

      //чтение страниц
      _no_sheets := true;
      if (SheetRelationNumber > 0) then
      begin
        for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = 2) then
        begin
          FreeAndNil(stream);
          stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
          if (not ZEXSLXReadWorkBook(XMLSS, stream, RelationsArray[SheetRelationNumber], RelationsCounts[SheetRelationNumber])) then
          begin
            result := 3;
            exit;
          end;
          break;
        end; //if

        //for i := 1 to RelationsCounts[SheetRelationNumber] do
        XLSXSortRelationArray(RelationsArray[SheetRelationNumber], RelationsCounts[SheetRelationNumber]);
        for j := 0 to RelationsCounts[SheetRelationNumber] - 1 do
          if (RelationsArray[SheetRelationNumber][j].sheetid > 0) then
          begin
            b := _CheckSheetRelations(FileArray[RelationsArray[SheetRelationNumber][j].fileid].name);
            FreeAndNil(stream);
            stream := TFileStream.Create(DirName + FileArray[RelationsArray[SheetRelationNumber][j].fileid].name, fmOpenRead or fmShareDenyNone);
            if (not ZEXSLXReadSheet(XMLSS, stream, RelationsArray[SheetRelationNumber][j].name, StrArray, StrCount, SheetRelations, SheetRelationsCount, RH)) then
              result := result or 4;
            if (b) then
              _ReadComments();
            _no_sheets := false;
          end; //if
      end;
      //если прочитано 0 листов - пробуем прочитать все (не удалось прочитать workbook/rel)
      if (_no_sheets) then
      for i := 0 to FilesCount - 1 do
      if (FileArray[i].ftype = 0) then
      begin
        b := _CheckSheetRelations(FileArray[i].name);
        FreeAndNil(stream);
        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZEXSLXReadSheet(XMLSS, stream, '', StrArray, StrCount, SheetRelations, SheetRelationsCount, RH)) then
          result := result or 4;
        if (b) then
            _ReadComments();
      end;
    except
      result := 2;
    end;
  finally
    if (Assigned(stream)) then
      FreeAndNil(stream);

    SetLength(FileArray, 0);
    FileArray := nil;
    SetLength(StrArray, 0);
    StrArray := nil;
    for i := 0 to RelationsCount - 1 do
    begin
      Setlength(RelationsArray[i], 0);
      RelationsArray[i] := nil;
    end;
    SetLength(RelationsCounts, 0);
    RelationsCounts := nil;
    SetLength(RelationsArray, 0);
    RelationsArray := nil;
    SetLength(ThemaColor, 0);
    ThemaColor := nil;
    SetLength(SheetRelations, 0);
    SheetRelations := nil;
    if (Assigned(RH)) then
      FreeAndNil(RH);
  end;
end; //ReadXLSXPath

//Читает xlsx
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//  FileName: string    - имя файла
//RETURN
//      integer - номер ошибки (0 - всё OK)
function ReadXLSX(var XMLSS: TZEXMLSS; FileName: string): integer;
var
  i, j: integer;
  u_zip: TUnZipper;
  lst: TStringList;
  ZH: TXSLXZipHelper;
  _no_sheets: boolean;

  procedure _getFile(num: integer);
  begin
    lst.Clear();
    lst.Add(ZH.FileArray[num].name); //список файлов
    u_zip.UnZipFiles(lst);
  end;

  //Пытается прочитать rel для листа
  //INPUT
  //     fnum: integer - Номер файла с листом
  //RETURN
  //      boolean - true - прочитал успешно
  function _CheckSheetRelations(fnum: integer): boolean;
  var
    i, t, l: integer;
    s, s1: string;

  begin
    ZH.ClearCurrentSheetRelations();
    result := false;
    s1 := '';
    t := -1;
    s := ZH.FileArray[fnum].original;
    l := length(s);
    for i := l downto 1 do
      if (s[i] in ['/', '\']) then
      begin
        t := i + 1;
        break;
      end;
    if (t > 0) then
      s1 := copy(s, t, l - t + 1)
    else
      s1 := s;
    if (s1 > '') then
    begin
      s1 := s1 + '.rels';
      for i := 0 to ZH.ArchFilesCount - 1 do
      if (i <> fnum) then
        if (pos(s1, ZH.ArchFile[i]) <> 0) then
        begin
          ZH.AddToFiles(ZH.ArchFile[i], 99);
          ZH.FileType := 111;
          _getFile(ZH.FilesCount - 1);
          break;
        end;
    end; //if
  end; //_CheckSheetRelations

  //Прочитать примечания
  procedure _ReadComments();
  var
    _num: integer;

  begin
    if (ZH.isNeedReadComments) then
    begin
      _num := ZH.GetCurrentPageCommentsNumber();
      if (_num >= 0) then
      begin
        ZH.FileType := 113;
        _getFile(i);
      end;
    end;
  end; //_ReadComments

begin
  result := 0;

  if (not FileExists(FileName)) then
  begin
    result := -1;
    exit;
  end;

  XMLSS.Styles.Clear();
  XMLSS.Sheets.Count := 0;

  u_zip := nil;
  lst := nil;
  ZH := nil;

  try
    try
      lst := TStringList.Create();
      lst.Clear();
      lst.Add('[Content_Types].xml'); //список файлов
      ZH := TXSLXZipHelper.Create();
      ZH.XMLSS := XMLSS;
      u_zip := TUnZipper.Create();
      u_zip.FileName := FileName;

      u_zip.Examine();
      for i := 0 to u_zip.Entries.Count - 1 do
        ZH.AddArchFile(u_zip.Entries[i].ArchiveFileName);

      u_zip.OnCreateStream := ZH.DoCreateOutZipStream;
      u_zip.OnDoneStream := ZH.DoDoneOutZipStream;
      ZH.FileType := -2;
      u_zip.UnZipFiles(lst);
      if (ZH.RetCode <> 0) then
      begin
        result := ZH.RetCode;
        exit;
      end;

      ZH.FileType := 3;
      for i := 0 to ZH.FilesCount - 1 do
      if (ZH.FileArray[i].ftype = 3) then
      begin
        ZH.FileNumber := i;
        _getFile(i);
        if (ZH.RetCode <> 0) then
        begin
          result := ZH.RetCode;
          exit;
        end;
      end;

      //sharedStrings.xml
      ZH.FileType := 4;
      for i:= 0 to ZH.FilesCount - 1 do
      if (ZH.FileArray[i].ftype = 4) then
      begin
        _getFile(i);
        if (ZH.RetCode <> 0) then
        begin
          result := ZH.RetCode;
          exit;
        end;
        break;
      end;

      //тема (если есть)
      ZH.FileType := 5;
      for i := 0 to ZH.FilesCount - 1 do
      if (ZH.FileArray[i].ftype = 8) then
      begin
        _getFile(i);
        if (ZH.RetCode <> 0) then
        begin
          result := ZH.RetCode;
          exit;
        end;
        break;
      end;

      //стили (styles.xml)
      ZH.FileType := 1;
      for i := 0 to ZH.FilesCount - 1 do
      if (ZH.FileArray[i].ftype = 1) then
      begin
        _getFile(i);
        if (ZH.RetCode <> 0) then
        begin
          result := ZH.RetCode;
          exit;
        end;
      end;

      _no_sheets := true;
      //чтение страниц
      if (ZH.SheetRelationNumber > 0) then
      begin
        ZH.FileType := 2;
        for i := 0 to ZH.FilesCount - 1 do
        if (ZH.FileArray[i].ftype = 2) then
        begin
          _getFile(i);
          if (ZH.RetCode <> 0) then
          begin
            result := ZH.RetCode;
            exit;
          end;
          break;
        end; //if

        //for i := 1 to ZH.RelationsCounts[ZH.SheetRelationNumber] do
        ZH.SortRelationArray();
        //XLSXSortRelationArray(ZH.RelationsArray[ZH.SheetRelationNumber], ZH.RelationsCounts[ZH.SheetRelationNumber]);
        for j := 0 to ZH.RelationsCounts[ZH.SheetRelationNumber] - 1 do
        if (ZH.RelationsArray[ZH.SheetRelationNumber][j].sheetid > 0) then
        begin
          ZH.ListName := ZH.RelationsArray[ZH.SheetRelationNumber][j].name;

          _CheckSheetRelations(ZH.RelationsArray[ZH.SheetRelationNumber][j].fileid);

          ZH.FileType := 0;
          _getFile(ZH.RelationsArray[ZH.SheetRelationNumber][j].fileid);
          if (ZH.RetCode <> 0) then
            result := result or ZH.RetCode;
          _ReadComments();
          _no_sheets := false;
        end; //if
      end;
      if (_no_sheets) then
      begin
        i := 0;
        while (i < ZH.FilesCount) do
        begin
          if (ZH.FileArray[i].ftype = 0) then
          begin
            _CheckSheetRelations(i);
            ZH.FileType := 0;
            _getFile(i);
            if (ZH.RetCode <> 0) then
              result := result or ZH.RetCode;
           _ReadComments();
          end;
          inc(i)
        end; //while
      end;
    except
      result := 2;
    end;
  finally
    if (Assigned(u_zip)) then
      FreeAndNil(u_zip);
    if (Assigned(lst)) then
       FreeAndNil(lst);
    if (Assigned(ZH)) then
       FreeAndNil(ZH);
  end;
end; //ReadXLSX

/////////////////////////////////////////////////////////////////////////////
/////////////                    запись                         /////////////
/////////////////////////////////////////////////////////////////////////////

//Создаёт [Content_Types].xml 
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    PageCount: integer                - кол-во страниц
//    CommentCount: integer             - кол-во страниц с комментариями
//  const PagesComments: TIntegerDynArray- номера страниц с комментариями (нумеряция с нуля)
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//  const WriteHelper: TZEXLSXWriteHelper - additional data
//RETURN
//      integer
function ZEXLSXCreateContentTypes(var XMLSS: TZEXMLSS; Stream: TStream; PageCount: integer; CommentCount: integer; const PagesComments: TIntegerDynArray;
                                  TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring;
                                  const WriteHelper: TZEXLSXWriteHelper): integer;
var
  _xml: TZsspXMLWriterH;    //писатель
  s: string;

  procedure _WriteOverride(const PartName: string; ct: integer);
  begin
    _xml.Attributes.Clear();
    _xml.Attributes.Add('PartName', PartName);
    case ct of
      0: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
      1: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
      2: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
      3: s := 'application/vnd.openxmlformats-package.relationships+xml';
      4: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
      5: s := 'application/vnd.openxmlformats-package.core-properties+xml';
      6: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
      7: s := 'application/vnd.openxmlformats-officedocument.vmlDrawing';
      8: s := 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
      9: s := 'application/vnd.openxmlformats-officedocument.drawing+xml';
    end;
    _xml.Attributes.Add('ContentType', s, false);
    _xml.WriteEmptyTag('Override', true);
  end; //_WriteOverride

  procedure _WriteTypes();
  var
    i, ii: integer;
    {$IFDEF ZUSE_DRAWINGS}
    _drawing: TZEDrawing;
    _picture: TZEPicture;
    {$ENDIF}

  begin
    _xml.Attributes.Clear();
    _xml.Attributes.Add('Extension', 'rels');
    _xml.Attributes.Add('ContentType', 'application/vnd.openxmlformats-package.relationships+xml', false);
    _xml.WriteEmptyTag('Default', true);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('Extension', 'xml');
    _xml.Attributes.Add('ContentType', 'application/xml', false);
    _xml.WriteEmptyTag('Default', true);

    {$IFDEF ZUSE_DRAWINGS}
    _xml.Attributes.Clear();
    _xml.Attributes.Add('Extension', 'png');
    _xml.Attributes.Add('ContentType', 'image/png', false);
    _xml.WriteEmptyTag('Default', true);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('Extension', 'jpeg');
    _xml.Attributes.Add('ContentType', 'image/jpeg', false);
    _xml.WriteEmptyTag('Default', true);
    {$ENDIF}

    //Страницы
    //_WriteOverride('/_rels/.rels', 3);
    //_WriteOverride('/xl/_rels/workbook.xml.rels', 3);
    for i := 0 to PageCount - 1 do
    begin
      _WriteOverride('/xl/worksheets/sheet' + IntToStr(i + 1) + '.xml', 0);
      if (WriteHelper.IsSheetHaveHyperlinks(i)) then
        _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(i + 1) + '.xml.rels', 3);
    end;
    //комментарии
    for i := 0 to CommentCount - 1 do
    begin
      _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(PagesComments[i] + 1) + '.xml.rels', 3);
      _WriteOverride('/xl/comments' + IntToStr(PagesComments[i] + 1) + '.xml', 6);
    end;
    {$IFDEF ZUSE_DRAWINGS}
    // картинки
    for i := 0 to XMLSS.DrawingCount - 1 do
    begin
      _drawing := XMLSS.GetDrawing(i);
      if Assigned(_drawing) and (_drawing.PictureStore.Count > 0) then
      begin
        _WriteOverride('/xl/drawings/drawing' + IntToStr(_drawing.Id) + '.xml', 9);
        _WriteOverride('/xl/drawings/_rels/drawing' + IntToStr(_drawing.Id) + '.xml.rels', 3);
        for ii := 0 to _drawing.PictureStore.Count - 1 do
        begin
          _picture := _drawing.PictureStore.Items[ii];
          // image/ override
          _xml.Attributes.Clear();
          _xml.Attributes.Add('PartName', '/xl/media/' + _picture.Name);
          _xml.Attributes.Add('ContentType', 'image/' + Copy(ExtractFileExt(_picture.Name), 2, 99), false);
          _xml.WriteEmptyTag('Override', true);
        end;
      end;
    end;
    {$ENDIF}
    _WriteOverride('/xl/workbook.xml', 2);
    _WriteOverride('/xl/styles.xml', 1);
    _WriteOverride('/xl/sharedStrings.xml', 4);
    _WriteOverride('/docProps/app.xml', 8);
    _WriteOverride('/docProps/core.xml', 5);
  end; //_WriteTypes

begin
  result := 0;
  _xml := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', SCHEMA_PACKAGE + '/content-types');
    _xml.WriteTagNode('Types', true, true, true);
    _WriteTypes();
    _xml.WriteEndTagNode(); //Types
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ZEXLSXCreateContentTypes


{$IFDEF ZUSE_DRAWINGS}
//Создаёт рисунок (drawing*.xml)
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//    Drawing: TZEDrawing               - набор рисунков
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateDrawing(var XMLSS: TZEXMLSS; Stream: TStream; Drawing: TZEDrawing; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;
  SheetNum: integer;
  pic: TZEPicture;

  function PtToStr(Value: Real): string;
  begin
    Result := IntToStr(Trunc(Value * 1000));
  end;

  // Адрес ячейки определяется по координатам
  procedure _WriteCellForXY(PageNum: Integer; X, Y: Real);
  var
    sh: TZSheet;
    n: Integer;
    Total: Real;
  begin
    sh := XMLSS.Sheets.Sheet[PageNum];
    if not Assigned(sh) then Exit;
    _xml.Attributes.Clear();

    // column
    n := 0;
    Total := 0;
    while (n < sh.ColCount) and (Total < X) do
    begin
      Total := Total + sh.ColWidths[n];
      Inc(n);
    end;
    _xml.WriteTag('xdr:col', IntToStr(n), false, false);
    _xml.WriteTag('xdr:colOff', PtToStr(Total), false, false);

    // row
    n := 0;
    Total := 0;
    while (n < sh.RowCount) and (Total < Y) do
    begin
      Total := Total + sh.RowHeights[n];
      Inc(n);
    end;
    _xml.WriteTag('xdr:row', IntToStr(n), false, false);
    _xml.WriteTag('xdr:rowOff', PtToStr(Total), false, false);
  end;

  // Если размер картинки указан, то адрес ячейки сдвигается на заданный размер
  procedure _WriteCellForRowCol(PageNum, Row, Col: Integer; Width, Height: Real);
  var
    sh: TZSheet;
    n: Integer;
    Total, X, Y: Real;
  begin
    sh := XMLSS.Sheets.Sheet[PageNum];
    if not Assigned(sh) then Exit;
    _xml.Attributes.Clear();

    // column
    n := 0;
    Total := 0;
    while (n < sh.ColCount) and (n < Col) do
    begin
      Total := Total + sh.ColWidths[n];
      Inc(n);
    end;
    // plus width
    X := Total + Width - 0.1;
    while (n < sh.ColCount) and (Total < X) do
    begin
      Total := Total + sh.ColWidths[n];
      Inc(n);
    end;
    _xml.WriteTag('xdr:col', IntToStr(n), false, false);
    _xml.WriteTag('xdr:colOff', PtToStr(Total), false, false);

    // row
    n := 0;
    Total := 0;
    while (n < sh.RowCount) and (n < Row) do
    begin
      Total := Total + sh.RowHeights[n];
      Inc(n);
    end;
    // plus height
    Y := Total + Height - 0.1;
    while (n < sh.ColCount) and (Total < Y) do
    begin
      Total := Total + sh.RowHeights[n];
      Inc(n);
    end;
    _xml.WriteTag('xdr:row', IntToStr(n), false, false);
    _xml.WriteTag('xdr:rowOff', PtToStr(Total), false, false);
  end;

var
  i: Integer;
begin
  result := 0;
  _xml := nil;

  if not Assigned(Drawing) then Exit;
  SheetNum := XMLSS.GetDrawingSheetNum(Drawing);

  _xml := TZsspXMLWriterH.Create();
  try
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    //_xml.NewLine := false;
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
    _xml.Attributes.Add('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
    _xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL, false);
    _xml.WriteTagNode('xdr:wsDr', false, false, false);

    for i := 0 to Drawing.PictureStore.Count - 1 do
    begin
      pic := Drawing.PictureStore.Items[i];
      // cell anchor
      _xml.Attributes.Clear();
      if pic.CellAnchor = ZAAbsolute then
        _xml.Attributes.Add('editAs', 'absolute')
      else
        _xml.Attributes.Add('editAs', 'oneCell');
      _xml.WriteTagNode('xdr:twoCellAnchor', false, false, false);

      // - xdr:from
      _xml.Attributes.Clear();
      _xml.WriteTagNode('xdr:from', false, false, false);
      if (pic.CellAnchor = ZACell) then
        _WriteCellForRowCol(SheetNum, pic.Row, pic.Col, 0, 0)
      else
        _WriteCellForXY(SheetNum, pic.X, pic.Y);
      _xml.WriteEndTagNode(); // xdr:from
      // - xdr:to
      _xml.Attributes.Clear();
      _xml.WriteTagNode('xdr:to', false, false, false);
      if (pic.CellAnchor = ZACell) then
        _WriteCellForRowCol(SheetNum, pic.Row, pic.Col, pic.Width, pic.Height)
      else
        _WriteCellForXY(SheetNum, (pic.X + pic.Width), (pic.Y + pic.Height));
      _xml.Attributes.Clear();
      _xml.WriteEndTagNode(); // xdr:from
      // - xdr:pic
      _xml.WriteTagNode('xdr:pic', false, false, false);
      // -- xdr:nvPicPr
      _xml.WriteTagNode('xdr:nvPicPr', false, false, false);
      // --- xdr:cNvPr
      _xml.Attributes.Clear();
      _xml.Attributes.Add('descr', pic.Description);
      _xml.Attributes.Add('name', pic.Name);
      _xml.Attributes.Add('id', IntToStr(pic.RelId));  // 1
      _xml.WriteEmptyTag('xdr:cNvPr', false);
      // --- xdr:cNvPicPr
      _xml.Attributes.Clear();
      _xml.WriteEmptyTag('xdr:cNvPicPr', false);
      _xml.WriteEndTagNode(); // -- xdr:nvPicPr

      // -- xdr:blipFill
      _xml.Attributes.Clear();
      _xml.WriteTagNode('xdr:blipFill', false, false, false);
      // --- a:blip
      _xml.Attributes.Clear();
      _xml.Attributes.Add('r:embed', pic.RelIdStr); // rId1
      _xml.WriteEmptyTag('a:blip', false);
      // --- a:stretch
      _xml.Attributes.Clear();
      _xml.WriteEmptyTag('a:stretch', false);
      _xml.WriteEndTagNode(); // -- xdr:blipFill

      // -- xdr:spPr
      _xml.Attributes.Clear();
      _xml.WriteTagNode('xdr:spPr', false, false, false);
      // --- a:xfrm
      _xml.WriteTagNode('a:xfrm', false, false, false);
      // ----
      _xml.Attributes.Clear();
      _xml.Attributes.Add('x', PtToStr(pic.X));
      _xml.Attributes.Add('y', PtToStr(pic.Y));
      _xml.WriteEmptyTag('a:off', false);
      // ----
      _xml.Attributes.Clear();
      _xml.Attributes.Add('cx', PtToStr(pic.X + pic.Width));
      _xml.Attributes.Add('cy', PtToStr(pic.Y + pic.Height));
      _xml.WriteEmptyTag('a:ext', false);
      _xml.Attributes.Clear();
      _xml.WriteEndTagNode(); // --- a:xfrm

      // --- a:prstGeom
      _xml.Attributes.Clear();
      _xml.Attributes.Add('prst', 'rect');
      _xml.WriteTagNode('a:prstGeom', false, false, false);
      _xml.Attributes.Clear();
      _xml.WriteEmptyTag('a:avLst', false);
      _xml.WriteEndTagNode(); // --- a:prstGeom

      // --- a:ln
      _xml.WriteTagNode('a:ln', false, false, false);
      _xml.WriteEmptyTag('a:noFill', false);
      _xml.WriteEndTagNode(); // --- a:ln

      _xml.WriteEndTagNode(); // -- xdr:spPr

      _xml.WriteEndTagNode(); // - xdr:pic

      _xml.WriteEmptyTag('xdr:clientData', false);

      _xml.WriteEndTagNode(); // xdr:twoCellAnchor
    end;
    _xml.WriteEndTagNode(); // xdr:wsDr
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end;

function ZEXLSXCreateDrawingRels(var XMLSS: TZEXMLSS; Stream: TStream; Drawing: TZEDrawing; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;
  pic: TZEPicture;
  i: integer;
begin
  result := 0;
  _xml := nil;

  if not Assigned(Drawing) then Exit;

  _xml := TZsspXMLWriterH.Create();
  try
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL, false);
    _xml.WriteTagNode('Relationships', false, false, false);

    for i := 0 to Drawing.PictureStore.Count - 1 do
    begin
      pic := Drawing.PictureStore[i];
      _xml.Attributes.Clear();
      _xml.Attributes.Add('Id', pic.RelIdStr);
      _xml.Attributes.Add('Type', SCHEMA_DOC_REL + '/image');
      _xml.Attributes.Add('Target', '../media/' + pic.Name);
      _xml.WriteEmptyTag('Relationship', false, true);
    end;
    _xml.WriteEndTagNode(); // Relationships
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end;
{$ENDIF} // ZUSE_DRAWINGS

//Создаёт лист документа (sheet*.xml)
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//    SheetNum: integer                 - номер листа в документе
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//  var isHaveComments: boolean         - возвращает true, если были комментарии (чтобы создать comments*.xml)
//  const WriteHelper: TZEXLSXWriteHelper - additional data
//RETURN
//      integer
function ZEXLSXCreateSheet(var XMLSS: TZEXMLSS; Stream: TStream; SheetNum: integer; TextConverter: TAnsiToCPConverter;
                                     CodePageName: String; BOM: ansistring; const WriteHelper: TZEXLSXWriteHelper): integer;
var
  _xml: TZsspXMLWriterH;    //писатель
  _sheet: TZSheet;

  {
  procedure _AddSelection(const _Cell, _Pane: string);
  begin
    _xml.Attributes.Clear();
    _xml.Attributes.Add('activeCell', _Cell);
    _xml.Attributes.Add('activeCellId', '0', false);
    _xml.Attributes.Add('pane', _Pane, false);
    _xml.Attributes.Add('sqref', _Cell, false);
    _xml.WriteEmptyTag('selection', true, false);
  end;
  }

  procedure WriteXLSXSheetHeader();
  var
    s: string;
    b: boolean;
    _SOptions: TZSheetOptions;

    procedure _AddSplitValue(const SplitMode: TZSplitMode; const SplitValue: integer; const AttrName: string);
    var
      s: string;
      b: boolean;

    begin
      s := '0';
      b := true;
      case SplitMode of
        ZSplitFrozen:
          begin
            s := IntToStr(SplitValue);
            if (SplitValue = 0) then
              b := false;
          end;
        ZSplitSplit: s := IntToStr(round(PixelToPoint(SplitValue) * 20));
        ZSplitNone: b := false;
      end;
      if (b) then
        _xml.Attributes.Add(AttrName, s);
    end; //_AddSplitValue

    procedure _AddTopLeftCell(const VMode: TZSplitMode; const VValue: integer; const HMode: TZSplitMode; const HValue: integer);
    var
      _isProblem: boolean;

    begin
      _isProblem := (VMode = ZSplitSplit) or (HMode = ZSplitSplit);
      _isProblem := _isProblem or (VValue > 1000) or (HValue > 100);

      if (not _isProblem) then
      begin
        s := ZEGetA1byCol(VValue) + IntToSTr(HValue + 1);
        _xml.Attributes.Add('topLeftCell', s);
      end;
    end; //_AddTopLeftCell

  begin
    _xml.Attributes.Clear();
    _xml.Attributes.Add('filterMode', 'false');
    _xml.WriteTagNode('sheetPr', true, true, false);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('fitToPage', 'false');
    _xml.WriteEmptyTag('pageSetUpPr', true, false);

    _xml.WriteEndTagNode(); //sheetPr

    _xml.Attributes.Clear();
    s := 'A1';
    if (_sheet.ColCount > 0) then
      s := s + ':' + ZEGetA1byCol(_sheet.ColCount - 1) + IntToStr(_sheet.RowCount);
    _xml.Attributes.Add('ref', s);
    _xml.WriteEmptyTag('dimension', true, false);

    _xml.Attributes.Clear();
    _xml.WriteTagNode('sheetViews', true, true, true);

    _xml.Attributes.Add('colorId', '64');
    _xml.Attributes.Add('defaultGridColor', 'true', false);
    _xml.Attributes.Add('rightToLeft', 'false', false);
    _xml.Attributes.Add('showFormulas', 'false', false);
    _xml.Attributes.Add('showGridLines', 'true', false);
    _xml.Attributes.Add('showOutlineSymbols', 'true', false);
    _xml.Attributes.Add('showRowColHeaders', 'true', false);
    _xml.Attributes.Add('showZeros', 'true', false);

    if (_sheet.Selected) then
      _xml.Attributes.Add('tabSelected', 'true', false);

    _xml.Attributes.Add('topLeftCell', 'A1', false);
    _xml.Attributes.Add('view', 'normal', false);
    _xml.Attributes.Add('windowProtection', 'false', false);
    _xml.Attributes.Add('workbookViewId', '0', false);
    _xml.Attributes.Add('zoomScale', '100', false);
    _xml.Attributes.Add('zoomScaleNormal', '100', false);
    _xml.Attributes.Add('zoomScalePageLayoutView', '100', false);
    _xml.WriteTagNode('sheetView', true, true, false);

    _SOptions := _sheet.SheetOptions;

    b := (_SOptions.SplitVerticalMode <> ZSplitNone) or
         (_SOptions.SplitHorizontalMode <> ZSplitNone);
    if (b) then
    begin
      _xml.Attributes.Clear();
      _AddSplitValue(_SOptions.SplitVerticalMode,
                     _SOptions.SplitVerticalValue,
                     'xSplit');
      _AddSplitValue(_SOptions.SplitHorizontalMode,
                     _SOptions.SplitHorizontalValue,
                     'ySplit');

      _AddTopLeftCell(_SOptions.SplitVerticalMode, _SOptions.SplitVerticalValue,
                      _SOptions.SplitHorizontalMode, _SOptions.SplitHorizontalValue);

      _xml.Attributes.Add('activePane', 'topLeft');

      s := 'split';
      if ((_SOptions.SplitVerticalMode = ZSplitFrozen) or (_SOptions.SplitHorizontalMode = ZSplitFrozen)) then
        s := 'frozen';
      _xml.Attributes.Add('state', s);

      _xml.WriteEmptyTag('pane', true, false);
    end; //if
    {
    <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
    activePane (Active Pane) The pane that is active.
                The possible values for this attribute are
                defined by the ST_Pane simple type (§18.18.52).
                  bottomRight	Bottom Right Pane
                  topRight	Top Right Pane
                  bottomLeft	Bottom Left Pane
                  topLeft	Top Left Pane

    state (Split State) Indicates whether the pane has horizontal / vertical
                splits, and whether those splits are frozen.
                The possible values for this attribute are defined by the ST_PaneState simple type (§18.18.53).
                   Split
                   Frozen
                   FrozenSplit

    topLeftCell (Top Left Visible Cell) Location of the top left visible
                cell in the bottom right pane (when in Left-To-Right mode).
                The possible values for this attribute are defined by the
                ST_CellRef simple type (§18.18.7).

    xSplit (Horizontal Split Position) Horizontal position of the split,
                in 1/20th of a point; 0 (zero) if none. If the pane is frozen,
                this value indicates the number of columns visible in the
                top pane. The possible values for this attribute are defined
                by the W3C XML Schema double datatype.

    ySplit (Vertical Split Position) Vertical position of the split, in 1/20th
                of a point; 0 (zero) if none. If the pane is frozen, this
                value indicates the number of rows visible in the left pane.
                The possible values for this attribute are defined by the
                W3C XML Schema double datatype.
    }

    {
    _xml.Attributes.Clear();
    _xml.Attributes.Add('activePane', 'topLeft');
    _xml.Attributes.Add('topLeftCell', 'A1', false);
    _xml.Attributes.Add('xSplit', '0', false);
    _xml.Attributes.Add('ySplit', '-1', false);
    _xml.WriteEmptyTag('pane', true, false);
    }

    {
    _AddSelection('A1', 'bottomLeft');
    _AddSelection('F16', 'topLeft');
    }
    
    s := ZEGetA1byCol(_sheet.SheetOptions.ActiveCol) + IntToSTr(_sheet.SheetOptions.ActiveRow + 1);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('activeCell', s);
    _xml.Attributes.Add('sqref', s);
    _xml.WriteEmptyTag('selection', true, false);

    _xml.WriteEndTagNode(); //sheetView
    _xml.WriteEndTagNode(); //sheetViews
  end; //WriteXLSXSheetHeader

  procedure WriteXLSXSheetCols();
  var
    i: integer;
    s: string;
    ProcessedColumn: TZColOptions;

  begin
    _xml.Attributes.Clear();
    _xml.WriteTagNode('cols', true, true, true);
    for i := 0 to _sheet.ColCount - 1 do
    begin
      _xml.Attributes.Clear();
      _xml.Attributes.Add('collapsed', 'false', false);
      _xml.Attributes.Add('hidden', XLSXBoolToStr(_sheet.Columns[i].Hidden), false);
      s := IntToStr(i + 1);
      //Как там эти max / min считаются?
      if (i = _sheet.ColCount - 1) then
      begin
        if (i + 1 >= 1025) then
          s := '1025';
      end;
      _xml.Attributes.Add('max', s, false); //??
      _xml.Attributes.Add('min', IntToStr(i + 1), false); //??
      s := '0';
      ProcessedColumn := _sheet.Columns[i];
      if ((ProcessedColumn.StyleID >= -1) and (ProcessedColumn.StyleID < XMLSS.Styles.Count)) then
        s := IntToStr(ProcessedColumn.StyleID + 1);
      _xml.Attributes.Add('style', s, false);
      _xml.Attributes.Add('width', ZEFloatSeparator(FormatFloat('0.##########', ProcessedColumn.WidthMM * 5.14509803921569 / 10)), false);
      if ProcessedColumn.AutoFitWidth then
        _xml.Attributes.Add('bestFit', '1', false);
      _xml.WriteEmptyTag('col', true, false);
    end;
    _xml.WriteEndTagNode(); //cols
  end; //WriteXLSXSheetCols

  procedure WriteXLSXSheetData();
  var
    i, j, n: integer;
    b: boolean;
    s: string;
    _r: TRect;
    k1, k2, k: integer;
    _in_merge_not_top: boolean; //if cell in merge area, but not top left
    
  begin
    _xml.Attributes.Clear();
    _xml.WriteTagNode('sheetData', true, true, true);
    n := _sheet.ColCount - 1;
    for i := 0 to _sheet.RowCount - 1 do
    begin
      _xml.Attributes.Clear();
      _xml.Attributes.Add('collapsed', 'false', false); //?
      _xml.Attributes.Add('customFormat', 'false', false); //?
      _xml.Attributes.Add('customHeight', XLSXBoolToStr((abs(_sheet.DefaultRowHeight - _sheet.Rows[i].Height) > 0.001)){'true'}, false); //?
      _xml.Attributes.Add('hidden', XLSXBoolToStr(_sheet.Rows[i].Hidden), false);
      _xml.Attributes.Add('ht', ZEFloatSeparator(FormatFloat('0.##', _sheet.Rows[i].HeightMM * 2.835)), false);
      _xml.Attributes.Add('outlineLevel', '0', false);
      _xml.Attributes.Add('r', IntToStr(i + 1), false);
      _xml.WriteTagNode('row', true, true, false);
      for j := 0 to n do
      begin
        _xml.Attributes.Clear();
        if (not WriteHelper.isHaveComments) then
          if (_sheet.Cell[j, i].Comment > '') then
            WriteHelper.isHaveComments := true;
        b := (_sheet.Cell[j, i].Data > '') or
             (_sheet.Cell[j, i].Formula > '');
        s := ZEGetA1byCol(j) + IntToStr(i + 1);

        if (_sheet.Cell[j, i].HRef <> '') then
          WriteHelper.AddHyperLink(s, _sheet.Cell[j, i].HRef, _sheet.Cell[j, i].HRefScreenTip, 'External');

        _xml.Attributes.Add('r', s);
        _in_merge_not_top := false;
        k := _sheet.MergeCells.InMergeRange(j, i);
        if (k >= 0) then
          _in_merge_not_top := _sheet.MergeCells.InLeftTopCorner(j, i) < 0;

        if (_in_merge_not_top) then
        begin
          k1 := _sheet.MergeCells.Items[k].Left;
          k2 := _sheet.MergeCells.Items[k].top;
        end
        else
        begin
          k1 := j;
          k2 := i;
        end;

        if (_sheet.Cell[k1, k2].CellStyle >= -1) and (_sheet.Cell[k1, k2].CellStyle < XMLSS.Styles.Count) then
          s := IntToStr(_sheet.Cell[k1, k2].CellStyle + 1)
        else
          s := '0';
        _xml.Attributes.Add('s', s, false);

        case _sheet.Cell[j, i].CellType of
          ZENumber: s := 'n';
          ZEDateTime: s := 'd'; //??
          ZEBoolean: s := 'b';
          ZEString: s := 'str';
          ZEError: s := 'e';
        end;
        
        // если тип ячейки ZEGeneral, то атрибут опускаем
        if _sheet.Cell[j, i].CellType <> ZEGeneral then
          _xml.Attributes.Add('t', s, false);

        if (b) then
        begin
          _xml.WriteTagNode('c', true, true, false);
          if (_sheet.Cell[j, i].Formula > '') then
          begin
            _xml.Attributes.Clear();
            _xml.Attributes.Add('aca', 'false');
            _xml.WriteTag('f', _sheet.Cell[j, i].Formula, true, false, true);
          end;
          if (_sheet.Cell[j, i].Data > '') then
          begin
            _xml.Attributes.Clear();
            _xml.WriteTag('v', _sheet.Cell[j, i].Data, true, false, true);
          end;
          _xml.WriteEndTagNode();
        end else
          _xml.WriteEmptyTag('c', true);
      end;
      _xml.WriteEndTagNode(); //row
    end; //for i

    _xml.WriteEndTagNode(); //sheetData

    // autoFilter
    if trim(_sheet.AutoFilter)<>'' then
    begin
      _xml.Attributes.Clear;
      _xml.Attributes.Add('ref',_sheet.AutoFilter);
      _xml.WriteEmptyTag('autoFilter', true, false);
    end;

    //Merge Cells
    if (_sheet.MergeCells.Count > 0) then
    begin
      _xml.Attributes.Clear();
      _xml.Attributes.Add('count', IntToStr(_sheet.MergeCells.Count));
      _xml.WriteTagNode('mergeCells', true, true, false);
      for i := 0 to _sheet.MergeCells.Count - 1 do
      begin
        _xml.Attributes.Clear();
        _r := _sheet.MergeCells.Items[i];
        s := ZEGetA1byCol(_r.Left) + IntToStr(_r.Top + 1) + ':' + ZEGetA1byCol(_r.Right) + IntToStr(_r.Bottom + 1);
        _xml.Attributes.Add('ref', s);
        _xml.WriteEmptyTag('mergeCell', true, false);
      end;
      _xml.WriteEndTagNode(); //mergeCells
    end; //if

    WriteHelper.WriteHyperLinksTag(_xml);
  end; //WriteXLSXSheetData

  procedure WriteXLSXSheetFooter();
  var
    s: string;

  begin
    _xml.Attributes.Clear();
    _xml.Attributes.Add('headings', 'false', false);
    _xml.Attributes.Add('gridLines', 'false', false);
    _xml.Attributes.Add('gridLinesSet', 'true', false);
    _xml.Attributes.Add('horizontalCentered', XLSXBoolToStr(_sheet.SheetOptions.CenterHorizontal), false);
    _xml.Attributes.Add('verticalCentered', XLSXBoolToStr(_sheet.SheetOptions.CenterVertical), false);
    _xml.WriteEmptyTag('printOptions', true, false);

    _xml.Attributes.Clear();
    s := '0.##########';
    _xml.Attributes.Add('left',   ZEFloatSeparator(FormatFloat(s, _sheet.SheetOptions.MarginLeft / ZE_MMinInch)),   false);
    _xml.Attributes.Add('right',  ZEFloatSeparator(FormatFloat(s, _sheet.SheetOptions.MarginRight / ZE_MMinInch)),  false);
    _xml.Attributes.Add('top',    ZEFloatSeparator(FormatFloat(s, _sheet.SheetOptions.MarginTop / ZE_MMinInch)),    false);
    _xml.Attributes.Add('bottom', ZEFloatSeparator(FormatFloat(s, _sheet.SheetOptions.MarginBottom / ZE_MMinInch)), false);
    _xml.Attributes.Add('header', ZEFloatSeparator(FormatFloat(s, _sheet.SheetOptions.HeaderMargins.Height / ZE_MMinInch)), false);
    _xml.Attributes.Add('footer', ZEFloatSeparator(FormatFloat(s, _sheet.SheetOptions.FooterMargins.Height / ZE_MMinInch)), false);
    _xml.WriteEmptyTag('pageMargins', true, false);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('blackAndWhite', 'false', false);
    _xml.Attributes.Add('cellComments', 'none', false);
    _xml.Attributes.Add('copies', '1', false);
    _xml.Attributes.Add('draft', 'false', false);
    _xml.Attributes.Add('firstPageNumber', '1', false);
    _xml.Attributes.Add('fitToHeight', '1', false);
    _xml.Attributes.Add('fitToWidth', '1', false);
    _xml.Attributes.Add('horizontalDpi', '300', false);
//    if (_sheet.SheetOptions.PortraitOrientation) then
//      s := 'portrait'
//    else
//      s := 'album';
//    _xml.Attributes.Add('orientation', s, false);

    // ECMA 376 ed.4 part1 18.18.50: default|portrait|landscape
    _xml.Attributes.Add('orientation',
        IfThen(_sheet.SheetOptions.PortraitOrientation, 'portrait', 'landscape'),
        false);

    _xml.Attributes.Add('pageOrder', 'downThenOver', false);
    _xml.Attributes.Add('paperSize', intToStr(_sheet.SheetOptions.PaperSize), false);
    _xml.Attributes.Add('scale', '100', false);
    _xml.Attributes.Add('useFirstPageNumber', 'true', false);
    //_xml.Attributes.Add('usePrinterDefaults', 'false', false); //do not use!
    _xml.Attributes.Add('verticalDpi', '300', false);
    _xml.WriteEmptyTag('pageSetup', true, false);
    //  <headerFooter differentFirst="false" differentOddEven="false">
    //    <oddHeader> ... </oddHeader>
    //    <oddFooter> ... </oddFooter>
    //  </headerFooter>
    //  <legacyDrawing r:id="..."/>
  end; //WriteXLSXSheetFooter

  {$IFDEF ZUSE_DRAWINGS}
  procedure WriteXLSXSheetDrawings();
  var
    rId: Integer;
  begin
    // drawings
    if (not _sheet.Drawing.IsEmpty) then
    begin
      // rels to helper
      rId := WriteHelper.AddDrawing('../drawings/drawing' + IntToStr(_sheet.Drawing.Id) + '.xml');

      _xml.Attributes.Clear();
      _xml.Attributes.Add('r:id', 'rId' + IntToStr(rId));
      _xml.WriteEmptyTag('drawing');
    end;
  end;
  {$ENDIF} // ZUSE_DRAWINGS

begin
  WriteHelper.Clear();
  result := 0;
  _xml := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    _xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL);
    _xml.WriteTagNode('worksheet', true, true, false);

    _sheet := XMLSS.Sheets[SheetNum];
    WriteXLSXSheetHeader();
    if (_sheet.ColCount > 0) then
      WriteXLSXSheetCols();
    WriteXLSXSheetData();
    WriteXLSXSheetFooter();

    {$IFDEF ZUSE_DRAWINGS}
    WriteXLSXSheetDrawings();
    {$ENDIF} // ZUSE_DRAWINGS

    _xml.WriteEndTagNode(); //worksheet
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ZEXLSXCreateSheet

//Создаёт workbook.xml
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//  const _pages: TIntegerDynArray       - массив страниц
//  const _names: TStringDynArray       - массив имён страниц
//    PageCount: integer                - кол-во страниц
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateWorkBook(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                              const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;    //писатель

  //<sheets> ... </sheets>
  procedure WriteSheets();
  var
    i: integer;

  begin
    _xml.Attributes.clear();
    _xml.WriteTagNode('sheets', true, true, true);

    for i := 0 to PageCount - 1 do
    begin
      _xml.Attributes.Clear();
      _xml.Attributes.Add('name', _names[i], false);
      _xml.Attributes.Add('sheetId', IntToStr(i + 1), false);
      _xml.Attributes.Add('state', 'visible', false);
      _xml.Attributes.Add('r:id', 'rId' + IntToStr(i + 2), false);
      _xml.WriteEmptyTag('sheet', true);
    end; //for i

    _xml.WriteEndTagNode(); //sheets
  end; //WriteSheets

begin
  result := 0;
  _xml := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    _xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL, false);
    _xml.WriteTagNode('workbook', true, true, true);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('appName', 'ZEXMLSSlib');
    _xml.WriteEmptyTag('fileVersion', true);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('backupFile', 'false');
    _xml.Attributes.Add('showObjects', 'all', false);
    _xml.Attributes.Add('date1904', 'false', false);
    _xml.WriteEmptyTag('workbookPr', true);

    _xml.Attributes.Clear();
    _xml.WriteEmptyTag('workbookProtection', true);

    _xml.WriteTagNode('bookViews', true, true, true);

    _xml.Attributes.Add('activeTab', '0');
    _xml.Attributes.Add('firstSheet', '0', false);
    _xml.Attributes.Add('showHorizontalScroll', 'true', false);
    _xml.Attributes.Add('showSheetTabs', 'true', false);
    _xml.Attributes.Add('showVerticalScroll', 'true', false);
    _xml.Attributes.Add('tabRatio', '600', false);
    _xml.Attributes.Add('windowHeight', '8192', false);
    _xml.Attributes.Add('windowWidth', '16384', false);
    _xml.Attributes.Add('xWindow', '0', false);
    _xml.Attributes.Add('yWindow', '0', false);
    _xml.WriteEmptyTag('workbookView', true);
    _xml.WriteEndTagNode(); // bookViews

    WriteSheets();

    _xml.Attributes.Clear();
    _xml.Attributes.Add('iterateCount', '100');
    _xml.Attributes.Add('refMode', 'A1', false); //{tut}
    _xml.Attributes.Add('iterate', 'false', false);
    _xml.Attributes.Add('iterateDelta', '0.001', false);
    _xml.WriteEmptyTag('calcPr', true);

    _xml.WriteEndTagNode(); //workbook
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ZEXLSXCreateWorkBook

//Создаёт styles.xml 
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateStyles(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;        //писатель
  _FontIndex: TIntegerDynArray;  //соответствия шрифтов
  _FillIndex: TIntegerDynArray;  //заливки
  _BorderIndex: TIntegerDynArray;//границы
  _StylesCount: integer;
  _NumFmtIndexes: array of integer;
  _FmtParser: TNumFormatParser;
  _DateParser: TZDateTimeODSFormatParser;

  // <numFmts> .. </numFmts>
  procedure WritenumFmts();
  var
    kol: integer;
    i: integer;
    _nfmt: TZEXLSXNumberFormats;
    _is_dateTime: array of boolean;
    s: string;
    _count: integer;
    _idx: array of integer;
    _fmt: array of string;
    _style: TZStyle;
    _currSheet: integer;
    _currRow, _currCol: integer;
    _sheet: TZSheet;
    _currstylenum: integer;
    _numfmt_counter: integer;

    function _GetNumFmt(StyleNum: integer): integer;
    var
      i, j, k: integer;
      b: boolean;
      _cs, _cr, _cc: integer;

    begin
      Result := 0;
      _style := XMLSS.Styles[StyleNum];

      //If cell type is datetime and cell style is empty then need write default NumFmtId = 14.
      if ((Trim(_style.NumberFormat) = '') or (UpperCase(_style.NumberFormat) = 'GENERAL')) then
      begin
        if (_is_dateTime[StyleNum + 1]) then
          Result := 14
        else
        begin
          b := false;
          _cs := _currSheet;
          for i := _cs to XMLSS.Sheets.Count - 1 do
          begin
            _sheet := XMLSS.Sheets[i];
            _cr := _currRow;
            for j := _cr to _sheet.RowCount - 1 do
            begin
              _cc := _currCol;
              for k := _cc to _sheet.ColCount - 1 do
              begin
                _currstylenum := _sheet[k, j].CellStyle + 1;
                if (_currstylenum >= 0) and (_currstylenum < kol) then
                  if (_sheet[k, j].CellType = ZEDateTime) then
                  begin
                    _is_dateTime[_currstylenum] := true;
                    if (_currstylenum = StyleNum + 1) then
                    begin
                      b := true;
                      break;
                    end;
                  end;
              end; //for k
              _currRow := j + 1;
              _currCol := 0;
              if (b) then
                break;
            end; //for j

            _currSheet := i + 1;
            _currRow := 0;
            _currCol := 0;
            if (b) then
              break;
          end; //for i

          if (b) then
            Result := 14;
        end;
      end //if
      else
      begin
        s := ConvertFormatNativeToXlsx(_style.NumberFormat, _FmtParser, _DateParser);
        i := _nfmt.FindFormatID(s);
        if (i < 0) then
        begin
          i := _numfmt_counter;
          _nfmt.Format[i] := s;
          inc(_numfmt_counter);

          SetLength(_idx, _count + 1);
          SetLength(_fmt, _count + 1);
          _idx[_count] := i;
          _fmt[_count] := s;

          inc(_count);
        end;
        Result := i;
      end;
    end; //_GetNumFmt

  begin
    kol := XMLSS.Styles.Count + 1;
    SetLength(_NumFmtIndexes, kol);
    SetLength(_is_dateTime, kol);
    for i := 0 to kol - 1 do
      _is_dateTime[i] := false;

    _nfmt := nil;
    _count := 0;

    _numfmt_counter := 164;

    _currSheet := 0;
    _currRow := 0;
    _currCol := 0;

    try
      _nfmt := TZEXLSXNumberFormats.Create();
      for i := -1 to kol - 2 do
        _NumFmtIndexes[i + 1] := _GetNumFmt(i);

      if (_count > 0) then
      begin
        _xml.Attributes.Clear();
        _xml.Attributes.Add('count', IntToStr(_count));
        _xml.WriteTagNode('numFmts', true, true, false);

        for i := 0 to _count - 1 do
        begin
          _xml.Attributes.Clear();
          _xml.Attributes.Add('numFmtId', IntToStr(_idx[i]));
          _xml.Attributes.Add('formatCode', _fmt[i]);
          _xml.WriteEmptyTag('numFmt', true, true);
        end;

        _xml.WriteEndTagNode(); //numFmts
      end;
    finally
      FreeAndNil(_nfmt);
      SetLength(_idx, 0);
      SetLength(_fmt, 0);
      SetLength(_is_dateTime, 0);
    end;
  end; //WritenumFmts

  //Являются ли шрифты одинаковыми
  function _isFontsEqual(const fnt1, fnt2: TFont): boolean;
  begin
    result := False;
    if (fnt1.Color <> fnt2.Color) then
      exit;

    if (fnt1.Height <> fnt2.Height) then
      exit;

    if (fnt1.Name <> fnt2.Name) then
      exit;

    if (fnt1.Pitch <> fnt2.Pitch) then
      exit;

    if (fnt1.Size <> fnt2.Size) then
      exit;

    if (fnt1.Style <> fnt2.Style) then
      exit;

    Result := true; // если уж до сюда добрались
  end; //_isFontsEqual

  //Обновляет индексы в массиве
  //INPUT
  //  var arr: TIntegerDynArray  - массив
  //      cnt: integer          - номер последнего элемента в массиве (начинает с 0)
  //                              (предполагается, что возникнет ситуация, когда нужно будет использовать только часть массива)
  procedure _UpdateArrayIndex(var arr: TIntegerDynArray; cnt: integer); // deprecated {$IFDEF USE_DEPRECATED_STRING}'remove CNT parameter!'{$ENDIF};
  var
    res: TIntegerDynArray;
    i, j: integer;
    num: integer;

  begin
    //Assert( Length(arr) - cnt = 2, 'Wow! We really may need this parameter!');
    //cnt := Length(arr) - 2;   // get ready to strip it
    SetLength(res, Length(arr));

    num := 0;
    for i := -1 to cnt do
    if (arr[i + 1] = -2) then
    begin
      res[i + 1] := num;
      for j := i + 1 to cnt do
      if (arr[j + 1] = i) then
        res[j + 1] := num;
      inc(num);
    end; //if

    arr := res;
  end; //_UpdateArrayIndex

  //<fonts>...</fonts>
  procedure WriteXLSXFonts();
  var
    i, j, n: integer;
    _fontCount: integer;
    fnt: TFont;

  begin
    _fontCount := 0;
    SetLength(_FontIndex, _StylesCount + 1);
    for i := 0 to _StylesCount do
      _FontIndex[i] := -2;

    for i := -1 to _StylesCount - 1 do
    if (_FontIndex[i + 1] = -2) then
    begin
      inc (_fontCount);
      n := i + 1;
      for j := n to _StylesCount - 1 do
      if (_FontIndex[j + 1] = -2) then
        if (_isFontsEqual(XMLSS.Styles[i].Font, XMLSS.Styles[j].Font)) then
          _FontIndex[j + 1] := i;
    end; //if

    _xml.Attributes.Clear();
    _xml.Attributes.Add('count', IntToStr(_fontCount));
    _xml.WriteTagNode('fonts', true, true, true);

    for i := 0 to _StylesCount do
    if (_FontIndex[i] = - 2) then
    begin
      fnt := XMLSS.Styles[i - 1].Font;
      _xml.Attributes.Clear();
      _xml.WriteTagNode('font', true, true, true);

      _xml.Attributes.Clear();
      _xml.Attributes.Add('val', fnt.Name);
      _xml.WriteEmptyTag('name', true);

      _xml.Attributes.Clear();
      _xml.Attributes.Add('val', IntToStr(fnt.Charset));
      _xml.WriteEmptyTag('charset', true);

      _xml.Attributes.Clear();
      _xml.Attributes.Add('val', IntToStr(fnt.Size));
      _xml.WriteEmptyTag('sz', true);

      if (fnt.Color <> clWindowText) then
      begin
        _xml.Attributes.Clear();
        _xml.Attributes.Add('rgb', '00' + ColorToHTMLHex(fnt.Color));
        _xml.WriteEmptyTag('color', true);
      end;

      if (fsBold in fnt.Style) then
      begin
        _xml.Attributes.Clear();
        _xml.Attributes.Add('val', 'true');
        _xml.WriteEmptyTag('b', true);
      end;

      if (fsItalic in fnt.Style) then
      begin
        _xml.Attributes.Clear();
        _xml.Attributes.Add('val', 'true');
        _xml.WriteEmptyTag('i', true);
      end;

      if (fsStrikeOut in fnt.Style) then
      begin
        _xml.Attributes.Clear();
        _xml.Attributes.Add('val', 'true');
        _xml.WriteEmptyTag('strike', true);
      end;

      if (fsUnderline in fnt.Style) then
      begin
        _xml.Attributes.Clear();
        _xml.Attributes.Add('val', 'single');
        _xml.WriteEmptyTag('u', true);
      end;

      _xml.WriteEndTagNode(); //font
    end; //if

    _UpdateArrayIndex(_FontIndex, _StylesCount - 1);

    _xml.WriteEndTagNode(); //fonts
  end; //WriteXLSXFonts

  //Являются ли заливки одинаковыми
  function _isFillsEqual(style1, style2: TZStyle): boolean;
  begin
    result := (style1.BGColor = style2.BGColor) and
              (style1.PatternColor = style2.PatternColor) and
              (style1.CellPattern = style2.CellPattern);
  end; //_isFillsEqual

  procedure _WriteBlankFill(const st: string);
  begin
    _xml.Attributes.Clear();
    _xml.WriteTagNode('fill', true, true, true);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('patternType', st);
    _xml.WriteEmptyTag('patternFill', true, false);
    _xml.WriteEndTagNode(); //fill
  end; //_WriteBlankFill

  //<fills> ... </fills>
  procedure WriteXLSXFills();
  var
    i, j: integer;
    _fillCount: integer;
    s: string;
    b: boolean;
    _tmpColor: TColor;
    _reverse: boolean;

  begin
    _fillCount := 0;
    SetLength(_FillIndex, _StylesCount + 1);
    for i := -1 to _StylesCount - 1 do
      _FillIndex[i + 1] := -2;
    for i := -1 to _StylesCount - 1 do
    if (_FillIndex[i + 1] = - 2) then
    begin
      inc(_fillCount);
      for j := i + 1 to _StylesCount - 1 do
      if (_FillIndex[j + 1] = -2) then
        if (_isFillsEqual(XMLSS.Styles[i], XMLSS.Styles[j])) then
          _FillIndex[j + 1] := i;
    end; //if

    _xml.Attributes.Clear();
    _xml.Attributes.Add('count', IntToStr(_fillCount + 2));
    _xml.WriteTagNode('fills', true, true, true);

    //по какой-то непонятной причине, если в начале нету двух стилей заливок (none + gray125),
    //в грёбаном 2010-ом офисе глючат заливки (то-ли чтобы сложно было сделать экспорт в xlsx, то-ли
    //кривые руки у мелкомягких программеров). LibreOffice открывает нормально.
    _WriteBlankFill('none');
    _WriteBlankFill('gray125');

    //TODO:
    //ВНИМАНИЕ!!! //{tut}
    //в некоторых случаях fgColor - это цвет заливки (вроде для solid), а в некоторых - bgColor.
    //Потом не забыть разобраться.
    for i := -1 to _StylesCount - 1 do
    if (_FillIndex[i + 1] = -2) then
    begin
      _xml.Attributes.Clear();
      _xml.WriteTagNode('fill', true, true, true);

      case XMLSS.Styles[i].CellPattern of
        ZPSolid:      s := 'solid';
        ZPNone:       s := 'none';
        ZPGray125:    s := 'gray125';
        ZPGray0625:   s := 'gray0625';
        ZPDiagStripe: s := 'darkUp';
        ZPGray50:     s := 'mediumGray';
        ZPGray75:     s := 'darkGray';
        ZPGray25:     s := 'lightGray';
        ZPHorzStripe: s := 'darkHorizontal';
        ZPVertStripe: s := 'darkVertical';
        ZPReverseDiagStripe:  s := 'darkDown';
//        ZPDiagStripe:         s := 'darkUpDark'; ??
        ZPDiagCross:          s := 'darkGrid';
        ZPThickDiagCross:     s := 'darkTrellis';
        ZPThinHorzStripe:     s := 'lightHorizontal';
        ZPThinVertStripe:     s := 'lightVertical'; 
        ZPThinReverseDiagStripe:  s := 'lightDown';
        ZPThinDiagStripe:         s := 'lightUp';
        ZPThinHorzCross:          s := 'lightGrid';
        ZPThinDiagCross:          s := 'lightTrellis';
        else
          s := 'solid';
      end; //case
      _xml.Attributes.Clear();
      _xml.Attributes.Add('patternType', s);
      b := (XMLSS.Styles[i].PatternColor <> clWindow) or (XMLSS.Styles[i].BGColor <> clWindow);

      if (b) then
        _xml.WriteTagNode('patternFill', true, true, false)
      else
        _xml.WriteEmptyTag('patternFill', true, false);

      _reverse := not (XMLSS.Styles[i].CellPattern in [ZPNone, ZPSolid]);
      if (XMLSS.Styles[i].BGColor <> clWindow) then
      begin
        _xml.Attributes.Clear();
        if (_reverse) then
          _tmpColor := XMLSS.Styles[i].PatternColor
        else
          _tmpColor := XMLSS.Styles[i].BGColor;
        _xml.Attributes.Add('rgb', '00' + ColorToHTMLHex(_tmpColor));
        _xml.WriteEmptyTag('fgColor', true);
      end;

      if (XMLSS.Styles[i].PatternColor <> clWindow) then
      begin
        _xml.Attributes.Clear();
        if (_reverse) then
          _tmpColor := XMLSS.Styles[i].BGColor
        else
          _tmpColor := XMLSS.Styles[i].PatternColor;
        _xml.Attributes.Add('rgb', '00' + ColorToHTMLHex(_tmpColor));
        _xml.WriteEmptyTag('bgColor', true);
      end;

      if (b) then
        _xml.WriteEndTagNode(); //patternFill

      _xml.WriteEndTagNode(); //fill
    end; //if

    _UpdateArrayIndex(_FillIndex, _StylesCount - 1);

    _xml.WriteEndTagNode(); //fills
  end; //WriteXLSXFills();

  //единичная граница
  procedure _WriteBorderItem(StyleNum: integer; BorderNum: integer);
  var
    s, s1: string;
    _border: TZBorderStyle;
    n: integer;

  begin
    _xml.Attributes.Clear();
    case BorderNum of
      0: s := 'left';
      1: s := 'top';
      2: s := 'right';
      3: s := 'bottom';
      else
        s := 'diagonal';
    end;
    _border := XMLSS.Styles[StyleNum].Border[BorderNum];
    s1 := '';
    case _border.LineStyle of
      ZEContinuous:
        begin
          if (_border.Weight = 1) then
            s1 := 'thin'
          else
          if (_border.Weight = 2) then
            s1 := 'medium'
          else
            s1 := 'thick';
        end;
      ZEHair:
        begin
          s1 := 'hair';
        end; 
      ZEDash:
        begin
          if (_border.Weight = 1) then
            s1 := 'dashed'
          else
          if (_border.Weight >= 2) then
            s1 := 'mediumDashed';
        end;
      ZEDot:
        begin
          if (_border.Weight = 1) then
            s1 := 'dotted'
          else
          if (_border.Weight >= 2) then
            s1 := 'mediumDotted';
        end;
      ZEDashDot:
        begin
          if (_border.Weight = 1) then
            s1 := 'dashDot'
          else
          if (_border.Weight >= 2) then
            s1 := 'mediumDashDot';
        end;
      ZEDashDotDot:
        begin
          if (_border.Weight = 1) then
            s1 := 'dashDotDot'
          else
          if (_border.Weight >= 2) then
            s1 := 'mediumDashDotDot';
        end;
      ZESlantDashDot:
        begin
          s1 := 'slantDashDot';
        end;
      ZEDouble:
        begin
          s1 := 'double';
        end;
      ZENone:
        begin
        end;
    end; //case

    n := length(s1);

    if (n > 0) then
      _xml.Attributes.Add('style', s1);

    if ((_border.Color <> clBlack) and (n > 0)) then
    begin
      _xml.WriteTagNode(s, true, true, true);
      _xml.Attributes.Clear();
      _xml.Attributes.Add('rgb', '00' + ColorToHTMLHex(_border.Color));
      _xml.WriteEmptyTag('color', true);
      _xml.WriteEndTagNode();
    end else
      _xml.WriteEmptyTag(s, true);
  end; //_WriteBorderItem

  //<borders> ... </borders>
  procedure WriteXLSXBorders();
  var
    i, j: integer;
    _borderCount: integer;
    s: string;

  begin
    _borderCount := 0;
    SetLength(_BorderIndex, _StylesCount + 1);
    for i := -1 to _StylesCount - 1 do
      _BorderIndex[i + 1] := -2;
    for i := -1 to _StylesCount - 1 do
    if (_BorderIndex[i + 1] = - 2) then
    begin
      inc(_borderCount);
      for j := i + 1 to _StylesCount - 1 do
      if (_BorderIndex[j + 1] = -2) then
        if (XMLSS.Styles[i].Border.isEqual(XMLSS.Styles[j].Border)) then
          _BorderIndex[j + 1] := i;
    end; //if

    _xml.Attributes.Clear();
    _xml.Attributes.Add('count', IntToStr(_borderCount));
    _xml.WriteTagNode('borders', true, true, true);

    for i := -1 to _StylesCount - 1 do
    if (_BorderIndex[i + 1] = -2) then
    begin
      _xml.Attributes.Clear();
      s := 'false';
      if (XMLSS.Styles[i].Border[4].Weight > 0) and (XMLSS.Styles[i].Border[4].LineStyle <> ZENone) then
        s := 'true';
      _xml.Attributes.Add('diagonalDown', s);
      s := 'false';
      if (XMLSS.Styles[i].Border[5].Weight > 0) and (XMLSS.Styles[i].Border[5].LineStyle <> ZENone) then
        s := 'true';
      _xml.Attributes.Add('diagonalUp', s, false);
      _xml.WriteTagNode('border', true, true, true);

      _WriteBorderItem(i, 0);
      _WriteBorderItem(i, 2);
      _WriteBorderItem(i, 1);
      _WriteBorderItem(i, 3);
      _WriteBorderItem(i, 4);

      _xml.WriteEndTagNode(); //border
    end; //if

    _UpdateArrayIndex(_BorderIndex, _StylesCount - 1);

    _xml.WriteEndTagNode(); //borders
  end; //WriteXLSXBorders

  //Добавляет <xf> ... </xf>
  //INPUT
  //      NumStyle: integer - номер стиля
  //      isxfId: boolean   - нужно ли добавлять атрибут "xfId"
  //      xfId: integer     - значение "xfId"
  procedure _WriteXF(NumStyle: integer; isxfId: boolean; xfId: integer);
  var
    _addalignment: boolean;
    _style: TZStyle;
    s: string;
    i: integer;
    _num: integer;

  begin
    _xml.Attributes.Clear();
    _style := XMLSS.Styles[NumStyle];
    _addalignment := _style.Alignment.WrapText or
                     _style.Alignment.VerticalText or
                    (_style.Alignment.Rotate <> 0) or
                    (_style.Alignment.Indent <> 0) or
                    _style.Alignment.ShrinkToFit or
                    (_style.Alignment.Vertical <> ZVAutomatic) or
                    (_style.Alignment.Horizontal <> ZHAutomatic);

    _xml.Attributes.Add('applyAlignment', XLSXBoolToStr(_addalignment));
    _xml.Attributes.Add('applyBorder', 'true', false);
    _xml.Attributes.Add('applyFont', 'true', false);
    _xml.Attributes.Add('applyProtection', 'true', false);
    _xml.Attributes.Add('borderId', IntToStr(_BorderIndex[NumStyle + 1]), false);
    _xml.Attributes.Add('fillId', IntToStr(_FillIndex[NumStyle + 1] + 2), false); //+2 т.к. первыми всегда идут 2 левых стиля заливки
    _xml.Attributes.Add('fontId', IntToStr(_FontIndex[NumStyle + 1]), false);

    // ECMA 376 Ed.4:  12.3.20 Styles Part; 17.9.17 numFmt (Numbering Format); 18.8.30 numFmt (Number Format)
    // http://social.msdn.microsoft.com/Forums/sa/oxmlsdk/thread/3919af8c-644b-4d56-be65-c5e1402bfcb6
    if (isxfId) then
      _num := _NumFmtIndexes[NumStyle + 1]
    else
      _num := 0;

    _xml.Attributes.Add('numFmtId', IntToStr(_num) {'164'}, false); // TODO: support formats

    if (_num > 0) then
      _xml.Attributes.Add('applyNumberFormat', '1', false);

    if (isxfId) then
      _xml.Attributes.Add('xfId', IntToStr(xfId), false);

    _xml.WriteTagNode('xf', true, true, true);

    if (_addalignment) then
    begin
      _xml.Attributes.Clear();
      case (_style.Alignment.Horizontal) of
        ZHLeft: s := 'left';
        ZHRight: s := 'right';
        ZHCenter: s := 'center';
        ZHFill: s := 'fill';
        ZHJustify: s := 'justify';
        ZHDistributed: s := 'distributed';
        ZHAutomatic:   s := 'general';
        else
          s := 'general';
(*  The standard does not specify a default value for the horizontal attribute.
        Excel uses a default value of general for this attribute.
    MS-OI29500: Microsoft Office Implementation Information for ISO/IEC-29500, 18.8.1.d *)
      end; //case
      _xml.Attributes.Add('horizontal', s);
      _xml.Attributes.Add('indent', IntToStr(_style.Alignment.Indent), false);
      _xml.Attributes.Add('shrinkToFit', XLSXBoolToStr(_style.Alignment.ShrinkToFit), false);


      if _style.Alignment.VerticalText then i := 255
         else i := ZENormalizeAngle180(_style.Alignment.Rotate);
      _xml.Attributes.Add('textRotation', IntToStr(i), false);

      case (_style.Alignment.Vertical) of
        ZVCenter: s := 'center';
        ZVTop: s := 'top';
        ZVBottom: s := 'bottom';
        ZVJustify: s := 'justify';
        ZVDistributed: s := 'distributed';
        else
          s := 'bottom';
(*  The standard does not specify a default value for the vertical attribute.
        Excel uses a default value of bottom for this attribute.
    MS-OI29500: Microsoft Office Implementation Information for ISO/IEC-29500, 18.8.1.e *)
      end; //case
      _xml.Attributes.Add('vertical', s, false);
      _xml.Attributes.Add('wrapText', XLSXBoolToStr(_style.Alignment.WrapText), false);
      _xml.WriteEmptyTag('alignment', true);
    end; //if (_addalignment)

    _xml.Attributes.Clear();
    _xml.Attributes.Add('hidden', XLSXBoolToStr(XMLSS.Styles[NumStyle].Protect));
    _xml.Attributes.Add('locked', XLSXBoolToStr(XMLSS.Styles[NumStyle].HideFormula));
    _xml.WriteEmptyTag('protection', true);

    _xml.WriteEndTagNode(); //xf
  end; //_WriteXF

  //<cellStyleXfs> ... </cellStyleXfs> / <cellXfs> ... </cellXfs>
  procedure WriteCellStyleXfs(const TagName: string; isxfId: boolean);
  var
    i: integer;

  begin
    _xml.Attributes.Clear();
    _xml.Attributes.Add('count', IntToStr(XMLSS.Styles.Count + 1));
    _xml.WriteTagNode(TagName, true, true, true);
    for i := -1 to XMLSS.Styles.Count - 1 do
    begin
      //Что-то не совсем понятно, какой именно xfId нужно ставить. Пока будет 0 для всех.
      _WriteXF(i, isxfId, 0{i + 1});
    end;
    _xml.WriteEndTagNode(); //cellStyleXfs
  end; //WriteCellStyleXfs

  //<cellStyles> ... </cellStyles>
  procedure WriteCellStyles();
  begin
  end; //WriteCellStyles 

begin
  result := 0;
  _xml := nil;
  _FmtParser := nil;
  _DateParser := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    _FmtParser := TNumFormatParser.Create();
    _DateParser := TZDateTimeODSFormatParser.Create();

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    _StylesCount := XMLSS.Styles.Count;

    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    _xml.WriteTagNode('styleSheet', true, true, true);

    WritenumFmts();

    WriteXLSXFonts();
    WriteXLSXFills();
    WriteXLSXBorders();
    //DO NOT remove cellStyleXfs!!!
    WriteCellStyleXfs('cellStyleXfs', false);
    WriteCellStyleXfs('cellXfs', true);

    // experiment: do not need styles, ergo do not need fake xfId
    //Result: experiment failed!
    //Libre/Open Office needs cellStyleXfs for reading background colors!
    // (Libre/Open Office have highest priority)
    //WriteCellStyleXfs('cellXfs', false);
    // experiment end

    WriteCellStyles(); //??

    _xml.WriteEndTagNode(); //styleSheet
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
    if (Assigned(_FmtParser)) then
      FreeAndNil(_FmtParser);
    if (Assigned(_DateParser)) then
      FreeAndNil(_DateParser);
    SetLength(_FontIndex, 0);
    SetLength(_FillIndex, 0);
    SetLength(_BorderIndex, 0);
    SetLength(_NumFmtIndexes, 0);
  end;
end; //ZEXLSXCreateStyles

//Добавить Relationship для rels
//INPUT
//      xml: TZsspXMLWriterH  - писалка
//  const rid: string         - rid
//      ridType: integer      - rIdType (0..8)
//  const Target: string      -
//  const TargetMode: string  -
procedure ZEAddRelsRelation(xml: TZsspXMLWriterH; const rid: string; ridType: integer; const Target: string; const TargetMode: string = '');
begin
  xml.Attributes.Clear();
  xml.Attributes.Add('Id', rid);
  xml.Attributes.Add('Type',  ZEXLSXGetRelationName(ridType), false);
  xml.Attributes.Add('Target', Target, false);
  if (TargetMode <> '') then
     xml.Attributes.Add('TargetMode', TargetMode, true);
  xml.WriteEmptyTag('Relationship', true, true);
end; //ZEAddRelsID

//Создаёт _rels/.rels
//INPUT
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateRelsMain(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;

begin
  result := 0;
  _xml := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    _xml.WriteTagNode('Relationships', true, true, false);

    ZEAddRelsRelation(_xml, 'rId1', 3, 'xl/workbook.xml');
    ZEAddRelsRelation(_xml, 'rId2', 5, 'docProps/app.xml');
    ZEAddRelsRelation(_xml, 'rId3', 4, 'docProps/core.xml');

    _xml.WriteEndTagNode(); //Relationships

  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ZEXLSXCreateRelsMain

//Создаёт xl/_rels/workbook.xml.rels   
//INPUT
//    PageCount: integer                - кол-во страниц
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateRelsWorkBook(PageCount: integer; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;
  i: integer;

begin
  result := 0;
  _xml := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    _xml.WriteTagNode('Relationships', true, true, false);

    ZEAddRelsRelation(_xml, 'rId1', 1, 'styles.xml');

    for i := 0 to PageCount - 1 do
      ZEAddRelsRelation(_xml, 'rId' + IntToStr(i + 2), 0, 'worksheets/sheet' + IntToStr(i + 1) + '.xml');

    ZEAddRelsRelation(_xml, 'rId' + IntToStr(PageCount + 2), 2, 'sharedStrings.xml');

    _xml.WriteEndTagNode(); //Relationships

  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ZEXLSXCreateRelsWorkBook

//Создаёт sharedStrings.xml (пока не реализовано)
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateSharedStrings(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;

begin
  result := 0;
  _xml := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('count', '0');
    _xml.Attributes.Add('uniqueCount', '0', false);
    _xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN, false);
    _xml.WriteTagNode('sst', true, true, false);

    //Пока не заполняется

    _xml.WriteEndTagNode(); //Relationships

  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ZEXLSXCreateSharedStrings

//Создаёт app.xml
//INPUT
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateDocPropsApp(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;
//  s: string;

begin
  result := 0;
  _xml := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', SCHEMA_DOC + '/extended-properties');
    _xml.Attributes.Add('xmlns:vt', SCHEMA_DOC + '/docPropsVTypes', false);
    _xml.WriteTagNode('Properties', true, true, false);

    _xml.Attributes.Clear();
    _xml.WriteTag('TotalTime', '0', true, false, false);

//    {$IFDEF FPC}
//    s := 'FPC';
//    {$ELSE}
//    s := 'DELPHI_or_CBUILDER';
//    {$ENDIF}
//    _xml.WriteTag('Application', 'ZEXMLSSlib/0.0.5$' + s, true, false, true);
    _xml.WriteTag('Application', ZELibraryName(), true, false, true);

    _xml.WriteEndTagNode(); //Properties

  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ZEXLSXCreateDocPropsApp

//Создаёт app.xml 
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateDocPropsCore(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;
  s: string;

begin
  result := 0;
  _xml := nil;
  try
    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.TabSymbol := ' ';
    if (not _xml.BeginSaveToStream(Stream)) then
    begin
      result := 2;
      exit;
    end;

    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns:cp', SCHEMA_PACKAGE + '/metadata/core-properties');
    _xml.Attributes.Add('xmlns:dc', 'http://purl.org/dc/elements/1.1/', false);
    _xml.Attributes.Add('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/', false);
    _xml.Attributes.Add('xmlns:dcterms', 'http://purl.org/dc/terms/', false);
    _xml.Attributes.Add('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance', false);
    _xml.WriteTagNode('cp:coreProperties', true, true, false);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('xsi:type', 'dcterms:W3CDTF');
    s := ZEDateTimeToStr(XMLSS.DocumentProperties.Created) + 'Z';
    _xml.WriteTag('dcterms:created', s, true, false, false);
    _xml.WriteTag('dcterms:modified', s, true, false, false);

    _xml.Attributes.Clear();
    _xml.WriteTag('cp:revision', '1', true, false, false);

    _xml.WriteEndTagNode(); //cp:coreProperties

  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ZEXLSXCreateDocPropsCore

//Сохраняет незапакованный документ в формате Office Open XML (OOXML)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//      TextConverter: TAnsiToCPConverter - конвертер
//      CodePageName: string              - имя кодировки
//      BOM: ansistring                   - Byte Order Mark
//RETURN
//      integer
function ExportXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string;
  const SheetsNumbers:array of integer;
  const SheetsNames: array of string;
  TextConverter: TAnsiToCPConverter;
  CodePageName: string;
  BOM: ansistring = '';
  AllowUnzippedFolder: boolean = false): integer;
var
  _pages: TIntegerDynArray;      //номера страниц
  _names: TStringDynArray;      //названия страниц
  kol, i, ii: integer;
  zip: TZipper;
  _WriteHelper: TZEXLSXWriteHelper;
  path_xl, path_sheets, path_relsmain, path_relsw, path_docprops: string;
  _commentArray: TIntegerDynArray;
  StreamList: TStreamList;
  TmpStream: TStream;
  {$IFDEF ZUSE_DRAWINGS}
  iDrawingsCount: Integer;
  path_draw, path_draw_rel, path_media: string;
  _drawing: TZEDrawing;
  _pic: TZEPicture;
  {$ENDIF} // ZUSE_DRAWINGS

  function _AddFile(const AStreamFileName: string): TStream;
  var
    sFullFileName: string;
    sFullFilePath: string;
  begin
    if AllowUnzippedFolder then
    begin
      sFullFileName := IncludeTrailingPathDelimiter(FileName) + AStreamFileName;
      if PathDelim <> '/' then
      begin
        // make copy of string to avoid problems with old Delphi
        sFullFileName := StringReplace(Copy(sFullFileName, 1, Maxint), '/', PathDelim, [rfReplaceAll]);
      end;
      sFullFilePath := ExtractFilePath(sFullFileName);
      if (not DirectoryExists(sFullFilePath)) then
        ForceDirectories(sFullFilePath);

      Result := TFileStream.Create(sFullFileName, fmCreate);
      StreamList.Add(Result);
    end
    else
    begin
      Result := TMemoryStream.Create();
      StreamList.Add(Result);
      //Result := TFileStream.Create(AStreamFileName, fmCreate);
      zip.Entries.AddFileEntry(Result, AStreamFileName);
    end;
  end;

begin
  StreamList := TStreamList.Create();
  _WriteHelper := TZEXLSXWriteHelper.Create();
  zip := TZipper.Create();
  try

    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      result := 2;
      exit;
    end;

    zip.FileName := FileName;

    path_xl := 'xl/';

    //стили
    TmpStream := _AddFile(path_xl + 'styles.xml');
    ZEXLSXCreateStyles(XMLSS, TmpStream, TextConverter, CodePageName, BOM);

    //sharedStrings.xml
    TmpStream := _AddFile(path_xl + 'sharedStrings.xml');
    ZEXLSXCreateSharedStrings(XMLSS, TmpStream, TextConverter, CodePageName, BOM);

    // _rels/.rels
    path_relsmain := '_rels/';
    TmpStream := _AddFile(path_relsmain + '.rels');
    ZEXLSXCreateRelsMain(TmpStream, TextConverter, CodePageName, BOM);

    // xl/_rels/workbook.xml.rels
    path_relsw := path_xl + '_rels/';
    TmpStream := _AddFile(path_relsw + 'workbook.xml.rels');
    ZEXLSXCreateRelsWorkBook(kol, TmpStream, TextConverter, CodePageName, BOM);

    path_sheets := path_xl + 'worksheets/';
    SetLength(_commentArray, kol);

    {$IFDEF ZUSE_DRAWINGS}
    iDrawingsCount := XMLSS.DrawingCount();
    {$ENDIF} // ZUSE_DRAWINGS

    //Листы документа
    for i := 0 to kol - 1 do
    begin
      _commentArray[i] := 0;
      TmpStream := _AddFile(path_sheets + 'sheet' + IntToStr(i + 1) + '.xml');
      ZEXLSXCreateSheet(XMLSS, TmpStream, _pages[i], TextConverter, CodePageName, BOM, _WriteHelper);

      if (_WriteHelper.HyperLinksCount > 0) then
      begin
        _WriteHelper.AddSheetHyperlink(i);
        TmpStream := _AddFile(path_sheets + '_rels/sheet' + IntToStr(i + 1) + '.xml.rels');
        _WriteHelper.CreateSheetRels(TmpStream, TextConverter, CodePageName, BOM);
      end;

      if (_WriteHelper.isHaveComments) then
      begin
        _commentArray[i] := 1;
        //создать файл с комментариями
      end;
    end; //for i

    {$IFDEF ZUSE_DRAWINGS}
    //iDrawingsCount := XMLSS.DrawingCount();
    if iDrawingsCount <> 0 then
    begin
      path_draw := path_xl + 'drawings/';
      path_draw_rel := path_draw + '_rels/';
      path_media := path_xl + 'media/';

      for i := 0 to iDrawingsCount - 1 do
      begin
        _drawing := XMLSS.GetDrawing(i);
        // drawings/drawingN.xml
        TmpStream := _AddFile(path_draw + 'drawing' + IntToStr(_drawing.Id) + '.xml');
        ZEXLSXCreateDrawing(XMLSS, TmpStream, _drawing, TextConverter, CodePageName, BOM);

        // drawings/_rels/drawingN.xml.rels
        TmpStream := _AddFile(path_draw_rel + 'drawing' + IntToStr(i + 1) + '.xml.rels');
        ZEXLSXCreateDrawingRels(XMLSS, TmpStream, _drawing, TextConverter, CodePageName, BOM);

        // media/imageN.png
        for ii := 0 to _drawing.PictureStore.Count - 1 do
        begin
          _pic := _drawing.PictureStore[ii];
          if not Assigned(_pic.DataStream) then Continue;
          TmpStream := _AddFile(path_media + _pic.Name);
          _pic.DataStream.Position := 0;
          TmpStream.CopyFrom(_pic.DataStream, _pic.DataStream.Size);
        end;
      end;
    end;
    {$ENDIF} // ZUSE_DRAWINGS

    //workbook.xml - список листов
    TmpStream := _AddFile(path_xl + 'workbook.xml');
    ZEXLSXCreateWorkBook(XMLSS, TmpStream, _pages, _names, kol, TextConverter, CodePageName, BOM);

    //[Content_Types].xml
    TmpStream := _AddFile('[Content_Types].xml');
    ZEXLSXCreateContentTypes(XMLSS, TmpStream, kol, 0, nil, TextConverter, CodePageName, BOM, _WriteHelper);

    path_docprops :='docProps/';
    // docProps/app.xml
    TmpStream := _AddFile(path_docprops + 'app.xml');
    ZEXLSXCreateDocPropsApp(TmpStream, TextConverter, CodePageName, BOM);

    // docProps/core.xml
    TmpStream := _AddFile(path_docprops + 'core.xml');
    ZEXLSXCreateDocPropsCore(XMLSS, TmpStream, TextConverter, CodePageName, BOM);

    zip.ZipAllFiles();
  finally
    StreamList.Free();
    ZESClearArrays(_pages, _names);
    if (Assigned(zip)) then
      FreeAndNil(zip);

    SetLength(_commentArray, 0);
    _commentArray := nil;

    FreeAndNil(_WriteHelper);
  end;
  Result := 0;
end;

//Сохраняет незапакованный документ в формате Office Open XML (OOXML)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//      TextConverter: TAnsiToCPConverter - конвертер
//      CodePageName: string              - имя кодировки
//      BOM: ansistring                   - Byte Order Mark
//RETURN
//      integer
function SaveXmlssToXLSXPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
begin
  Result := ExportXmlssToXLSX(XMLSS, PathName, SheetsNumbers, SheetsNames, TextConverter, CodePageName, BOM, True);
end; //SaveXmlssToXLSXPath

//SaveXmlssToXLSXPath
//Сохраняет незапакованный документ в формате Office Open XML (OOXML)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//RETURN
//      integer
function SaveXmlssToXLSXPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers: array of integer;
                             const SheetsNames: array of string): integer; overload;

begin
  result := SaveXmlssToXLSXPath(XMLSS, PathName, SheetsNumbers, SheetsNames, ZEGetDefaultUTF8Converter(), 'UTF-8', '');
end; //SaveXmlssToXLSXPath

//SaveXmlssToXLSXPath
//Сохраняет незапакованный документ в формате Office Open XML (OOXML)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//RETURN
//      integer
function SaveXmlssToXLSXPath(var XMLSS: TZEXMLSS; PathName: string): integer; overload;
begin
  Result := SaveXmlssToXLSXPath(XMLSS, PathName, [], []);
end; //SaveXmlssToXLSXPath

//Сохраняет документ в формате Open Office XML (xlsx)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      FileName: string                  - имя файла для сохранения
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//      TextConverter: TAnsiToCPConverter - конвертер
//      CodePageName: string              - имя кодировки
//      BOM: ansistring                   - Byte Order Mark
//RETURN
//      integer
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer;
begin
  Result := ExportXmlssToXLSX(XMLSS, FileName, SheetsNumbers, SheetsNames, TextConverter, CodePageName, BOM, False);
end; //SaveXmlssToXSLX

//Сохраняет документ в формате Open Office XML (xlsx)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      FileName: string                  - имя файла для сохранения
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//RETURN
//      integer
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string): integer; overload;
begin
  result := SaveXmlssToXLSX(XMLSS, FileName, SheetsNumbers, SheetsNames, ZEGetDefaultUTF8Converter(), 'UTF-8', '');
end; //SaveXmlssToXLSX

//Сохраняет документ в формате Open Office XML (xlsx)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      FileName: string                  - имя файла для сохранения
//RETURN
//      integer
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
begin
  result := SaveXmlssToXLSX(XMLSS, FileName, [], []);
end; //SaveXmlssToXLSX

//Перепутал малость названия ^_^

function ReadXSLXPath(var XMLSS: TZEXMLSS; DirName: string): integer; //deprecated 'Use ReadXLSXPath!';
begin
  result := ReadXLSXPath(XMLSS, DirName);
end;

function ReadXSLX(var XMLSS: TZEXMLSS; FileName: string): integer; //deprecated 'Use ReadXLSX!';
begin
  result := ReadXLSX(XMLSS, FileName);
end;

{ TStreamList }

function TStreamList.GetItem(AIndex: Integer): TStream;
begin
  Result := TStream(Get(AIndex));
end;

procedure TStreamList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  if Action = lnDeleted then
    TStream(Ptr).Free();
  inherited;
end;

end.
