//****************************************************************
// Read/write xlsx (Office Open XML file format (Spreadsheet))
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2013.01.30
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
//****************************************************************
unit zexlsx;

interface

{$I zexml.inc}
{$I compver.inc}

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

uses
  SysUtils, Classes, Types, Graphics,
  zeformula, zsspxml, zexmlss, zesavecommon, zeZippy
  {$IFNDEF FPC}
  ,windows
  {$ELSE}
  ,LCLType,
  LResources
  {$ENDIF}
  {$IFDEF FPC},zipper{$ELSE}{$I xlsxzipuses.inc}{$ENDIF};

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

function ReadXSLXPath(var XMLSS: TZEXMLSS; DirName: string): integer; deprecated {$IFDEF FPC}'Use ReadXLSXPath!'{$ENDIF};
function ReadXLSXPath(var XMLSS: TZEXMLSS; DirName: string): integer;

{$IFDEF FPC}
function ReadXSLX(var XMLSS: TZEXMLSS; FileName: string): integer; deprecated {$IFDEF FPC}'Use ReadXLSX!'{$ENDIF};
function ReadXLSX(var XMLSS: TZEXMLSS; FileName: string): integer;
{$ENDIF}

function SaveXmlssToXLSXPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers: array of integer;
                             const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer;

{$IFDEF FPC}
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer;
{$ENDIF}

{$IFNDEF FPC}
{$I xlsxzipfunc.inc}
{$ENDIF}

function ExportXmlssToXLSX(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string;
                         BOM: ansistring = '';
                         AllowUnzippedFolder: boolean = false;
                         ZipGenerator: CZxZipGens = nil): integer;

//Дополнительные функции, на случай чтения отдельного файла
function ZEXSLXReadTheme(var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer): boolean;
function ZEXSLXReadContentTypes(var Stream: TStream; var FileArray: TZXLSXFileArray; var FilesCount: integer): boolean;
function ZEXSLXReadSharedStrings(var Stream: TStream; out StrArray: TStringDynArray; out StrCount: integer): boolean;
function ZEXSLXReadStyles(var XMLSS: TZEXMLSS; var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer): boolean;
function ZE_XSLXReadRelationships(var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer; var isWorkSheet: boolean; needReplaceDelimiter: boolean): boolean;
function ZEXSLXReadWorkBook(var XMLSS: TZEXMLSS; var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer): boolean;
function ZEXSLXReadSheet(var XMLSS: TZEXMLSS; var Stream: TStream; const SheetName: string; var StrArray: TStringDynArray; StrCount: integer; var Relations: TZXLSXRelationsArray; RelationsCount: integer): boolean;
function ZEXSLXReadComments(var XMLSS: TZEXMLSS; var Stream: TStream): boolean;

//Дополнительные функции для экспорта отдельных файлов
function ZEXLSXCreateStyles(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateWorkBook(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                              const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
function ZEXLSXCreateSheet(var XMLSS: TZEXMLSS; Stream: TStream; SheetNum: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring; out isHaveComments: boolean): integer;
function ZEXLSXCreateContentTypes(var XMLSS: TZEXMLSS; Stream: TStream; PageCount: integer; CommentCount: integer; const PagesComments: TIntegerDynArray;

                                  TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateRelsMain(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateSharedStrings(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateDocPropsApp(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateDocPropsCore(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;

implementation uses
{$IfDef DELPHI_UNICODE}
  AnsiStrings,  // AnsiString targeted overloaded versions of Pos, Trim, etc
{$EndIf}
  StrUtils;

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
    fontsize: integer;
  end;

  TZEXLSXFontArray = array of TZEXLSXFont; //массив шрифтов

{$IFDEF FPC}
  //Для распаковки в поток

  { TXSLXZipHelper }

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
    function GetFileItem(num: integer): TZXLSXFileItem;
    function GetRelationsCounts(num: integer): integer;
    function GetRelationsArray(num: integer): TZXLSXRelationsArray;
    function GetArchFileItem(num: integer): string;
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
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
  inherited;
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
  AStream := TMemorystream.Create;
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
        if (not ZEXSLXReadStyles(FXMLSS, AStream, FThemaColor, FThemaColorCount)) then
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
        if (not ZEXSLXReadSheet(FXMLSS, AStream, ListName, FStrArray, FStrCount, FSheetRelations, FSheetRelationsCount)) then
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
{$ENDIF}

//Возвращает номер Relations из rels
//INPUT
//  const name: string - текст отношения
//RETURN
//      integer - номер отношения. -1 - не определено
function ZEXLSXGetRelationNumber(const name: string): integer;
begin
  result := -1;
  if (name = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet') then
    result := 0
  else
  if (name = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles') then
    result := 1
  else
  if (name = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings') then
    result := 2
  else
  if (name = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument') then
    result := 3
  else
  if (name = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties') then
    result := 4
  else
  if (name = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties') then
    result := 5
  else
  if (name = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') then
    result := 6
  else
  if (name = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments') then
    result := 7
  else
  if (name = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing') then
    result := 8;        
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
    0: result := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
    1: result := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
    2: result := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
    3: result := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
    4: result := 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
    5: result := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
    6: result := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
    7: result := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
    8: result := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
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
        if (xml.TagType = 4) then
          _b := true;
        if (xml.TagType = 6) then
          _b := false;
      end else
      if ((xml.TagName = 'a:sysClr') and (_b) and (xml.TagType in [4, 5])) then
      begin
        _addFillColor(xml.Attributes.ItemsByName['lastClr']);
      end else
      if ((xml.TagName = 'a:srgbClr') and (_b) and (xml.TagType in [4, 5])) then
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
      if ((xml.TagName = 'Override') and (xml.TagType = 5)) then
      begin
        s := xml.Attributes.ItemsByName['PartName'];
        if (length(s) > 0) then
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
          if (s = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml') then
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
      if ((xml.TagName = 'si') and (xml.TagType = 4)) then
      begin
        s := '';
        k := 0;
        while (not((xml.TagName = 'si') and (xml.TagType = 6))) and (not xml.Eof) do
        begin
          xml.ReadTag();
          if ((xml.TagName = 't') and (xml.TagType = 6)) then
          begin
            if (k > 1) then
              s := s + {$IFDEF FPC} LineEnding {$ELSE} sLineBreak {$ENDIF};
            s := s + xml.TextBeforeTag;
          end;
          if (xml.TagName = 'r') and (xml.TagType = 6) then
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

//Читает страницу документа
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//  var Stream: TStream                 - поток для чтения
//  const SheetName: string             - название страницы
//  var StrArray: TStringDynArray       - строки для подстановки
//      StrCount: integer               - кол-во строк подстановки
//  var Relations: TZXLSXRelationsArray - отношения
//      RelationsCount: integer         - кол-во отношений
//RETURN
//      boolean - true - страница прочиталась успешно
function ZEXSLXReadSheet(var XMLSS: TZEXMLSS; var Stream: TStream; const SheetName: string; var StrArray: TStringDynArray; StrCount: integer;
                         var Relations: TZXLSXRelationsArray; RelationsCount: integer): boolean;
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
    while (not ((xml.TagName = 'sheetData') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof) then
        break;

      //ячейка
      if (xml.TagName = 'c') then
      begin
        s := xml.Attributes.ItemsByName['r']; //номер
        if (length(s) > 0) then
          if (ZEGetCellCoords(s, _cc, _cr)) then
          begin
            _currCol := _cc;
            CheckCol(_cc + 1);
          end;

        _type := xml.Attributes.ItemsByName['t']; //тип
        //s := xml.Attributes.ItemsByName['cm'];
        //s := xml.Attributes.ItemsByName['ph'];
        //s := xml.Attributes.ItemsByName['vm'];
        v := '';
        _num := 0;
        _currCell := _currSheet.Cell[_currCol, _currRow];
        s := xml.Attributes.ItemsByName['s']; //стиль
        if (length(s) > 0) then
          if (tryStrToInt(s, t)) then
            _currCell.CellStyle := t;
        if (xml.TagType = 4) then
        while (not ((xml.TagName = 'c') and (xml.TagType = 6))) do
        begin
          xml.ReadTag();
          if (xml.Eof) then
            break;

          //is пока игнорируем

          if (((xml.TagName = 'v') or (xml.TagName = 't')) and (xml.TagType = 6)) then
          begin
            if (_num > 0) then
              v := v + {$IFDEF FPC} LineEnding {$ELSE} sLineBreak {$ENDIF};
            v := v + xml.TextBeforeTag;  
            inc(_num);
          end else
          if ((xml.TagName = 'f') and (xml.TagType = 6)) then
            _currCell.Formula := ZEReplaceEntity(xml.TextBeforeTag);

        end; //while

        //s - sharedstring
        //b - boolean
        //n - number
        //e - error
        //str - string
        //inlineStr - inline string ??
        if (_type = 's') then
        begin
          _currCell.CellType := ZEString;
          if (TryStrToInt(v, t)) then
            if ((t >= 0) and (t < StrCount)) then
              v := StrArray[t];
        end;

        _currCell.Data := ZEReplaceEntity(v);
        inc(_currCol);
        CheckCol(_currCol + 1);
      end else
      //строка
      if ((xml.TagName = 'row') and (xml.TagType in [4, 5])) then
      begin
        _currCol := 0;
        s := xml.Attributes.ItemsByName['r']; //индекс строки
        if (length(s) > 0) then
          if (TryStrToInt(s, t)) then
          begin
            _currRow := t - 1;
            CheckRow(t);
          end;
        //s := xml.Attributes.ItemsByName['collapsed'];
        //s := xml.Attributes.ItemsByName['customFormat'];
        //s := xml.Attributes.ItemsByName['customHeight'];
        s := xml.Attributes.ItemsByName['hidden'];
        if (length(s) > 0) then
          _currSheet.Rows[_currRow].Hidden := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['ht']; //в поинтах?
        if (length(s) > 0) then
        begin
          _tmpr := ZETryStrToFloat(s, 10);
          _tmpr := _tmpr / 2.835; //???
          _currSheet.Rows[_currRow].HeightMM := _tmpr;
        end;

        //s := xml.Attributes.ItemsByName['outlineLevel'];
        //s := xml.Attributes.ItemsByName['ph'];

        s := xml.Attributes.ItemsByName['s']; //номер стиля
        if (length(s) > 0) then
          if (TryStrToInt(s, t)) then
          begin
            //нужно подставить нужный стиль
          end;
        //s := xml.Attributes.ItemsByName['spans'];
        //s := xml.Attributes.ItemsByName['thickBot'];
        //s := xml.Attributes.ItemsByName['thickTop'];

        if (xml.TagType = 5) then
        begin
          inc(_currRow);
          CheckRow(_currRow + 1);
        end;
      end else
      //конец строки
      if ((xml.TagName = 'row') and (xml.TagType = 6)) then
      begin
        inc(_currRow);
        CheckRow(_currRow + 1);
      end;
    end; //while
  end; //_ReadSheetData

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
    while (not ((xml.TagName = 'mergeCells') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof) then
        break;

      if ((xml.TagName = 'mergeCell') and (xml.TagType in [4, 5])) then
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

  begin
    num := 0;
    while (not ((xml.TagName = 'cols') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof) then
        break;

      if ((xml.TagName = 'col') and (xml.TagType in [4, 5])) then
      begin
        CheckCol(num + 1);
        s := xml.Attributes.ItemsByName['bestFit'];
        if (length(s) > 0) then
          _currSheet.Columns[num].AutoFitWidth := ZETryStrToBoolean(s);

        //s := xml.Attributes.ItemsByName['collapsed'];
        //s := xml.Attributes.ItemsByName['customWidth'];
        s := xml.Attributes.ItemsByName['hidden'];
        if (length(s) > 0) then
          _currSheet.Columns[num].Hidden := ZETryStrToBoolean(s);

        //s := xml.Attributes.ItemsByName['max'];
        //s := xml.Attributes.ItemsByName['min'];
        //s := xml.Attributes.ItemsByName['outlineLevel'];
        //s := xml.Attributes.ItemsByName['phonetic'];
        s := xml.Attributes.ItemsByName['style'];

        s := xml.Attributes.ItemsByName['width'];
        if (length(s) > 0) then
        begin
          t := ZETryStrToFloat(s, 5.14509803921569);
          t := 10 * t / 5.14509803921569;
          _currSheet.Columns[num].WidthMM := t;
        end;

        inc(num);
      end; //if
    end; //while    
  end; //_ReadCols

  function _StrToMM(const st: string; var retFloat: real): boolean;
  begin
    result := false;
    if (length(s) > 0) then
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
      if (st[l] = ':') then
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
        s := s + st[l];
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
    while (not ((xml.TagName = 'hyperlinks') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'hyperlink') and (xml.TagType = 5)) then
      begin
        s := xml.Attributes.ItemsByName['ref'];
        if (length(s) > 0) then
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
    while (not ((xml.TagName = 'sheetViews') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'pane') and (xml.TagType = 5)) then
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
      if ((xml.TagName = 'sheetData') and (xml.TagType = 4)) then
        _ReadSheetData()
      else
      if ((xml.TagName = 'mergeCells') and (xml.TagType = 4)) then
        _ReadMerge()
      else
      if ((xml.TagName = 'cols') and (xml.TagType = 4)) then
        _ReadCols()
      else
      //Отступы
      if ((xml.TagName = 'pageMargins') and (xml.TagType = 5)) then
      begin
        //в дюймах
        s := xml.Attributes.ItemsByName['bottom'];
        if (_StrToMM(s, _tmpr)) then
          _currSheet.SheetOptions.MarginBottom := round(_tmpr);
        s := xml.Attributes.ItemsByName['footer'];
        if (_StrToMM(s, _tmpr)) then
          _currSheet.SheetOptions.FooterMargin := round(_tmpr);
        s := xml.Attributes.ItemsByName['header'];
        if (_StrToMM(s, _tmpr)) then
          _currSheet.SheetOptions.HeaderMargin := round(_tmpr);
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
      if ((xml.TagName = 'pageSetup') and (xml.TagType = 5)) then
      begin
        //s := xml.Attributes.ItemsByName['blackAndWhite'];
        //s := xml.Attributes.ItemsByName['cellComments'];
        //s := xml.Attributes.ItemsByName['copies'];
        //s := xml.Attributes.ItemsByName['draft'];
        //s := xml.Attributes.ItemsByName['errors'];
        s := xml.Attributes.ItemsByName['firstPageNumber'];
        if (length(s) > 0) then
          if (TryStrToInt(s, _t)) then
            _currSheet.SheetOptions.StartPageNumber := _t;
            
        //s := xml.Attributes.ItemsByName['fitToHeight'];
        //s := xml.Attributes.ItemsByName['fitToWidth'];
        //s := xml.Attributes.ItemsByName['horizontalDpi'];
        //s := xml.Attributes.ItemsByName['id'];
        s := xml.Attributes.ItemsByName['orientation'];
        if (length(s) > 0) then
        begin
          _currSheet.SheetOptions.PortraitOrientation := false;
          if (s = 'portrait') then
            _currSheet.SheetOptions.PortraitOrientation := true;
        end;
        
        //s := xml.Attributes.ItemsByName['pageOrder'];

        s := xml.Attributes.ItemsByName['paperSize'];
        if (length(s) > 0) then
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
      if ((xml.TagName = 'printOptions') and (xml.TagType = 5)) then
      begin
        //s := xml.Attributes.ItemsByName['gridLines'];
        //s := xml.Attributes.ItemsByName['gridLinesSet'];
        //s := xml.Attributes.ItemsByName['headings'];
        s := xml.Attributes.ItemsByName['horizontalCentered'];
        if (length(s) > 0) then
          _currSheet.SheetOptions.CenterHorizontal := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['verticalCentered'];
        if (length(s) > 0) then
          _currSheet.SheetOptions.CenterVertical := ZEStrToBoolean(s);
      end else
      if ((xml.TagName = 'dimension') and (xml.TagType = 5)) then
        _GetDimension()
      else
      if ((xml.TagName = 'hyperlinks') and (xml.TagType = 4)) then
        _ReadHyperLinks()
      else
      if ((xml.TagName = 'sheetViews') and (xml.TagType = 4)) then
        _ReadSheetViews();

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
//RETURN
//      boolean - true - стили прочитались без ошибок
function ZEXSLXReadStyles(var XMLSS: TZEXMLSS; var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer): boolean;
type

  TZXLSXBorderItem = record
    color: TColor;
    isColor: boolean;
    isEnabled: boolean;
    style: TZBorderType;
    width: integer;
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
  end;

  TZXLSXFillArray = array of TZXLSXFill;

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
      border[i].width := 0;
    end;
  end; //ZEXLSXZeroBorder

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

  procedure _ReadFonts();
  var
    _currFont: integer;

  begin
    _currFont := -1;
    while (not ((xml.TagName = 'fonts') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof) then
        break;

      s := xml.Attributes.ItemsByName['val'];
      if ((xml.TagName = 'font') and (xml.TagType = 4)) then
      begin
        _currFont := FontCount;
        inc(FontCount);
        SetLength(FontArray, FontCount);
        ZEXLSXZeroFont(FontArray[_currFont]);
      end else
      if (_currFont >= 0) then
      begin
        if ((xml.TagName = 'name') and (xml.TagType = 5)) then
          FontArray[_currFont].name := s
        else
        if ((xml.TagName = 'b') and (xml.TagType = 5)) then
          FontArray[_currFont].bold := ZEStrToBoolean(s)
        else
        if ((xml.TagName = 'charset') and (xml.TagType = 5)) then
        begin
          if (TryStrToInt(s, t)) then
            FontArray[_currFont].charset := t;
        end else
        if ((xml.TagName = 'color') and (xml.TagType = 5)) then
        begin
          s := xml.Attributes.ItemsByName['rgb'];
          if (length(s) > 2) then
            delete(s, 1, 2);
          if (length(s) > 0) then
            FontArray[_currFont].color := HTMLHexToColor(s);
        end else
        if ((xml.TagName = 'i') and (xml.TagType = 5)) then
          FontArray[_currFont].italic := ZEStrToBoolean(s)
        else
  //      if ((xml.TagName = 'outline') and (xml.TagType = 5)) then
  //      begin
  //      end else
        if ((xml.TagName = 'strike') and (xml.TagType = 5)) then
          FontArray[_currFont].strike := ZEStrToBoolean(s)
        else
        if ((xml.TagName = 'sz') and (xml.TagType = 5)) then
        begin
          if (TryStrToInt(s, t)) then
            FontArray[_currFont].fontsize := t;
        end else
        if ((xml.TagName = 'u') and (xml.TagType = 5)) then
        begin
          if (length(s) > 0) then
            if (s <> 'none') then
              FontArray[_currFont].underline := true;
        end
        else
        if ((xml.TagName = 'vertAlign') and (xml.TagType = 5)) then
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

  procedure _ReadBorders();
  var
    _diagDown, _diagUP: boolean;
    _currBorder: integer; //текущий набор границ
    _currBorderItem: integer; //текущая граница (левая/правая ...)
    _color: TColor;
    _isColor: boolean;

    procedure _SetCurBorder(borderNum: integer);
    var
      b: boolean;

    begin
      _currBorderItem := borderNum;
      s := xml.Attributes.ItemsByName['style'];
      if (Length(s) > 0) then
      begin
        b := true;
        BorderArray[_currBorder][borderNum].width := 1;
        if (s = 'thin') then
        begin
          BorderArray[_currBorder][borderNum].style := ZEContinuous;
        end else
        if (s = 'hair') then
          BorderArray[_currBorder][borderNum].style := ZEContinuous
        else
        if (s = 'dashed') then
          BorderArray[_currBorder][borderNum].style := ZEDash
        else
        if (s = 'dotted') then
          BorderArray[_currBorder][borderNum].style := ZEDot
        else
        if (s = 'dashDot') then
          BorderArray[_currBorder][borderNum].style := ZEDashDot
        else
        if (s = 'dashDotDot') then
          BorderArray[_currBorder][borderNum].style := ZEDashDotDot
        else
        if (s = 'slantDashDot') then
          BorderArray[_currBorder][borderNum].style := ZESlantDashDot
        else
        if (s = 'double') then
          BorderArray[_currBorder][borderNum].style := ZEDouble
        else
        if (s = 'medium') then
        begin
          BorderArray[_currBorder][borderNum].style := ZEContinuous;
          BorderArray[_currBorder][borderNum].width := 2;
        end else
        if (s = 'thick') then
        begin
          BorderArray[_currBorder][borderNum].style := ZEContinuous;
          BorderArray[_currBorder][borderNum].width := 3;
        end else
        if (s = 'mediumDashed') then
        begin
          BorderArray[_currBorder][borderNum].style := ZEDash;
          BorderArray[_currBorder][borderNum].width := 2;
        end else
        if (s = 'mediumDashDot') then
        begin
          BorderArray[_currBorder][borderNum].style := ZEDashDot;
          BorderArray[_currBorder][borderNum].width := 2;
        end else
        if (s = 'mediumDashDotDot') then
        begin
          BorderArray[_currBorder][borderNum].style := ZEDashDotDot;
          BorderArray[_currBorder][borderNum].width := 2;
        end else
        if (s = 'none') then
          BorderArray[_currBorder][borderNum].style := ZENone
        else
          b := false;

        BorderArray[_currBorder][borderNum].isEnabled := b;
      end;
    end; //_SetCurBorder

  begin
    _currBorderItem := -1;
    _diagDown := false;
    _diagUP := false;
    _color := clBlack;
    while (not ((xml.TagName = 'borders') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'border') and (xml.TagType = 4)) then
      begin
        _currBorder := BorderCount;
        inc(BorderCount);
        SetLength(BorderArray, BorderCount);
        ZEXLSXZeroBorder(BorderArray[_currBorder]);
        _diagDown := false;
        _diagUP := false;
        s := xml.Attributes.ItemsByName['diagonalDown'];
        if (length(s) > 0) then
          _diagDown := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['diagonalUp'];
        if (length(s) > 0) then
          _diagUP := ZEStrToBoolean(s);
      end else
      begin
        if (xml.TagType in [4, 5]) then
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
            if (length(s) > 0) then
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
    _l: integer;
    _b: byte;
    _t: TColor;

  begin
    _currFill := -1;
    while (not ((xml.TagName = 'fills') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'fill') and (xml.TagType = 4)) then
      begin
        _currFill := FillCount;
        inc(FillCount);
        SetLength(FillArray, FillCount);
        FillArray[_currFill].patternfill := ZPNone;
        FillArray[_currFill].bgcolor := clWindow;
        FillArray[_currFill].patterncolor := clWindow;
        FillArray[_currFill].bgColorType := 0;
        FillArray[_currFill].patternColorType := 0;
      end else
      if ((xml.TagName = 'patternFill') and (xml.TagType in [4, 5])) then
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

          if (length(s) > 0) then
          begin
            if (s = 'solid') then
              FillArray[_currFill].patternfill := ZPSolid
            else
            if (s = 'none') then
              FillArray[_currFill].patternfill := ZPNone
            else
            if (s = 'gray125') then
              FillArray[_currFill].patternfill := ZPGray125
            else
            if (s = 'gray0625') then
              FillArray[_currFill].patternfill := ZPGray0625
            else
            if (s = 'darkUp') then
              FillArray[_currFill].patternfill := ZPDiagStripe
            else
            if (s = 'mediumGray') then
              FillArray[_currFill].patternfill := ZPGray50
            else
            if (s = 'darkGray') then
              FillArray[_currFill].patternfill := ZPGray75
            else
            if (s = 'lightGray') then
              FillArray[_currFill].patternfill := ZPGray25
            else
            if (s = 'darkHorizontal') then
              FillArray[_currFill].patternfill := ZPHorzStripe
            else
            if (s = 'darkVertical') then
              FillArray[_currFill].patternfill := ZPVertStripe
            else
            if (s = 'darkDown') then
              FillArray[_currFill].patternfill := ZPReverseDiagStripe
            else
            if (s = 'darkUpDark') then
              FillArray[_currFill].patternfill := ZPDiagStripe
            else
            if (s = 'darkGrid') then
              FillArray[_currFill].patternfill := ZPDiagCross
            else
            if (s = 'darkTrellis') then
              FillArray[_currFill].patternfill := ZPThickDiagCross
            else
            if (s = 'lightHorizontal') then
              FillArray[_currFill].patternfill := ZPThinHorzStripe
            else
            if (s = 'lightVertical') then
              FillArray[_currFill].patternfill := ZPThinVertStripe
            else
            if (s = 'lightDown') then
              FillArray[_currFill].patternfill := ZPThinReverseDiagStripe
            else
            if (s = 'lightUp') then
              FillArray[_currFill].patternfill := ZPThinDiagStripe
            else
            if (s = 'lightGrid') then
              FillArray[_currFill].patternfill := ZPThinHorzCross
            else
            if (s = 'lightTrellis') then
              FillArray[_currFill].patternfill := ZPThinDiagCross
            //
            else
              FillArray[_currFill].patternfill := ZPSolid; //{tut} потом подумать насчёт стилей границ
          end;
        end;
      end else
      if ((xml.TagName = 'bgColor') and (xml.TagType = 5)) then
      begin
        if (_currFill >= 0) then
        begin
          s := xml.Attributes.ItemsByName['rgb'];
          if (length(s) > 2) then
          begin
            delete(s, 1, 2);
            if (length(s) > 0) then
              FillArray[_currFill].patterncolor := HTMLHexToColor(s);
          end;
          s := xml.Attributes.ItemsByName['theme'];
          if (length(s) > 0) then
          begin
            if (TryStrToInt(s, _l)) then
            begin
              FillArray[_currFill].patterncolor := _l;
              FillArray[_currFill].patternColorType := 2;
            end;
          end;

          //если не сплошная заливка - нужно поменять местами цвета (bgColor <-> fgColor)
          if (not (FillArray[_currFill].patternfill in [ZPNone, ZPSolid])) then
          begin
            _t := FillArray[_currFill].patterncolor;
            FillArray[_currFill].patterncolor := FillArray[_currFill].bgcolor;
            FillArray[_currFill].bgColor := _t;

            _b := FillArray[_currFill].patternColorType;
            FillArray[_currFill].patternColorType := FillArray[_currFill].bgColorType;
            FillArray[_currFill].bgColorType := _b;
          end; //if

        end;
      end else
      if ((xml.TagName = 'fgColor') and (xml.TagType = 5)) then
      begin
        if (_currFill >= 0) then
        begin
          s := xml.Attributes.ItemsByName['rgb'];
          if (length(s) > 2) then
          begin
            delete(s, 1, 2);
            if (length(s) > 0) then
              FillArray[_currFill].bgcolor := HTMLHexToColor(s);
          end;
          s := xml.Attributes.ItemsByName['theme'];
          if (length(s) > 0) then
          begin
            if (TryStrToInt(s, _l)) then
            begin
              FillArray[_currFill].bgColorType := 2;
              FillArray[_currFill].bgcolor := _l;
            end;  
          end;
        end;
      end;
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
    while (not ((xml.TagName = TagName) and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      b := false;  

      if ((xml.TagName = 'xf') and (xml.TagType in [4, 5])) then
      begin
        _currCell := StyleCount;
        inc(StyleCount);
        SetLength(CSA, StyleCount);
        ZEXLSXZeroCellStyle(CSA[_currCell]);
        s := xml.Attributes.ItemsByName['applyAlignment'];
        if (length(s) > 0) then
          CSA[_currCell].applyAlignment := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['applyBorder'];
        if (length(s) > 0) then
          CSA[_currCell].applyBorder := ZEStrToBoolean(s)
        else
          b := true;

        s := xml.Attributes.ItemsByName['applyFont'];
        if (length(s) > 0) then
          CSA[_currCell].applyFont := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['applyProtection'];
        if (length(s) > 0) then
          CSA[_currCell].applyProtection := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['borderId'];
        if (length(s) > 0) then
          if (TryStrToInt(s, t)) then
          begin
            CSA[_currCell].borderId := t;
            if (b and (t >= 1)) then
              CSA[_currCell].applyBorder := true;
          end;

        s := xml.Attributes.ItemsByName['fillId'];
        if (length(s) > 0) then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].fillId := t;

        s := xml.Attributes.ItemsByName['fontId'];
        if (length(s) > 0) then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].fontId := t;

        s := xml.Attributes.ItemsByName['numFmtId'];
        if (length(s) > 0) then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].numFmtId := t;

        s := xml.Attributes.ItemsByName['xfId'];
        if (length(s) > 0) then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].xfId := t;
      end else
      if ((xml.TagName = 'alignment') and (xml.TagType = 5)) then
      begin
        if (_currCell >= 0) then
        begin
          s := xml.Attributes.ItemsByName['horizontal'];
          if (length(s) > 0) then
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
          if (length(s) > 0) then
            if (TryStrToInt(s, t)) then
              CSA[_currCell].alignment.indent := t;

          s := xml.Attributes.ItemsByName['shrinkToFit'];
          if (length(s) > 0) then
            CSA[_currCell].alignment.shrinkToFit := ZEStrToBoolean(s);

          s := xml.Attributes.ItemsByName['textRotation'];
          if (length(s) > 0) then
            if (TryStrToInt(s, t)) then
              CSA[_currCell].alignment.textRotation := t;

          s := xml.Attributes.ItemsByName['vertical'];
          if (length(s) > 0) then
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
          if (length(s) > 0) then
            CSA[_currCell].alignment.wrapText := ZEStrToBoolean(s);
        end; //if
      end else
      if ((xml.TagName = 'protection') and (xml.TagType = 5)) then
      begin
        if (_currCell >= 0) then
        begin
          s := xml.Attributes.ItemsByName['hidden'];
          if (length(s) > 0) then
            CSA[_currCell].hidden := ZEStrToBoolean(s);

          s := xml.Attributes.ItemsByName['locked'];
          if (length(s) > 0) then
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
    while (not ((xml.TagName = 'cellStyles') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'cellStyle') and (xml.TagType = 5)) then
      begin
        b := false;
        SetLength(StyleArray, StyleCount + 1);
        s := xml.Attributes.ItemsByName['builtinId']; //?
        if (length(s) > 0) then
          if (TryStrToInt(s, t)) then
            StyleArray[StyleCount].builtinId := t;

        s := xml.Attributes.ItemsByName['customBuiltin']; //?
        if (length(s) > 0) then
          StyleArray[StyleCount].customBuiltin := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['name']; //?
          StyleArray[StyleCount].name := s;

        s := xml.Attributes.ItemsByName['xfId'];
        if (length(s) > 0) then
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
    while (not ((xml.TagName = 'colors') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'rgbColor') and (xml.TagType = 5)) then
      begin
        s := xml.Attributes.ItemsByName['rgb'];
        if (length(s) > 2) then
          delete(s, 1, 2);
        if (length(s) > 0) then
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

  //Применить стиль
  //INPUT
  //  var XMLSSStyle: TZStyle         - стиль в хранилище
  //  var XLSXStyle: TZXLSXCellStyle  - стиль в xlsx
  procedure _ApplyStyle(var XMLSSStyle: TZStyle; var XLSXStyle: TZXLSXCellStyle);
  var
    i: integer;

  begin
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
          XMLSSStyle.Border[i].Weight := BorderArray[n][i].width;
          if (BorderArray[n][i].isColor) then
            XMLSSStyle.Border[i].Color := BorderArray[n][i].color;
        end;
    end;

    if (XLSXStyle.applyFont) then
    begin
      n := XLSXStyle.fontId;
      if ((n >= 0) and (n < FontCount)) then
      begin
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

      if ((xml.TagName = 'fonts') and (xml.TagType = 4)) then
        _ReadFonts()
      else
      if ((xml.TagName = 'borders') and (xml.TagType = 4)) then
        _ReadBorders()
      else
      if ((xml.TagName = 'fills') and (xml.TagType = 4)) then
        _ReadFills()
      else
      //TODO: разобраться, чем отличаются cellStyleXfs и cellXfs. Что за cellStyles?
      if ((xml.TagName = 'cellStyleXfs') and (xml.TagType = 4)) then
        _ReadCellCommonStyles('cellStyleXfs', CellStyleArray, CellStyleCount)//_ReadCellStyleXfs()
      else
      if ((xml.TagName = 'cellXfs') and (xml.TagType = 4)) then  //сами стили?
        _ReadCellCommonStyles('cellXfs', CellXfsArray, CellXfsCount) //_ReadCellXfs()
      else
      if ((xml.TagName = 'cellStyles') and (xml.TagType = 4)) then //??
        _ReadCellStyles()
      else
      if ((xml.TagName = 'colors') and (xml.TagType = 4)) then
        _ReadColors();
    end; //while

    //тут незабыть применить номера цветов, если были введены

    //
    for i := 0 to FillCount - 1 do
    begin
      if (FillArray[i].bgColorType = 2) then
      begin
        t := FillArray[i].bgcolor;
        if ((t >= 0) and (t < ThemaColorCount)) then
          FillArray[i].bgcolor := ThemaFillsColors[t];
      end;
      if (FillArray[i].patternColorType = 2) then
      begin
        t := FillArray[i].patterncolor;
        if ((t >= 0) and (t < ThemaColorCount)) then
          FillArray[i].patterncolor := ThemaFillsColors[t];
      end;
    end; //for

    //{tut}

    XMLSS.Styles.Count := CellXfsCount;
    for i := 0 to CellXfsCount - 1 do
    begin
      t := CellXfsArray[i].xfId;
      _Style := XMLSS.Styles[i];
      if ((t >= 0) and (t < CellStyleCount)) then
        _ApplyStyle(_Style, CellStyleArray[t]);
      _ApplyStyle(_Style, CellXfsArray[i]);
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
    Setlength(indexedColor, 0);
    indexedColor := nil;
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

      if ((xml.TagName = 'sheet') and (xml.TagType = 5)) then
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
      if ((xml.TagName = 'workbookView') and (xml.TagType = 5)) then
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

      if ((xml.TagName = 'Relationship') and (xml.TagType = 5)) then
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
    while (not((xml.TagName = 'authors') and (xml.TagType = 6))) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if ((xml.TagName = 'author') and (xml.TagType = 6)) then
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
    if (length(s) = 0) then
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
      while (not((xml.TagName = 'comment') and (xml.TagType = 6))) do
      begin
        xml.ReadTag();
        if (xml.Eof()) then
          break;
        if ((xml.TagName = 't') and (xml.TagType = 6)) then
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

      if ((xml.TagName = 'authors') and (xml.TagType = 4)) then
        _ReadAuthors()
      else
      if ((xml.TagName = 'comment') and (xml.TagType = 4)) then
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
        if (not ZEXSLXReadStyles(XMLSS, stream, ThemaColor, ThemaColorCount)) then
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

        for i := 1 to RelationsCounts[SheetRelationNumber] do
        for j := 0 to RelationsCounts[SheetRelationNumber] - 1 do
        if (RelationsArray[SheetRelationNumber][j].sheetid = i) then
        begin
          b := _CheckSheetRelations(FileArray[RelationsArray[SheetRelationNumber][j].fileid].name);
          FreeAndNil(stream);
          stream := TFileStream.Create(DirName + FileArray[RelationsArray[SheetRelationNumber][j].fileid].name, fmOpenRead or fmShareDenyNone);
          if (not ZEXSLXReadSheet(XMLSS, stream, RelationsArray[SheetRelationNumber][j].name, StrArray, StrCount, SheetRelations, SheetRelationsCount)) then
            result := result or 4;
          if (b) then
            _ReadComments();
          _no_sheets := false;
          break;
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
        if (not ZEXSLXReadSheet(XMLSS, stream, '', StrArray, StrCount, SheetRelations, SheetRelationsCount)) then
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
  end;
end; //ReadXLSXPath

{$IFDEF FPC}
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
    if (length(s1) > 0) then
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

      u_zip.OnCreateStream := @ZH.DoCreateOutZipStream;
      u_zip.OnDoneStream := @ZH.DoDoneOutZipStream;
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

        for i := 1 to ZH.RelationsCounts[ZH.SheetRelationNumber] do
        for j := 0 to ZH.RelationsCounts[ZH.SheetRelationNumber] - 1 do
        if (ZH.RelationsArray[ZH.SheetRelationNumber][j].sheetid = i) then
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
{$ENDIF}

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
//RETURN
//      integer
function ZEXLSXCreateContentTypes(var XMLSS: TZEXMLSS; Stream: TStream; PageCount: integer; CommentCount: integer; const PagesComments: TIntegerDynArray;
                                  TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
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
    end;
    _xml.Attributes.Add('ContentType', s, false);
    _xml.WriteEmptyTag('Override', true);
  end; //_WriteOverride

  procedure _WriteTypes();
  var
    i: integer;

  begin
    //Страницы
    _WriteOverride('/_rels/.rels', 3);
    _WriteOverride('/xl/_rels/workbook.xml.rels', 3);
    for i := 0 to PageCount - 1 do
      _WriteOverride('/xl/worksheets/sheet' + IntToStr(i + 1) + '.xml', 0);
    //комментарии
    for i := 0 to CommentCount - 1 do
    begin
      _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(PagesComments[i] + 1) + '.xml.rels', 3);
      _WriteOverride('/xl/comments' + IntToStr(PagesComments[i] + 1) + '.xml', 6);
    end;
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
    _xml.Attributes.Add('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');
    _xml.WriteTagNode('Types', true, true, true);
    _WriteTypes();
    _xml.WriteEndTagNode(); //Types
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ZEXLSXCreateContentTypes

//Создаёт лист документа (sheet*.xml)
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//    SheetNum: integer                 - номер листа в документе
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//  var isHaveComments: boolean         - возвращает true, если были комментарии (чтобы создать comments*.xml)
//RETURN
//      integer
function ZEXLSXCreateSheet(var XMLSS: TZEXMLSS; Stream: TStream; SheetNum: integer; TextConverter: TAnsiToCPConverter;
                                     CodePageName: String; BOM: ansistring; out isHaveComments: boolean): integer;
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
    _xml.Attributes.Add('ref', 'A1:' + ZEGetA1byCol(_sheet.ColCount - 1) + IntToStr(_sheet.RowCount));
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
    _xml.Attributes.Add('tabSelected', 'true', false);
    _xml.Attributes.Add('topLeftCell', 'A1', false);
    _xml.Attributes.Add('view', 'normal', false);
    _xml.Attributes.Add('windowProtection', 'false', false);
    _xml.Attributes.Add('workbookViewId', '0', false);
    _xml.Attributes.Add('zoomScale', '100', false);
    _xml.Attributes.Add('zoomScaleNormal', '100', false);
    _xml.Attributes.Add('zoomScalePageLayoutView', '100', false);
    _xml.WriteTagNode('sheetView', true, true, false);

    _SOptions := XMLSS.Sheets[SheetNum].SheetOptions;

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
    
    s := ZEGetA1byCol(XMLSS.Sheets[SheetNum].SheetOptions.ActiveCol) + IntToSTr(XMLSS.Sheets[SheetNum].SheetOptions.ActiveRow + 1);
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
        if (i + 1 <= 1025) then
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
        if (not isHaveComments) then
          if (length(_sheet.Cell[j, i].Comment) > 0) then
            isHaveComments := true;
        b := (length(_sheet.Cell[j, i].Data) > 0) or
             (length(_sheet.Cell[j, i].Formula) > 0);
        _xml.Attributes.Add('r', ZEGetA1byCol(j) + IntToStr(i + 1));
        if (_sheet.Cell[j, i].CellStyle >= -1) and (_sheet.Cell[j, i].CellStyle < XMLSS.Styles.Count) then
          s := IntToStr(_sheet.Cell[j, i].CellStyle + 1)
        else
          s := '0';
        _xml.Attributes.Add('s', s, false);

        case _sheet.Cell[j, i].CellType of
          ZENumber: s := 'n';
          ZEDateTime: s := 'str'; //??
          ZEBoolean: s := 'b';
          ZEString: s := 'str';
          ZEError: s := 'e';
        end;
        
        _xml.Attributes.Add('t', s, false);
        if (b) then
        begin
          _xml.WriteTagNode('c', true, true, false);
          if (length(_sheet.Cell[j, i].Formula) > 0) then
          begin
            _xml.Attributes.Clear();
            _xml.Attributes.Add('aca', 'false');
            _xml.WriteTag('f', _sheet.Cell[j, i].Formula, true, false, true);
          end;
          if (length(_sheet.Cell[j, i].Data) > 0) then
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
    _xml.Attributes.Add('header', ZEFloatSeparator(FormatFloat(s, _sheet.SheetOptions.HeaderMargin / ZE_MMinInch)), false);
    _xml.Attributes.Add('footer', ZEFloatSeparator(FormatFloat(s, _sheet.SheetOptions.FooterMargin / ZE_MMinInch)), false);
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
        IfThen(_sheet.SheetOptions.PortraitOrientation, 'portrait','landscape'),
        false);

    _xml.Attributes.Add('pageOrder', 'downThenOver', false);
    _xml.Attributes.Add('paperSize', intToStr(_sheet.SheetOptions.PaperSize), false);
    _xml.Attributes.Add('scale', '100', false);
    _xml.Attributes.Add('useFirstPageNumber', 'true', false);
    _xml.Attributes.Add('usePrinterDefaults', 'false', false);
    _xml.Attributes.Add('verticalDpi', '300', false);
    _xml.WriteEmptyTag('pageSetup', true, false);
    //  <headerFooter differentFirst="false" differentOddEven="false">
    //    <oddHeader> ... </oddHeader>
    //    <oddFooter> ... </oddFooter>
    //  </headerFooter>
    //  <legacyDrawing r:id="..."/>
  end; //WriteXLSXSheetFooter

begin
  isHaveComments := false;
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
    _xml.Attributes.Add('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
    _xml.Attributes.Add('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
    _xml.WriteTagNode('worksheet', true, true, false);

    _sheet := XMLSS.Sheets[SheetNum];
    WriteXLSXSheetHeader();
    WriteXLSXSheetCols();
    WriteXLSXSheetData();
    WriteXLSXSheetFooter();

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
    _xml.Attributes.Add('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
    _xml.Attributes.Add('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', false);
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
    _xml.Attributes.Add('tabRatio', '200', false);
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
  procedure _UpdateArrayIndex(var arr: TIntegerDynArray; cnt: integer); deprecated {$IFDEF USE_DEPRECATED_STRING}'remove CNT parameter!'{$ENDIF};
  var
    res: TIntegerDynArray;
    i, j: integer;
    num: integer;
  begin
    Assert( Length(arr) - cnt = 2, 'Wow! We really may need this parameter!');
    cnt := Length(arr) - 2;   // get ready to strip it
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
          //thin ??
          if (_border.Weight = 1) then
            s1 := 'hair'
          else
          if (_border.Weight = 2) then
            s1 := 'medium'
          else
            s1 := 'thick';
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
    s: string; i: integer;

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
    _xml.Attributes.Add('numFmtId', '0' {'164'}, false); // TODO: support formats

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
    _StylesCount := XMLSS.Styles.Count;

    _xml.Attributes.Clear();
    _xml.Attributes.Add('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
    _xml.WriteTagNode('styleSheet', true, true, true);

    WriteXLSXFonts();
    WriteXLSXFills();
    WriteXLSXBorders();
//    WriteCellStyleXfs('cellStyleXfs', false);
//    WriteCellStyleXfs('cellXfs', true);

    // experiment: do not need styles, ergo do not need fake xfId
    WriteCellStyleXfs('cellXfs', false);

    WriteCellStyles(); //??

    _xml.WriteEndTagNode(); //styleSheet
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
    SetLength(_FontIndex, 0);
    _FontIndex := nil;
    SetLength(_FillIndex, 0);
    _FillIndex := nil;
    SetLength(_BorderIndex, 0);
    _BorderIndex := nil;
  end;
end; //ZEXLSXCreateStyles

//Добавить Relationship для rels
//INPUT
//      xml: TZsspXMLWriterH  - писалка
//  const rid: string         - rid
//      ridType: integer      -
//  const Target: string      -
procedure ZEAddRelsRelation(xml: TZsspXMLWriterH; const rid: string; ridType: integer; const Target: string);
begin
  xml.Attributes.Clear();
  xml.Attributes.Add('Id', rid);
  xml.Attributes.Add('Type',  ZEXLSXGetRelationName(ridType), false);
  xml.Attributes.Add('Target', Target, false);
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
    _xml.Attributes.Add('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
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
    _xml.Attributes.Add('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
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
    _xml.Attributes.Add('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', false);
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
    _xml.Attributes.Add('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
    _xml.Attributes.Add('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes', false);
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
    _xml.Attributes.Add('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
    _xml.Attributes.Add('xmlns:dc', 'http://purl.org/dc/elements/1.1/', false);
    _xml.Attributes.Add('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/', false);
    _xml.Attributes.Add('xmlns:dcterms', 'http://purl.org/dc/terms/', false);
    _xml.Attributes.Add('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance', false);
    _xml.WriteTagNode('cp:coreProperties', true, true, false);

    _xml.Attributes.Clear();
    _xml.Attributes.Add('xsi:type', 'dcterms:W3CDTF');
    s := ZEDateToStr(XMLSS.DocumentProperties.Created) + 'Z';
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
function ExportXmlssToXLSX(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string;
                         BOM: ansistring = '';
                         AllowUnzippedFolder: boolean = false;
                         ZipGenerator: CZxZipGens = nil): integer;
var
  _pages: TIntegerDynArray;      //номера страниц
  _names: TStringDynArray;      //названия страниц
  kol, i: integer;
  Stream: TStream;
  need_comments: boolean;
  path_xl, path_sheets, path_relsmain, path_relsw, path_docprops: string;
  _commentArray: TIntegerDynArray;

  azg: TZxZipGen; // Actual Zip Generator

begin
  Result := 0;
  Stream := nil;
  kol := 0;
  need_comments := false;

  azg := nil;
  try
    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      result := 2;
      exit;
    end;

    if nil = ZipGenerator then begin
      ZipGenerator := TZxZipGen.QueryZipGen;
      if nil = ZipGenerator then
        if AllowUnzippedFolder
           then ZipGenerator := TZxZipGen.QueryDummyZipGen
           else raise EZxZipGen.Create('No zip generators registered, folder output disabled.');
           // result := 3 ????
    end;
    azg := ZipGenerator.Create(PathName);

//    if (not ZE_CheckDirExist(PathName)) then
//    begin
//      result := 3;
//      exit;
//    end;

//    path_xl := PathName + 'xl' + PathDelim;
//    if (not DirectoryExists(path_xl)) then
//      ForceDirectories(path_xl);

    path_xl := 'xl' + PathDelim;
    //стили
    Stream := azg.NewStream(path_xl + 'styles.xml');
    ZEXLSXCreateStyles(XMLSS, Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);   Stream := nil;
//    FreeAndNil(Stream);

    //sharedStrings.xml
    Stream := azg.NewStream(path_xl + 'sharedStrings.xml');
    ZEXLSXCreateSharedStrings(XMLSS, Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);    Stream := nil;
//    FreeAndNil(Stream);

    // _rels/.rels
    path_relsmain := {PathName + PathDelim +} '_rels' + PathDelim;
//    if (not DirectoryExists(path_relsmain)) then
//      ForceDirectories(path_relsmain);
    Stream := azg.NewStream(path_relsmain + '.rels');
    ZEXLSXCreateRelsMain(Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);   Stream := nil;

    // xl/_rels/workbook.xml.rels 
    path_relsw := path_xl + {PathDelim +} '_rels' + PathDelim;
//    if (not DirectoryExists(path_relsw)) then
//      ForceDirectories(path_relsw);
    Stream := azg.NewStream(path_relsw + 'workbook.xml.rels');
    ZEXLSXCreateRelsWorkBook(kol, Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);   Stream := nil;

    path_sheets := path_xl + 'worksheets' + PathDelim;
//    if (not DirectoryExists(path_sheets)) then
//      ForceDirectories(path_sheets);

    SetLength(_commentArray, kol);

    //Листы документа
    for i := 0 to kol - 1 do
    begin
      _commentArray[i] := 0;
      Stream := azg.NewStream(path_sheets + 'sheet' + IntToStr(i + 1) + '.xml');
      ZEXLSXCreateSheet(XMLSS, Stream, _pages[i], TextConverter, CodePageName, BOM, need_comments);
      if (need_comments) then
        _commentArray[i] := 1;
      azg.SealStream(Stream);   Stream := nil;

      if (need_comments) then
      begin
        _commentArray[i] := 1;
        //создать файл с комментариями
      end;  
    end; //for i

    //workbook.xml - список листов
    Stream := azg.NewStream(path_xl + 'workbook.xml');
    ZEXLSXCreateWorkBook(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);  Stream := nil;

    //[Content_Types].xml
    Stream := azg.NewStream({PathName +} '[Content_Types].xml');
    ZEXLSXCreateContentTypes(XMLSS, Stream, kol, 0, nil, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);  Stream := nil;

    path_docprops := {PathName +} 'docProps' + PathDelim;
//    if (not DirectoryExists(path_docprops)) then
//      ForceDirectories(path_docprops);

    // docProps/app.xml
    Stream := azg.NewStream(path_docprops + 'app.xml');
    ZEXLSXCreateDocPropsApp(Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);  Stream := nil;

    // docProps/core.xml
    Stream := azg.NewStream(path_docprops + 'core.xml');
    ZEXLSXCreateDocPropsCore(XMLSS, Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream); Stream := nil;

    azg.SaveAndSeal;
  finally
    ZESClearArrays(_pages, _names);
    if (Assigned(Stream)) then
      FreeAndNil(Stream);
    SetLength(_commentArray, 0);
    _commentArray := nil;
    azg.Free;
  end;
end; //SaveXmlssToXLSXPath     //Сохраняет незапакованный документ в формате Office Open XML (OOXML)
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
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer;
var
  _pages: TIntegerDynArray;      //номера страниц
  _names: TStringDynArray;      //названия страниц
  kol, i: integer;
  Stream: TStream;
  need_comments: boolean;
  path_xl, path_sheets, path_relsmain, path_relsw, path_docprops: string;
  _commentArray: TIntegerDynArray;

begin
  Result := 0;
  Stream := nil;
  kol := 0;
  need_comments := false;
  try
    if (not ZE_CheckDirExist(PathName)) then
    begin
      result := 3;
      exit;
    end;

    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      result := 2;
      exit;
    end;

    path_xl := PathName + 'xl' + PathDelim;
    if (not DirectoryExists(path_xl)) then
      ForceDirectories(path_xl);

    //стили
    Stream := TFileStream.Create(path_xl + 'styles.xml', fmCreate);
    ZEXLSXCreateStyles(XMLSS, Stream, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    //sharedStrings.xml
    Stream := TFileStream.Create(path_xl + 'sharedStrings.xml', fmCreate);
    ZEXLSXCreateSharedStrings(XMLSS, Stream, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    // _rels/.rels
    path_relsmain := PathName + PathDelim + '_rels' + PathDelim;
    if (not DirectoryExists(path_relsmain)) then
      ForceDirectories(path_relsmain);
    Stream := TFileStream.Create(path_relsmain + '.rels', fmCreate);
    ZEXLSXCreateRelsMain(Stream, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    // xl/_rels/workbook.xml.rels 
    path_relsw := path_xl + '_rels' + PathDelim;
    if (not DirectoryExists(path_relsw)) then
      ForceDirectories(path_relsw);
    Stream := TFileStream.Create(path_relsw + 'workbook.xml.rels', fmCreate);
    ZEXLSXCreateRelsWorkBook(kol, Stream, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    path_sheets := path_xl + 'worksheets' + PathDelim;
    if (not DirectoryExists(path_sheets)) then
      ForceDirectories(path_sheets);

    SetLength(_commentArray, kol);

    //Листы документа
    for i := 0 to kol - 1 do
    begin
      _commentArray[i] := 0;
      Stream := TFileStream.Create(path_sheets + 'sheet' + IntToStr(i + 1) + '.xml', fmCreate);
      ZEXLSXCreateSheet(XMLSS, Stream, _pages[i], TextConverter, CodePageName, BOM, need_comments);
      if (need_comments) then
        _commentArray[i] := 1;
      FreeAndNil(Stream);

      if (need_comments) then
      begin
        _commentArray[i] := 1;
        //создать файл с комментариями
      end;  
    end; //for i

    //workbook.xml - список листов
    Stream := TFileStream.Create(path_xl + 'workbook.xml', fmCreate);
    ZEXLSXCreateWorkBook(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    //[Content_Types].xml
    Stream := TFileStream.Create(PathName + '[Content_Types].xml', fmCreate);
    ZEXLSXCreateContentTypes(XMLSS, Stream, kol, 0, nil, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    path_docprops := PathName + 'docProps' + PathDelim;
    if (not DirectoryExists(path_docprops)) then
      ForceDirectories(path_docprops);

    // docProps/app.xml
    Stream := TFileStream.Create(path_docprops + 'app.xml', fmCreate);
    ZEXLSXCreateDocPropsApp(Stream, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    // docProps/core.xml
    Stream := TFileStream.Create(path_docprops + 'core.xml', fmCreate);
    ZEXLSXCreateDocPropsCore(XMLSS, Stream, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

  finally
    ZESClearArrays(_pages, _names);
    if (Assigned(Stream)) then
      FreeAndNil(Stream);
    SetLength(_commentArray, 0);
    _commentArray := nil;
  end;
end; //SaveXmlssToXLSXPath

{$IFDEF FPC}
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
var
  _pages: TIntegerDynArray;      //номера страниц
  _names: TStringDynArray;      //названия страниц
  kol, i: integer;
  zip: TZipper;
  Stream: array of TStream;
  StreamCount: integer;
  MaxStreamCount: integer;
  need_comments: boolean;
  path_xl, path_sheets, path_relsmain, path_relsw, path_docprops: string;
  _commentArray: TIntegerDynArray;

  procedure _AddFile(const _fname: string);
  begin
    Stream[StreamCount - 1].Position := 0;
    zip.Entries.AddFileEntry(Stream[StreamCount - 1], _fname);
  end;

  procedure _AddStream();
  var
    i: integer;

  begin
    inc(StreamCount);
    if (StreamCount >= MaxStreamCount) then
    begin
      MaxStreamCount := StreamCount + 100;
      SetLength(Stream, MaxStreamCount);
      for i := StreamCount - 1 to MaxStreamCount - 1 do
        Stream[i] := nil;
    end;
    Stream[StreamCount - 1] := TMemoryStream.Create();
  end;

begin
  zip := nil;
  StreamCount := 0;
  MaxStreamCount := -1;
  try

    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      result := 2;
      exit;
    end;

    zip := TZipper.Create();
    zip.FileName := FileName;

    path_xl := 'xl/';

    //стили
    _AddStream();
    ZEXLSXCreateStyles(XMLSS, Stream[StreamCount - 1], TextConverter, CodePageName, BOM);
    _AddFile(path_xl + 'styles.xml');

    //sharedStrings.xml
    _AddStream();
    ZEXLSXCreateSharedStrings(XMLSS, Stream[StreamCount - 1], TextConverter, CodePageName, BOM);
    _AddFile(path_xl + 'sharedStrings.xml');

    // _rels/.rels
    path_relsmain := '_rels/';
    _AddStream();
    ZEXLSXCreateRelsMain(Stream[StreamCount - 1], TextConverter, CodePageName, BOM);
    _AddFile(path_relsmain + '.rels');

    // xl/_rels/workbook.xml.rels
    path_relsw := path_xl + '_rels/';
    _AddStream();
    ZEXLSXCreateRelsWorkBook(kol, Stream[StreamCount - 1], TextConverter, CodePageName, BOM);
    _AddFile(path_relsw + 'workbook.xml.rels');

    path_sheets := path_xl + 'worksheets/';
    SetLength(_commentArray, kol);

    //Листы документа
    for i := 0 to kol - 1 do
    begin
      _commentArray[i] := 0;
      _AddStream();
      ZEXLSXCreateSheet(XMLSS, Stream[StreamCount - 1], _pages[i], TextConverter, CodePageName, BOM, need_comments);
      if (need_comments) then
        _commentArray[i] := 1;
      _AddFile(path_sheets + 'sheet' + IntToStr(i + 1) + '.xml');

      if (need_comments) then
      begin
        _commentArray[i] := 1;
        //создать файл с комментариями
      end;
    end; //for i

    //workbook.xml - список листов
    _AddStream();
    ZEXLSXCreateWorkBook(XMLSS, Stream[StreamCount - 1], _pages, _names, kol, TextConverter, CodePageName, BOM);
    _AddFile(path_xl + 'workbook.xml');

    //[Content_Types].xml
    _AddStream();
    ZEXLSXCreateContentTypes(XMLSS, Stream[StreamCount - 1], kol, 0, nil, TextConverter, CodePageName, BOM);
    _AddFile('[Content_Types].xml');

    path_docprops :='docProps/';
    // docProps/app.xml
    _AddStream();
    ZEXLSXCreateDocPropsApp(Stream[StreamCount - 1], TextConverter, CodePageName, BOM);
    _AddFile(path_docprops + 'app.xml');

    // docProps/core.xml
    _AddStream();
    ZEXLSXCreateDocPropsCore(XMLSS, Stream[StreamCount - 1], TextConverter, CodePageName, BOM);
    _AddFile(path_docprops + 'core.xml');

    zip.ZipAllFiles();
  finally
    ZESClearArrays(_pages, _names);
    if (Assigned(zip)) then
      FreeAndNil(zip);

    SetLength(_commentArray, 0);
    _commentArray := nil;

    for i := 0 to StreamCount - 1 do
    if (Assigned(Stream[i])) then
      FreeAndNil(Stream[i]);
    SetLength(Stream, 0);
    Stream := nil;
  end;
  Result := 0;
end; //SaveXmlssToXSLX

{$ENDIF}

//Перепутал малость названия ^_^

function ReadXSLXPath(var XMLSS: TZEXMLSS; DirName: string): integer; //deprecated 'Use ReadXLSXPath!';
begin
  result := ReadXLSXPath(XMLSS, DirName);
end;

{$IFDEF FPC}
function ReadXSLX(var XMLSS: TZEXMLSS; FileName: string): integer; //deprecated 'Use ReadXLSX!';
begin
  result := ReadXLSX(XMLSS, FileName);
end;
{$ENDIF}

{$IFNDEF FPC}
{$I xlsxzipfuncimpl.inc}
{$ENDIF}

end.
