//****************************************************************
// Read/write Open Document Format (spreadsheet)
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2013.11.05
//----------------------------------------------------------------
// Modified by the_Arioch@nm.ru - added uniform save API
//     to create ODS in Delphi/Windows
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
//****************************************************************


unit zeodfs;

interface

{$I zexml.inc}
{$I compver.inc}

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

uses
  SysUtils, Graphics, Classes, Types, 
  zsspxml, zexmlss, zesavecommon, zeZippy
  {$IFDEF FPC},zipper {$ELSE}{$I odszipuses.inc}{$ENDIF};

//Сохраняет незапакованный документ в формате Open Document
function SaveXmlssToODFSPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;

{$IFDEF FPC}
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
{$ENDIF}

function ExportXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                           const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: String;
                           BOM: ansistring = '';
                           AllowUnzippedFolder: boolean = false; ZipGenerator: CZxZipGens = nil): integer; overload;


function ReadODFSPath(var XMLSS: TZEXMLSS; DirName: string): integer;

{$IFDEF FPC}
function ReadODFS(var XMLSS: TZEXMLSS; FileName: string): integer;
{$ENDIF}

{$IFNDEF FPC}
{$I odszipfunc.inc}
{$ENDIF}

//////////////////// Дополнительные функции, если понадобится читать/писать отдельные файлы или ещё для чего
{для записи}
//Записывает в поток стили документа (styles.xml)
function ODFCreateStyles(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;

//Записывает в поток настройки (settings.xml)
function ODFCreateSettings(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;

//Записывает в поток документ + автоматические стили (content.xml)
function ODFCreateContent(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;

//Записывает в поток метаинформацию (meta.xml)
function ODFCreateMeta(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;

{для чтения}
//Чтение содержимого документа ODS (content.xml)
function ReadODFContent(var XMLSS: TZEXMLSS; stream: TStream): boolean;

//Чтение настроек документа ODS (settings.xml)
function ReadODFSettings(var XMLSS: TZEXMLSS; stream: TStream): boolean;

implementation
uses StrUtils;

const
  ZETag_StyleFontFace       = 'style:font-face';      //style:font-face
  ZETag_Attr_StyleName      = 'style:name';           //style:name

type
  TZODFColumnStyle = record
    name: string;     //имя стиля строки
    width: real;      //ширина
    breaked: boolean; //разрыв
  end;

  TZODFColumnStyleArray = array of TZODFColumnStyle;

  TZODFRowStyle = record
    name: string;
    height: real;
    breaked: boolean;
    color: TColor;
  end;

  TZODFRowStyleArray = array of TZODFRowStyle;

  TZODFStyle = record
    name: string; 
    index: integer;
  end;

  TZODFStyleArray = array of TZODFStyle;

  TZODFTableStyle = record
    name: string;
    isColor: boolean;
    Color: TColor;
  end;

  TZODFTableArray = array of TZODFTableStyle;

{$IFDEF FPC}
  //Для распаковки в поток
  TODFZipHelper = class
  private
    FXMLSS: TZEXMLSS;
    FRetCode: integer;
    FFileType: integer;
  protected
  public
    constructor Create(); virtual;
    procedure DoCreateOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
    procedure DoDoneOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
    property XMLSS: TZEXMLSS read FXMLSS write FXMLSS;
    property RetCode: integer read FRetCode;
    property FileType: integer read FFileType write FFileType;
  end;

constructor TODFZipHelper.Create();
begin
  inherited;
  FXMLSS := nil;
  FRetCode := 0;
end;

procedure TODFZipHelper.DoCreateOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
begin
  AStream := TMemoryStream.Create();
end;

procedure TODFZipHelper.DoDoneOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
begin
  if (Assigned(AStream)) then
  try
    AStream.Position := 0;

    if (FileType = 0) then
    begin
      if (not ReadODFContent(FXMLSS, AStream)) then
        FRetCode := FRetCode or 2;
    end else
    if (FileType = 1) then
    begin
      if (not ReadODFSettings(FXMLSS, AStream)) then
        FRetCode := FRetCode or 2;
    end;
  finally
    FreeAndNil(AStream)
  end;
end; //DoDoneOutZipStream

{$ENDIF}

//BooleanToStr для ODF //TODO: потом заменить
function ODFBoolToStr(value: boolean): string;
begin
  if (value) then
    result := 'true'
  else
    result := 'false';
end;

//Переводит тип значения ODF в нужный
function ODFTypeToZCellType(const value: string): TZCellType;
var
  s: string;

begin
  s := UpperCase(value);
  if (s = 'FLOAT') then
    result := ZENumber
  else
  if (s = 'PERCENTAGE') then
    result := ZENumber
  else
  if (s = 'CURRENCY') then
    result := ZENumber
  else
  if (s = 'DATE') then
    result := ZEDateTime
  else
  if (s = 'TIME') then
    result := ZEDateTime
  else
  if (s = 'BOOLEAN') then
    result := ZEBoolean
  else
  if (s = 'STRING') then
    result := ZEString
  else
    result := ZEString;
end; //ODFTypeToZCellType


//Добавляет атрибуты для тэга office:document-content
procedure GenODContentAttr(Attr: TZAttributesH);
begin
  Attr.Add('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', false);
  Attr.Add('xmlns:style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0', false);
  Attr.Add('xmlns:text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0', false);
  Attr.Add('xmlns:table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0', false);
  Attr.Add('xmlns:draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', false);
  Attr.Add('xmlns:fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', false);
  Attr.Add('xmlns:xlink', 'http://www.w3.org/1999/xlink', false);
  Attr.Add('xmlns:dc', 'http://purl.org/dc/elements/1.1/', false);
  Attr.Add('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0', false);
  Attr.Add('xmlns:number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0', false);
  Attr.Add('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0', false);
  Attr.Add('xmlns:svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', false);
  Attr.Add('xmlns:chart', 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0', false);
  Attr.Add('xmlns:dr3d', 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0', false);
  Attr.Add('xmlns:math', 'http://www.w3.org/1998/Math/MathML', false);
  Attr.Add('xmlns:form', 'urn:oasis:names:tc:opendocument:xmlns:form:1.0', false);
  Attr.Add('xmlns:script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0', false);
  Attr.Add('xmlns:ooo', 'http://openoffice.org/2004/office', false);
  Attr.Add('xmlns:ooow', 'http://openoffice.org/2004/writer', false);
  Attr.Add('xmlns:oooc', 'http://openoffice.org/2004/calc', false);
  Attr.Add('xmlns:dom', 'http://www.w3.org/2001/xml-events', false);
  Attr.Add('xmlns:xforms', 'http://www.w3.org/2002/xforms', false);
  Attr.Add('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema', false);
  Attr.Add('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance', false);
  Attr.Add('xmlns:rpt', 'http://openoffice.org/2005/report', false);
  Attr.Add('xmlns:of', 'urn:oasis:names:tc:opendocument:xmlns:of:1.2', false);
  Attr.Add('xmlns:xhtml', 'http://www.w3.org/1999/xhtml', false);
  Attr.Add('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#', false);
  Attr.Add('xmlns:tableooo', 'http://openoffice.org/2009/table', false);
  Attr.Add('xmlns:field', 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0', false);
  Attr.Add('xmlns:formx', 'urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0', false);
  Attr.Add('xmlns:css3t', 'http://www.w3.org/TR/css3-text/', false);
  Attr.Add('office:version', '1.2', false);
end; //GenODContentAttr

//добавляет атрибуты для тэга office:document-meta
procedure GenODMetaAttr(Attr: TZAttributesH);
begin
  Attr.Add('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', false);
  Attr.Add('xmlns:xlink', 'http://www.w3.org/1999/xlink', false);
  Attr.Add('xmlns:dc', 'http://purl.org/dc/elements/1.1/', false);
  Attr.Add('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0', false);
  Attr.Add('xmlns:ooo', 'http://openoffice.org/2004/office', false);
  Attr.Add('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#', false);
  Attr.Add('office:version', '1.2', false);
end; //GenODMetaAttr

//добавляет атрибуты для тэга office:document-styles (styles.xml)
procedure GenODStylesAttr(Attr: TZAttributesH);
begin
  Attr.Add('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
  Attr.Add('xmlns:style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0');
  Attr.Add('xmlns:text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
  Attr.Add('xmlns:table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0');
  Attr.Add('xmlns:draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0');
  Attr.Add('xmlns:fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0');
  Attr.Add('xmlns:xlink', 'http://www.w3.org/1999/xlink');
  Attr.Add('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
  Attr.Add('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
  Attr.Add('xmlns:number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0');
  Attr.Add('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0');
  Attr.Add('xmlns:svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0');
  Attr.Add('xmlns:chart', 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0');
  Attr.Add('xmlns:dr3d', 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0');
  Attr.Add('xmlns:math', 'http://www.w3.org/1998/Math/MathML');
  Attr.Add('xmlns:form', 'urn:oasis:names:tc:opendocument:xmlns:form:1.0');
  Attr.Add('xmlns:script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0');
  Attr.Add('xmlns:ooo', 'http://openoffice.org/2004/office');
  Attr.Add('xmlns:ooow', 'http://openoffice.org/2004/writer'); //???
  Attr.Add('xmlns:oooc', 'http://openoffice.org/2004/calc');
  Attr.Add('xmlns:dom', 'http://www.w3.org/2001/xml-events');
  Attr.Add('xmlns:rpt', 'http://openoffice.org/2005/report');
  Attr.Add('xmlns:of', 'urn:oasis:names:tc:opendocument:xmlns:of:1.2');
  Attr.Add('xmlns:xhtml', 'http://www.w3.org/1999/xhtml');
  Attr.Add('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#');
  Attr.Add('xmlns:tableooo', 'http://openoffice.org/2009/table');
  Attr.Add('xmlns:css3t', 'http://www.w3.org/TR/css3-text/');
  Attr.Add('office:version', '1.2');
end; //GenODStylesAttr

//<office:font-face-decls> ... </office:font-face-decls>
procedure ZEWriteFontFaceDecls(XMLSS: TZEXMLSS; _xml: TZsspXMLWriterH);
var
  kol, maxkol: integer;
  _fn: array of string;

  procedure _addFonts();
  var
    i, j: integer;
    s: string;
    b: boolean;
    _name: string;

  begin
    for i := -1 to XMLSS.Styles.Count - 1 do
    begin
      _name := XMLSS.Styles[i].Font.Name;
      s := UpperCase(_name);
      b := true;
      for j := 0 to kol - 1 do
        if (_fn[j] = s) then
        begin
          b := false;
          break;
        end;
      if (b) then
      begin
        _xml.Attributes.ItemsByNum[0] := _name;
        if (pos(' ', _name) = 0) then
          _xml.Attributes.ItemsByNum[1] := _name
        else
          _xml.Attributes.ItemsByNum[1] := '''' + _name + '''';

        _xml.WriteEmptyTag(ZETag_StyleFontFace, true, true);
        inc(kol);
        if (kol + 1 >= maxkol) then
        begin
          inc(maxkol, 10);
          SetLength(_fn, maxkol);
        end;
        _fn[kol - 1] := s;
      end; //if
    end; //for
  end; //_addFonts

begin
  maxkol := 10;
  SetLength(_fn, maxkol);
  kol := 3;
  _fn[0] := 'ARIAL';
  _fn[1] := 'MANGAL';
  _fn[2] := 'TAHOMA';
  try
    _xml.WriteTagNode('office:font-face-decls', true, true, true);
    _xml.Attributes.Add(ZETag_Attr_StyleName, 'Arial', true);
    _xml.Attributes.Add('svg:font-family', 'Arial', true);
    _xml.Attributes.Add('style:font-family-generic', 'swiss', true);
    _xml.Attributes.Add('style:font-pitch', 'variable', true);
    _xml.WriteEmptyTag(ZETag_StyleFontFace, true, false);

    _xml.Attributes.ItemsByNum[0] := 'Mangal';
    _xml.Attributes.ItemsByNum[1] := 'Mangal';
    _xml.Attributes.ItemsByNum[2] := 'system';
    _xml.WriteEmptyTag(ZETag_StyleFontFace, true, false);

    _xml.Attributes.ItemsByNum[0] := 'Tahoma';
    _xml.Attributes.ItemsByNum[1] := 'Tahoma';
    _xml.WriteEmptyTag(ZETag_StyleFontFace, true, false);

    _addFonts();

    _xml.Attributes.Clear();
    _xml.WriteEndTagNode(); //office:font-face-decls
  finally
    SetLength(_fn, 0);
    _fn := nil;
  end;
end; //WriteFontFaceDecls

//Получит нужный цвет для цвета фона
//TODO: нужно будет восстанавливать цвет по названию
//INPUT
//  const value: string - цвет фона
//RETURN
//      TColor - цвет фона (clwindow = transparent)
function GetBGColorForODS(const value: string): TColor;
var
  l: integer;

begin
  l := length(value);
  result := 0;
  if (l >= 1) then
    if (value[1] = '#') then
      result := HTMLHexToColor(value)
    else
    begin
      if (value = 'transparent') then
        result := clWindow
      //добавить нужные цвета
    end;
end; //GetBGColorForODS

//Переводит стиль границы в строку для ODF
//INPUT
//      BStyle: TZBorderStyle - стиль границы
function ZEODFBorderStyleTostr(BStyle: TZBorderStyle): string;
var
  s: string;

begin
  result := '';
  if (Assigned(BStyle)) then
  begin
    //не забыть потом толщину подправить {tut}
    case (BStyle.Weight) of
      0: result := '0pt';
      1: result := '0.26pt';
      2: result := '2.49pt';
      3: result := '4pt';
      else
       result := '0.99pt';
    end;
    s := '';
    case (BStyle.LineStyle) of
      ZENone: s := ' ';
      ZEContinuous: s := ' solid';
      ZEDot: s := ' dotted';
      ZEDash: s := ' dashed';
      ZEDashDot: s := ' ';
      ZEDashDotDot: s := ' ';
      ZESlantDashDot: s := ' ';
      ZEDouble: s := ' double';
    end;
    result := result + s + ' #' + ColorToHTMLHex(BStyle.Color);
  end;
end; //ZEODFBorderStyleTostr

//Выдирает из строки параметры стиля границы
//INPUT
//  const st: string          - строка с параметрами
//  BStyle: TZBorderStyle - стиль границы
procedure ZEStrToODFBorderStyle(const st: string; BStyle: TZBorderStyle);
var
  i: integer;
  s: string;

  procedure _CheckStr();
  begin
    if (s > '') then
    begin
      if (s[1] = '#') then
        BStyle.Color := HTMLHexToColor(s)
      else
      {$IFDEF DELPHI_UNICODE}
      if (CharInSet(s[1], ['0'..'9'])) then
      {$ELSE}
      if (s[1] in ['0'..'9']) then //толщина
      {$ENDIF}
      begin
        if (s = '0pt') then
          BStyle.Weight := 0
        else
        if (s = '0.26pt') then
          BStyle.Weight := 1
        else
        if (s = '2.49pt') then
          BStyle.Weight := 2
        else
        if (s = '4pt') then
          BStyle.Weight := 3
        else
          BStyle.Weight := 1;
      end else
      begin
        if (s = 'solid') then
          BStyle.LineStyle := ZEContinuous
        else
        if (s = 'dotted') then
          BStyle.LineStyle := ZEDot
        else
        if (s = 'dashed') then
          BStyle.LineStyle := ZEDash
        else
        if (s = 'double') then
          BStyle.LineStyle := ZEDouble
        else
          BStyle.LineStyle := ZENone;
      end;
      s := '';
    end;
  end; //_CheckStr

begin
  if (Assigned(BStyle)) then
  begin
    s := '';
    for i := 1 to length(st) do
    if (st[i] = ' ') then
      _CheckStr()
    else
      s := s + st[i];
    _CheckStr();  
  end;
end; //ZEStrToODFBorderStyle

//Записывает настройки стиля
//INPUT
//  var XMLSS: TZEXMLSS           - хранилище
//      _xml: TZsspXMLWriterH     - писатель
//      StyleNum: integer         - номер стиля
//      isDefaultStyle: boolean   - является-ли данный стиль стилем по-умолчанию
procedure ODFWriteTableStyle(var XMLSS: TZEXMLSS; _xml: TZsspXMLWriterH; const StyleNum: integer; isDefaultStyle: boolean);
var
  b: boolean;
  s, satt: string;
  j, n: integer;
  ProcessedStyle: TZStyle;
  ProcessedFont: TFont;

begin
{
     <attribute name="style:family"><value>table-cell</value>
     Доступны тэги:
        style:table-cell-properties
        style:paragraph-properties
        style:text-properties
}

  //Тэг style:table-cell-properties
  //Возможные атрибуты:
  //    style:vertical-align - выравниванние по вертикали (top | middle | bottom | automatic)
  //??  style:text-align-source - источник выравнивания текста (fix | value-type)
  //??  style:direction - направление символов в ячейке (ltr | ttb) слева-направо и сверху-вниз
  //??  style:glyph-orientation-vertical - ориентация глифа по вертикали
  //??  style:shadow - применяется эффект тени
  //    fo:background-color - цвет фона ячейки
  //    fo:border           - [
  //    fo:border-top       -
  //    fo:border-bottom    -   обрамление ячейки
  //    fo:border-left      -
  //    fo:border-right     -  ]
  //    style:diagonal-tl-br - диагональ верхний левый правый нижний
  //       style:diagonal-bl-tr-widths
  //    style:diagonal-bl-tr - диагональ нижний левый правый верхний угол
  //       style:diagonal-tl-br-widths
  //    style:border-line-width         -  [
  //    style:border-line-width-top     -
  //    style:border-line-width-bottom  -   толщина линии обрамления
  //    style:border-line-width-left    -
  //    style:border-line-width-right   -  ]
  //    fo:padding          - [
  //    fo:padding-top      -
  //    fo:padding-bottom   -  отступы
  //    fo:padding-left     -
  //    fo:padding-right    - ]
  //    fo:wrap-option  - свойство переноса по словам (no-wrap | wrap)
  //    style:rotation-angle - угол поворота (int >= 0)
  //??  style:rotation-align - выравнивание после поворота (none | bottom | top | center)
  //??  style:cell-protect - (none | hidden­and­protected ?? protected | formula­hidden)
  //??  style:print-content - выводить ли на печать содержимое ячейки (bool)
  //??  style:decimal-places - кол-во дробных разрядов
  //??  style:repeat-content - повторять-ли содержимое ячейки (bool)
  //    style:shrink-to-fit - подгонять ли содержимое по размеру, если текст не помещается (bool)

  _xml.Attributes.Clear();
  ProcessedStyle := XMLSS.Styles[StyleNum];

  b := true;
  for j := 1 to 3 do
    if (not ProcessedStyle.Border[j].IsEqual(ProcessedStyle.Border[0])) then
    begin
      b := false;
      break;
    end;

  n := 0;
  if (b) then
  begin
    n := 4;
    if (not((ProcessedStyle.Border[0].LineStyle = ZEContinuous) and (ProcessedStyle.Border[0].Weight = 0))) then
    begin
      s := ZEODFBorderStyleTostr(ProcessedStyle.Border[0]);
      _xml.Attributes.Add('fo:border', s);
    end;
  end;

  //TODO: по умолчанию no-wrap?
  if (ProcessedStyle.Alignment.WrapText) then
    _xml.Attributes.Add('fo:wrap-option', 'wrap');

  for j := n to 5 do
  begin
    case (j) of
      0: satt := 'fo:border-left';
      1: satt := 'fo:border-top';
      2: satt := 'fo:border-right';
      3: satt := 'fo:border-bottom';
      4: satt := 'style:diagonal-bl-tr';
      5: satt := 'style:diagonal-tl-br';
    end;
    if (not((ProcessedStyle.Border[j].LineStyle = ZEContinuous) and (ProcessedStyle.Border[j].Weight = 0))) then
    begin
      s := ZEODFBorderStyleTostr(ProcessedStyle.Border[j]);
      _xml.Attributes.Add(satt, s, false);
    end;
  end;

  //Выравнивание по вертикали
// при загрузке XML в Excel 2010  ZVJustify почти везде (кроме ZHFill) рисуется вверху,
// a ZV[Justified]Distributed рисуются в по центру - но это для одного "слова"
  case (ProcessedStyle.Alignment.Vertical) of
    ZVAutomatic: s := 'automatic';
    ZVTop, ZVJustify: s := 'top';
    ZVBottom: s := 'bottom';
    ZVCenter, ZVDistributed, ZVJustifyDistributed: s := 'middle';
    //(top | middle | bottom | automatic)
  end;
  _xml.Attributes.Add('style:vertical-align', s);

//??  style:text-align-source - источник выравнивания текста (fix | value-type)
  _xml.Attributes.Add('style:text-align-source',
      IfThen( ZHAutomatic = ProcessedStyle.Alignment.Horizontal,
              'value-type', 'fix') );

   if ZHFill = ProcessedStyle.Alignment.Horizontal then
     _xml.Attributes.Add('style:repeat-content', ODFBoolToStr(true), false);

  //Поворот текста
  _xml.Attributes.Add('style:direction',
       IfThen(ProcessedStyle.Alignment.VerticalText, 'ttb', 'ltr'), false);
  if (ProcessedStyle.Alignment.Rotate <> 0) then
    _xml.Attributes.Add('style:rotation-angle',
        IntToStr( ProcessedStyle.Alignment.Rotate mod 360 ));

  //Цвет фона ячейки
  if ((ProcessedStyle.BGColor <> XMLSS.Styles.DefaultStyle.BGColor) or (isDefaultStyle)) then
    if (ProcessedStyle.BGColor <> clWindow) then
      _xml.Attributes.Add('fo:background-color', '#' + ColorToHTMLHex(ProcessedStyle.BGColor));

  //style:shrink-to-fit
  if (ProcessedStyle.Alignment.ShrinkToFit) then
    _xml.Attributes.Add('style:shrink-to-fit', ODFBoolToStr(true));

  _xml.WriteEmptyTag('style:table-cell-properties', true, true);

  //*************
  //Тэг style-paragraph-properties
  //Возможные атрибуты:
  //??  fo:line-height - фиксированная высота строки
  //??  style:line-height-at-least - минимальная высотка строки
  //??  style:line-spacing - межстрочный интервал
  //??  style:font-independent-line-spacing - независимый от шрифта межстрочный интервал (bool)
  //    fo:text-align - выравнивание текста (start | end | left | right | center | justify)
  //??  fo:text-align-last - выравнивание текста в последней строке (start | center | justify)
  //??  style:justify-single-word - выравнивать-ли последнее слово (bool)
  //??  fo:keep-together - не разрывать (auto | always)
  //??  fo:widows - int > 0
  //??  fo:orphans - int > 0
  //??  style:tab-stop-distance - int > 0
  //??  fo:hyphenation-keep
  //??  fo:hyphenation-ladder-count

  // ZHFill -> style:repeat-content  ?
  if ZHAutomatic <> ProcessedStyle.Alignment.Horizontal then begin
    _xml.Attributes.Clear();
    //Выравнивание по горизонтали
    case (ProcessedStyle.Alignment.Horizontal) of
      //ZHAutomatic: s := 'start';//'right';
      ZHLeft: s := 'start';//'left';
      ZHCenter: s := 'center';
      ZHRight: s := 'end';//'right';
      ZHFill: s := 'start';
      ZHJustify: s := 'justify';
      ZHCenterAcrossSelection: s := 'center';
      ZHDistributed: s := 'center';
      ZHJustifyDistributed: s := 'justify';
    end;
    _xml.Attributes.Add('fo:text-align', s);
    _xml.WriteEmptyTag('style:paragraph-properties', true, true);
  end;

  //*************
  //Тэг style:text-properties
  //Возможные атрибуты:
  //??  fo:font-variant - отображение текста прописныим буквами (normal | small-caps). Взаимоисключается с fo:text-transform
  //??  fo:text-transform - преобразование текста (none | lowercase | uppercase | capitalize)
  //    fo:color - цвет переднего плана текста
  //    fo:text-indent - первая строка параграфа, с единицами изм.
  //??  style:use-window-font-color - должен ли цвет переднего плана окна быть цветом фона для светлого фона и белый для тёмного фона (bool)
  //??  style:text-outline - показывать ли структуру текста или текст (bool)
  //    style:text-line-through-type - тип линии зачёркивания текста (none | single | double)
  //    style:text-line-through-style - (none | single | double) ??
  //??  style:text-line-through-width - толщина зачёркивания
  //??  style:text-line-through-color - цвет зачёркивания
  //??  style:text-line-through-text
  //??  style:text-line-through-text-style
  //??  style:text-position - (super | sub ?? percent)
  //     style:font-name          - [
  //     style:font-name-asian    - название шрифта
  //     style:font-name-complex  - ]
  //      fo:font-family            - [
  //      style:font-family-asian   -  семейство шрифтов
  //      style:font-family-complex - ]
  //      style:font-family-generic         - [
  //      style:font-family-generic-asian   - Группа семейства шрифтов (roman | swiss | modern | decorative | script | system)
  //      style:font-family-generic-complex - ]
  //    style:font-style-name         - [
  //    style:font-style-name-asian   - стиль шрифта
  //    style:font-style-name-complex - ]
  //??  style:font-pitch          - [
  //??  style:font-pitch-asian    - шаг шрифта (fixed | variable)
  //??  style:font-pitch-complex  - ]
  //??  style:font-charset          - [
  //??  style:font-charset-asian    - набор символов
  //??  style:font-charset-complex  - ]
  //    fo:font-size            - [
  //    style:font-size-asian   -  размер шрифта
  //    style:font-size-complex - ]
  //??  style:font-size-rel         - [
  //??  style:font-size-rel-asian   - масштаб шрифта
  //??  style:font-size-rel-complex - ]
  //??  style:script-type
  //??  fo:letter-spacing - межбуквенный интервал
  //??  fo:language         - [
  //??  fo:language-asian   - код языка
  //??  fo:language-complex - ]
  //??  fo:country            - [
  //??  style:country-asian   - код страны
  //??  style:country-complex - ]
  //    fo:font-style             - [
  //    style:font-style-asian    - стиль шрифта (normal | italic | oblique)
  //    style:font-style-complex  - ]
  //??  style:font-relief - рельефтность (выпуклый, высеченый, плоский) (none | embossed | engraved)
  //??  fo:text-shadow - тень
  //    style:text-underline-type - тип подчёркивания (none | single | double)
  //??  style:text-underline-style - стиль подчёркивания (none | solid | dotted | dash | long-dash | dot-dash | dot-dot-dash | wave)
  //??  style:text-underline-width - толщина подчёркивания (auto | norma | bold | thin | dash | medium | thick ?? int>0)
  //??  style:text-underline-color - цвет подчёркивания
  //    fo:font-weight            - [
  //    style:font-weight-asian   - жирность (normal | bold | 100 | 200 | 300 | 400 | 500 | 600 | 700 | 800 | 900)
  //    style:font-weight-complex - ]
  //??  style:text-underline-mode - режим подчёркивания слов (continuous | skip-white-space)
  //??  style:text-line-through-mode - режим зачёркивания слов (continuous | skip-white-space)
  //??  style:letter-kerning - кернинг (bool)
  //??  style:text-blinking - мигание текста (bool)
  //**  fo:background-color - цвет фона текста
  //??  style:text-combine - объединение текста (none | letters | lines)
  //??  style:text-combine-start-char
  //??  style:text-combine-end-char
  //??  *tyle:text-emphasize - вроде как для иероглифов выделение (none | accent | dot | circle | disc) + (above | below) (пример: "dot above")
  //??  style:text-scale - масштаб
  //??  style:text-rotation-angle - угол вращения текста (0 | 90 | 270)
  //??  style:text-rotation-scale - масштабирование при вращении (fixed | line-height)
  //??  fo:hyphenate - расстановка переносов
  //    text:display - показывать-ли текст (true - да, none - скрыть, condition - в зависимости от атрибута text:condition)
  _xml.Attributes.Clear();
  ProcessedFont := ProcessedStyle.Font;

  //style:font-name
  if ((ProcessedFont.Name <> XMLSS.Styles.DefaultStyle.Font.Name) or (isDefaultStyle)) then
  begin
    s := ProcessedFont.Name;
    _xml.Attributes.Add('style:font-name', s);
    _xml.Attributes.Add('style:font-name-asian', s, false);
    _xml.Attributes.Add('style:font-name-complex', s, false);
  end;

  //размер шрифта
  if ((ProcessedFont.Size <> XMLSS.Styles.DefaultStyle.Font.Size) or (isDefaultStyle)) then
  begin
    s := IntToStr(ProcessedFont.Size) + 'pt';
    _xml.Attributes.Add('fo:font-size', s, false);
    _xml.Attributes.Add('style:font-size-asian', s, false);
    _xml.Attributes.Add('style:font-size-complex', s, false);
  end;

  //Жирность
  if (fsBold in ProcessedFont.Style) then
  begin
    s := 'bold';
    _xml.Attributes.Add('fo:font-weight', s, false);
    _xml.Attributes.Add('style:font-weight-asian', s, false);
    _xml.Attributes.Add('style:font-weight-complex', s, false);
  end;

  //перечёркнутый текст
  if (fsStrikeOut in ProcessedFont.Style) then
    _xml.Attributes.Add('style:text-line-through-type', 'single', false); //(none | single | double)

  //Подчёркнутый текст
  if (fsUnderline in ProcessedFont.Style) then
    _xml.Attributes.Add('style:text-underline-type', 'single', false); //(none | single | double)

  //цвет шрифта
  if ((ProcessedFont.Color <> XMLSS.Styles.DefaultStyle.Font.Color) or (isDefaultStyle)) then
    _xml.Attributes.Add('fo:color', '#' + ColorToHTMLHex(ProcessedFont.Color), false);

  if (fsItalic in ProcessedFont.Style) then
  begin
    s := 'italic';
    _xml.Attributes.Add('fo:font-style', s, false);
    _xml.Attributes.Add('style:font-style-asian', s, false);
    _xml.Attributes.Add('style:font-style-complex', s, false);
  end;

  _xml.WriteEmptyTag('style:text-properties', true, true);
end; //ODFWriteTableStyle

//Записывает в поток стили документа (styles.xml)
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
function ODFCreateStyles(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
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
    GenODStylesAttr(_xml.Attributes);
    _xml.WriteTagNode('office:document-styles', true, true, true);
    _xml.Attributes.Clear();
    ZEWriteFontFaceDecls(XMLSS, _xml);

    //office:styles
    _xml.Attributes.Clear();
    _xml.WriteTagNode('office:styles', true, true, true);

    //Стиль по-умолчанию
    _xml.Attributes.Clear();
    _xml.Attributes.Add(ZETag_Attr_StyleName, 'Default');
    _xml.Attributes.Add('style:family', 'table-cell', false);
    _xml.WriteTagNode('style:style', true, true, true);
    ODFWriteTableStyle(XMLSS, _xml, -1, true);
    _xml.WriteEndTagNode();

    _xml.WriteEndTagNode(); //office:styles

    _xml.WriteEndTagNode(); //office:document-styles
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ODFCreateStyles

//Записывает в поток настройки (settings.xml)
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//  const _pages: TIntegerDynArray      - массив страниц
//  const _names: TStringDynArray       - массив имён страниц
//    PageCount: integer                - кол-во страниц
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ODFCreateSettings(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;
  i: integer;

  //<config:config-item config:name="ConfigName" config:type="ConfigType">ConfigValue</config:config-item>
  procedure _AddConfigItem(const ConfigName, ConfigType, ConfigValue: string);
  begin
    _xml.Attributes.Clear();
    _xml.Attributes.Add('config:name', ConfigName);
    _xml.Attributes.Add('config:type', ConfigType);
    _xml.WriteTag('config:config-item', ConfigValue, true, false, true);
  end; //_AddConfigItem

  procedure _WriteSplitValue(const SPlitMode: TZSplitMode; const SplitValue: integer; const SplitModeName, SplitValueName: string; NeedAdd: boolean = false);
  var
    s: string;

  begin
    if ({(SplitMode <> ZSplitNone) and} (SplitValue <> 0) or (NeedAdd)) then
    begin
      s := '0';
      case (SPlitMode) of
        ZSplitFrozen: s := '2';
        ZSplitSplit: s := '1';
      end;
      _AddConfigItem(SplitModeName, 'short', s);
      _AddConfigItem(SplitValueName, 'int', IntToStr(SplitValue));
    end;
  end; //_WriteSplitValue

  procedure _WritePageSettings(const num: integer);
  var
    _PageNum: integer;
    _SheetOptions: TZSheetOptions;
    b: boolean;

  begin
    _PageNum := _pages[num];
    _xml.Attributes.Clear();
    _xml.Attributes.Add('config:name', _names[num]);
    _xml.WriteTagNode('config:config-item-map-entry', true, true, true);
    _SheetOptions := XMLSS.Sheets[_PageNum].SheetOptions;

    _AddConfigItem('CursorPositionX', 'int', IntToStr(_SheetOptions.ActiveCol));
    _AddConfigItem('CursorPositionY', 'int', IntToStr(_SheetOptions.ActiveRow));

    b := (_SheetOptions.SplitHorizontalMode = ZSplitSplit) or
         (_SheetOptions.SplitHorizontalMode = ZSplitSplit);
    //это не ошибка (_SheetOptions.SplitHorizontalMode = VerticalSplitMode)
    _WriteSplitValue(_SheetOptions.SplitHorizontalMode, _SheetOptions.SplitHorizontalValue, 'VerticalSplitMode', 'VerticalSplitPosition', b);
    _WriteSplitValue(_SheetOptions.SplitVerticalMode, _SheetOptions.SplitVerticalValue, 'HorizontalSplitMode', 'HorizontalSplitPosition', b);

    _AddConfigItem('ActiveSplitRange', 'short', '2');
    _AddConfigItem('PositionLeft', 'int', '0');
    _AddConfigItem('PositionRight', 'int', '1');
    _AddConfigItem('PositionTop', 'int', '0');
    _AddConfigItem('PositionBottom', 'int', '1');

    _xml.WriteEndTagNode(); //config:config-item-map-entry
  end; //_WritePageSettings

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
    _xml.Attributes.Add('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
    _xml.Attributes.Add('xmlns:xlink', 'http://www.w3.org/1999/xlink', false);
    _xml.Attributes.Add('xmlns:config', 'urn:oasis:names:tc:opendocument:xmlns:config:1.0', false);
    _xml.Attributes.Add('xmlns:ooo', 'http://openoffice.org/2004/office', false);
    _xml.Attributes.Add('office:version', '1.2', false);
    _xml.WriteTagNode('office:document-settings', true, true, true);
    _xml.Attributes.Clear();
    _xml.WriteTagNode('office:settings', true, true, true);

    _xml.Attributes.Add('config:name', 'ooo:view-settings');
    _xml.WriteTagNode('config:config-item-set', true, true, false);

    _AddConfigItem('VisibleAreaTop', 'int', '0');
    _AddConfigItem('VisibleAreaLeft', 'int', '0');
    _AddConfigItem('VisibleAreaWidth', 'int', '6773');
    _AddConfigItem('VisibleAreaHeight', 'int', '1813');

    _xml.Attributes.Clear();
    _xml.Attributes.Add('config:name', 'Views');
    _xml.WriteTagNode('config:config-item-map-indexed', true, true, false);

    _xml.Attributes.Clear();
    _xml.WriteTagNode('config:config-item-map-entry', true, true, false);

    _xml.Attributes.Add('config:name', 'Tables');
    _xml.WriteTagNode('config:config-item-map-named', true, true, false);

    for i := 0 to PageCount - 1 do
      _WritePageSettings(i);

    _xml.WriteEndTagNode(); //config:config-item-map-named
    _xml.WriteEndTagNode(); //config:config-item-map-entry
    _xml.WriteEndTagNode(); //config:config-item-map-indexed
    _xml.WriteEndTagNode(); //config:config-item-set
    _xml.WriteEndTagNode(); //office:settings
    _xml.WriteEndTagNode(); //office:document-settings
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ODFCreateSettings

//Записывает в поток документ + автоматические стили (content.xml)
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
function ODFCreateContent(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;
  ColumnStyle, RowStyle: array of array of integer;  //стили столбцов/строк
  i: integer;

  //Заголовок для content.xml
  procedure WriteHeader();
  var
    i, j: integer;
    kol: integer;
    n: integer;
    ColStyleNumber, RowStyleNumber: integer;
    
    //Стили для колонок
    procedure WriteColumnStyle(now_i, now_j, now_StyleNumber, count_i, count_j: integer);
    var
      i, j: integer;
      start_j: integer;
      b: boolean;
      s: string;

    begin
      if (ColumnStyle[now_i][now_j] > -1) then
        exit;
      _xml.Attributes.Clear();
      _xml.Attributes.Add(ZETag_Attr_StyleName, 'co' + IntToStr(now_StyleNumber));
      _xml.Attributes.Add('style:family', 'table-column', false);
      _xml.WriteTagNode('style:style', true, true, false);

      _xml.Attributes.Clear();
      //разрыв страницы (fo:break-before = auto | column | page)
      s := 'auto';
      if (XMLSS.Sheets[_pages[now_i]].Columns[now_j].Breaked) then
        s := 'column';
      _xml.Attributes.Add('fo:break-before', s);
      //Ширина колонки style:column-width
      _xml.Attributes.Add('style:column-width', ZEFloatSeparator(FormatFloat('0.###', XMLSS.Sheets[_pages[now_i]].Columns[now_j].WidthMM / 10)) + 'cm', false);
      _xml.WriteEmptyTag('style:table-column-properties', true, false);

      _xml.WriteEndTagNode(); //style:style
      start_j := now_j + 1;
      ColumnStyle[now_i][now_j] := now_StyleNumber;
      for i := now_i to count_i do
      begin
        for j := start_j to XMLSS.Sheets[_pages[i]].ColCount - 1{count_j} do
          if (ColumnStyle[i][j] = -1) then
          begin
            b := true;
            //style:column-width
            if (XMLSS.Sheets[_pages[i]].Columns[j].WidthPix <> XMLSS.Sheets[_pages[now_i]].Columns[now_j].WidthPix) then
              b := false;
            //fo:break-before
            if (XMLSS.Sheets[_pages[i]].Columns[j].Breaked <> XMLSS.Sheets[_pages[now_i]].Columns[now_j].Breaked) then
              b := false;

            if (b) then
              ColumnStyle[i][j] := now_StyleNumber;
          end;

        start_j := 0;
      end;
    end; //WriteColumnStyle

    //Стили для строк
    procedure WriteRowStyle(now_i, now_j, now_StyleNumber, count_i, count_j: integer);
    var
      i, j, k: integer;
      start_j: integer;
      b: boolean;
      s: string;

    begin
      if (RowStyle[now_i][now_j] > -1) then
        exit;
      _xml.Attributes.Clear();
      _xml.Attributes.Add(ZETag_Attr_StyleName, 'ro' + IntToStr(now_StyleNumber));
      _xml.Attributes.Add('style:family', 'table-row', false);
      _xml.WriteTagNode('style:style', true, true, false);

      _xml.Attributes.Clear();
      //разрыв страницы (fo:break-before = auto | column | page)
      s := 'auto';
      if (XMLSS.Sheets[_pages[now_i]].Rows[now_j].Breaked) then
        s := 'page';
      _xml.Attributes.Add('fo:break-before', s);
      //высота строки style:row-height
      _xml.Attributes.Add('style:row-height', ZEFloatSeparator(FormatFloat('0.###', XMLSS.Sheets[_pages[now_i]].Rows[now_j].HeightMM / 10)) + 'cm', false);
       //?? style:min-row-height
      //style:use-optimal-row-height - пересчитывать ли высоту, если содержимое ячеек изменилось
      if (abs(XMLSS.Sheets[_pages[now_i]].Rows[now_j].Height - XMLSS.Sheets[_pages[now_i]].DefaultRowHeight) < 0.001) then
        _xml.Attributes.Add('style:use-optimal-row-height', ODFBoolToStr(true), false);
      //fo:background-color - цвет фона
      k := XMLSS.Sheets[_pages[now_i]].Rows[now_j].StyleID;
      if (k > -1) then
        if (XMLSS.Styles.Count - 1 >= k) then
          if (XMLSS.Styles[k].BGColor <> XMLSS.Styles[-1].BGColor) then
            _xml.Attributes.Add('fo:background-color', '#' + ColorToHTMLHex(XMLSS.Styles[k].BGColor), false);

      //?? fo:keep-together - неразрывные строки (auto | always)
      _xml.WriteEmptyTag('style:table-row-properties', true, false);

      _xml.WriteEndTagNode(); //style:style
      start_j := now_j + 1;
      RowStyle[now_i][now_j] := now_StyleNumber;
      for i := now_i to count_i do
      begin
        for j := start_j to XMLSS.Sheets[_pages[i]].RowCount - 1{count_j} do
          if (RowStyle[i][j] = -1) then
          begin
            b := true;
            //style:row-height
            if (XMLSS.Sheets[_pages[i]].Rows[j].HeightPix <> XMLSS.Sheets[_pages[now_i]].Rows[now_j].HeightPix) then
              b := false;
            //fo:break-before
            if (XMLSS.Sheets[_pages[i]].Rows[j].Breaked <> XMLSS.Sheets[_pages[now_i]].Rows[now_j].Breaked) then
              b := false;

            if (b) then
              RowStyle[i][j] := now_StyleNumber;
          end;

        start_j := 0;
      end;
    end; //WriteRowStyle

  begin
    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    GenODContentAttr(_xml.Attributes);
    _xml.WriteTagNode('office:document-content', true, true, false);
    _xml.Attributes.Clear();
    _xml.WriteEmptyTag('office:scripts', true, false);  //потом на досуге можно подумать
    ZEWriteFontFaceDecls(XMLSS, _xml);

    ///********   Automatic Styles   ********///
    _xml.WriteTagNode('office:automatic-styles', true, true, false);
    //******* стили столбцов
    kol := High(_pages);
    SetLength(ColumnStyle, kol + 1);
    for i := 0 to kol do
    begin
      n := XMLSS.Sheets[_pages[i]].ColCount;
      SetLength(ColumnStyle[i], n);
      for j := 0 to n - 1 do
        ColumnStyle[i][j] := -1;
    end;
    ColStyleNumber := 0;
    for i := 0 to kol do
    begin
      n := XMLSS.Sheets[_pages[i]].ColCount;
      for j := 0 to n - 1 do
      begin
        WriteColumnStyle(i, j, ColStyleNumber, kol, n - 1);
        inc(ColStyleNumber);
      end;
    end;

    //******* стили строк
    SetLength(RowStyle, kol + 1);
    for i := 0 to kol do
    begin
      n := XMLSS.Sheets[_pages[i]].RowCount;
      SetLength(RowStyle[i], n);
      for j := 0 to n - 1 do
        RowStyle[i][j] := -1;
    end;
    RowStyleNumber := 0;
    for i := 0 to kol do
    begin
      n := XMLSS.Sheets[_pages[i]].RowCount - 1;
      for j := 0 to n - 1 do
      begin
        WriteRowStyle(i, j, RowStyleNumber, kol, n);
        inc(RowStyleNumber);
      end;
    end;

    //******* остальные стили
    for i := 0 to XMLSS.Styles.Count - 1 do
    begin
      _xml.Attributes.Clear();
      _xml.Attributes.Add(ZETag_Attr_StyleName, 'ce' + IntToStr(i));
      _xml.Attributes.Add('style:family', 'table-cell', false);
        //??style:parent-style-name = Default
      _xml.WriteTagNode('style:style', true, true, false);
      ODFWriteTableStyle(XMLSS, _xml, i, false);
      _xml.WriteEndTagNode(); //style:style
    end;

    _xml.WriteEndTagNode(); //office:automatic-styles
  end; //WriteHeader

  //<table:table> ... </table:table>
  //INPUT
  //      PageNum: integer    - номер страницы в хранилище
  //  const TableName: string - название страницы
  //      PageIndex: integer  - номер в массиве страниц
  procedure WriteODFTable(const PageNum: integer; const TableName: string; PageIndex: integer);
  var
    b: boolean;
    i, j: integer;
    s, ss: string;
    k, t: integer;
    NumTopLeft: integer;  //Номер объединённой области, в которой ячейка является верхней левой
    NumArea: integer;     //Номер объединённой области, в которую входит ячейка
    isNotEmpty: boolean;
    _StyleID: integer;
    _CellData: string;
    ProcessedSheet: TZSheet;
    DivedIntoHeader: boolean; // начали запись повторяющегося на печати столбца

    //Выводит содержимое ячейки с учётом переноса строк
    procedure WriteTextP(xml: TZsspXMLWriterH; const CellData: string; const href: string = '');
    var
      s: string;
      i: integer;

    begin
      //ссылка <text:p><text:a xlink:href="http://google.com/" office:target-frame-name="_blank">Some_text</text:a></text:p>
      //Ссылка имеет больший приоритет
      if (href > '') then
      begin
        xml.Attributes.Clear();
        xml.WriteTagNode('text:p', true, false, true);
        xml.Attributes.Add('xlink:type', 'simple'); // mandatory for ODF 1.2 validator
        xml.Attributes.Add('xlink:href', href);
        //office:target-frame-name='_blank' - открывать в новом фрейме
        xml.WriteTag('text:a', CellData, false, false, true);
        xml.WriteEndTagNode(); //text:p
      end else
      begin
        s := '';
        for i := 1 to length(CellData) do
        begin
          if CellData[i] = AnsiChar(#10) then
          begin
            xml.WriteTag('text:p', s, true, false, true);
            s := '';
          end else
            if (CellData[i] <> AnsiChar(#13)) then
              s := s + CellData[i];
        end;
        if (s > '') then
          xml.WriteTag('text:p', s, true, false, true);
      end;
    end; //WriteTextP

  begin
    ProcessedSheet := XMLSS.Sheets[PageNum];
    _xml.Attributes.Clear();
    //Атрибуты для таблицы:
    //    table:name        - название таблицы
    //    table:style-name  - стиль таблицы
    //    table:protected   - признак защищённой таблицы (true/false)
    //?   table:protection-key - ХЭШ пароля (если таблица защищённая)
    //?   table:print       - является-ли таблица печатаемой (true - по-умолчанию)
    //?   table:display     - признак отображаемости таблицы (мощнее печати, true - по-умолчани.)
    //?   table:print-ranges - диапазон печати
    _xml.Attributes.Add('table:name', Tablename, false);
    b := ProcessedSheet.Protect;
    if (b) then
      _xml.Attributes.Add('table:protected', ODFBoolToStr(b), false);
    _xml.WriteTagNode('table:table', true, true, false);

    //::::::: колонки :::::::::
    //table:table-column - описание колонок
    //Атрибуты
    //    table:number-columns-repeated   - кол-во столбцов, в которых повторяется описание столбца (кол-во - 1 последующих пропускать) - потом надо подумать {tut}
    //    table:style-name                - стиль столбца
    //    table:visibility                - видимость столбца (по-умолчанию visible);
    //    table:default-cell-style-name   - стиль ячеек по умолчанию (если не задан стиль строки и ячейки)
    DivedIntoHeader := False;
    for i := 0 to ProcessedSheet.ColCount - 1 do
    begin
      _xml.Attributes.Clear();
      _xml.Attributes.Add('table:style-name', 'co' + IntToStr(ColumnStyle[PageIndex][i]));
      //Видимость: table:visibility (visible | collapse | filter)
      if (ProcessedSheet.Columns[i].Hidden) then
        _xml.Attributes.Add('table:visibility', 'collapse', false); //или всё-таки filter?
      s := 'Default';
      k := ProcessedSheet.Columns[i].StyleID;
      if (k >= 0) then
        s := 'ce' + IntToStr(k);
      _xml.Attributes.Add('table:default-cell-style-name', s, false);
      //table:default-cell-style-name

      // на память: учёт столбцов-заголовков сам по себе
      //  должен влиять на table:number-columns-repeated
      with ProcessedSheet.ColsToRepeat do
        if Active and (i = From) then begin // входим в зону заголовка
           _xml.WriteTagNode('table:table-header-columns', []);
           DivedIntoHeader := true;
        end;

      _xml.WriteEmptyTag('table:table-column', true, false);

      if DivedIntoHeader and (i = ProcessedSheet.ColsToRepeat.Till) then begin
         // выходим из зоны заголовка
           _xml.WriteEndTagNode;//  TagNode('table:table-header-columns', []);
           DivedIntoHeader := False;
        end;
    end;
    if DivedIntoHeader then // может кто-то уменьшил ColCount после установки ColsToRepeat ?
       _xml.WriteEndTagNode;//  TagNode('table:table-header-columns', []);

    //::::::: строки :::::::::
    DivedIntoHeader := False;
    for i := 0 to ProcessedSheet.RowCount - 1 do
    begin
      //table:table-row
      _xml.Attributes.Clear();
      _xml.Attributes.Add('table:style-name', 'ro' + IntToStr(RowStyle[PageIndex][i]));
      //?? table:number-rows-repeated - кол-во повторяемых строк
      // table:default-cell-style-name - стиль ячейки по-умолчанию
      k := ProcessedSheet.Rows[i].StyleID;
      if (k >= 0) then
      begin
        s := 'ce' + IntToStr(k);
        _xml.Attributes.Add('table:default-cell-style-name', s, false);
      end;
      // table:visibility - видимость строки
      if (ProcessedSheet.Rows[i].Hidden) then
        _xml.Attributes.Add('table:visibility', 'collapse', false);

      // на память: учёт строк-заголовков сам по себе
      //  должен влиять на table:number-rows-repeated
      with ProcessedSheet.RowsToRepeat do
        if Active and (i = From) then begin // входим в зону заголовка
           _xml.WriteTagNode('table:table-header-rows', []);
           DivedIntoHeader := true;
        end;

      _xml.WriteTagNode('table:table-row', true, true, false);
      {ячейки}
      //**** пробегаем по всем ячейкам
      for j := 0 to ProcessedSheet.ColCount - 1 do
      begin
        NumTopLeft := ProcessedSheet.MergeCells.InLeftTopCorner(j, i);
        NumArea := ProcessedSheet.MergeCells.InMergeRange(j, i);
        s := 'table:table-cell';
        _xml.Attributes.Clear();
        isNotEmpty := false;
        //Возможные атрибуты для ячейки:
        //    table:number-columns-repeated   - кол-во повторяемых ячеек
        //    table:number-columns-spanned    - кол-во объединённых столбцов
        //    table:number-rows-spanned       - кол-во объединённых строк
        //    table:style-name                - стиль ячейки
        //??  table:content-validation-name   - проводится ли в данной ячейке проверка правильности
        //    table:formula                   - формула
        //      office:value                  -  текущее числовое значение (для float | percentage | currency)
        //??    office:date-value             -  текущее значение даты
        //??    office:time-value             -  текущее значение времени
        //??    office:boolean-value          -  текущее логическое значение
        //      office:string-value           -  текущее строковое значение
        //     table:value-type               - тип значения в ячейке (float | percentage | currency,
        //                                        date, time, boolean, string
        //??     tableoffice:currency         - текущая денежная единица (только для currency)
        //     table:protected                - защищённость ячейки
        if ((NumTopLeft < 0) and (NumArea >= 0)) then //скрытая ячейка в объединённой области
        begin
          s := 'table:covered-table-cell';
        end else
        if (NumTopLeft >= 0) then   //объединённая ячейка (левая верхняя)
        begin
          t := ProcessedSheet.MergeCells.Items[NumTopLeft].Right -
               ProcessedSheet.MergeCells.Items[NumTopLeft].Left + 1;
          _xml.Attributes.Add('table:number-columns-spanned', IntToStr(t), false);
          t := ProcessedSheet.MergeCells.Items[NumTopLeft].Bottom -
               ProcessedSheet.MergeCells.Items[NumTopLeft].Top + 1;
          _xml.Attributes.Add('table:number-rows-spanned', IntToStr(t), false);
        end;

        //стиль
        _StyleID := ProcessedSheet.Cell[j, i].CellStyle;
        if ((_StyleID >= 0) and (_StyleID < XMLSS.Styles.Count)) then
          _xml.Attributes.Add('table:style-name', 'ce' + IntToStr(_StyleID), false);

        //защита ячейки
        b := XMLSS.Styles[_StyleID].Protect;
        if (b) then
          _xml.Attributes.Add('table:protected', ODFBoolToStr(b), false);

        _CellData := ProcessedSheet.Cell[j, i].Data;
        //table:value-type + office:some_type_value
        //  ZENumber -> float
        ss := '';
        case (ProcessedSheet.Cell[j, i].CellType) of
          ZENumber:
            begin
              ss := 'float';
              _xml.Attributes.Add('office:value', ZEFloatSeparator(FormatFloat('0.#######', ZETryStrToFloat(_CellData))), false);
            end;
          ZEBoolean:
            begin
              ss := 'boolean';
              _xml.Attributes.Add('office:boolean-value', ODFBoolToStr(ZETryStrToBoolean(_CellData)), false);
            end;
          else
            // всё остальное считаем строкой (потом подправить, возможно, добавить новые типы)
            {ZEansistring ZEError ZEDateTime}
        end;
        if (ss > '') then
          _xml.Attributes.Add('office:value-type', ss, false);

        //формула  
        ss := ProcessedSheet.Cell[j, i].Formula;
        if (ss > '') then
          _xml.Attributes.Add('table:formula', ss, false);

        //Примечание
        //office:annotation
        //Атрибуты:
        //    office:display - отображать ли (true | false)
        //??  draw:style-name
        //??  draw:text-style-name
        //??  svg:width
        //??  svg:height
        //??  svg:x
        //??  svg:y
        //??  draw:caption-point-x
        //??  draw:caption-point-y
        if (ProcessedSheet.Cell[j, i].ShowComment) then
        begin
          if (not isNotEmpty) then
            _xml.WriteTagNode(s, true, true, true);
          isNotEmpty := true;
          _xml.Attributes.Clear();
          b := ProcessedSheet.Cell[j, i].AlwaysShowComment;
          if (b) then
            _xml.Attributes.Add('office:display', ODFBoolToStr(b), false);
          _xml.WriteTagNode('office:annotation', true, true, false);
          //автор примечания
          if (ProcessedSheet.Cell[j, i].CommentAuthor > '') then
          begin
            _xml.Attributes.Clear();
            _xml.WriteTag('dc:creator', ProcessedSheet.Cell[j, i].CommentAuthor, true, false, true);
          end;
          _xml.Attributes.Clear();
          WriteTextP(_xml, ProcessedSheet.Cell[j, i].Comment);
          _xml.WriteEndTagNode(); //office:annotation
        end;

        //Содержимое ячейки
        if (_CellData > '') then
        begin
          if (not isNotEmpty) then
            _xml.WriteTagNode(s, true, true, true);
          isNotEmpty := true;
          _xml.Attributes.Clear();
          WriteTextP(_xml, _CellData, ProcessedSheet.Cell[j, i].Href);
        end;

        if (isNotEmpty) then
          _xml.WriteEndTagNode() //ячейка  table:table-cell | table:covered-table-cell
        else
          _xml.WriteEmptyTag(s, true, true);
      end;
      {/ячейки}

      if DivedIntoHeader and (i = ProcessedSheet.RowsToRepeat.Till) then begin
         // выходим из зоны заголовка
           _xml.WriteEndTagNode;//  TagNode('table:table-header-rows', []);
           DivedIntoHeader := False;
        end;

      _xml.WriteEndTagNode(); //table:table-row
    end;
    if DivedIntoHeader then // может кто-то уменьшил RowCount после установки RowsToRepeat ?
       _xml.WriteEndTagNode;//  TagNode('table:table-header-rows', []);

    _xml.WriteEndTagNode(); //table:table
  end; //WriteODFTable

  //Сами таблицы
  procedure WriteBody();
  var
    i: integer;

  begin
    _xml.Attributes.Clear();
    _xml.WriteTagNode('office:body', true, true, false);
    _xml.WriteTagNode('office:spreadsheet', true, false);
    for i := Low(_pages) to High(_pages) do
      WriteODFTable(_pages[i], _names[i], i);
    _xml.WriteEndTagNode(); //office:spreadsheet
    _xml.WriteEndTagNode(); //office:body
  end; //WriteBody

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
    WriteHeader();
    WriteBody();
    _xml.EndSaveTo();
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
    for i := Low(ColumnStyle) to High(ColumnStyle) do
    begin
      SetLength(ColumnStyle[i], 0);
      ColumnStyle[i] := nil;
    end;
    SetLength(ColumnStyle, 0);
    for i := Low(RowStyle) to High(RowStyle) do
    begin
      SetLength(RowStyle[i], 0);
      RowStyle[i] := nil;
    end;
    RowStyle := nil;
  end;
end; //ODFCreateContent

//Записывает в поток метаинформацию (meta.xml)
//INPUT
//  var XMLSS: TZEXMLSS                 - хранилище
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ODFCreateMeta(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var
  _xml: TZsspXMLWriterH;    //писатель
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
    GenODMetaAttr(_xml.Attributes);
    _xml.WriteTagNode('office:document-meta', true, true, false);
    _xml.Attributes.Clear();
    _xml.WriteTagNode('office:meta', true, true, true);
    //дата создания
    s := ZEDateToStr(XMLSS.DocumentProperties.Created);
    _xml.WriteTag('meta:creation-date', s, true, false, true);
    //Дата последнего редактирования пусть будет равна дате создания
    _xml.WriteTag('dc:date', s, true, false, true);

    {
    //Длительность редактирования PnYnMnDTnHnMnS
    //Пока не используется
    <meta:editing-duration>PT2M2S</meta:editing-duration>
    }

    //Кол-во циклов редактирования (каждый раз, когда сохраняется, нужно увеличивать на 1) > 0
    //Пока считаем, что документ создаётся только 1 раз
    //Потом можно будет добавить поле в хранилище
    _xml.WriteTag('meta:editing-cycles', '1', true, false, true);

    {
    //Статистика документа (кол-во страниц и др)
    //Пока не используется
    (*
    meta:page-count
    meta:table-count
    meta:image-count
    meta:cell-count
    meta:object-count
    *)
    <meta:document-statistic meta:table-count="3" meta:cell-count="7" meta:object-count="0"/>
    }
    //Генератор - какое приложение создало или редактировало документ
    //Потом надо добавить такое поле в хранилище
//    {$IFDEF FPC}
//    s := 'FPC';
//    {$ELSE}
//    s := 'DELPHI_or_CBUILDER';
//    {$ENDIF}
//    _xml.WriteTag('meta:generator', 'ZEXMLSSlib/0.0.5$' + s, true, false, true);
    _xml.WriteTag('meta:generator', ZELibraryName, true, false, true);

    _xml.WriteEndTagNode(); // office:meta
    _xml.WriteEndTagNode(); // office:document-meta

  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ODFCreateMeta

//Создаёт манифест
//INPUT
//      Stream: TStream                   - поток для записи
//      TextConverter: TAnsiToCPConverter - конвертер
//      CodePageName: string              - имя кодировки
//      BOM: ansistring                   - BOM
//RETURN
//      integer
function ODFCreateManifest(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer;
var
  _xml: TZsspXMLWriterH;    //писатель
  tag_name, att1, att2, s: string;

  procedure _writetag(const s1, s2: string);
  begin
    _xml.Attributes.Clear();
    _xml.Attributes.Add(att1, s1);
    _xml.Attributes.Add(att2, s2, false);
    _xml.WriteEmptyTag(tag_name, true, false);
  end;

begin
  _xml := nil;
  result := 0;
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
    _xml.Attributes.Add('xmlns:manifest', 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0');
    _xml.Attributes.Add('manifest:version', '1.2');
    _xml.WriteTagNode('manifest:manifest', true, true, true);

    tag_name := 'manifest:file-entry';
    att1 := 'manifest:media-type';
    att2 := 'manifest:full-path';

    _xml.Attributes.Clear();
    _xml.Attributes.Add(att1, 'application/vnd.oasis.opendocument.spreadsheet');
    _xml.Attributes.Add('manifest:version', '1.2');
    _xml.Attributes.Add(att2, '/');
    _xml.WriteEmptyTag(tag_name);
    s := 'text/xml';

    _writetag(s, 'meta.xml');
    _writetag(s, 'settings.xml');
    _writetag(s, 'content.xml');
//    _writetag('image/png', 'Thumbnails/thumbnail.png'); - not implemented
//    _writetag('', 'Configurations2/accelerator/current.xml'); - not implemented
//    _writetag('application/vnd.sun.xml.ui.configuration', 'Configurations2/'); - no such folder
    _writetag(s, 'styles.xml');

    _xml.WriteEndTagNode(); //manifest:manifest
    _xml.EndSaveTo();
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ODFCreateManifest

//Сохраняет незапакованный документ в формате Open Document
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
function SaveXmlssToODFSPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
var
  _pages: TIntegerDynArray;      //номера страниц
  _names: TStringDynArray;      //названия страниц
  kol: integer;
  Stream: TStream;
  s: string;

begin
  result := 0;
  Stream := nil;
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

    Stream := TFileStream.Create(PathName + 'content.xml', fmCreate);
    ODFCreateContent(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    Stream := TFileStream.Create(PathName + 'meta.xml', fmCreate);
    ODFCreateMeta(XMLSS, Stream, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    Stream := TFileStream.Create(PathName + 'settings.xml', fmCreate);
    ODFCreateSettings(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    Stream := TFileStream.Create(PathName + 'styles.xml', fmCreate);
    ODFCreateStyles(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    s := PathName + 'META-INF' + PathDelim;
    if (not DirectoryExists(s)) then
       ForceDirectories(s);

    Stream := TFileStream.Create(s + 'manifest.xml', fmCreate);
    ODFCreateManifest(Stream, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

  finally
    ZESClearArrays(_pages, _names);
    if (Assigned(Stream)) then
      FreeAndNil(Stream);
  end;
end; //SaveXmlssToODFSPath

{$IFDEF FPC}
//Сохраняет документ в формате Open Document
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
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
var
  _pages: TIntegerDynArray;      //номера страниц
  _names: TStringDynArray;      //названия страниц
  kol: integer;
  zip: TZipper;
  StreamC, StreamME, StreamS, StreamST, StreamMA: TStream;

begin
  zip := nil;
  StreamC := nil;
  StreamME := nil;
  StreamS := nil;
  StreamST := nil;
  StreamMA := nil;
  result := 0;
  try

    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      result := 2;
      exit;
    end;

    zip := TZipper.Create();

    StreamC := TMemoryStream.Create();
    ODFCreateContent(XMLSS, StreamC, _pages, _names, kol, TextConverter, CodePageName, BOM);

    StreamME := TMemoryStream.Create();
    ODFCreateMeta(XMLSS, StreamME, TextConverter, CodePageName, BOM);

    StreamS := TMemoryStream.Create();
    ODFCreateSettings(XMLSS, StreamS, _pages, _names, kol, TextConverter, CodePageName, BOM);

    StreamST := TMemoryStream.Create();
    ODFCreateStyles(XMLSS, StreamST, _pages, _names, kol, TextConverter, CodePageName, BOM);

    StreamMA := TMemoryStream.Create();
    ODFCreateManifest(StreamMA, TextConverter, CodePageName, BOM);

    zip.FileName := FileName;

    StreamC.Position := 0;
    StreamME.Position := 0;
    StreamS.Position := 0;
    StreamST.Position := 0;
    StreamMA.Position := 0;

    zip.Entries.AddFileEntry(StreamC, 'content.xml');
    zip.Entries.AddFileEntry(StreamME, 'meta.xml');
    zip.Entries.AddFileEntry(StreamS, 'settings.xml');
    zip.Entries.AddFileEntry(StreamST, 'styles.xml');
    zip.Entries.AddFileEntry(StreamMA, 'META-INF/manifest.xml');
    zip.ZipAllFiles();

  finally
    ZESClearArrays(_pages, _names);
    if (Assigned(zip)) then
      FreeAndNil(zip);
    if (Assigned(StreamC)) then
      FreeAndNil(StreamC);
    if (Assigned(StreamME)) then
      FreeAndNil(StreamME);
    if (Assigned(StreamS)) then
      FreeAndNil(StreamS);
    if (Assigned(StreamST)) then
      FreeAndNil(StreamST);
    if (Assigned(StreamMA)) then
      FreeAndNil(StreamMA);
  end;

end; //SaveXmlssToODFS

{$ENDIF}

function ExportXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                           const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: String;
                           BOM: ansistring = '';
                           AllowUnzippedFolder: boolean = false; ZipGenerator: CZxZipGens = nil): integer; overload;
var
  _pages: TIntegerDynArray;      //номера страниц
  _names: TStringDynArray;      //названия страниц
  kol: integer;
  Stream: TStream;
  azg: TZxZipGen; // Actual Zip Generator
  mime: AnsiString;

begin
  azg := nil;
  try

    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      result := 2;
      exit;
    end;

    // Todo - common block and exception const with XLSX output in zexlsx unit => need merge
    if nil = ZipGenerator then begin
      ZipGenerator := TZxZipGen.QueryZipGen;
      if nil = ZipGenerator then
        if AllowUnzippedFolder
           then ZipGenerator := TZxZipGen.QueryDummyZipGen
           else raise EZxZipGen.Create('No zip generators registered, folder output disabled.');
           // result := 3 ????
    end;
    azg := ZipGenerator.Create(FileName);

// этот файл долен быть первым
// а еще он не должен был сжат! http://odf-validator.rhcloud.com/
    Stream := azg.NewStream('mimetype');
    mime := 'application/vnd.oasis.opendocument.spreadsheet';
    Stream.WriteBuffer(mime[1], Length(mime));
    azg.SealStream(Stream);

    Stream := azg.NewStream('content.xml');
    ODFCreateContent(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);

    Stream := azg.NewStream('meta.xml');
    ODFCreateMeta(XMLSS, Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);

    Stream := azg.NewStream('settings.xml');
    ODFCreateSettings(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);

    Stream := azg.NewStream('styles.xml');
    ODFCreateStyles(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);

    Stream := azg.NewStream('META-INF/manifest.xml');
    ODFCreateManifest(Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);

    azg.SaveAndSeal;
  finally
    ZESClearArrays(_pages, _names);
    azg.Free;
  end;
  Result := 0;
end; //ExportXmlssToODFS



/////////////////// чтение

//Возвращает размер измерения в ММ
//INPUT
//  const value: string     - строка со значением
//  var RetSize: real       - возвращаемое значение
//      isMultiply: boolean - флаг необходимости умножать значение с учётом единицы измерения
//RETURN
//      boolean - true - размер определён успешно
function ODFGetValueSizeMM(const value: string; out RetSize: real; isMultiply: boolean = true): boolean;
var
  i: integer;
  sv, su: string;
  ch: char;
  _isU: boolean;
  r: double;

begin
  result := true;
  sv := '';
  su := '';
  _isU := false;
  for i := 1 to length(value) do
  begin
    ch := value[i];
    case ch of
      '0'..'9':
        begin
          if (_isU) then
            su := su + ch
          else
            sv := sv + ch;
        end;
      '.', ',':
        begin
          if (_isU) then
            su := su + ch
          else
            {$IFDEF DELPHI_UNICODE}
            sv := sv + FormatSettings.DecimalSeparator
            {$ELSE}
            sv := sv + DecimalSeparator
            {$ENDIF}
        end;
      else
        begin
          _isU := true;
          su := su + ch;
        end;
    end;
  end; //for
  if (not TryStrToFloat(sv, r)) then
  begin
    result := false;
    exit;
  end;
  if (not isMultiply) then
  begin
    RetSize := r;
    exit;
  end;
//  su := UpperCase(su);
  if (su = 'cm') then
    RetSize := r * 10
  else
  if (su = 'mm') then
    RetSize := r
  else
  if (su = 'dm') then
    RetSize := r * 100
  else
  if (su = 'm') then
    RetSize := r * 1000
  else
  if (su='pt') then
    RetSize := r * _PointToMM
  else
  if (su='in') then
    RetSize := r * 25.4
  else
    result := false;
end; //ODFGetValueSizeMM

//Чтение автоматических стилей
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//      stream: TStream - поток для чтения
//RETURN
//      boolean - true - всё ок
function ReadODFStyles(var XMLSS: TZEXMLSS; stream: TStream): boolean;
begin
  result := false;
end; //ReadODFStyles

//Чтение содержимого документа ODS (content.xml)
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//      stream: TStream - поток для чтения
//RETURN
//      boolean - true - всё ок
function ReadODFContent(var XMLSS: TZEXMLSS; stream: TStream): boolean;
var
  xml: TZsspXMLReaderH;
  ErrorReadCode: integer;
  ODFColumnStyles: TZODFColumnStyleArray;
  ODFRowStyles: TZODFRowStyleArray;
  ODFStyles: TZODFStyleArray;
  ODFTableStyles: TZODFTableArray;
  ColStyleCount, MaxColStyleCount: integer;
  RowStyleCount, MaxRowStyleCount: integer;
  StyleCount, MaxStyleCount: integer;
  TableStyleCount, MaxTableStyleCount: integer;

  function IfTag(const TgName: string; const TgType: integer): boolean;
  begin
    result := (xml.TagName = TgName) and (xml.TagType = TgType);
  end;

  //Ищет номер стиля по названию
  function _FindStyleID(const st: string): integer;
  var
    i: integer;

  begin
    result := -1;
    for i := 0 to StyleCount - 1 do
    if (ODFStyles[i].name = st) then
    begin
      result := ODFStyles[i].index;
      break;
    end;
  end; //_FindStyleID

  //Автоматические стили
  procedure _ReadAutomaticStyle();
  var
    _stylefamily: string;
    _stylename: string;
    s: string;
    _style: TZSTyle;

    //Чтение стиля для ячейки
    procedure _ReadCellStyle();
    var
      t: integer; HAutoForced: boolean;
      r: real;
      s: string;
     function ODF12AngleUnit(const un: string; const scale: double): boolean;
     var err: integer; d: double;
     begin
       Result := AnsiEndsStr(un, s);
       if Result then begin
          Val( Trim(Copy( s, 1, length(s) - length(un))), d, err);
          Result := err = 0;
          if Result then
             t := round(d * scale);
       end;
     end;
    begin
      if (StyleCount >= MaxStyleCount) then
      begin
        MaxStyleCount := StyleCount + 20;
        SetLength(ODFStyles, MaxStyleCount);
      end;  
      ODFStyles[StyleCount].index := -1;
      ODFStyles[StyleCount].name := _stylename;
      _style.Assign(XMLSS.Styles.DefaultStyle);
      HAutoForced := false; // pre-clean: paragraph properties may come before cell properties

      while (not IfTag('style:style', 6)) do
      begin
        if (xml.Eof()) then
          break;
        xml.ReadTag();

        if ((xml.TagName = 'style:table-cell-properties') and (xml.TagType in [4, 5])) then
        begin
          //Выравнивание по вертикали
          s := xml.Attributes.ItemsByName['style:vertical-align'];
          if (s > '') then
          begin
            if (s = 'automatic') then
              _style.Alignment.Vertical := ZVAutomatic
            else
            if (s = 'top') then
              _style.Alignment.Vertical := ZVTop
            else
            if (s = 'bottom') then
              _style.Alignment.Vertical := ZVBottom
            else
            if (s = 'middle') then
              _style.Alignment.Vertical := ZVCenter;
          end;

          HAutoForced := 'value-type' = xml.Attributes['style:text-align-source'];
          If HAutoForced then _style.Alignment.Horizontal := ZHAutomatic;

          //Угол поворота текста
          s := xml.Attributes.ItemsByName['style:rotation-angle'];
          if (s > '') then begin
            if not TryStrToInt(s, t) // ODS 1.1 - pure integer - failed
            then begin // ODF 1.2+ ? float with units ?
              s := LowerCase(Trim(s));
              if not ODF12AngleUnit('deg', 1) then
                 if not ODF12AngleUnit('grad', 90 / 100) then
                    if not ODF12AngleUnit('rad', 180 / Pi ) then
                       if not ODF12AngleUnit('', 1) then // just unit-less float ?
                          s := ''; // not parsed
            end;
            if s > '' then begin  // need reduce to -180 to +180
               t := t mod 360;
               if t > +180 then t := t - 360;
               if t < -180 then t := t + 360;
               _style.Alignment.Rotate := t;
            end;
          end;
          _style.Alignment.VerticalText :=
               'ttb' = xml.Attributes['style:direction'];

          //цвет фона
          s := xml.Attributes.ItemsByName['fo:background-color'];
          if (s > '') then
            _style.BGColor := GetBGColorForODS(s);//HTMLHexToColor(s);

          //подгонять ли, если текст не помещается
          s := xml.Attributes.ItemsByName['style:shrink-to-fit'];
          if (s > '') then
            _style.Alignment.ShrinkToFit := ZEStrToBoolean(s);

          ///обрамление
          s := xml.Attributes.ItemsByName['fo:border'];
          if (s > '') then
          begin
            ZEStrToODFBorderStyle(s, _style.Border[0]);
            for t := 1 to 3 do
              _style.Border[t].Assign(_style.Border[0]);
          end;

          s := xml.Attributes.ItemsByName['fo:border-left'];
          if (s > '') then
            ZEStrToODFBorderStyle(s, _style.Border[0]);
          s := xml.Attributes.ItemsByName['fo:border-top'];
          if (s > '') then
            ZEStrToODFBorderStyle(s, _style.Border[1]);
          s := xml.Attributes.ItemsByName['fo:border-right'];
          if (s > '') then
            ZEStrToODFBorderStyle(s, _style.Border[2]);
          s := xml.Attributes.ItemsByName['fo:border-bottom'];
          if (s > '') then
            ZEStrToODFBorderStyle(s, _style.Border[3]);
          s := xml.Attributes.ItemsByName['style:diagonal-bl-tr'];
          if (s > '') then
            ZEStrToODFBorderStyle(s, _style.Border[4]);
          s := xml.Attributes.ItemsByName['style:diagonal-tl-br'];
          if (s > '') then
            ZEStrToODFBorderStyle(s, _style.Border[5]);

          //Перенос по словам (wrap no-wrap)
          s := xml.Attributes.ItemsByName['fo:wrap-option'];
          if (s > '') then
            if (UpperCase(s) = 'WRAP') then
              _style.Alignment.WrapText := true;
        end else //if

        if ((xml.TagName = 'style:paragraph-properties') and (xml.TagType in [4, 5])) then
        begin
          if not HAutoForced then begin
              s := xml.Attributes.ItemsByName['fo:text-align'];
              if (s > '') then
              begin
                if ((s = 'start') or (s = 'left')) then
                  _style.Alignment.Horizontal := ZHLeft
                else
                if (s = 'center') then
                  _style.Alignment.Horizontal := ZHCenter
                else
                if (s = 'justify') then
                  _style.Alignment.Horizontal := ZHJustify
                else
                if ((s = 'end') or (s = 'right')) then
                  _style.Alignment.Horizontal := ZHRight
                else
                  _style.Alignment.Horizontal := ZHAutomatic;
              end;
          end;
        end else //if

        if ((xml.TagName = 'style:text-properties') and (xml.TagType in [4, 5])) then
        begin
          //style:font-name (style:font-name-asian style:font-name-complex)
          s := xml.Attributes.ItemsByName['style:font-name'];
          if (s > '') then
            _style.Font.Name := s;

          //fo:font-size (style:font-size-asian style:font-size-complex)
          s := xml.Attributes.ItemsByName['fo:font-size'];
          if (s > '') then
            if (ODFGetValueSizeMM(s, r, false)) then
              _style.Font.Size := round(r);

          //fo:font-weight (style:font-weight-asian style:font-weight-complex)
          s := xml.Attributes.ItemsByName['fo:font-weight'];
          if (s > '') then
            if (s <> 'normal') then
              _style.Font.Style := _style.Font.Style + [fsBold];

          s := xml.Attributes.ItemsByName['style:text-line-through-type'];
          if (s > '') then
            if (s <> 'none') then
              _style.Font.Style := _style.Font.Style + [fsStrikeOut];

          s := xml.Attributes.ItemsByName['style:text-underline-type'];
          if (s > '') then
            if (s <> 'none') then
              _style.Font.Style := _style.Font.Style + [fsUnderline];

          //fo:font-style (style:font-style-asian style:font-style-complex)
          s := xml.Attributes.ItemsByName['fo:font-style'];
          if (s > '') then
            if (s = 'italic') then
              _style.Font.Style := _style.Font.Style + [fsItalic];

          //цвет fo:color
          s := xml.Attributes.ItemsByName['fo:color'];
          if (s > '') then
            _style.Font.Color := HTMLHexToColor(s);
        end; //if

      end; //while
      ODFStyles[StyleCount].index := XMLSS.Styles.Add(_style, true);
      inc(StyleCount);
    end; //_ReadCellStyle

  begin
    _style := nil;
    try
      _style := TZStyle.Create();
      while (not IfTag('office:automatic-styles', 6)) do
      begin
        if (xml.Eof()) then
          break;
        xml.ReadTag();

        if (IfTag('style:style', 4)) then
        begin
          _stylefamily := xml.Attributes.ItemsByName['style:family'];
          _stylename := xml.Attributes.ItemsByName[ZETag_Attr_StyleName];

          if (_stylefamily = 'table-column') then //столбец
          begin
            if (ColStyleCount >= MaxColStyleCount) then
            begin
              MaxColStyleCount := ColStyleCount + 20;
              SetLength(ODFColumnStyles, MaxColStyleCount);
            end;
            ODFColumnStyles[ColStyleCount].name := '';
            ODFColumnStyles[ColStyleCount].breaked := false;
            ODFColumnStyles[ColStyleCount].width := 25;

            while (not IfTag('style:style', 6)) do
            begin
              if (xml.Eof()) then
                break;
              xml.ReadTag();
              if ((xml.TagName ='style:table-column-properties') and (xml.TagType in [4, 5])) then
              begin
                ODFColumnStyles[ColStyleCount].name := _stylename;
                s := xml.Attributes.ItemsByName['fo:break-before'];
                if (s = 'column') then
                  ODFColumnStyles[ColStyleCount].breaked := true;
                s := xml.Attributes.ItemsByName['style:column-width'];
                if (s > '') then
                  ODFGetValueSizeMM(s, ODFColumnStyles[ColStyleCount].width);
              end;
            end; //while
            inc(ColStyleCount);

          end else
          if (_stylefamily = 'table-row') then //строка
          begin
            if (RowStyleCount >= MaxRowStyleCount) then
            begin
              MaxRowStyleCount := RowStyleCount + 20;
              SetLength(ODFRowStyles, MaxRowStyleCount);
            end;
            ODFRowStyles[RowStyleCount].name := '';
            ODFRowStyles[RowStyleCount].breaked := false;
            ODFRowStyles[RowStyleCount].color := clBlack;
            ODFRowStyles[RowStyleCount].height := 10;

            while (not ifTag('style:style', 6)) do
            begin
              if (xml.Eof()) then
                break;
              xml.ReadTag();

              if ((xml.TagName = 'style:table-row-properties') and (xml.TagType in [4, 5])) then
              begin
                ODFRowStyles[RowStyleCount].name := _stylename;
                s := xml.Attributes.ItemsByName['fo:break-before'];
                if (s = 'page') then
                  ODFRowStyles[RowStyleCount].breaked := true;
                s := xml.Attributes.ItemsByName['style:row-height'];
                if (s > '') then
                  ODFGetValueSizeMM(s, ODFRowStyles[RowStyleCount].height);
                s := xml.Attributes.ItemsByName['fo:background-color'];
               if (s > '') then
                 ODFRowStyles[RowStyleCount].color := HTMLHexToColor(s);
              end;
            end; //while
            inc(RowStyleCount);

          end else
          if (_stylefamily = 'table') then //таблица
          begin
            if (TableStyleCount >= MaxTableStyleCount) then
            begin
              MaxTableStyleCount := TableStyleCount + 20;
              SetLength(ODFTableStyles, MaxTableStyleCount);
            end;
            ODFTableStyles[TableStyleCount].name := _stylename;
            ODFTableStyles[TableStyleCount].isColor := false;
            while (not ifTag('style:style', 6)) do
            begin
              if (xml.Eof()) then
                break;
              xml.ReadTag();

              if ((xml.TagName = 'style:table-properties') and (xml.TagType in [4, 5])) then
              begin
                s := xml.Attributes.ItemsByName['tableooo:tab-color'];
                if (s > '') then
                begin
                  ODFTableStyles[TableStyleCount].isColor := true;
                  ODFTableStyles[TableStyleCount].Color := HTMLHexToColor(s);
                end;
              end;

            end; //while
            inc(TableStyleCount);
          end else
          if (_stylefamily = 'table-cell') then //ячейка
          begin
            _ReadCellStyle();
          end;

        end;
      end; //while
    finally
      if (Assigned(_style)) then
        FreeAndNil(_style);
    end;
  end; //_ReadAutomaticStyle

  //Проверить кол-во строк
  procedure CheckRow(const PageNum: integer; const RowCount: integer);
  begin
    if XMLSS.Sheets[PageNum].RowCount < RowCount then
      XMLSS.Sheets[PageNum].RowCount := RowCount
  end;

  //Проверить кол-во столбцов
  procedure CheckCol(const PageNum: integer; const ColCount: integer);
  begin
    if XMLSS.Sheets[PageNum].ColCount < ColCount then
      XMLSS.Sheets[PageNum].ColCount := ColCount
  end;

  //Чтение таблицы
  procedure _ReadTable();
  var
    isRepeatRow: boolean;     //нужно ли повторять строку
    isRepeatCell: boolean;    //нужно ли повторять ячейку
    _RepeatRowCount: integer; //кол-во повторений строки
    _RepeatCellCount: integer;//кол-во повторений ячейки
    _CurrentRow, _CurrentCol: integer; //текущая строка/столбец
    _CurrentPage: integer;    //текущая страница
    _MaxCol: integer;       
    s: string;                
    _CurrCell: TZCell;
    i, t: integer;
    _IsHaveTextInRow: boolean;

    //Повторить строку
    procedure _RepeatRow();
    var
      i, j, n, z: integer;

    begin
      //Максимум строка может повторятся 2000 раз
      if ((isRepeatRow) and (_RepeatRowCount <= 2000)) then
      begin
        //Если в строке небыло текста и её нужно повторить более 512 раз - уменьшаем кол-во
        // до 512 раз
        if ((not _IsHaveTextInRow) and (_RepeatRowCount > 512)) then
          _RepeatRowCount := 512;
        n := XMLSS.Sheets[_CurrentPage].ColCount - 1;
        z := _CurrentRow - 1;
        CheckRow(_CurrentPage, _CurrentRow + _RepeatRowCount);
        for i := 1 to _RepeatRowCount do
        begin
          for j := 0 to n do
          with XMLSS.Sheets[_CurrentPage] do
            Cell[j, _CurrentRow].Assign(Cell[j, z]);
        end;
        inc(_CurrentRow);
      end; //if
    end; //_RepeatRow

    //Чтение ячейки
    procedure _ReadCell();
    var
      _celltext: string;
      i: integer;
      _isnf: boolean;
      _numX, _numY: integer;
      _isHaveTextCell: boolean;
      _kol: integer;

    begin
      if (((xml.TagName = 'table:table-cell') or (xml.TagName = 'table:covered-table-cell')) and (xml.TagType in [4, 5])) then
      begin
        CheckCol(_CurrentPage, _CurrentCol + 1);
        _CurrCell := XMLSS.Sheets[_CurrentPage].Cell[_CurrentCol, _CurrentRow];
        s := xml.Attributes.ItemsByName['table:number-columns-repeated'];
        isRepeatCell := TryStrToInt(s, _RepeatCellCount);
        //кол-во объединённых столбцов
        s := xml.Attributes.ItemsByName['table:number-columns-spanned'];
        if (TryStrToInt(s, _numX)) then
          dec(_numX)
        else
          _numX := 0;
        if (_numX < 0) then
          _numX := 0;
        //Кол-во объединённых строк
        s := xml.Attributes.ItemsByName['table:number-rows-spanned'];
        if (TryStrToInt(s, _numY)) then
          dec(_numY)
        else
          _numY := 0;
        if (_numY < 0) then
          _numY := 0;
        if (_numX + _numY > 0) then
        begin
          CheckCol(_CurrentPage, _CurrentCol + _numX + 1);
          CheckRow(_CurrentPage, _CurrentRow + _numY + 1);
          XMLSS.Sheets[_CurrentPage].MergeCells.AddRectXY(_CurrentCol, _CurrentRow, _CurrentCol + _numX, _CurrentRow + _numY);
        end;

        //стиль ячейки
        s := xml.Attributes.ItemsByName['table:style-name'];
        _CurrCell.CellStyle := XMLSS.Sheets[_CurrentPage].Columns[_CurrentCol].StyleID;
        if (s > '') then
          _CurrCell.CellStyle := _FindStyleID(s);
        //Проверка правильности наполнения
        //*s := xml.Attributes.ItemsByName['table:cell-content-validation'];
        //формула
        _CurrCell.Formula := ZEReplaceEntity(xml.Attributes.ItemsByName['table:formula']);
        //текущее числовое значение (для float | percentage | currency)
        //*s := xml.Attributes.ItemsByName['office:value'];
        {
        //текущее значение даты
        s := xml.Attributes.ItemsByName['office:date-value'];
        //текущее значение времени
        s := xml.Attributes.ItemsByName['office:time-value'];
        //текущее логическое значение
        s := xml.Attributes.ItemsByName['office:boolean-value'];
        //текущая денежная единица
        s := xml.Attributes.ItemsByName['tableoffice:currency'];
        }
        //текущее строковое значение
        //*s := xml.Attributes.ItemsByName['office:string-value'];
        //Тип значения в ячейке
        s := xml.Attributes.ItemsByName['table:value-type'];
        if (s = '') then
        begin
          s := xml.Attributes.ItemsByName['office:value-type'];
          _CurrCell.CellType := ODFTypeToZCellType(s);
        end else
          _CurrCell.CellType := ODFTypeToZCellType(s);

        //защищённость ячейки
        s := xml.Attributes.ItemsByName['table:protected']; //{tut} надо будет добавить ещё один стиль
        //table:number-matrix-rows-spanned ??
        //table:number-matrix-columns-spanned ??

        _celltext := '';
        _isHaveTextCell := false;
        if (xml.TagType = 4) then
        begin
          _isnf := false;
          while (not(((xml.TagName = 'table:table-cell') or (xml.TagName = 'table:covered-table-cell')) and (xml.TagType = 6))) do
          begin
            if (xml.Eof()) then
              break;
            xml.ReadTag();

            //Текст ячейки
            if (IfTag('text:p', 4)) then
            begin
              _IsHaveTextInRow := true;
              _isHaveTextCell := true;
              if (_isnf) then
                _celltext := _celltext + {$IFDEF FPC} LineEnding {$ELSE} sLineBreak {$ENDIF};
              while (not IfTag('text:p', 6)) do
              begin
                if (xml.Eof()) then
                  break;
                xml.ReadTag();

                //text:a - ссылка
                if (IfTag('text:a', 4)) then
                begin
                  _CurrCell.Href := xml.Attributes.ItemsByName['xlink:href'];
                  //Дополнительные атрибуты: (пока игнорируются)
                  //  office:name - название ссылки
                  //  office:target-frame-name - фрэйм назначения
                  //            _self   - документ по ссылке заменяет текущий
                  //            _blank  - открывается в новом фрэйме
                  //            _parent - открывается в родительском текущего
                  //            _top    - самый верхний
                  //            название_фрэйма
                  //?? xlink:show - (new | replace)
                  //  text:style-name - стиль непосещённой ссылки
                  //  text:visited-style-name - стиль посещённой ссылки
                  s := '';
                  while (not IfTag('text:a', 6)) do
                  begin
                    if (xml.Eof()) then
                      break;
                    xml.ReadTag();
                    s := s + xml.TextBeforeTag;
                    if (xml.TagName <> 'text:a') then
                      s := s + xml.RawTextTag;
                  end;
                  _CurrCell.HRefScreenTip := s;
                end; //if

                //TODO: <text:span> - в будущем нужно будет как-то обрабатывать текст с
                //      форматированием, сейчас игнорируем
                _celltext := _celltext + xml.TextBeforeTag;
                if ((xml.TagName <> 'text:p') and (xml.TagName <> 'text:a') and (xml.TagName <> 'text:s') and
                    (xml.TagName <> 'text:span')) then
                  _celltext := _celltext +  xml.RawTextTag;
              end; //while
              _isnf := true;
            end; //if

            //Комментарий к ячейке
            if (IfTag('office:annotation', 4)) then
            begin
              s := xml.Attributes.ItemsByName['office:display'];
              _CurrCell.AlwaysShowComment := ZEStrToBoolean(s);
              s := '';
              _kol := 0;
              while (not IfTag('office:annotation', 6)) do
              begin
                if (xml.Eof()) then
                  break;
                xml.ReadTag();

                if (IfTag('dc:creator', 6)) then
                  _CurrCell.CommentAuthor := xml.TextBeforeTag;

                //dc:date - дата комментария, пока игнорируется

                //Текст примечания
                if (IfTag('text:p', 4)) then
                begin
                  while (not IfTag('text:p', 6)) do
                  begin
                    if (xml.Eof()) then
                      break;
                    xml.ReadTag();
                    if (_kol > 0) then
                      s := s + {$IFDEF FPC} LineEnding {$ELSE} sLineBreak {$ENDIF};
                    s := s + xml.TextBeforeTag;
                    inc(_kol);
                    {
                    if (xml.TagName <> 'text:p') then
                      s := s +  xml.RawTextTag;
                    }
                  end; //while
                end; //if
              end; //while
              _CurrCell.Comment := s;
              _currCell.ShowComment := s > '';
            end; //if

          end; //while *table-cell
        end; //if

        _CurrCell.Data := ZEReplaceEntity(_celltext);

        //Если ячейку нужно повторить
        if (isRepeatCell) then
          //Если в ячейке небыло текста и нужно повторить её более 255 раз - игнорируется повторение
          if (_isHaveTextCell) or (_RepeatCellCount < 255) then
          begin
            CheckCol(_CurrentPage, _CurrentCol + _RepeatCellCount + 1);
            for i := 1 to _RepeatCellCount do
              XMLSS.Sheets[_CurrentPage].Cell[_CurrentCol + i, _CurrentRow].Assign(_CurrCell);
            //-1, т.к. нужно учитывать, что номер ячейки увеличивается на 1 каждый раз
            inc(_CurrentCol, _RepeatCellCount - 1);
          end;

        inc(_CurrentCol);
      end; //if
    end; //_ReadCell

  begin
    _CurrentRow := 0; 
    _CurrentCol := 0;
    _MaxCol := 0;
    _CurrentPage := XMLSS.Sheets.Count;
    XMLSS.Sheets.Count := _CurrentPage + 1;
    XMLSS.Sheets[_CurrentPage].RowCount := 1;
    XMLSS.Sheets[_CurrentPage].Title := xml.Attributes.ItemsByName['table:name'];
    XMLSS.Sheets[_CurrentPage].Protect := ZEStrToBoolean(xml.Attributes.ItemsByName['table:protected']);

    s := xml.Attributes.ItemsByName['table:style-name'];
    if (s > '') then
      for i := 0 to TableStyleCount - 1 do
      if (ODFTableStyles[i].name = s) then
      begin
        if (ODFTableStyles[i].isColor) then
          XMLSS.Sheets[_CurrentPage].TabColor := ODFTableStyles[i].Color;
        break;
      end;

    while (not IfTag('table:table', 6)) do
    begin
      if (xml.Eof()) then
        break;
      xml.ReadTag();

      //Строка
      if ((xml.TagName = 'table:table-row') and (xml.TagType in [4, 5])) then
      begin
        //кол-во повторений строки
        s := xml.Attributes.ItemsByName['table:number-rows-repeated'];
        isRepeatRow := TryStrToInt(s, _RepeatRowCount);
        _IsHaveTextInRow := false;
        //стиль строки
        s := xml.Attributes.ItemsByName['table:style-name'];

        if (s > '') then
          for i := 0 to RowStyleCount - 1 do
            if (ODFRowStyles[i].name = s) then
            begin
              CheckRow(_CurrentPage, _CurrentRow + 1);
              XMLSS.Sheets[_CurrentPage].Rows[_CurrentRow].Breaked := ODFRowStyles[i].breaked;
              if (ODFRowStyles[i].height >= 0) then
                XMLSS.Sheets[_CurrentPage].Rows[_CurrentRow].HeightMM := ODFRowStyles[i].height;
            end;

        //стиль ячейки по умолчанию
        s := xml.Attributes.ItemsByName['table:default-cell-style-name'];

        //Видимость: visible | collapse | filter
        s := xml.Attributes.ItemsByName['table:visibility'];
        if (s = 'collapse') then
          XMLSS.Sheets[_CurrentRow].Rows[_CurrentRow].Hidden := true;

        if (xml.TagType = 5) then
        begin
          inc(_CurrentRow);
          CheckRow(_CurrentPage, _CurrentRow + 1);
          _RepeatRow();
        end;
        _CurrentCol := 0;
      end; //if
      if (IfTag('table:table-row', 6)) then
      begin
        inc(_CurrentRow);
        CheckRow(_CurrentPage, _CurrentRow + 1);
        _RepeatRow();
      end;

      //Ширина строки
      if ((xml.TagName = 'table:table-column') and (xml.TagType in [4, 5])) then
      begin
        CheckCol(_CurrentPage, _MaxCol + 1);
        s := xml.Attributes.ItemsByName['table:style-name'];
        for i := 0 to ColStyleCount - 1 do
          if (ODFColumnStyles[i].name = s) then
          begin
            XMLSS.Sheets[_CurrentPage].Columns[_MaxCol].Breaked := ODFColumnStyles[i].breaked;
            if (ODFColumnStyles[i].width >= 0) then
              XMLSS.Sheets[_CurrentPage].Columns[_MaxCol].WidthMM := ODFColumnStyles[i].width;
            break;
          end;

        s := xml.Attributes.ItemsByName['table:default-cell-style-name'];
        if (s > '') then
          XMLSS.Sheets[_CurrentPage].Columns[_MaxCol].StyleID := _FindStyleID(s);

        s := xml.Attributes.ItemsByName['table:number-columns-repeated'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            if (t < 255) then
            begin
              dec(t); //т.к. один столбец уже есть
              CheckCol(_CurrentPage, _MaxCol + t + 1);
              for i := 1 to t do
                XMLSS.Sheets[_CurrentPage].Columns[_MaxCol + i].Assign(XMLSS.Sheets[_CurrentPage].Columns[_MaxCol]);
              inc(_MaxCol, t)
            end;

        inc(_MaxCol);
      end;

      //ячейка
      _ReadCell();

    end; //while
    for i := 0 to XMLSS.Sheets[_CurrentPage].ColCount - 1 do
      XMLSS.Sheets[_CurrentPage].Columns[i].StyleID := -1;
  end; //_ReadTable

  procedure _ReadDocument();
  begin
    while (not xml.Eof()) do
    begin
      xml.ReadTag();
      ErrorReadCode := ErrorReadCode or xml.ErrorCode;
      if (ifTag('office:automatic-styles', 4)) then
        _ReadAutomaticStyle();
      //if (ifTag('office:styles', 4)) then
      //if (ifTag('office:master-styles', 4)) then

      if (ifTag('table:table', 4)) then
        _ReadTable();

    end;
  end; //_ReadDocument

begin
  result := false;
  xml := nil;
  ErrorReadCode := 0;
  ColStyleCount := 0;
  RowStyleCount := 0;
  MaxColStyleCount := -1;
  MaxRowStyleCount := -1;
  StyleCount := 0;
  MaxStyleCount := -1;
  TableStyleCount := 0;
  MaxTableStyleCount := -1;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(stream) <> 0) then
      exit;
    _ReadDocument();
    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
    SetLength(ODFColumnStyles, 0);
    ODFColumnStyles := nil;
    SetLength(ODFRowStyles, 0);
    ODFRowStyles := nil;
    SetLength(ODFStyles, 0);
    ODFStyles := nil;
    SetLength(ODFTableStyles, 0);
    ODFTableStyles := nil;
  end;
end; //ReadODFContent

//Чтение настроек документа ODS (settings.xml)
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//      stream: TStream - поток для чтения
//RETURN
//      boolean - true - всё ок
function ReadODFSettings(var XMLSS: TZEXMLSS; stream: TStream): boolean;
var
  xml: TZsspXMLReaderH;

  function _GetSplitModeByNum(const num: integer): TZSplitMode;
  begin
    result := ZSplitNone;
    case (num) of
      1: result := ZSplitSplit;
      2: result := ZSplitFrozen
    end;
  end; //_GetSplitModeByNum

  procedure _ReadSettingsPage();
  var
    _ConfigName: string;
    _ConfigType: string;
    _ConfigValue: string;
    _Sheet: TZSheet;
    s: string;
    i: integer;
    b: boolean;
    _intValue: integer;

    procedure _FindParam();
    begin
      if (_ConfigValue > '') then
      begin
        if (_ConfigName = 'CursorPositionX') then
        begin
          _Sheet.SheetOptions.ActiveCol := _intValue;
        end else
        if (_ConfigName = 'CursorPositionY') then
        begin
          _Sheet.SheetOptions.ActiveRow := _intValue;
        end else
        if (_ConfigName = 'HorizontalSplitMode') then
        begin
          _Sheet.SheetOptions.SplitVerticalMode := _GetSplitModeByNum(_intValue);
        end else
        if (_ConfigName = 'HorizontalSplitPosition') then
        begin
          _Sheet.SheetOptions.SplitVerticalValue := _intValue;
        end else
        if (_ConfigName = 'VerticalSplitMode') then
        begin
          _Sheet.SheetOptions.SplitHorizontalMode := _GetSplitModeByNum(_intValue);
        end else
        if (_ConfigName = 'VerticalSplitPosition') then
        begin
          _Sheet.SheetOptions.SplitHorizontalValue := _intValue;
        end;
      end; //if
    end; //_FindParam

  begin
    b := true;
    s := xml.Attributes.ItemsByName['config:name'];
    if (s = '') then
      exit;

    for i := 0 to XMLSS.Sheets.Count - 1 do
      if (XMLSS.Sheets[i].Title = s) then
      begin
        _Sheet := XMLSS.Sheets[i];
        b := false;
        break;
      end;
    if (b) then
      exit;

    while not ((xml.TagName = 'config:config-item-map-entry') and (xml.TagType = 6)) do
    begin
      if (xml.Eof()) then
        break;

      if (xml.TagName = 'config:config-item') then
      begin
        if (xml.TagType = 4) then
        begin
          _ConfigName := xml.Attributes.ItemsByName['config:name'];
          _ConfigType := xml.Attributes.ItemsByName['config:type'];
        end else
        begin
          _ConfigValue := xml.TextBeforeTag;
          if (TryStrToInt(_ConfigValue, _intValue)) then
            _FindParam();
        end; //if
      end; //if
      xml.ReadTag();
    end; //while
  end; //_ReadSettingsPage

  procedure _ReadSettings();
  begin
    while not ((xml.TagName = 'config:config-item-map-named') and (xml.TagType = 6)) do
    begin
      if (xml.Eof) then
        break;

      xml.ReadTag();
      _ReadSettingsPage();
    end; //while
  end; //_ReadSettings

begin
  result := false;
  xml := nil;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(stream) <> 0) then
      exit;

    while (not xml.Eof()) do
    begin
      xml.ReadTag();
      if ((xml.TagName = 'config:config-item-map-named') and (xml.TagType = 4)) then
        if (xml.Attributes.ItemsByName['config:name'] = 'Tables') then
          _ReadSettings();
    end; //while

    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
  end;
end; //ReadODFSettings

//Читает распакованный ODS
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//  DirName: string     - имя папки
//RETURN
//      integer - номер ошибки (0 - всё OK)
function ReadODFSPath(var XMLSS: TZEXMLSS; DirName: string): integer;
var
  stream: TStream;

begin
  result := 0;

  if (not ZE_CheckDirExist(DirName)) then
  begin
    result := -1;
    exit;
  end;

  XMLSS.Styles.Clear();
  XMLSS.Sheets.Count := 0;
  stream := nil;

  try
    //стили (styles.xml)
    ReadODFStyles(XMLSS, nil);

    //содержимое (content.xml)
    try
      stream := TFileStream.Create(DirName + 'content.xml', fmOpenRead or fmShareDenyNone);
    except
      result := 2;
      exit;
    end;
    if (not ReadODFContent(XMLSS, stream)) then
      result := result or 2;
    FreeAndNil(stream);

    //метаинформация (meta.xml)

    //настройки (settings.xml)
    try
      stream := TFileStream.Create(DirName + 'settings.xml', fmOpenRead or fmShareDenyNone);
    except
      result := 2;
      exit;
    end;
    if (not ReadODFSettings(XMLSS, stream)) then
      result := result or 2;
    FreeAndNil(stream);

  finally
    if (Assigned(stream)) then
      FreeAndNil(stream);
  end;
end; //ReadODFPath

{$IFDEF FPC}
//Читает ODS
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//  FileName: string    - имя файла
//RETURN
//      integer - номер ошибки (0 - всё OK)
function ReadODFS(var XMLSS: TZEXMLSS; FileName: string): integer;
var
  u_zip: TUnZipper;
  ZH: TODFZipHelper;
  lst: TStringList;

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
  ZH := nil;
  lst := nil;

  try
    lst := TStringList.Create();
    lst.Clear();
    lst.Add('content.xml'); //содержимое
    ZH := TODFZipHelper.Create();
    ZH.XMLSS := XMLSS;
    u_zip := TUnZipper.Create();
    u_zip.FileName := FileName;
    u_zip.OnCreateStream := @ZH.DoCreateOutZipStream;
    u_zip.OnDoneStream := @ZH.DoDoneOutZipStream;
    ZH.FileType := 0;
    u_zip.UnZipFiles(lst);
    result := result or ZH.RetCode;

    //метаинформация (meta.xml)

    //настройки (settings.xml)
    lst.Clear();
    lst.Add('settings.xml'); //настройки
    ZH.FileType := 1;
    u_zip.UnZipFiles(lst);

    result := result or ZH.RetCode;

  finally
    if (Assigned(u_zip)) then
      FreeAndNil(u_zip);
    if (Assigned(ZH)) then
      FreeAndNil(ZH);
    if (Assigned(lst)) then
      FreeAndNil(lst);
  end;
end; //ReadODFS
{$ENDIF}

{$IFNDEF FPC}
{$I odszipfuncimpl.inc}
{$ENDIF}

end.
