//****************************************************************
// Read/write Open Document Format (spreadsheet)
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2016.09.10
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
  SysUtils,
  Graphics,
  Classes,
  Types,
  zsspxml,
  zexmlss,
  zesavecommon,
  zenumberformats,        //TZEODSNumberFormatReader etc
  zeZippy
  {$IFDEF FPC},zipper {$ELSE}{$I odszipuses.inc}{$ENDIF};

type
  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  TZEConditionMap = record
    ConditionValue: string;       //Условие
    ApplyStyleName: string;       //Применяемый стиль
    ApplyStyleIDX: integer;       //Номер применяемого стиля
    ApplyBaseCellAddres: string   //Адрес ячейки
  end;
  {$ENDIF}

  //Доп. свойства стиля
  TZEODFStyleProperties = record
    name: string;           //Имя стиля
    index: integer;
    ParentName: string;     //Имя родителя
    isHaveParent: boolean;  //Флаг наличая родительского стиля
    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    ConditionsCount: integer;             //Кол-во условий (признак условного форматирования)
    Conditions: array of TZEConditionMap; //Условия
    {$ENDIF}
  end;

  TZODFStyleArray = array of TZEODFStyleProperties;

{$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  TZEODFCFLine = record
    CellNum: integer;
    StyleNumber: integer;
    Count: integer;
  end;

  TZEODFCFAreaItem = record
    RowNum: integer;
    ColNum: integer;
    Width: integer;
    Height: integer;
    CFStyleNumber: integer;
  end;

  TZEODFReadHelper = class;

  //Для чтения условного форматирования в ODS
  TZODFConditionalReadHelper = class
  private
    FXMLSS: TZEXMLSS;
    FCountInLine: integer;            //Кол-во условных стилей в текущей линии
    FMaxCountInLine: integer;
    FCurrentLine: array of TZEODFCFLine; //Текущая строка с условными стилями
    FColumnSCFNumbers: array of array [0..1] of integer;
    FColumnsCount: integer;           //Кол-во столбцов в листе
    FMaxColumnsCount: integer;        //Максимальное кол-во листов

    FAreas: array of TZEODFCFAreaItem;//Области с условным морматированием
    FAreasCount: integer;             //Кол-во областей с условным форматированием
    FMaxAreasCount: integer;

    FLineItemWidth: integer;          //Ширина текущей линии с одинаковым StyleID
    FLineItemStartCell: integer;      //Номер начальной ячейки в строке
    FLineItemStyleCFNumber: integer;  //Номер стиля
    FReadHelper: TZEODFReadHelper;
  protected
    procedure AddToLine(const CellNum: integer; const AStyleCFNumber: integer; const ACount: integer);
    function ODFReadGetConditional(const ConditionalValue: string;
                                out Condition: TZCondition;
                                out ConditionOperator: TZConditionalOperator;
                                out Value1: string;
                                out Value2: string): boolean;
  public
    constructor Create(XMLSS: TZEXMLSS);
    destructor Destroy(); override;
    procedure CheckCell(CellNum: integer; AStyleCFNumber: integer; RepeatCount: integer = 1);
    procedure ApplyConditionStylesToSheet(SheetNumber: integer;
                                          var DefStylesArray: TZODFStyleArray; DefStylesCount: integer;
                                          var StylesArray: TZODFStyleArray; StylesCount: integer);
    procedure AddColumnCF(ColumnNumber: integer; StyleCFNumber: integer);
    function GetColumnCF(ColumnNumber: integer): integer;
    procedure ApplyBaseCellAddr(const BaseCellTxt: string; const ACFStyle: TZConditionalStyleItem; PageNum: integer);
    procedure Clear();
    procedure ClearLine();
    procedure ProgressLine(RowNumber: integer; RepeatCount: integer = 1);
    procedure ReadCalcextTag(var xml: TZsspXMLReaderH; SheetNum: integer);
    property ColumnsCount: integer read FColumnsCount;
    property LineItemWidth: integer read FLineItemWidth write FLineItemWidth;
    property LineItemStartCell: integer read FLineItemStartCell write FLineItemStartCell;
    property LineItemStyleID: integer read FLineItemStyleCFNumber write FLineItemStyleCFNumber;
    property ReadHelper: TZEODFReadHelper read FReadHelper write FReadHelper;
  end;
{$ENDIF} //ZUSE_CONDITIONAL_FORMATTING read

  //Условное форматирование (для записи)
{$IFDEF ZUSE_CONDITIONAL_FORMATTING}

  TODFCFAreas = array of TZConditionalAreas;

  TODFCFmatch = record
    StyleID: integer;
    StyleCFID: integer;
  end;

  TODFStyleCFID = record
    Count: integer;
    ID: array of TODFCFmatch;
  end;

  TODFCFWriterArray = record
    CountCF: integer;                   //кол-во стилей
    StyleCFID: array of TODFStyleCFID;  //номер стиля c условным форматированием
    Areas: TODFCFAreas;                 //Области
  end;

  //Помошник для записи условного форматирования
  TZODFConditionalWriteHelper = class
  private
    FPagesCount: integer;                 //кол-во страниц
    FPageIndex: TIntegerDynArray;
    FPageNames: TStringDynArray;
    FPageCF: array of TODFCFWriterArray;
    FFirstCFIdInPage: TIntegerDynArray;
    FXMLSS: TZEXMLSS;
    FStylesCount: integer;                //кол-во стилей (которые с style:map)
    FApplyCFStylesCount: integer;         //кол-во стилей ConditionalStyle_f (в styles.xml)
    FMaxApplyCFStylesCount: integer;
    FApplyCFStyles: array of integer;     //массив с применяемымы стилями
  protected
    function GetBaseCellAddr(const StCondition: TZConditionalStyleItem;
                             const CurrPageName: string): string;
    function AddBetweenCond(const ConditName, Value1, Value2: string; out retCondition: string): boolean;
  public
    constructor Create(ZEXMLSS: TZEXMLSS;
                       const _pages: TIntegerDynArray;
                       const _names: TStringDynArray;
                       PageCount: integer);
    function TryAddApplyCFStyle(AStyleIndex: integer; out retCFIndex: integer): boolean;
    function GetApplyCFStyle(AStyleIndex: integer): integer;
    procedure WriteCFStyles(xml: TZsspXMLWriterH);
    procedure WriteCalcextCF(xml: TZsspXMLWriterH; PageIndex: integer);
    function ODSGetOperatorStr(AOperator: TZConditionalOperator): string;
    function GetStyleNum(const PageIndex, Col, Row: integer): integer;
    destructor Destroy(); override;
    property StylesCount: integer read FStylesCount;
  end;

{$ENDIF} //ZUSE_CONDITIONAL_FORMATTING write

  //предок для помошников чтения/записи
  TZEODFReadWriteHelperParent = class
  private
    FXMLSS: TZEXMLSS;
  protected
    property XMLSS: TZEXMLSS read FXMLSS;
  public
    constructor Create(AXMLSS: TZEXMLSS); virtual;
  end;

  TODSManifestMediaType = (
                          ZEODSMediaTypeSpreadSheet,
                          ZEODSMediaTypeTextXml,
                          ZEODSMediaTypeRdfXml,
                          ZEODSMediaTypeConfig,
                          ZEODSMediaTypeImagePng,
                          ZEODSMediaTypeChart,
                          ZEODSMediaTypeGdiMetaFile,
                          ZEODSMediaTypeNull,
                          ZEODSMediaTypeUnknown
                          );

  //Для чтения стилей и всего такого
  TZEODFReadHelper = class(TZEODFReadWriteHelperParent)
  private
    FStylesCount: integer;                  //Кол-во стилей
    FStyles: array of TZStyle;              //Стили из styles.xml
    FMasterPagesCount: integer;
    FMasterPages: array of TZSheetOptions;  //Array of master pages
    FMasterPagesNames: array of string;
    FPageLayoutsCount: integer;
    FPageLayouts: array of TZSheetOptions;
    FPageLayoutsNames: array of string;

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    FConditionReader: TZODFConditionalReadHelper; //Для чтения условного форматирования
    {$ENDIF}
    FNumberStylesHelper: TZEODSNumberFormatReader;
    function GetStyle(num: integer): TZStyle;
  protected
  public
    StylesProperties: TZODFStyleArray;
    constructor Create(AXMLSS: TZEXMLSS); override;
    destructor Destroy(); override;
    procedure ReadAutomaticStyles(xml: TZsspXMLReaderH);
    procedure ReadMasterStyles(xml: TZsspXMLReaderH);
    function ODSReadManifest(const stream: TStream): boolean;
    procedure AddStyle();
    procedure ApplyMasterPageStyle(SheetOptions: TZSheetOptions; const MasterPageName: string);
    property StylesCount: integer read FStylesCount;
    property Style[num: integer]: TZStyle read GetStyle;
    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    property ConditionReader: TZODFConditionalReadHelper read FConditionReader;
    {$ENDIF}
    property NumberStylesHelper: TZEODSNumberFormatReader read FNumberStylesHelper;
  end;

  //Для записи (пока нужно только условное форматирование)
  TZEODFWriteHelper = class(TZEODFReadWriteHelperParent)
  private
    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    FConditionWriter: TZODFConditionalWriteHelper;
    {$ENDIF}
    FUniquePageLayouts: array of integer;     //array of links to pages with unique page layouts
                                              //  -1 - default style
    FUniquePageLayoutsCount: integer;         //count of unique Page layouts
    FPageLayoutsIndexes: array of integer;    //Links to unique page layouts for each sheet

    FMasterPagesIndexes: array of integer;    //Links to unique masterpage style for each sheet
    FMasterPagesCount: integer;               //Count of unique masterpages styles
    FMasterPages: array of integer;           //array of links to pages with unique masterpages styles
                                              // -1 - default style
    FMasterPagesNames: array of string;       //array of masterpages styles names
    FNumberFormatWriter: TZEODSNumberFormatWriter; //Write and store numbers formats
  protected
  public
    constructor Create(AXMLSS: TZEXMLSS;
                       const _pages: TIntegerDynArray;
                       const _names: TStringDynArray;
                       PagesCount: integer); reintroduce; overload; virtual;
    destructor Destroy(); override;

    procedure WriteStylesPageLayouts(xml: TZsspXMLWriterH; const _pages: TIntegerDynArray);
    procedure WriteStylesMasterPages(xml: TZsspXMLWriterH; const _pages: TIntegerDynArray);
    function GetMasterPageName(PageNum: integer): string;

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    property ConditionWriter: TZODFConditionalWriteHelper read FConditionWriter;
    {$ENDIF}
    property NumberFormatWriter: TZEODSNumberFormatWriter read FNumberFormatWriter;
  end;

//Сохраняет незапакованный документ в формате Open Document
function SaveXmlssToODFSPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
//Сохраняет незапакованный документ в формате Open Document
function SaveXmlssToODFSPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string): integer; overload;
//Сохраняет незапакованный документ в формате Open Document
function SaveXmlssToODFSPath(var XMLSS: TZEXMLSS; PathName: string): integer; overload;

{$IFDEF FPC}
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string): integer; overload;
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
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
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring;
                          const WriteHelper: TZEODFWriteHelper): integer;

//Записывает в поток настройки (settings.xml)
function ODFCreateSettings(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;

//Записывает в поток документ + автоматические стили (content.xml)
function ODFCreateContent(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring;
                          const WriteHelper: TZEODFWriteHelper): integer;

//Записывает в поток метаинформацию (meta.xml)
function ODFCreateMeta(var XMLSS: TZEXMLSS; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;

{для чтения}
//Чтение содержимого документа ODS (content.xml)
function ReadODFContent(var XMLSS: TZEXMLSS; stream: TStream; var ReadHelper: TZEODFReadHelper): boolean;

//Чтение настроек документа ODS (settings.xml)
function ReadODFSettings(var XMLSS: TZEXMLSS; stream: TStream): boolean;

function GetODSMediaTypeByStr(const MediaType: string): TODSManifestMediaType;

function GetODSStrByMediaType(const MediaType: TODSManifestMediaType): string;

implementation

uses
   StrUtils
   {$IFDEF ZUSE_CONDITIONAL_FORMATTING}, zeformula {$ENDIF} //пока формулы нужны только для условного форматирования
   ;

const
  ZETag_text_p              = 'text:p';
  ZETag_StyleFontFace       = 'style:font-face';
  ZETag_StyleStyle          = 'style:style';
  ZETag_config_name         = 'config:name';
  ZETag_config_config_item_map_named = 'config:config-item-map-named';
  ZETag_office_automatic_styles = 'office:automatic-styles';
  ZETag_style_page_layout   = 'style:page-layout';
  ZETag_fo_margin_top       = 'fo:margin-top';
  ZETag_fo_margin_bottom    = 'fo:margin-bottom';
  ZETag_fo_margin_left      = 'fo:margin-left';
  ZETag_fo_margin_right     = 'fo:margin-right';
  ZETag_portrait            = 'portrait';
  ZETag_landscape           = 'landscape';
  ZETag_style_print_orientation = 'style:print-orientation';
  ZETag_fo_page_width       = 'fo:page-width';
  ZETag_fo_page_height      = 'fo:page-height';
  ZETag_fo_min_height       = 'fo:min-height';
  ZETag_style_header_style  = 'style:header-style';
  ZETag_style_footer_style  = 'style:footer-style';
  ZETag_fo_background_color = 'fo:background-color';
  ZETag_style_header_footer_properties = 'style:header-footer-properties';
  ZETag_style_page_layout_properties = 'style:page-layout-properties';
  ZETag_office_master_styles = 'office:master-styles';
  ZETag_style_master_page   = 'style:master-page';
  ZETag_style_page_layout_name = 'style:page-layout-name';
  ZETag_style_header        = 'style:header';
  ZETag_style_header_left   = 'style:header-left';
  ZETag_style_footer        = 'style:footer';
  ZETag_style_footer_left   = 'style:footer-left';
  ZETag_style_region_left   = 'style:region-left';
  ZETag_style_region_center = 'style:region-center';
  ZETag_style_region_right  = 'style:region-right';
  ZETag_style_family        = 'style:family';
  ZETag_style_display       = 'style:display';
  ZETag_style_master_page_name = 'style:master-page-name';
  ZETag_svg_height          = 'svg:height';

  ZETag_style_scale_to      = 'style:scale-to';
  ZETag_style_scale_to_pages = 'style:scale-to-pages';

  ZETag_table_style_name    = 'table:style-name';
  ZETag_style_table_properties = 'style:table-properties';
  ZETag_tableooo_tab_color  = 'tableooo:tab-color';

  ZETag_style_data_style_name = 'style:data-style-name';

  ZETag_style_use_optimal_column_width = 'style:use-optimal-column-width';
  ZETag_style_use_optimal_row_height = 'style:use-optimal-row-height';

  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  const_ConditionalStylePrefix      = 'ConditionalStyle_f';
  const_calcext_conditional_formats = 'calcext:conditional-formats';
  const_calcext_conditional_format  = 'calcext:conditional-format';
  const_calcext_value               = 'calcext:value';
  const_calcext_apply_style_name    = 'calcext:apply-style-name';
  const_calcext_base_cell_address   = 'calcext:base-cell-address';
  const_calcext_condition           = 'calcext:condition';
  const_calcext_target_range_address= 'calcext:target-range-address';
  {$ENDIF}

  //tags and atributes for manifest.xml
  ZETag_manifest_file_entry         = 'manifest:file-entry';
  ZETag_manifest_full_path          = 'manifest:full-path';
  ZETag_manifest_version            = 'manifest:version';
  ZETag_manifest_media_type         = 'manifest:media-type';

  const_ODS_paper_sizes_count = 42;

  //Sizes in 10^(-4) m! A4 = 2100*10^(-4) m x 2970*10^(-4) m
  const_ODS_paper_sizes: array [0..const_ODS_paper_sizes_count - 1] of array [0..1] of integer =
                  (
                    (0, 0),             // 0        Undefined
                    (2159, 2794),       // 1        Letter                   8 1/2" x 11"
                    (2159, 2794),       // 2        Letter small             8 1/2" x 11"
                    (2794, 4318),       // 3        Tabloid                     11" x 17"
                    (4318, 2794),       // 4        Ledger                      17" x 11"
                    (2159, 3556),       // 5        Legal                    8 1/2" x 14"
                    (1397, 2159),       // 6        Statement                5 1/2" x 8 1/2"
                    (1842, 2667),       // 7        Executive                7 1/4" x 10 1/2"
                    (2970, 4200),       // 8        A3                        297mm x 420mm
                    (2100, 2970),       // 9        A4                        210mm x 297mm
                    (2100, 2970),       // 10       A4 small                  210mm x 297mm
                    (1480, 2100),       // 11       A5                        148mm x 210mm
                    (2500, 3540),       // 12       B4                        250mm x 354mm
                    (1820, 2570),       // 13       B5                        182mm x 257mm
                    (2159, 3302),       // 14       Folio                    8 1/2" x 13"
                    (2150, 2750),       // 15       Quarto                    215mm x 275mm
                    (2540, 3556),       // 16                                   10" x 14"
                    (2794, 4318),       // 17                                   11" x 17"
                    (2159, 2794),       // 18       Note                     8 1/2" x 11"
                    (984, 2254),        // 19       #9 Envelope              3 7/8" x 8 7/8"
                    (1048, 2413),       // 20       #10 Envelope             4 1/8" x 9 1/2"
                    (1143, 2635),       // 21       #11 Envelope             4 1/2" x 10 3/8"
                    (1207, 2794),       // 22       #12 Envelope             4 3/4" x 11"
                    (1270, 2921),       // 23       #14 Envelope                 5" x 11 1/2"
                    (4318, 5588),       // 24       C Sheet                     17" x 22"
                    (5588, 8636),       // 25       D Sheet                     22" x 34"
                    (8636, 11176),      // 26       E Sheet                     34" x 44"
                    (1100, 2200),       // 27       DL Envelope               110mm x 220mm
                    (1620, 2290),       // 28       C5 Envelope               162mm x 229mm
                    (3240, 4580),       // 29       C3 Envelope               324mm x 458mm
                    (2290, 3240),       // 30       C4 Envelope               229mm x 324mm
                    (1140, 1620),       // 31       C6 Envelope               114mm x 162mm
                    (1140, 2290),       // 32       C65 Envelope              114mm x 229mm
                    (2500, 3530),       // 33       B4 Envelope               250mm x 353mm
                    (1760, 2500),       // 34       B5 Envelope               176mm x 250mm
                    (1250, 1760),       // 35       B6 Envelope               125mm x 176mm
                    (1100, 2300),       // 36       Italy Envelope            110mm x 230mm
                    (984, 1905),        // 37       Monarch Envelope         3 7/8" x 7 1/2"
                    (921, 1651),        // 38       6 3/4 Envelope           3 5/8" x 6 1/2"
                    (3778, 2794),       // 39       US Standard Fanfold     14 7/8" x 11"
                    (2159, 3048),       // 40       German Std. Fanfold      8 1/2" x 12"
                    (2159, 3302)        // 41       German Legal Fanfold     8 1/2" x 13"
                  );

   const_ODS_manifest_mediatypes_str: array [0..7] of string =
                  (
                    'application/vnd.oasis.opendocument.spreadsheet',  //ZEODSMediaTypeSpreadSheet
                    'text/xml',                                        //ZEODSMediaTypeTextXml
                    'application/rdf+xml',                             //ZEODSMediaTypeRdfXml
                    'application/vnd.sun.xml.ui.configuration',        //ZEODSMediaTypeConfig
                    'image/png',                                       //ZEODSMediaTypeImagePng,
                    'application/vnd.oasis.opendocument.chart',        //ZEODSMediaTypeChart,
                    'application/x-openoffice-gdimetafile;windows_formatname=&quot;GDIMetaFile&quot;',  //ZEODSMediaTypeGdiMetaFile,
                    ''                                                 //ZEODSMediaTypeNull,
                  );

  const_ODS_manifest_mediatypes: array [0..8] of TODSManifestMediaType =
                  (
                          ZEODSMediaTypeSpreadSheet,
                          ZEODSMediaTypeTextXml,
                          ZEODSMediaTypeRdfXml,
                          ZEODSMediaTypeConfig,
                          ZEODSMediaTypeImagePng,
                          ZEODSMediaTypeChart,
                          ZEODSMediaTypeGdiMetaFile,
                          ZEODSMediaTypeNull,
                          ZEODSMediaTypeUnknown
                  );

type
  TZODFColumnStyle = record
    name: string;     //имя стиля строки
    width: real;      //ширина
    breaked: boolean; //разрыв
    AutoWidth: boolean; //Optimal width
  end;

  TZODFColumnStyleArray = array of TZODFColumnStyle;

  TZODFRowStyle = record
    name: string;
    height: real;
    breaked: boolean;
    color: TColor;
    AutoHeight: boolean; //Optimal Height
  end;

  TZODFRowStyleArray = array of TZODFRowStyle;

  TZODFTableStyle = record
    name: string;
    isColor: boolean;
    Color: TColor;
    MasterPageName: string;
  end;

  TZODFTableArray = array of TZODFTableStyle;

{$IFDEF FPC}
  function ReadODFStyles(var XMLSS: TZEXMLSS; stream: TStream; var ReadHelper: TZEODFReadHelper): boolean; forward;

type
  //Для распаковки в поток
  TODFZipHelper = class
  private
    FXMLSS: TZEXMLSS;
    FRetCode: integer;
    FFileType: integer;
    FReadHelper: TZEODFReadHelper;
    procedure SetXMLSS(AXMLSS: TZEXMLSS);
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure DoCreateOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
    procedure DoDoneOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
    property XMLSS: TZEXMLSS read FXMLSS write SetXMLSS;
    property RetCode: integer read FRetCode;
    property FileType: integer read FFileType write FFileType;
  end;

constructor TODFZipHelper.Create();
begin
  inherited;
  FXMLSS := nil;
  FRetCode := 0;
  FReadHelper := nil;
end;

destructor TODFZipHelper.Destroy();
begin
  if (Assigned(FReadHelper)) then
    FreeAndNil(FReadHelper);

  inherited;
end;

procedure TODFZipHelper.SetXMLSS(AXMLSS: TZEXMLSS);
begin
  if (not Assigned(FReadHelper)) then
    FReadHelper := TZEODFReadHelper.Create(AXMLSS);

  FXMLSS := AXMLSS;
end;

procedure TODFZipHelper.DoCreateOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
begin
  AStream := TMemoryStream.Create();
end;

procedure TODFZipHelper.DoDoneOutZipStream(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry);
var
  _isError: boolean;

begin
  if (Assigned(AStream)) then
  try
    _isError := false;
    AStream.Position := 0;

    case (FileType) of
      0: _isError := not ReadODFContent(FXMLSS, AStream, FReadHelper);
      1: _isError := not ReadODFSettings(FXMLSS, AStream);
      2: _isError := not ReadODFStyles(FXMLSS, AStream, FReadHelper);
    end;

    if (_isError) then
      FRetCode := FRetCode or 2;
  finally
    FreeAndNil(AStream)
  end;
end; //DoDoneOutZipStream

{$ENDIF}

//Convert double to string for ODF sizes (like 23.95cm etc)
//INPUT
//      AValue: double
//  const AFormat: string
//  const APostfix: string
//RETURN
//      string
function ODFGetSizeToStr(AValue: double; const AFormat: string = '0.###'; const APostfix: string = 'cm'): string;
begin
  Result := ZEFloatSeparator(FormatFloat(AFormat, AValue)) + APostfix;
end; //ODFGetSizeToStr

{$IFDEF ZUSE_CONDITIONAL_FORMATTING}

procedure ODFWriteTableStyle(var XMLSS: TZEXMLSS; _xml: TZsspXMLWriterH; const StyleNum: integer; isDefaultStyle: boolean); forward;
function ODFGetValueSizeMM(const value: string; out RetSize: real; isMultiply: boolean = true): boolean; forward;
function GetBGColorForODS(const value: string): TColor; forward;

////::::::::::::: TZODFConditionalWriteHelper :::::::::::::::::////

//конструктор
//INPUT
//      ZEXMLSS: TZEXMLSS           - хранилище
//  const _pages: TIntegerDynArray  - индексы страниц
//  const _names: TStringDynArray   - названия страниц
//      PageCount: integer          - кол-во страниц
constructor TZODFConditionalWriteHelper.Create(ZEXMLSS: TZEXMLSS;
                                               const _pages: TIntegerDynArray;
                                               const _names: TStringDynArray;
                                               PageCount: integer);
var
  i: integer;

begin
  FStylesCount := 0;
  SetLength(FPageIndex, PageCount);
  Setlength(FPageNames, PageCount);
  SetLength(FPageCF, PageCount);
  SetLength(FFirstCFIdInPage, PageCount);
  FPagesCount := PageCount;
  for i := 0 to FPagesCount - 1 do
  begin
    FPageindex[i] := _pages[i];
    FPageNames[i] := _names[i];
    FPageCF[i].CountCF := 0;
    FFirstCFIdInPage[i] := 0;
  end;
  FXMLSS := ZEXMLSS;
  FApplyCFStylesCount := 0;
  FMaxApplyCFStylesCount := 20;
  SetLength(FApplyCFStyles, FMaxApplyCFStylesCount);
end; //Create

destructor TZODFConditionalWriteHelper.Destroy();
var
  i, j: integer;

begin
  SetLength(FPageIndex, 0);
  FPageIndex := nil;
  SetLength(FPageNames, 0);
  FPageNames := nil;

  for i := 0 to FPagesCount - 1 do
  begin
    for j := 0 to FPageCF[i].CountCF - 1 do
    begin
      SetLength(FPageCF[i].StyleCFID[j].ID, 0);
      FPageCF[i].StyleCFID[j].ID := nil;
    end;

    SetLength(FPageCF[i].StyleCFID, 0);
    FPageCF[i].StyleCFID := nil;
    SetLength(FPageCF[i].Areas, 0);
    FPageCF[i].Areas := nil;
  end;

  SetLength(FPageCF, 0);
  FPageCF := nil;
  SetLength(FFirstCFIdInPage, 0);
  SetLength(FApplyCFStyles, 0);
  inherited
end; //Destroy

//Получить текст оператора
function TZODFConditionalWriteHelper.ODSGetOperatorStr(AOperator: TZConditionalOperator): string;
begin
  case (AOperator) of
    ZCFOpGT: result := '>';
    ZCFOpLT: result := '<';
    ZCFOpGTE: result := '>=';
    ZCFOpLTE: result := '<=';
    ZCFOpEqual: result := '=';
    ZCFOpNotEqual: result := '!=';
    else
      result := '';
  end;
end; //ODSGetOperatorStr

//style:base-cell-address / calcext:base-cell-address
//INPUT
//  const StCondition: TZConditionalStyleItem - условие
//  const CurrPageName: string                - название листа по дефолту
//RETURN
//      string - текст базовой ячейки
function TZODFConditionalWriteHelper.GetBaseCellAddr(const StCondition: TZConditionalStyleItem;
                                                     const CurrPageName: string): string;
var
  s: string;
  b: boolean;
  i: integer;

begin
  if ((StCondition.BaseCellPageIndex < 0) or (StCondition.BaseCellPageIndex >= FPagesCount)) then
    s := CurrPageName
  else
  begin
    b := false;
    for i := 0 to FPagesCount - 1 do
      if (FPageIndex[i] = StCondition.BaseCellPageIndex) then
      begin
        s := FPageNames[i];
        b := true;
        break;
      end;
    if (not b) then
      s := currPageName;
  end;
  if (pos(' ', s) <> 0) then
    s := '''' + s + '''';
  //listname.ColRow
  s := s + '.' + ZEGetA1byCol(StCondition.BaseCellColumnIndex) + IntToStr(StCondition.BaseCellRowIndex + 1);
  result := s;
end; //GetBaseCellAddr

function TZODFConditionalWriteHelper.AddBetweenCond(const ConditName, Value1, Value2: string; out retCondition: string): boolean;
var
  t1, t2: double;
  b: boolean;

begin
  result := false;
  retCondition := '';
  t1 := ZETryStrToFloat(Value1, b);
  if (b) then
  begin
    t2 := ZETryStrToFloat(Value2, b);
    if (b and (t1 <= t2)) then
    begin
      result := true;
      retCondition := ConditName + '(' +
           ZEFloatSeparator(FormatFloat('', t1)) + ',' +
           ZEFloatSeparator(FormatFloat('', t2)) + ')'
    end;
  end;
end; //AddBetweenCond

//Получить применяемый стиль для условного форматирования по индексу
function TZODFConditionalWriteHelper.GetApplyCFStyle(AStyleIndex: integer): integer;
var
  i: integer;

begin
  result := -1;
  for i := 0 to FApplyCFStylesCount - 1 do
    if (FApplyCFStyles[i] = AStyleIndex) then
    begin
      result := i;
      break;
    end;
end; //GetApplyCfStyle

//Добавить в массив FApplyCFStyles (уникальные CF в styles.xml) стиль
//INPUT
//      AStyleIndex: integer  - применяемый стиль
//  out retCFIndex: integer   - возвращаемый индекс условного стиля в styles.xml
//                              (ConditionalStyle_f + IntToStr(retCFIndex))
//RETURN
//      boolean - true - стиль успешно добавился, можно писать в xml
function TZODFConditionalWriteHelper.TryAddApplyCFStyle(AStyleIndex: integer; out retCFIndex: integer): boolean;
var
  i: integer;

begin
  result := true;
  retCFIndex := -1;
  for i := 0 to FApplyCFStylesCount - 1 do
    if (FApplyCFStyles[i] = AStyleIndex) then
    begin
      result := false;
      break;
    end;

  if (result) then
  begin
    retCFIndex := FApplyCFStylesCount;
    inc(FApplyCFStylesCount);
    if (FApplyCFStylesCount >= FMaxApplyCFStylesCount) then
    begin
      inc(FMaxApplyCFStylesCount, 20);
      SetLength(FApplyCFStyles, FMaxApplyCFStylesCount);
    end;
    FApplyCFStyles[retCFIndex] := AStyleIndex;
  end;
end; //AddApplyCFStyle

//Запись стилей с условным форматированием в content.xml
// (тэги <style:style ...>...<style:map .../>...</style:style>)
//INPUT
//      xml: TZsspXMLWriterH - куда записывать
procedure TZODFConditionalWriteHelper.WriteCFStyles(xml: TZsspXMLWriterH);
var
  i, j: integer;
  _sheet: TZSheet;
  _CFStyle: TZConditionalStyle;
  _StyleID: integer;
  _kol: integer;
  _StCondition: TZConditionalStyleItem;
  _att: TZAttributesH;
  v1, v2: string;
  d1: Double;
  b: boolean;
  _currPageName: string;
  s: string;
  num: integer;
  _BeforeCFStyles: TIntegerDynArray;
  _BeforeCFStyleCount: integer;
  _BeforeCFStyleMaxCount: integer;

  CFMaps: array of array [0..2] of string;
  CFMapsCount: integer;
  CFMapsMaxCount: integer;

  //Добавить условие
  //INPUT
  //      mapnum: integer - номер условия
  //RETURN
  //      boolean - true - условие адекватное, добавляем
  function _AddMapCondition(mapnum: integer): boolean;

    function _AddContentOperator(): boolean;
    begin
      result := true;
      d1 := ZETryStrToFloat(v1, b);
      if (b) then
        s := ZEFloatSeparator(FormatFloat('', d1))
      else
        //TODO: на случай, если ввели v1 с кавычками, нужно будет потом сделать проверку
        s := '''' + v1 + '''';
      s := 'cell-content()' + ODSGetOperatorStr(_StCondition.ConditionOperator) + s;
    end; //_AddContentOperator()

    function _AddIsFormula(): boolean;
    begin
      result := false;
      //TODO: добавить проверку на валидность формулы!!!
      if (length(v1) > 0) then
      begin
        s := 'is-true-formula(' + v1 + ')';
        result := true;
      end;
    end; //_AddIsFormula

  begin
    result := false;
    _StCondition := _CFStyle[mapnum];
    inc(num);

    v1 := _StCondition.Value1;
    v2 := _StCondition.Value2;

    s := '';
    case _StCondition.Condition of
      ZCFIsTrueFormula:
        result := _AddIsFormula();
      ZCFCellContentIsBetween:
        result := AddBetweenCond('cell-content-is-between', v1, v2, s);
      ZCFCellContentIsNotBetween:
        result := AddBetweenCond('cell-content-is-not-between', v1, v2, s);
      ZCFCellContentOperator: result := _AddContentOperator();
      ZCFNumberValue:;
      ZCFString:;
      ZCFBoolTrue:;
      ZCFBoolFalse:;
      ZCFFormula:;
    end; //case

    if (result) then
    begin
      if (CFMapsCount + 1 >= CFMapsMaxCount) then
      begin
        inc(CFMapsMaxCount, 10);
        SetLength(CFMaps, CFMapsMaxCount);
      end;
      CFMaps[CFMapsCount][0] := s;
      if ((_StCondition.ApplyStyleID >= 0) and (_StCondition.ApplyStyleID < FXMLSS.Styles.Count)) then
      begin
        CFMaps[CFMapsCount][1] := const_ConditionalStylePrefix + IntToStr(GetApplyCFStyle(_StCondition.ApplyStyleID));
        CFMaps[CFMapsCount][2] := GetBaseCellAddr(_StCondition, _currPageName);
        inc(CFMapsCount);
      end else
        result := false;
    end;
  end; //_AddMapCondition

  procedure _CheckBorders(var brd: integer; maxborder: integer);
  begin
    if (brd >= maxborder) then
        brd := maxborder - 1;
      if (brd < 0) then
        brd := 0;
  end; //_CheckBorders

  //Получает все стили из области с условным форматированием
  procedure _GetAreasStyles();
  var
    i, j, k: integer;
    _cs, _ce: integer;
    _rs, _re: integer;
    t: integer;
    b: boolean;
    _StyleID: integer;

  begin
    _BeforeCFStyleCount := 0;
    for i := 0 to _CFStyle.Areas.Count - 1 do
    begin
      _cs := _CFStyle.Areas[i].Column;
      _CheckBorders(_cs, _sheet.ColCount);
      _ce := _cs + _CFStyle.Areas[i].Width - 1;
      _CheckBorders(_ce, _sheet.ColCount);

      _rs := _CFStyle.Areas[i].Row;
      _CheckBorders(_cs, _sheet.RowCount);
      _re := _rs + _CFStyle.Areas[i].Height - 1;
      _CheckBorders(_re, _sheet.RowCount);

      for j := _cs to _ce do
      for k := _rs to _re do
      begin
        b := true;
        _StyleID := _sheet.Cell[j, k].CellStyle;
        if (_styleID < 0) or (_StyleID >= FXMLSS.Styles.Count) then
          _StyleID := -1;
        for t := 0 to _BeforeCFStyleCount - 1 do
          if (_BeforeCFStyles[t] = _StyleID) then
          begin
            b := false;
            break;
          end;
        if (b) then
        begin
          if (_BeforeCFStyleCount + 1 >= _BeforeCFStyleMaxCount) then
          begin
            inc(_BeforeCFStyleMaxCount);
            SetLength(_BeforeCFStyles, _BeforeCFStyleMaxCount);
          end;
          _BeforeCFStyles[_BeforeCFStyleCount] := _StyleID;
          inc(_BeforeCFStyleCount);
        end; //if
      end; //for k
    end; // for i
  end; //_GetAreasStyles

  //Добавить стиль условного форматирования
  //  (текущий итем условного форматирования берётся из _CFStyle)
  //INPUT
  //      idx: integer - "индекс" страницы
  procedure _AddCFStyle(idx: integer);
  var
    i, j: integer;
    _addedCount: integer; //кол-во добавленных условий

  begin
    _addedCount := 0;
    CFMapsCount := 0;
    //добавляем стиль для условного форматирования
    //  обычно будет максимум 1-2 условное форматирование,
    //    поэтому пока не паримся насчёт SetLength.
    //  TODO: если в будущем понадобится добавлять большое кол-во
    //        условных форматирований, то нужно будет чуть оптимизировать.
    inc(FPageCF[idx].CountCF);
    _kol := FPageCF[idx].CountCF;
    SetLength(FPageCF[idx].StyleCFID, _kol);
    SetLength(FPageCF[idx].Areas, _kol);

    _GetAreasStyles();

    FPageCF[idx].StyleCFID[_kol - 1].Count := _BeforeCFStyleCount;
    SetLength(FPageCF[idx].StyleCFID[_kol - 1].ID, _BeforeCFStyleCount);
    FPageCF[idx].Areas[_kol - 1] := _CFStyle.Areas;

    for i := 0 to _CFStyle.Count - 1 do
      if (_AddMapCondition(i)) then
        inc(_addedCount);

    if (_addedCount > 0) then
    begin
      for i := 0 to _BeforeCFStyleCount - 1 do
      begin
        FPageCF[idx].StyleCFID[_kol - 1].ID[i].StyleID := _BeforeCFStyles[i];
        FPageCF[idx].StyleCFID[_kol - 1].ID[i].StyleCFID := _StyleID;

        _att.Clear();
        _att.Add(ZETag_Attr_StyleName, 'ce' + IntToStr(_StyleID));
        _att.Add(ZETag_style_family, 'table-cell');
        _att.Add('style:parent-style-name', 'Default');
        xml.WriteTagNode(ZETag_StyleStyle, _att, true, true, false);
        ODFWriteTableStyle(FXMLSS, xml, _BeforeCFStyles[i], false);

        for j := 0 to CFMapsCount - 1 do
        begin
          xml.Attributes.Clear();
          xml.Attributes.Add(ZETag_style_condition, CFMaps[j][0]);
          xml.Attributes.Add(ZETag_style_apply_style_name, CFMaps[j][1]);
          xml.Attributes.Add('style:base-cell-address', CFMaps[j][2]);
          xml.WriteEmptyTag(ZETag_style_map, true, true);
        end;

        xml.WriteEndTagNode(); //style:style
        inc(_StyleID);
        inc(FStylesCount);
      end; //for i
    end else
    begin
      //уменьшаем кол-во стилей
      dec(FPageCF[idx].CountCF);
    end;
  end; //_AddCFStyle

begin
  if (not Assigned(xml)) then
    exit;
  if (not Assigned(FXMLSS)) then
    exit;
  _att := nil;
  num := 0;

  _BeforeCFStyleMaxCount := 10;
  SetLength(_BeforeCFStyles, _BeforeCFStyleMaxCount);

  CFMapsMaxCount := 10;
  SetLength(CFMaps, CFMapsMaxCount);

  try
    _att := TZAttributesH.Create();
    _StyleID := FXMLSS.Styles.Count;
    for i := 0 to FPagesCount - 1 do
    begin
      _sheet := FXMLSS.Sheets[FPageIndex[i]];
      _currPageName := FPageNames[i];
      FFirstCFIdInPage[i] := num;
      if (_sheet.ConditionalFormatting.Count > 0) then
        for j := 0 to _sheet.ConditionalFormatting.Count - 1 do
        begin
          _CFStyle := _sheet.ConditionalFormatting[j];
          if ((_CFStyle.Count > 0) and (_CFStyle.Areas.Count > 0)) then
            _AddCFStyle(i);
        end; //for
    end; //for i
  finally
    if (Assigned(_att)) then
      FreeAndNil(_att);
   SetLength(_BeforeCFStyles, 0);
   _BeforeCFStyles := nil;

   Setlength(CFMaps, 0);
   CFMaps := nil;
  end;
end; //WriteCFStyles

//Пишет условное форматирование <calcext:conditional-formats> </calcext:conditional-formats>
// для LibreOffice
//INPUT
//      xml: TZsspXMLWriterH  - писатель
//      PageIndex: integer    - номер страницы
procedure TZODFConditionalWriteHelper.WriteCalcextCF(xml: TZsspXMLWriterH; PageIndex: integer);
var
  i: integer;
  _cfCount: integer;
  _CF: TZConditionalFormatting;
  StartStyleNum: integer;
  _PageName: string;

  procedure _WriteFormat(Num: integer);
  var
    i: integer;
    _StyleItem: TZConditionalStyle;
    kol: integer;

    procedure _WriteFormatItem(StyleItem: TZConditionalStyleItem);
    var
      _condition: string;

      function _GetTextCondition(const CName: string; out retCondition: string): boolean;
      var
        l: integer;
        s: string;
        b: boolean;

      begin
        result := true;
        s := StyleItem.Value1;
        l := Length(s);
        b := true;
        if (l >= 2) then
          if ((s[1] = '"') and (s[l] = '"')) then
            b := false;
        if (b) then
          s := '"' + s + '"';
        retCondition := CName + '(' + s + ')';
      end; //_GetTextCondition

      function _GetContentOperator(out retCondition: string): boolean;
      var
        t: double;
        b: boolean;

      begin
        result := true;
        t := ZETryStrToFloat(StyleItem.Value1, b);
        if (b) then
          retCondition := ZEFloatSeparator(FormatFloat('', t))
        else
          retCondition := '''' + StyleItem.Value1 + '''';

        retCondition := ODSGetOperatorStr(StyleItem.ConditionOperator) + retCondition;
      end; //_GetContentOperator

      function _GetSimpleText(const ConditName: string; out retCondition: string): boolean;
      begin
        result := true;
        retCondition := ConditName;
      end; //_GetSimpleText

      function _GetConditionOneNumber(const ConditName: string; out retCondition: string; isFloat: boolean = false): boolean;
      var
        tI: integer;
        tF: double;

      begin
        if (isFloat) then
        begin
          tF := ZETryStrToFloat(StyleItem.Value1, result);
          if (result) then
            retCondition := ConditName + '(' + ZEFloatSeparator(FormatFloat('', tF)) + ')';
        end else
        begin
          result := TryStrToInt(StyleItem.Value1, tI);
          if (result) then
            retCondition := ConditName + '(' + StyleItem.Value1 + ')';
        end;
      end; //_GetConditionOneNumber

      function _GetIsTrueFormula(out retCondition: string): boolean;
      begin
        //TODO: добавить проверку на валидность формулы!
        if (Length(StyleItem.Value1) > 0) then
        begin
          result := true;
          retCondition := 'formula-is(' + StyleItem.Value1 + ')';
        end else
          result := false;
      end; //_GetIsTrueFormula

      function isGetCondition(out retCondition: string): boolean;
      begin
        result := false;

        case (Styleitem.Condition) of
          ZCFIsTrueFormula:           result := _GetIsTrueFormula(retCondition);
          ZCFCellContentIsBetween:    result := AddBetweenCond('between', StyleItem.Value1, StyleItem.Value2, retCondition);
          ZCFCellContentIsNotBetween: result := AddBetweenCond('not-between', StyleItem.Value1, StyleItem.Value2, retCondition);
          ZCFCellContentOperator:     result := _GetContentOperator(retCondition);
          ZCFNumberValue:;
          ZCFString:;
          ZCFBoolTrue:;
          ZCFBoolFalse:;
          ZCFFormula:;
          ZCFContainsText:    result := _GetTextCondition('contains-text', retCondition);
          ZCFNotContainsText: result := _GetTextCondition('not-contains-text', retCondition);
          ZCFBeginsWithText:  result := _GetTextCondition('begins-with', retCondition);
          ZCFEndsWithText:    result := _GetTextCondition('ends-with', retCondition);
          ZCFCellIsEmpty:;
          ZCFDuplicate:       result := _GetSimpleText('duplicate', retCondition);
          ZCFUnique:          result := _GetSimpleText('unique', retCondition);
          ZCFAboveAverage:    result := _GetSimpleText('above-average', retCondition);
          ZCFBellowAverage:   result := _GetSimpleText('below-average', retCondition);
          ZCFAboveEqualAverage: result := _GetSimpleText('above-equal-average', retCondition);
          ZCFBelowEqualAverage: result := _GetSimpleText('below-equal-average', retCondition);
          ZCFTopElements:     result := _GetConditionOneNumber('top-elements', retCondition);
          ZCFBottomElements:  result := _GetConditionOneNumber('bottom-elements', retCondition);
          ZCFTopPercent:      result := _GetConditionOneNumber('top-percent', retCondition, true);
          ZCFBottomPercent:   result := _GetConditionOneNumber('bottom-percent', retCondition, true);
          ZCFIsError:         result := _GetSimpleText('is-error', retCondition);
          ZCFIsNoError:       result := _GetSimpleText('is-no-error', retCondition);
        end;
      end; //_GetCondition

    begin
      if (isGetCondition(_condition)) then
      begin
        inc(StartStyleNum);
        xml.Attributes.Clear();
        xml.Attributes.Add(const_calcext_apply_style_name, const_ConditionalStylePrefix + IntToStr(GetApplyCFStyle(StyleItem.ApplyStyleID)));
        xml.Attributes.Add(const_calcext_value, _condition);
        xml.Attributes.Add(const_calcext_base_cell_address, GetBaseCellAddr(StyleItem, _PageName));
        xml.WriteEmptyTag(const_calcext_condition, true);
      end;

       {
       const_calcext_value               = 'calcext:value';
       const_calcext_apply_style_name    = 'calcext:apply-style-name';
       const_calcext_base_cell_address   = 'calcext:base-cell-address';
       const_calcext_condition           = 'calcext:condition';
       }
    end; //_WriteFormatItem

    function _GetRanges(): string;
    var
      i: integer;
      s: string;
      n: integer;

    begin
      result := '';
      s :=  _PageName;
      if (pos(' ', s) <> 0) then
        s := '''' + s + '''';
      n := _StyleItem.Areas.Count - 1;
      for i := 0 to n do
      begin
        result := result +
                  s + '.' + ZEGetA1byCol(_StyleItem.Areas[i].Column) + IntToStr(_StyleItem.Areas[i].Row + 1) +
                  ':' +
                  s + '.' + ZEGetA1byCol(_StyleItem.Areas[i].Column + _StyleItem.Areas[i].Width - 1) + IntToStr(_StyleItem.Areas[i].Row + _StyleItem.Areas[i].Height);
        if (i <> n) then
          result := result + ' ';
      end;
    end; //_GetRanges

  begin
    _StyleItem := _CF.Items[Num];
    kol := _StyleItem.Count;
    if (kol > 0) then
    begin
      xml.Attributes.Clear();
      xml.Attributes.Add(const_calcext_target_range_address, _GetRanges());
      xml.WriteTagNode(const_calcext_conditional_format, true, true, true);
      for i := 0 to kol - 1 do
        _WriteFormatItem(_StyleItem.Items[i]);

      xml.WriteEndTagNode(); //calcext:conditional-format
    end;
  end; //_WriteFormat

begin
  if (PageIndex >= 0) and (PageIndex < FPagesCount) then
  begin
    _PageName := FPageNames[PageIndex];
    _CF := FXMLSS.Sheets[FPageIndex[PageIndex]].ConditionalFormatting;
    StartStyleNum := FFirstCFIdInPage[PageIndex];
    _cfCount := _CF.Count;
    if (_cfCount > 0) then
    begin
       xml.Attributes.Clear();
       xml.WriteTagNode(const_calcext_conditional_formats, true, true, false);

       for i := 0 to _cfCount - 1 do
         _WriteFormat(i);

       xml.WriteEndTagNode(); //calcext:conditional-formats
    end;
  end;
end; //WriteCalcextCF

//Получить номер стиля (с условным форматированием или без) ячейки
//  подразумевается, что индекс страницы и координаты ячейки правильные и
//  не выходят за границы.
//INPUT
//  const PageIndex: integer  - индекс страницы
//  const Col: integer        - номер столбца
//  const Row: integer        - номер строки
//RETURN
//      integer - номер стиля (-1 - стиль по умолчанию)
function TZODFConditionalWriteHelper.GetStyleNum(const PageIndex, Col, Row: integer): integer;
var
  i, j: integer;

begin
  result := FXMLSS.Sheets[FPageIndex[PageIndex]].Cell[Col, Row].CellStyle;
  if (result < -1) then
    result := -1;
  if (FPageCF[PageIndex].CountCF > 0) then
    for i := 0 to FPageCF[PageIndex].CountCF - 1 do
      if (FPageCF[PageIndex].Areas[i].IsCellInArea(Col, Row)) then
      begin
        for j := 0 to FPageCF[PageIndex].StyleCFID[i].Count - 1 do
          if (FPageCF[PageIndex].StyleCFID[i].ID[j].StyleID = result) then
          begin
            result := FPageCF[PageIndex].StyleCFID[i].ID[j].StyleCFID;
            exit;
          end;
        exit;
      end;
end; //GetStyleNum

////::::::::::::: TZODFConditionalReadHelper :::::::::::::::::////

constructor TZODFConditionalReadHelper.Create(XMLSS: TZEXMLSS);
begin
  FXMLSS := XMLSS;
  FCountInLine := 0;
  FMaxCountInLine := 10;
  SetLength(FCurrentLine, FMaxCountInLine);
  FColumnsCount := 0;
  FMaxColumnsCount := 10;
  SetLength(FColumnSCFNumbers, FMaxColumnsCount);
  FAreasCount := 0;
  FMaxAreasCount := 10;
  SetLength(FAreas, FMaxAreasCount);
  FReadHelper := nil;
end;

destructor TZODFConditionalReadHelper.Destroy();
begin
  SetLength(FCurrentLine, 0);
  SetLength(FColumnSCFNumbers, 0);
  SetLength(FAreas, 0);
  inherited;
end;

//Проверить условный стиль для текущей ячейки и заполить линию условных стилей
//INPUT
//      CellNum: integer        - номер текущей ячейки
//      AStyleCFNumber: integer - номер условного стиля ячейки (из массивов)
//      RepeatCount: integer    - кол-во повторений
procedure TZODFConditionalReadHelper.CheckCell(CellNum: integer; AStyleCFNumber: integer; RepeatCount: integer = 1);
var
  _add: boolean;

  procedure _AddLineItem();
  begin
    if (AStyleCFNumber >= 0) then
    begin
      FLineItemWidth := RepeatCount;
      FLineItemStartCell := CellNum;
      FLineItemStyleCFNumber := AStyleCFNumber;
    end else
      FLineItemStartCell := -2;
  end;

begin
  //Если первая ячейка в строке или если предыдущий стиль был не условным
  if (FLineItemStartCell < 0) then
  begin
    //стиль должен быть только условным
    _AddLineItem();
  end else
  begin
    _add := false;
    if ((AStyleCFNumber >= 0) and (FLineItemStyleCFNumber = AStyleCFNumber)) then
      FLineItemWidth := FLineItemWidth + RepeatCount
    else
      _add := true;

    if (_add) then
    begin
      AddToLine(FLineItemStartCell, FLineItemStyleCFNumber, FLineItemWidth);
      _AddLineItem();
    end;
  end;
end; //CheckCell

//Добавить к текущей линии
//INPUT
//  const CellNum: integer        - номер начальной ячейки
//  const AStyleCFNumber: integer - номер стиля в хранилище
//  const ACount: integer         - кол-во ячеек
procedure TZODFConditionalReadHelper.AddToLine(const CellNum: integer; const AStyleCFNumber: integer; const ACount: integer);
var
  t: integer;

begin
  t := FCountInLine;
  inc(FCountInLine);
  if (FCountInLine >= FMaxCountInLine) then
  begin
    inc(FMaxCountInLine, 10);
    SetLength(FCurrentLine, FMaxCountInLine);
  end;
  FCurrentLine[t].CellNum := CellNum;
  FCurrentLine[t].StyleNumber := AStyleCFNumber;
  FCurrentLine[t].Count := ACount;
end; //AddToLine

//Получить из текста условия (style:condition) уловие, оператор и значения
//INPUT
//  const ConditionalValue: string
//  out Condition: TZCondition
//  out ConditionOperator: TZConditionalOperator
//  out Value1: string
//  out Value2: string
//RETURN
//      boolean - true - уловие успешно определено
function TZODFConditionalReadHelper.ODFReadGetConditional(const ConditionalValue: string;
                                out Condition: TZCondition;
                                out ConditionOperator: TZConditionalOperator;
                                out Value1: string;
                                out Value2: string): boolean;
var
  i: integer;
  len: integer;
  s: string;
  ch: char;
  kol: integer;
  _OCount: integer;
  _strArr: array of string;
  _maxKol: integer;
  _isFirstOperator: boolean;

  //Заполняет строку retStr до тех пор, пока не встретит символ=ch или не дойдёт до конца
  //INPUT
  //  var retStr: string  - результирующая строка
  //  var num: integer    - позиция текущего символа в исходной строке
  //      ch: char        - до кокого символа просматривать
  procedure _ReadWhileNotChar(var retStr: string; var num: integer; ch: char);
  begin
    while (num < len - 1) do
    begin
      inc(num);
      if (ConditionalValue[num] = ch) then
        break
      else
        retStr := retStr + ConditionalValue[num];
    end; //while
  end; //_ReadWhileNotChar

  procedure _ProcessBeforeDelimiter();
  begin
    if (length(trim(s)) > 0) then
    begin
      _strArr[kol] := s;
      inc(kol);
      if (kol >= _maxKol) then
      begin
        inc(_maxKol, 5);
        SetLength(_strArr, _maxKol);
      end;
      s := '';
    end;
  end; //_ProcessBeforeDelimiter

  //Определить оператор (для ODF)
  //INPUT
  //  const st: string                             - текст оператора
  //  var ConditionOperator: TZConditionalOperator - возвращаемый оператор
  //RETURN
  //      boolean - true - оператор успешно определён
  function ODFGetOperatorByStr(const st: string; var ConditionOperator: TZConditionalOperator): boolean;
  begin
    result := true;
    if (st = '<') then
       ConditionOperator := ZCFOpLT
    else
    if (st = '>') then
       ConditionOperator := ZCFOpGT
    else
    if (st = '<=') then
       ConditionOperator := ZCFOpLTE
    else
    if (st = '>=') then
       ConditionOperator := ZCFOpGTE
    else
    if (st = '=') then
       ConditionOperator := ZCFOpEqual
    else
    if (st = '!=') then
       ConditionOperator := ZCFOpNotEqual
    else
      result := false;
  end; //ODFGetOperatorByStr

  //Определение условия
  function _CheckCondition(): boolean;
  var
    v1, v2: double;
    t: integer;

    function _CheckBetween(val: TZCondition): boolean;
    begin
      result := false;
      if (kol = 3) then
        //can not use standart TryStrToFloat - no FS in C++ builder 6!
        if (ZEIsTryStrToFloat(_strArr[1], v1)) then
          if (ZEIsTryStrToFloat(_strArr[2], v2)) then
          begin
            result := true;
            Condition := val;
            Value1 := _strArr[1];
            Value2 := _strArr[2];
          end;
    end; //_CheckBetween

    function _CheckOperator(): boolean;
    begin
      result := false;
      if (kol >= 2) then
        if (ODFGetOperatorByStr(_strArr[0], ConditionOperator)) then
          if (ZEIsTryStrToFloat(_strArr[1], v1)) then
          begin
            result := true;
            Condition := ZCFCellContentOperator;
            Value1 := _strArr[1];
          end;
    end; //_CheckOperator

    function _CheckTextCondition(val: TZCondition): boolean;
    begin
      result := false;
      if (kol >= 2) then
        if (_strArr[1] <> '') then
        begin
          //TODO: потом проверить, нужно ли убирать кавычки и всё такое
          result := true;
          Condition := val;
          Value1 := _strArr[1];
        end;
    end; //_CheckTextCondition

    function _getSimpleCondition(val: TZCondition): boolean;
    begin
      result := true;
      Condition := val;
    end; //_getSimpleCondition

    function _getOneParamNumberCondition(val: TZCondition; isFloat: boolean = false): boolean;
    begin
      result := false;
      if (kol = 2) then
      begin
        if (isFloat) then
          result := ZEIsTryStrToFloat(_strArr[1], v1)
        else
          result := TryStrToInt(_strArr[1], t);
        if (result) then
        begin
          Condition := val;
          Value1 := _strArr[1];
        end;
      end;
    end; //_getOneParamNumberCondition

    function _CheckFormula(): boolean;
    var
      i: integer;
      _f, _l: integer;

    begin
      //Будем считать, что при чтении формула всегда валидная
      _f := 0;
      _l := 0;
      result := true;
      for i := 1 to len - 1 do
        if (ConditionalValue[i] = '(') then
        begin
          _f := i + 1;
          break;
        end;

      for i := len downto 2 do
        if (ConditionalValue[i] = ')') then
        begin
          _l := i;
          break;
        end;
     Condition := ZCFIsTrueFormula;
     Value1 := copy(ConditionalValue, _f, _l - _f);
    end; //_CheckFormula

  begin
    result := false;
    s := _strArr[0];

    //TODO: не забыть добавить все остальные условия
    //Для условных стилей из ODF (без расширения в LibreOffice):
    //  cell-content-is-between
    //  cell-content-is-not-between
    //  cell-content
    //  value
    //
    //Для LibreOffice (<calcext:conditional-formats>):
    //  between()
    //  not-between()
    //  begins-with()
    //  ends-with()
    //  contains-text()
    //  not-contains-text()
    //  operator value (>10 etc)
    //  begins-with
    //  ends-with
    //  contains-text
    //  not-contains-text
    //  duplicate
    //  unique
    //  above-average
    //  below-average
    //  above-equal-average
    //  below-equal-average
    //  top-elements
    //  bottom-elements
    //  top-percent
    //  bottom-percent
    //  is-error
    //  is-no-error
    //  formula-is

    if (_isFirstOperator or (s = 'cell-content')) then
      result := _CheckOperator()
    else
    if ((s = 'cell-content-is-between') or (s = 'between')) then
      result := _CheckBetween(ZCFCellContentIsBetween)
    else
    if ((s = 'cell-content-is-not-between') or (s = 'not-between')) then
      result := _CheckBetween(ZCFCellContentIsNotBetween)
    else
    if (s = 'value') then
    begin
    end else
    if (s = 'begins-with') then
      result := _CheckTextCondition(ZCFBeginsWithText)
    else
    if (s = 'ends-with') then
      result := _CheckTextCondition(ZCFEndsWithText)
    else
    if (s = 'contains-text') then
      result := _CheckTextCondition(ZCFContainsText)
    else
    if (s = 'not-contains-text') then
      result := _CheckTextCondition(ZCFNotContainsText)
    else
    if (s = 'duplicate') then
      result := _getSimpleCondition(ZCFDuplicate)
    else
    if (s = 'unique') then
      result := _getSimpleCondition(ZCFUnique)
    else
    if (s = 'above-average') then
      result := _getSimpleCondition(ZCFAboveAverage)
    else
    if (s = 'below-average') then
      result := _getSimpleCondition(ZCFBellowAverage)
    else
    if (s = 'above-equal-average') then
      result := _getSimpleCondition(ZCFAboveEqualAverage)
    else
    if (s = 'below-equal-average') then
      result := _getSimpleCondition(ZCFBelowEqualAverage)
    else
    if (s = 'top-elements') then
      result := _getOneParamNumberCondition(ZCFTopElements)
    else
    if (s = 'bottom-elements') then
      result := _getOneParamNumberCondition(ZCFBottomElements)
    else
    if (s = 'top-percent') then
      result := _getOneParamNumberCondition(ZCFTopPercent, true)
    else
    if (s = 'bottom-percent') then
      result := _getOneParamNumberCondition(ZCFBottomPercent, true)
    else
    if (s = 'is-error') then
      result := _getSimpleCondition(ZCFIsError)
    else
    if (s = 'is-no-error') then
      result := _getSimpleCondition(ZCFIsNoError)
    else
    if ((s = 'is-true-formula') or (s = 'formula-is')) then
      result := _CheckFormula();
  end; //_CheckCondition

  //Читает оператор
  //INPUT
  //  var retStr: string  - результирующая строка
  //  var num: integer    - позиция текущего символа в исходной строке
  procedure _ReadOperator(var retStr: string; var num: integer);
  var
    ch: char;

  begin
    if (num + 1 <= len) then
    begin
      ch := ConditionalValue[num + 1];
      // >, >=, <, <=, !=, =
      case (ch) of
        '>', '<', '=':
          begin
            retStr := retStr + ch;
            inc(num);
          end;
        //'!':;
      end;
    end;
    _ProcessBeforeDelimiter();
    if (kol = 1) then
      _isFirstOperator := true;
  end; //_ReadOperator

begin
  result := false;
  Value1 := '';
  Value2 := '';
  Condition := ZCFNumberValue;
  ConditionOperator := ZCFOpGT;
  len := length(ConditionalValue);
  _isFirstOperator := false;

  //TODO: нужно потом на досуге подумать более приличный способ разбора формул
  //      (перевести в обратную польскую запись и всё такое)
  {
  is-true-formula()                             (??)
  cell-content-is-between(value1, value2)
  cell-content-is-not-between(value1, value2)
  cell-content() operator value1
  string                                        (??)
  formula                                       (??)
  value() operator n                            (??)
  bool (true/false)                             (??)

  Example:
  cell-content-is-between(0,3)
  }

  if (len > 0) then
  try
    _maxKol := 4;
    SetLength(_strArr, _maxKol);
    result := true;
    s := '';
    kol := 0;
    _OCount := 0;
    i := 0;
    while (i < len) do
    begin
      inc(i);
      ch := ConditionalValue[i];
      case (ch) of
        ' ':
          begin
            _ProcessBeforeDelimiter();
          end;
        '''', '"':
          begin
            _ProcessBeforeDelimiter();
            _ReadWhileNotChar(s, i, ch);
          end;
        '(':
          begin
            _ProcessBeforeDelimiter();
            inc(_OCount);
          end;
        ')':
          begin
            _ProcessBeforeDelimiter();
            dec(_OCount);
            if (_OCount < 0) then
            begin
              result := false;
              break;
            end;
          end;
        ',':
          begin
            _ProcessBeforeDelimiter();
          end;
        '>', '<', '=', '!':
          begin
            _ProcessBeforeDelimiter();
            s := ch;
            _ReadOperator(s, i);
          end;
        else
          s := s + ch;
      end; //case
    end; //while

    _ProcessBeforeDelimiter();
    if (_OCount <> 0) then
      result := false;
    if (kol > 0) then
      result := _CheckCondition()
    else
      result := false;
  finally
    SetLength(_strArr, 0);
  end; //if
end; //ODFReadGetConditional

//Применить условные форматирования к листу
//INPUT
//      SheetNumber: integer            - номер листа
//  var DefStylesArray: TZODFStyleArray - массив со стилями (Styles.xml)
//      DefStylesCount: integer         - кол-во стилей в массиве
//  var StylesArray: TZODFStyleArray    - массив со стилями (основной из content)
//      StylesCount: integer            - кол-во стилей в массиве
procedure TZODFConditionalReadHelper.ApplyConditionStylesToSheet(SheetNumber: integer;
                                                     var DefStylesArray: TZODFStyleArray; DefStylesCount: integer;
                                                     var StylesArray: TZODFStyleArray; StylesCount: integer);
var
  i, j: integer;
  _CFCount: integer;            //сколько нужно добавить условных стилей
  _CFArray: array of integer;
  t: integer;
  _StartIDX: integer;
  _CF: TZConditionalFormatting;

  //заменить в условных стилях индексы на нужный
  //INPUT
  //  const StyleName: string - имя стиля
  //      StyleIndex: integer - номер стиля в хранилище
  procedure _ApplyCFDefIndexes(const StyleName: string; StyleIndex: integer);
  var
    i, j: integer;

  begin
    for i := 0 to DefStylesCount - 1 do
      for j := 0 to DefStylesArray[i].ConditionsCount - 1 do
        if (DefStylesArray[i].Conditions[j].ApplyStyleIDX < 0) then
          if (DefStylesArray[i].Conditions[j].ApplyStyleName = StyleName) then
            DefStylesArray[i].Conditions[j].ApplyStyleIDX := StyleIndex;
    for i := 0 to StylesCount - 1 do
      for j := 0 to StylesArray[i].ConditionsCount - 1 do
        if (StylesArray[i].Conditions[j].ApplyStyleIDX < 0) then
          if (StylesArray[i].Conditions[j].ApplyStyleName = StyleName) then
            StylesArray[i].Conditions[j].ApplyStyleIDX := StyleIndex
  end; //_ApplyCFDefIndexes

  //Добавить условное форматирование в документ
  //INPUT
  //      CFStyle: TZConditionalStyle       - добавляемый условный стиль
  //  var StyleItem: TZEODFStyleProperties  - свойства стилей
  procedure _AddCFItem(CFStyle: TZConditionalStyle; var StyleItem: TZEODFStyleProperties);
  var
    i, j: integer;
    _Condition: TZCondition;
    _ConditionOperator: TZConditionalOperator;
    _Value1: string;
    _Value2: string;
    t: integer;
    b: boolean;

  begin
    for i := 0 to StyleItem.ConditionsCount - 1 do
      if (ODFReadGetConditional(StyleItem.Conditions[i].ConditionValue,
                              _Condition,
                              _ConditionOperator,
                              _Value1,
                              _Value2)) then
      begin
        if (StyleItem.Conditions[i].ApplyStyleIDX < 0) then
        begin
          b := false;
          //Проверка по styles.xml
          for j := 0 to DefStylesCount - 1 do
            if (StyleItem.Conditions[i].ApplyStyleName = DefStylesArray[j].name) then
            begin
              b := true;
              if (DefStylesArray[j].index < 0) then
              begin
                DefStylesArray[j].index := FXMLSS.Styles.Add(FReadHelper.FStyles[j], true);
                _ApplyCFDefIndexes(DefStylesArray[j].name, DefStylesArray[j].index);
              end;

              StyleItem.Conditions[i].ApplyStyleIDX := DefStylesArray[j].index;
              break;
            end; //if
          //Проверка по стилям из content.xml
          if (not b) then
          for j := 0 to StylesCount - 1 do
            if (StyleItem.Conditions[i].ApplyStyleName = StylesArray[j].name) then
              if (StylesArray[j].index >= 0) then
              begin
                b := true;
                StyleItem.Conditions[i].ApplyStyleIDX := StylesArray[j].index;
                break;
              end;
        end else
          b := true;

        if (b) then
        begin
          t := CFStyle.Count;
          CFStyle.Add();
          CFStyle[t].Condition := _Condition;
          CFStyle[t].ConditionOperator := _ConditionOperator;
          CFStyle[t].Value1 := _Value1;
          CFStyle[t].Value2 := _Value2;
          CFStyle[t].ApplyStyleID := StyleItem.Conditions[i].ApplyStyleIDX;
        end; //if
      end; //if
  end; //_AddCFItem

  procedure _FillCF();
  var
    StyleIDX: integer;
    _numCF: integer;

  begin
    _numCF := _CFCount + _StartIDX;
    StyleIDX := _CFArray[_CFCount];
    if (StyleIDX >= StylesCount) then
    begin
      StyleIDX := StyleIDX - StylesCount;
      _AddCFItem(_CF[_numCF], DefStylesArray[StyleIDX]);
    end else
      _AddCFItem(_CF[_numCF], StylesArray[StyleIDX]);
  end; //_FillCF

  procedure _CheckAreas();
  var
    i, j: integer;
    b: boolean;

  begin
    _CFCount := 0;
    SetLength(_CFArray, FAreasCount);
    _CF := FXMLSS.Sheets[SheetNumber].ConditionalFormatting;
    _StartIDX := _CF.Count {- 1};
    for i := 0 to FAreasCount - 1 do
    begin
      b := true;
      t := FAreas[i].CFStyleNumber;
      for j := 0 to _CFCount - 1 do
        if (t = _CFArray[j]) then
        begin
          b := false;
          break;
        end;
      if (b) then
      begin
        _CF.Add();
        _CFArray[_CFCount] := t;
        _FillCF();
        inc(_CFCount);
      end;
    end; //for i
  end; //_CheckAreas

  //Добавить новую область к условному форматированию в хранилище
  //INPUT
  //      AreaNumber: integer   - номер области условного форматирования на текущем листе
  //      NewCFNumber: integer  - номер существующего условного форматирования в хранилище
  procedure _AddArea(AreaNumber: integer; NewCFNumber: integer);
  begin
    _cf.Items[NewCFNumber].Areas.Add(FAreas[AreaNumber].ColNum,
                                     FAreas[AreaNumber].RowNum,
                                     FAreas[AreaNumber].Width,
                                     FAreas[AreaNumber].Height);
  end; //_AddArea

begin
  if (FAreasCount > 0) then
    if (Assigned(FXMLSS)) then
      //если уже есть условное форматирование, то дополнительно не добавляем
      if (FXMLSS.Sheets[SheetNumber].ConditionalFormatting.Count = 0) then
      try
        _CheckAreas();
        for i := 0 to FAreasCount - 1 do
        for j := 0 to _CFCount - 1 do
          if (FAreas[i].CFStyleNumber = _CFArray[j]) then
          begin
            _AddArea(i, j + _StartIDX);
            break;
          end;
      finally
        SetLength(_CFArray, 0);
      end;
end; //ApplyConditionStylesToSheet

//Очистка всех условных форматирований (выполняется перед началом нового листа)
procedure TZODFConditionalReadHelper.Clear();
begin
  ClearLine();
  FColumnsCount := 0;
  FAreasCount := 0;
end; //Clear

procedure TZODFConditionalReadHelper.ClearLine();
begin
  FCountInLine := 0;
  FLineItemWidth := 0;
  FLineItemStartCell := -2; //Если < 0 - значит это первый раз в строке
  FLineItemStyleCFNumber := 0;
end; //ClearLine

//Добавить текущую строку с условными стилями в список условных стилей текущей страницы
//INPUT
//      RowNumber: integer    - текущая строка
//      RepeatCount: integer  - кол-во повторов строки
procedure TZODFConditionalReadHelper.ProgressLine(RowNumber: integer; RepeatCount: integer = 1);
var
  b: boolean;
  i, j: integer;
  LastAreaIndexBeforeProgress: integer;

begin
  if (FLineItemStartCell >= 0) then
  begin
    AddToLine(FLineItemStartCell, FLineItemStyleCFNumber, FLineItemWidth);
    FLineItemStartCell := -2;
  end;

  LastAreaIndexBeforeProgress := FAreasCount - 1;
  for i := 0 to FCountInLine - 1 do
  begin
    b := true;
    for j := 0 to LastAreaIndexBeforeProgress do
      if (FCurrentLine[i].StyleNumber = FAreas[j].CFStyleNumber) then
        if (FCurrentLine[i].CellNum = FAreas[j].ColNum) then
          if (FCurrentLine[i].Count = FAreas[j].Width) then
            if (FAreas[j].RowNum + FAreas[j].Height = RowNumber) then
            begin
              FAreas[j].Height := FAreas[j].Height + RepeatCount;
              b := false;
              break;
            end;

    if (b) then
    begin
      j := FAreasCount;
      inc(FAreasCount);
      if (FAreasCount >= FMaxAreasCount) then
      begin
        inc(FMaxAreasCount, 10);
        SetLength(FAreas, FMaxAreasCount);
      end;
      FAreas[j].RowNum := RowNumber;
      FAreas[j].ColNum := FCurrentLine[i].CellNum;
      FAreas[j].Width := FCurrentLine[i].Count;
      FAreas[j].Height := RepeatCount;
      FAreas[j].CFStyleNumber := FCurrentLine[i].StyleNumber;
    end;
  end;
end; //ProgressLine

//Применить адрес базовой ячейки
//INPUT
//  const BaseCellTxt: string         - текст базовой ячейки
//  const ACFStyle: TZConditionalStyle  - условный стиль
//      PageNum: integer              - номер текущей страницы
procedure TZODFConditionalReadHelper.ApplyBaseCellAddr(const BaseCellTxt: string; const ACFStyle: TZConditionalStyleItem; PageNum: integer);
var
  i: integer;
  _l: integer;
  _len: integer;
  s: string;
  _c, _r: integer;

begin
  if (Assigned(ACFStyle)) then
  begin
    _len := length(BaseCellTxt);
    _l := -1;
    for i := _len downto 1 do
      if (BaseCellTxt[i] = '.') then
      begin
        _l := i;
        break;
      end;
    if (_l > 0) then
    begin
      s := copy(BaseCellTxt, _l + 1, _len - _l);
      {$HINTS OFF} //It's ok
      ZEGetCellCoords(s, _c, _r);
      {$HINTS ON}
      ACFStyle.BaseCellColumnIndex := _c;
      ACFStyle.BaseCellRowIndex := _r;
      s := ZEReplaceEntity(copy(BaseCellTxt, 1, _l - 1));
      _l := Length(s);
      if (_l >= 2) then
        if ((s[1] = '''') and (s[1] = s[_l])) then
        begin
          delete(s, 1, 1);
          dec(_l);
          delete(s, _l, 1);
        end;

      _l := -1;
      for i := 0 to FXMLSS.Sheets.Count - 1 do
        if (FXMLSS.Sheets[i].Title = s) then
        begin
          _l := i;
          break;
        end;
      if (PageNum = _l) then
        _l := -1;
      ACFStyle.BaseCellPageIndex := _l;
    end;
  end;
end; //ApplyBaseCellAddr

//Читает <calcext:conditional-formats> .. </calcext:conditional-formats> - условное
//  форатирование для LibreOffice
//INPUT
//  var xml: TZsspXMLReaderH  - читатель (<> nil !!!)
//      SheetNum: integer     - номер страницы
procedure TZODFConditionalReadHelper.ReadCalcextTag(var xml: TZsspXMLReaderH; SheetNum: integer);
var
  _isCFItem: boolean;
  _CF: TZConditionalFormatting;
  _CFItem: TZConditionalStyle;
  s: string;
  tmpRec: array [0..1] of array [0..1] of integer;  //0 - c, 1 - r
  b: boolean;
  i: integer;
  _CFvalue: string;
  _stylename: string;
  _basecelladdr: string;

  procedure _AddAreas();
  var
    i: integer;
    ss: string;
    ch: char;
    _isQuote: boolean;
    _isApos: boolean;
    kol: integer;

    procedure _addSubArea(var str: string);
    begin
      if (str <> '') then
        if (kol < 2) then
        begin
          ZEGetCellCoords(str, tmpRec[kol][0], tmpRec[kol][1]);
          inc(kol);
        end;
      str := '';
    end; //_addSubArea

    procedure _PrepareAreaAndAdd(var RangeItem: string);
    var
      i: integer;
      s: string;
      _isQuote: boolean;
      _isApos: boolean;
      w, h: integer;

    begin
      if (RangeItem <> '') then
      begin
        kol := 0;
        RangeItem := RangeItem + ':';
        s := '';
        _isQuote := false;
        _isApos := false;
        for i := 1 to Length(RangeItem) do
        begin
          ch := RangeItem[i];
          case ch of
            '.': s := '';
            '"': if (not _isApos) then _isQuote := not _isQuote;
            '''': if (not _isQuote) then _isApos := not _isApos;
            ':':
              begin
                if (not (_isQuote or _isApos)) then
                  _addSubArea(s);
              end;
            else
              s := s + ch;
          end;
        end;

        if (kol > 0) then
        begin
          w := 1;
          h := 1;
          if (kol = 2) then
          begin
            w := tmpRec[1][0] - tmpRec[0][0] + 1;
            h := tmpRec[1][1] - tmpRec[0][1] + 1;
          end;
          _CFItem.Areas.Add(tmpRec[0][0], tmpRec[0][1], w, h);
        end;
      end; //if

      RangeItem := '';
    end; //_PrepareArea

  begin
    _CFItem.Areas.Count := 0;
    s := ZEReplaceEntity(xml.Attributes[const_calcext_target_range_address]);
    if (s <>  '') then
    begin
      s := s + ' ';
      ss := '';
      _isQuote := false;
      _isApos := false;
      for i := 1 to Length(s) do
        case s[i] of
          '"': if (not _isApos) then
               begin
                 _isQuote := not _isQuote;
                 ss := ss + s[i];
               end;
          '''': if (not _isQuote) then
                begin
                 _isApos := not _isApos;
                 ss := ss + s[i];
                end;
          ' ':
            if (_isQuote or _isApos) then
              ss := ss + ' '
            else
              _PrepareAreaAndAdd(ss)
          else
            ss := ss + s[i];
        end;
    end;
  end; //_AddAreas

  procedure _GetCondition();
  var
    _Condition: TZCondition;
    _ConditionOperator: TZConditionalOperator;
    _Value1: string;
    _Value2: string;
    num: integer;
    i: integer;
    _styleID: integer;

  begin
    _CFvalue := ZEReplaceEntity(xml.Attributes[const_calcext_value]);
    _stylename := xml.Attributes[const_calcext_apply_style_name];
    _basecelladdr := xml.Attributes[const_calcext_base_cell_address];
    if (_CFvalue <> '') then
      if (ODFReadGetConditional(_CFvalue,
                                _Condition,
                                _ConditionOperator,
                                _Value1,
                                _Value2)) then
      begin
        num := _CFItem.Count;
        _CFItem.Add();
        _CFItem[num].Condition := _Condition;
        _CFItem[num].ConditionOperator := _ConditionOperator;
        _CFItem[num].Value1 := _Value1;
        _CFItem[num].Value2 := _Value2;

        ApplyBaseCellAddr(_basecelladdr, _CFItem[num], SheetNum);

        _styleID := -1;
        for i := 0 to ReadHelper.StylesCount - 1 do
          if (ReadHelper.StylesProperties[i].name = _stylename) then
          begin
            if (ReadHelper.StylesProperties[i].index < 0) then
              ReadHelper.StylesProperties[i].index := FXMLSS.Styles.Add(ReadHelper.Style[i]);
            _styleID := ReadHelper.StylesProperties[i].index;
            break;
          end;

        _CFItem[num].ApplyStyleID := _styleID;
      end;
  end; //_GetCondition

begin
  _isCFItem := false;
  _CF := FXMLSS.Sheets[SheetNum].ConditionalFormatting;
  (*
   <calcext:conditional-format calcext:target-range-address="Лист1.A1:Лист1.D17 Лист1.E1:Лист1.F17">
      <calcext:condition calcext:apply-style-name="Безымянный1" calcext:value="begins-with(&quot;as&quot;)" calcext:base-cell-address="Лист1.A1"/>
      <calcext:condition calcext:apply-style-name="Безымянный2" calcext:value="ends-with(&quot;ey&quot;)" calcext:base-cell-address="Лист1.A1"/>
      <calcext:condition calcext:apply-style-name="Безымянный3" calcext:value="contains-text(&quot;et&quot;)" calcext:base-cell-address="Лист1.A1"/>
      <calcext:condition calcext:apply-style-name="Безымянный4" calcext:value="not-contains-text(&quot;rt&quot;)" calcext:base-cell-address="Лист1.A1"/>
      <calcext:condition calcext:apply-style-name="Безымянный5" calcext:value="between(1,20)" calcext:base-cell-address="Лист1.A1"/>
   </calcext:conditional-format>
  *)
  _CFItem := TZConditionalStyle.Create();
  try
    while ((xml.TagType <> 6) or (xml.TagName <> const_calcext_conditional_formats)) do
    begin
      xml.ReadTag();
      if (xml.Eof()) then
        break;

      if (xml.TagName = const_calcext_conditional_format) then
      begin
        if (xml.TagType = 4) then
        begin
          _isCFItem := true;
          _CFItem.Count := 0;
          _AddAreas();
        end else
        begin
          if (_isCFItem) then
            if (_CFItem.Count > 0) then
            begin
              b := true;
              //TODO: потом переделать сравнение условных форматирований
              //      (пересечение областей и др.)
              for i := 0 to _CF.Count - 1 do
                if (_CF[i].IsEqual(_CFItem)) then
                begin
                  b := false;
                  break;
                end;

              if (b) then
                _CF.Add(_CFItem);
            end;
          _isCFItem := false;
        end;
      end; //if

      if ((xml.TagType = 5) and (xml.TagName = const_calcext_condition)) then
        if (_isCFItem) then
          _GetCondition();
    end; //while
  finally
    FreeAndNil(_CFItem);
  end;
end; //ReadCalcextTag

//Добавить для указанного столбца индекс условного форматирования
//INPUT
//      ColumnNumber: integer   - номер столбца
//      StyleCFNumber: integer  - номер условного стиля (в массиве)
procedure TZODFConditionalReadHelper.AddColumnCF(ColumnNumber: integer; StyleCFNumber: integer);
var
  t: integer;

begin
  t := FColumnsCount;
  inc(FColumnsCount);
  if (FColumnsCount >= FMaxColumnsCount) then
  begin
    inc(FMaxColumnsCount, 10);
    SetLength(FColumnSCFNumbers, FMaxColumnsCount);
  end;
  FColumnSCFNumbers[t][0] := ColumnNumber;
  FColumnSCFNumbers[t][1] := StyleCFNumber;
end;  //AddColumnCF

//Получить номер стиля условного форматирования для колонки
//INPUT
//      ColumnNumber: integer - номер столбца
//RETURN
//      integer - >= 0 - номер стиля условного форматирования в массиве
//                < 0 - для данного столбца не применялся условный стиль по дефолту
function TZODFConditionalReadHelper.GetColumnCF(ColumnNumber: integer): integer;
var
  i: integer;

begin
  result := -2;
  for i := 0 to FColumnsCount - 1 do
    if (FColumnSCFNumbers[i][0] = ColumnNumber) then
    begin
      result := FColumnSCFNumbers[i][1];
      break;
    end;
end; //GetColumnCF

{$ENDIF} //ZUSE_CONDITIONAL_FORMATTING

//Очистка доп. свойств стиля
procedure ODFClearStyleProperties(var StyleProperties: TZEODFStyleProperties);
begin
  StyleProperties.name := '';
  StyleProperties.index := -2;
  StyleProperties.ParentName := '';
  StyleProperties.isHaveParent := false;
  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  StyleProperties.ConditionsCount := 0;
  SetLength(StyleProperties.Conditions, 0);
  {$ENDIF}
end; //ODFClearStyleProperties

///////////////////////////////////////////////////////////////////
////::::::::::::: TZEODFReadWriteHelperParent::::::::::::::::::////
///////////////////////////////////////////////////////////////////

constructor TZEODFReadWriteHelperParent.Create(AXMLSS: TZEXMLSS);
begin
  FXMLSS := AXMLSS;
end;

///////////////////////////////////////////////////////////////////
////::::::::::::::::::: TZEODFReadHelper ::::::::::::::::::::::////
///////////////////////////////////////////////////////////////////

constructor TZEODFReadHelper.Create(AXMLSS: TZEXMLSS);
begin
  inherited Create(AXMLSS);
  FStylesCount := 0;
  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  FConditionReader := TZODFConditionalReadHelper.Create(AXMLSS);
  FConditionReader.ReadHelper := self;
  {$ENDIF}
  FMasterPagesCount := 0;
  SetLength(FMasterPages, 0);
  SetLength(FMasterPagesNames, 0);
  FPageLayoutsCount := 0;
  SetLength(FPageLayouts, 0);
  SetLength(FPageLayoutsNames, 0);
  FNumberStylesHelper := TZEODSNumberFormatReader.Create();
end; //Create

destructor TZEODFReadHelper.Destroy();
var
  i: integer;

begin
  for i := 0 to FStylesCount - 1 do
    if (Assigned(FStyles[i])) then
      FreeAndNil(FStyles[i]);

  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  for i := 0 to FStylesCount - 1 do
    SetLength(StylesProperties[i].Conditions, 0);
  if (Assigned(FConditionReader)) then
    FreeAndNil(FConditionReader);
  {$ENDIF}

  for i := 0 to FMasterPagesCount - 1 do
    FreeAndNil(FMasterPages[i]);

  SetLength(FMasterPages, 0);
  SetLength(FMasterPagesNames, 0);

  for i := 0 to FPageLayoutsCount - 1 do
    FreeAndNil(FPageLayouts[i]);

  SetLength(FPageLayouts, 0);
  SetLength(FPageLayoutsNames, 0);

  SetLength(FStyles, 0);
  SetLength(StylesProperties, 0);

  FreeAndNil(FNumberStylesHelper);

  inherited;
end;

function TZEODFReadHelper.GetStyle(num: integer): TZStyle;
begin
  result := nil;
  if (num >= 0) and (num < FStylesCount) then
    result := FStyles[num];
end; //GetStyle

//Read <office:automatic-styles> </office:automatic-styles>
// (Page layouts: paper size, orientation etc)
procedure TZEODFReadHelper.ReadAutomaticStyles(xml: TZsspXMLReaderH);
var
  _PL: integer;
  _w, _h, t: integer; //page width and height
  s: string;
  r: real;
  _tmpColor: TColor;

  function _GetPaperSize(w, h: integer): byte;
  var
    a: array [0..3] of integer;
    r2: byte;

    function _Find(w, h: integer; out MinDeltaX, MinDeltaY: integer): byte;
    var
      i: byte;

    begin
      Result := 0;
      MinDeltaX := const_ODS_paper_sizes[1][0];
      MinDeltaY := const_ODS_paper_sizes[1][1];
      for i := 1 to const_ODS_paper_sizes_count - 1 do
      begin
        t := abs(w - const_ODS_paper_sizes[i][0]) + abs(h - const_ODS_paper_sizes[i][1]);

        if (t < MinDeltaX + MinDeltaY) then
        begin
          Result := i;
          MinDeltaX := abs(w - const_ODS_paper_sizes[i][0]);
          MinDeltaY := abs(h - const_ODS_paper_sizes[i][1]);
        end;
      end;
    end; //_Find

  begin
    Result := _Find(w, h, a[0], a[1]);
    if ((a[0] <> 0) or (a[1] <> 0)) then
    begin
      r2 := _Find(h, w, a[2], a[3]);
      if (a[0] + a[1] > a[2] + a[3]) then
        Result := r2;
    end;
  end; //_GetPaperSize

  procedure _GetAttrValue(const attrName: string; var retValue: integer; koef: integer = 10);
  begin
    s := xml.Attributes.ItemsByName[attrName];
    if (s > '') then
    begin
      if (ODFGetValueSizeMM(s, r)) then
        retValue := round(r * koef)
    end;
  end; //_GetAttrValue

  function _GetMarginValue(const attrName: string; DefValue: integer): integer;
  begin
    Result := DefValue;
    s := xml.Attributes.ItemsByName[attrName];
    if (s > '') then
    begin
      if (ODFGetValueSizeMM(s, r)) then
        result := round(r);
    end;
  end; //_GetMarginValue

  procedure _ReadFooterHeader(const tagName: string; HeaderItem: TZSheetFooterHeader; isHeader: boolean);
  begin
    while (not xml.Eof()) do
    begin
      if (not xml.ReadTag()) then
        break;

      if (xml.TagName = ZETag_style_header_footer_properties) then
      begin
        s := xml.Attributes[ZETag_fo_background_color];
        if (s > '') then
          _tmpColor := GetBGColorForODS(s);

        if (isHeader) then
        begin
          t := FPageLayouts[_PL].HeaderMargins.MarginTopBottom;
          _GetAttrValue(ZETag_fo_margin_top, t);
          FPageLayouts[_PL].HeaderMargins.MarginTopBottom := t;
        end
        else
        begin
          t := FPageLayouts[_PL].FooterMargins.MarginTopBottom;
          _GetAttrValue(ZETag_fo_margin_bottom, t);
          FPageLayouts[_PL].FooterMargins.MarginTopBottom := t;
        end;

        FPageLayouts[_PL].FooterMargins.UseAutoFitHeight := true;
        t := FPageLayouts[_PL].FooterMargins.Height;
        _GetAttrValue(ZETag_fo_min_height, t, 1);

        s := xml.Attributes[ZETag_svg_height];
        if (s <> '') then
        begin
          FPageLayouts[_PL].FooterMargins.UseAutoFitHeight := false;
          _GetAttrValue(ZETag_svg_height, t, 1);
        end;
        FPageLayouts[_PL].FooterMargins.Height := t;

        t := FPageLayouts[_PL].FooterMargins.MarginLeft;
        _GetAttrValue(ZETag_fo_margin_left, t, 1);
        FPageLayouts[_PL].FooterMargins.MarginLeft := t;

        t := FPageLayouts[_PL].FooterMargins.MarginRight;
        _GetAttrValue(ZETag_fo_margin_right, t, 1);
        FPageLayouts[_PL].FooterMargins.MarginRight := t;
      end;

      if (xml.TagType = 6) and (xml.TagName = tagName) then
        break;
    end;
  end; //_ReadFooterHeader

  //<style:page-layout> .. </style:page-layout>
  //  style:page-layout possible elements:
  //    style:footer-style
  //    style:header-style
  //    style:page-layout-properties
  procedure _ReadPageLayout();
  begin
    _PL := FPageLayoutsCount;
    inc(FPageLayoutsCount);
    SetLength(FPageLayouts, FPageLayoutsCount);
    SetLength(FPageLayoutsNames, FPageLayoutsCount);
    FPageLayoutsNames[_PL] := xml.Attributes.ItemsByName[ZETag_Attr_StyleName];
    FPageLayouts[_PL] := TZSheetOptions.Create();
    _w := -1;
    _h := -1;
    while (not xml.Eof()) do
    begin
      if (not xml.ReadTag()) then
        break;

      if ((xml.TagType in [4, 5]) and (xml.TagName = ZETag_style_page_layout_properties)) then
      begin
        //  fo:page-width
        //  fo:page-height
        //  style:num-format
        //  style:print-orientation
        //  fo:margin-top
        //  fo:margin-bottom
        //  fo:margin-left
        //  fo:margin-right
        //  fo:border
        //  fo:padding
        //  style:shadow
        //  fo:background-color
        //  style:writing-mode

        _GetAttrValue(ZETag_fo_page_width, _w);
        _GetAttrValue(ZETag_fo_page_height, _h);

        //portrait by default
        s := xml.Attributes[ZETag_style_print_orientation];
        FPageLayouts[_PL].PortraitOrientation := (s = '') or (s = ZETag_portrait);

        FPageLayouts[_PL].MarginTop := _GetMarginValue(ZETag_fo_margin_top, FPageLayouts[_PL].MarginTop);
        FPageLayouts[_PL].MarginBottom := _GetMarginValue(ZETag_fo_margin_bottom, FPageLayouts[_PL].MarginBottom);
        FPageLayouts[_PL].MarginLeft := _GetMarginValue(ZETag_fo_margin_left, FPageLayouts[_PL].MarginLeft);
        FPageLayouts[_PL].MarginRight := _GetMarginValue(ZETag_fo_margin_right, FPageLayouts[_PL].MarginRight);

        s := xml.Attributes[ZETag_style_scale_to];
        if (s <> '') then
          if (TryStrToIntPercent(s, t)) then
            FPageLayouts[_PL].ScaleToPercent := t;

        s := xml.Attributes[ZETag_style_scale_to_pages];
        if (s <> '') then
          if (TryStrToInt(s, t)) then
            FPageLayouts[_PL].ScaleToPages := t;
      end; //if

      if ((xml.TagType = 4) and (xml.TagName = ZETag_style_header_style)) then
      begin
        _tmpColor := FPageLayouts[_PL].HeaderBGColor;
        _ReadFooterHeader(ZETag_style_header_style, FPageLayouts[_PL].Header, true);
        FPageLayouts[_PL].HeaderBGColor := _tmpColor;
      end;

      if ((xml.TagType = 4) and (xml.TagName = ZETag_style_footer_style)) then
      begin
        _tmpColor := FPageLayouts[_PL].FooterBGColor;
        _ReadFooterHeader(ZETag_style_footer_style, FPageLayouts[_PL].Footer, true);
        FPageLayouts[_PL].FooterBGColor := _tmpColor;
      end;

      if ((xml.TagType = 6) and (xml.TagName = ZETag_style_page_layout)) then
        break;
    end; //while

    if ((_h > 0) and (_w > 0)) then
    begin
      FPageLayouts[_PL].PaperSize := _GetPaperSize(_w, _h);
      FPageLayouts[_PL].PaperWidth := _w div 10;
      FPageLayouts[_PL].PaperHeight := _h div 10;
    end;
  end; //_ReadPageLayout

begin
  while (not xml.Eof()) do
  begin
    if (not xml.ReadTag()) then
      break;

    if ((xml.TagType = 6) and (xml.TagName = ZETag_office_automatic_styles)) then
      break;

    if ((xml.TagType = 4) and (xml.TagName = ZETag_style_page_layout)) then
      _ReadPageLayout();
  end; //while
end; //ReadAutomaticStyles

//Read <office:master-styles> .. </office:master-styles> -
//  List of master pages
//INPUT
//    xml: TZsspXMLReaderH - reader
procedure TZEODFReadHelper.ReadMasterStyles(xml: TZsspXMLReaderH);
var
  _pagelayoutname: string;
  _MP: TZSheetOptions;
  s: string;

  //Read <text:p> .. </text:p> and return text
  function _ReadTextP(): string;
  begin
    //Possible child elements:
    //    text:date                       ??
    //    text:time                       ??
    //    text:page-number
    //    text:page-continuation          ??
    //    text:sender-firstname           ??
    //    text:sender-lastname            ??
    //    text:sender-initials            ??
    //    text:sender-title               ??
    //    text:sender-position            ??
    //    text:sender-email               ??
    //    text:sender-phone-private       ??
    //    text:sender-fax                 ??
    //    text:sender-company             ??
    //    text:sender-phone-work          ??
    //    text:sender-street              ??
    //    text:sender-city                ??
    //    text:sender-postal-code         ??
    //    text:sender-country             ??
    //    text:sender-state-or-province   ??
    //    text:author-name                ??
    //    text:author-initials            ??
    //    text:chapter                    ??
    //    text:file-name
    //    text:template-name              ??
    //    text:sheet-name
    Result := '';
    while (not xml.Eof()) do
    begin
      if (not xml.ReadTag()) then
        break;

      Result := Result + xml.TextBeforeTag;

      if ((xml.TagType = 6) and (xml.TagName = ZETag_text_p)) then
        break;
    end; //while
  end; //_ReadTextP()

  function _ReadLCR(const TagName: string): string;
  begin
    Result := '';
    while (not xml.Eof()) do
    begin
      if (not xml.ReadTag()) then
        break;

      if ((xml.TagType = 4) and (xml.TagName = ZETag_text_p)) then
        Result := _ReadTextP();

      if ((xml.TagType = 6) and (xml.TagName = TagName)) then
        break;
    end; //while
  end; //_ReadLCR

  procedure _ReadFooterHeader(const TagName: string; FooterHeader: TZSheetFooterHeader);
  begin
    FooterHeader.IsDisplay := true; //By default
    s := xml.Attributes[ZETag_style_display];
    if (s <> '') then
      FooterHeader.IsDisplay := ZEStrToBoolean(s);
    if (xml.TagType = 4) then
    begin
      while (not xml.Eof()) do
      begin
        if (not xml.ReadTag()) then
          break;

        if ((xml.TagType = 4) and (xml.TagName = ZETag_text_p)) then
          FooterHeader.Data := _ReadTextP();

        if (xml.TagType = 4) then
        begin
          if (xml.TagName = ZETag_style_region_left) then
            FooterHeader.DataLeft := _ReadLCR(ZETag_style_region_left);
          if (xml.TagName = ZETag_style_region_center) then
            FooterHeader.Data := _ReadLCR(ZETag_style_region_center);
          if (xml.TagName = ZETag_style_region_right) then
            FooterHeader.DataRight := _ReadLCR(ZETag_style_region_right);
        end;

        if ((xml.TagType = 6) and (xml.TagName = TagName)) then
          break;
      end;
    end;
  end; //_ReadFooterHeader

  //<style:master-page> .. </style:master-page>
  procedure _ReadMasterPageStyle();
  var
    i: integer;

  begin
    //Possible attributes:
    //    draw:style-name           ??
    //    style:display-name        ??
    //    style:name                *
    //    style:next-style-name     ??
    //    style:page-layout-name    *
    _pagelayoutname := xml.Attributes[ZETag_style_page_layout_name];

    inc(FMasterPagesCount);
    SetLength(FMasterPages, FMasterPagesCount);
    SetLength(FMasterPagesNames, FMasterPagesCount);
    FMasterPages[FMasterPagesCount - 1] := TZSheetOptions.Create();
    _MP := FMasterPages[FMasterPagesCount - 1];

    FMasterPagesNames[FMasterPagesCount - 1] := xml.Attributes[ZETag_Attr_StyleName];

    for i := 0 to FPageLayoutsCount - 1 do
      if (FPageLayoutsNames[i] = _pagelayoutname) then
      begin
        _MP.Assign(FPageLayouts[i]);
        break;
      end;

    //Possible elements:
    //      anim:animate              ??
    //      anim:animateColor         ??
    //      anim:animateMotion        ??
    //      anim:animateTransform     ??
    //      anim:audio                ??
    //      anim:command              ??
    //      anim:iterate              ??
    //      anim:par                  ??
    //      anim:seq                  ??
    //      anim:set                  ??
    //      anim:transitionFilter     ??
    //      dr3d:scene                ??
    //      draw:a                    ??
    //      draw:caption              ??
    //      draw:circle               ??
    //      draw:connector            ??
    //      draw:control              ??
    //      draw:custom-shape         ??
    //      draw:ellipse              ??
    //      draw:frame                ??
    //      draw:layer-set            ??
    //      draw:line                 ??
    //      draw:measure              ??
    //      draw:page-thumbnail       ??
    //      draw:path                 ??
    //      draw:polygon              ??
    //      draw:polyline             ??
    //      draw:rect                 ??
    //      draw:regular-polygon      ??
    //      office:forms              ??
    //      presentation:notes        ??
    //      style:footer              *
    //      style:footer-left         *
    //      style:header              *
    //      style:header-left         *
    while (not xml.Eof()) do
    begin
      if (not xml.ReadTag()) then
        break;

      if (xml.TagType in [4, 5]) then
      begin
        if (xml.TagName = ZETag_style_header) then
          _ReadFooterHeader(xml.TagName, _MP.Header);
        if (xml.TagName = ZETag_style_header_left) then
        begin
          _ReadFooterHeader(xml.TagName, _MP.EvenHeader);
          _MP.IsEvenHeaderEqual := false;
        end;
        if (xml.TagName = ZETag_style_footer) then
          _ReadFooterHeader(xml.TagName, _MP.Footer);
        if (xml.TagName = ZETag_style_footer_left) then
        begin
          _ReadFooterHeader(xml.TagName, _MP.EvenFooter);
          _MP.IsEvenFooterEqual := false;
        end;
      end;

      if ((xml.TagType = 6) and (xml.TagName = ZETag_style_master_page)) then
        break;
    end; //while
  end;  //_ReadMasterPageStyle

begin
  //TODO: there are potential trouble:
  //    if element "office:master-styles" will be before "office:automatic-styles"
  //    in xml, then master styles will be wrong page layouts! Solution: read master pages data
  //    to tmp variables and apply at the end of xml.
  while (not xml.Eof()) do
  begin
    if (not xml.ReadTag()) then
      break;

    if ((xml.TagType = 4) and (xml.TagName = ZETag_style_master_page)) then
      _ReadMasterPageStyle();

    if ((xml.TagType = 6) and (xml.TagName = ZETag_office_master_styles)) then
      break;
  end; //while
end; //ReadMasterStyles

//Read manifest.xml
//INPUT
//  const stream: TStream - stream with manifest.xml
//RETURN
//      boolean - true - manifest read ok
function TZEODFReadHelper.ODSReadManifest(const stream: TStream): boolean;
var
  _xml: TZsspXMLReaderH;
  kol, maxkol: integer;
  a: array of array [0..2] of string;

  procedure _AddFileEntry();
  begin
    if (kol + 1 >= maxkol) then
    begin
      inc(maxkol, 20);
      SetLength(a, maxkol);
    end;
    a[kol][0] := _xml.Attributes[ZETag_manifest_full_path];
    a[kol][1] := _xml.Attributes[ZETag_manifest_media_type];
    a[kol][2] := _xml.Attributes[ZETag_manifest_version];
    inc(kol);
  end;

begin
  Result := false;
  kol := 0;

  _xml := nil;
  try
    _xml := TZsspXMLReaderH.Create();
    if (_xml.BeginReadStream(stream) = 0) then
    begin
      maxkol := 30;
      SetLength(a, maxkol);
      while (not _xml.Eof()) do
        if (_xml.ReadTag()) then
          if ((_xml.TagName = ZETag_manifest_file_entry) and (_xml.TagType in [4, 5])) then
            _AddFileEntry();
      Result := true;
    end;
  finally
    SetLength(a, 0);
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
  end;
end; //ODSReadManifest

procedure TZEODFReadHelper.AddStyle();
var
  num: integer;

begin
  num := FStylesCount;
  inc(FStylesCount);
  SetLength(FStyles, FStylesCount);
  SetLength(StylesProperties, FStylesCount);
  FStyles[num] := TZStyle.Create();
  if (Assigned(XMLSS)) then
    FStyles[num].Assign(XMLSS.Styles.DefaultStyle);
  ODFClearStyleProperties(StylesProperties[num]);
end; //AddStyle

//Apply for SheetOptions master page settings by name
//INPUT
//    SheetOptions: TZSheetOptions  - sheet options
//  const MasterPageName: string    - name of Master Page
procedure TZEODFReadHelper.ApplyMasterPageStyle(SheetOptions: TZSheetOptions; const MasterPageName: string);
var
  i: integer;

begin
  for i := 0 to FMasterPagesCount - 1 do
    if (MasterPageName = FMasterPagesNames[i]) then
    begin
      SheetOptions.Assign(FMasterPages[i]);
      break;
    end;
end; //ApplyMasterPageStyle

///////////////////////////////////////////////////////////////////
////::::::::::::::::: END TZEODFReadHelper ::::::::::::::::::::////
///////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////
////:::::::::::::::::: TZEODFWriteHelper ::::::::::::::::::::::////
///////////////////////////////////////////////////////////////////

constructor TZEODFWriteHelper.Create(AXMLSS: TZEXMLSS;
                       const _pages: TIntegerDynArray;
                       const _names: TStringDynArray;
                       PagesCount: integer);
var
  i: integer;
  _a: array of integer;

  function _GetSheetOptions(PageNum: integer): TZSheetOptions;
  begin
    if (PageNum <= -1) then
      Result := AXMLSS.DefaultSheetOptions
    else
      Result := AXMLSS.Sheets[_pages[PageNum]].SheetOptions;
  end; //_GetSheetOptions

  //For Styles.xml, get number of <style:page-layout> .. </style:page-layout>
  //INPUT
  //      SheetOptions: TZSheetOptions  - sheet options
  //      PageNum: integer              - number of page in _pages array
  //RETURN
  //      integer - index of page layout
  function _GetPageLayoutIndex(SheetOptions: TZSheetOptions; PageNum: integer): integer;
  var
    i: integer;
    _SO: TZSheetOptions;
    b: boolean;

  begin
    Result := -1;
    for i := 0 to FUniquePageLayoutsCount - 1 do
    begin
      _SO := _GetSheetOptions(FUniquePageLayouts[i]);

      b := (_SO.MarginLeft = SheetOptions.MarginLeft) and
           (_SO.MarginBottom = SheetOptions.MarginBottom) and
           (_SO.MarginTop = SheetOptions.MarginTop) and
           (_SO.MarginRight = SheetOptions.MarginRight) and
           (_SO.PaperSize = SheetOptions.PaperSize) and
           (_SO.PortraitOrientation = SheetOptions.PortraitOrientation) and
           (_SO.HeaderMargins.IsEqual(SheetOptions.HeaderMargins)) and
           (_SO.FooterMargins.IsEqual(SheetOptions.FooterMargins)) and
           (_SO.HeaderBGColor = SheetOptions.HeaderBGColor) and
           (_SO.FooterBGColor = SheetOptions.FooterBGColor)
           ;
      if (b) then
      begin
        Result := i;
        exit;
      end;
    end;
    if (Result < 0) then
    begin
      Result := FUniquePageLayoutsCount;
      FUniquePageLayouts[FUniquePageLayoutsCount] := PageNum;
      inc(FUniquePageLayoutsCount);
    end;
  end; //_GetPageLayoutIndex

  //For Styles.xml, get number of <style:master-page> .. </style:master-page>
  //INPUT
  //      SheetOptions: TZSheetOptions  - sheet options
  //      PageNum: integer              - number of page in _pages array
  //RETURN
  //      integer - index of master page
  function _GetMasterPageIndex(SheetOptions: TZSheetOptions; PageNum: integer): integer;
  var
    i: integer;
    _SO: TZSheetOptions;
    b: boolean;

  begin
    Result := -1;
    for i := 0 to FMasterPagesCount - 1 do
    begin
      if (FMasterPages[i] <= -1) then
        _SO := AXMLSS.DefaultSheetOptions
      else
        _SO := AXMLSS.Sheets[_pages[FMasterPages[i]]].SheetOptions;

      b := (_SO.Header.IsEqual(SheetOptions.Header)) and
           (_SO.Footer.IsEqual(SheetOptions.Footer)) and
           (_SO.IsEvenFooterEqual = SheetOptions.IsEvenFooterEqual) and
           (_SO.IsEvenHeaderEqual = SheetOptions.IsEvenHeaderEqual) and
           (FPageLayoutsIndexes[FMasterPages[i + 1]] = _a[PageNum]);
      if (b) then
      begin
        Result := i;
        exit;
      end;
    end;
    if (Result < 0) then
    begin
      Result := FMasterPagesCount;
      FMasterPages[FMasterPagesCount] := PageNum;
      FPageLayoutsIndexes[FMasterPagesCount] := _a[PageNum];
      FMasterPagesNames[FMasterPagesCount] := 'MasterPage' + IntToStr(FMasterPagesCount);
      inc(FMasterPagesCount);
    end;
  end; //_GetMasterPageIndex

begin
  inherited Create(AXMLSS);
  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  FConditionWriter := TZODFConditionalWriteHelper.Create(AXMLSS, _pages, _names, PagesCount);
  {$ENDIF}

  FNumberFormatWriter := TZEODSNumberFormatWriter.Create(AXMLSS.Styles.Count + 1);

  SetLength(FUniquePageLayouts, PagesCount + 1);
  FUniquePageLayouts[0] := -1;  //Default page layout
  FUniquePageLayoutsCount := 1;
  SetLength(FPageLayoutsIndexes, PagesCount + 1);

  SetLength(_a, PagesCount);

  SetLength(FMasterPages, PagesCount + 1);
  SetLength(FMasterPagesNames, PagesCount + 1);
  SetLength(FMasterPagesIndexes, PagesCount + 1);
  FMasterPagesCount := 1;
  FMasterPages[0] := -1;        //Default masterpage
  FMasterPagesNames[0] := 'Default';

  try
    for i := 0 to PagesCount - 1 do
      _a[i] := _GetPageLayoutIndex(AXMLSS.Sheets[_pages[i]].SheetOptions, i);

    for i := 0 to PagesCount - 1 do
      FMasterPagesIndexes[i + 1] := _GetMasterPageIndex(AXMLSS.Sheets[_pages[i]].SheetOptions, i);
  finally
    SetLength(_a, 0);
  end;
end; //Create

destructor TZEODFWriteHelper.Destroy();
begin
  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  if (Assigned(FConditionWriter)) then
    FreeAndNil(FConditionWriter);
  {$ENDIF}

  FreeAndNil(FNumberFormatWriter);

  SetLength(FPageLayoutsIndexes, 0);
  SetLength(FUniquePageLayouts, 0);
  SetLength(FMasterPages, 0);
  SetLength(FMasterPagesNames, 0);
  SetLength(FMasterPagesIndexes, 0);
  inherited;
end;

function TZEODFWriteHelper.GetMasterPageName(PageNum: integer): string;
begin
  Result := FMasterPagesNames[FMasterPagesIndexes[PageNum + 1]];
end;

//Write to XML all page layouts styles (<style:page-layout> .. </style:page-layout>)
//INPUT
//      xml: TZsspXMLWriterH
//  const _pages: TIntegerDynArray
procedure TZEODFWriteHelper.WriteStylesPageLayouts(xml: TZsspXMLWriterH; const _pages: TIntegerDynArray);
var
  i: integer;
  _SO: TZSheetOptions;
  s: string;
  _w, _h: integer;
  tmp: integer;

  procedure _AddAttrSize(const AttrName: string; Value: integer; DefaultValue: integer = 20);
  begin
    if (Value <> DefaultValue) then
      xml.Attributes.Add(AttrName, ODFGetSizeToStr(Value * 0.1), false);
  end;

  procedure _AddHeaderFooter(const ATagName, AAttrName: string; const HF: TZHeaderFooterMargins; Color: TColor);
  begin
    xml.Attributes.Clear();
    xml.WriteTagNode(ATagName, true, true, false);
    tmp := HF.Height;
    if (tmp = 0) then
      tmp := 3;

    if (HF.UseAutoFitHeight) then
      _AddAttrSize(ZETag_fo_min_height, tmp, -1)
    else
      _AddAttrSize(ZETag_svg_height, tmp, -1);

    _AddAttrSize(AAttrName, HF.MarginTopBottom, -1);
    _AddAttrSize(ZETag_fo_margin_left, HF.MarginLeft, -1);
    _AddAttrSize(ZETag_fo_margin_right, HF.MarginRight, -1);

    if (Color <> clWindow) then
      xml.Attributes.Add(ZETag_fo_background_color, '#' + ColorToHTMLHex(Color), false);

    xml.WriteTagNode(ZETag_style_header_footer_properties, true, true, false);
    xml.WriteEndTagNode(); //style:header-footer-properties

    xml.WriteEndTagNode(); // ATagName
  end;

  //<style:page-layout-properties>..</style:page-layout-properties>
  procedure _AddPageLayoutProperties(num: integer);
  begin
    if (num < 0) then
      _SO := FXMLSS.DefaultSheetOptions
    else
      _SO := FXMLSS.Sheets[_pages[num]].SheetOptions;

    xml.Attributes.Clear();

    //Possible attributes (+ = implemented, ??? = not implemented):
    //  fo:page-width="21.001cm"            +
    //  fo:page-height="29.7cm"             +
    //  style:num-format="1"                ???
    //  style:print-orientation="portrait"  +
    //  fo:margin-top="2.2cm"               +
    //  fo:margin-bottom="1.799cm"          +
    //  fo:margin-left="1.9cm"              +
    //  fo:margin-right="2.101cm"           +
    //  fo:border="0.06pt solid #000000"    ???
    //  fo:padding="0cm"                    ???
    //  style:shadow="none"                 ???
    //  fo:background-color="#00ff00"       ???
    //  style:writing-mode="lr-tb"          ???

    //  fo:border-bottom                    ???
    //  fo:border-left                      ???
    //  fo:border-right                     ???
    //  fo:border-top                       ???
    //  fo:margin                           ???
    //  fo:padding-bottom                   ???
    //  fo:padding-left                     ???
    //  fo:padding-right                    ???
    //  fo:padding-top                      ???
    //  style:border-line-width             ???
    //  style:border-line-width-bottom      ???
    //  style:border-line-width-left        ???
    //  style:border-line-width-right       ???
    //  style:border-line-width-top         ???
    //  style:first-page-number             ???
    //  style:footnote-max-height           ???
    //  style:layout-grid-base-height       ???
    //  style:layout-grid-base-width        ???
    //  style:layout-grid-color             ???
    //  style:layout-grid-display           ???
    //  style:layout-grid-lines             ???
    //  style:layout-grid-mode              ???
    //  style:layout-grid-print             ???
    //  style:layout-grid-ruby-below        ???
    //  style:layout-grid-ruby-height       ???
    //  style:layout-grid-snap-to           ???
    //  style:layout-grid-standard-mode     ???
    //  style:num-letter-sync               ???
    //  style:num-prefix                    ???
    //  style:num-suffix                    ???
    //  style:paper-tray-name               ???
    //  style:print                         ???
    //  style:print-page-order              ???
    //  style:register-truth-ref-style-name ???

    //  style:scale-to                      +  for printing, default 100%
    //  style:scale-to-pages                +  for printing, number of pages on witch a document shoud be printed, default 1
    //  style:table-centering               ???

    //Default margins (t/r/b/l) in mm: 20/20/20/20
    _AddAttrSize(ZETag_fo_margin_top, _SO.MarginTop);
    _AddAttrSize(ZETag_fo_margin_bottom, _SO.MarginBottom);
    _AddAttrSize(ZETag_fo_margin_left, _SO.MarginLeft);
    _AddAttrSize(ZETag_fo_margin_right, _SO.MarginRight);

    s := IfThen(_SO.PortraitOrientation, ZETag_portrait, ZETag_landscape);
    xml.Attributes.Add(ZETag_style_print_orientation, s, false);

    if (_SO.ScaleToPercent <> 100) then
      xml.Attributes.Add(ZETag_style_scale_to, IntToStr(_SO.ScaleToPercent) + '%');
    if (_SO.ScaleToPages <> 1) then
      xml.Attributes.Add(ZETag_style_scale_to_pages, IntToStr(_SO.ScaleToPages));

    if (_SO.PaperSize > 0) and (_so.PaperSize < const_ODS_paper_sizes_count) then
    begin
      _w := const_ODS_paper_sizes[_SO.PaperSize][0];
      _h := const_ODS_paper_sizes[_SO.PaperSize][1];
    end
    else
    begin
      _w := _SO.PaperWidth * 10;
      _h := _So.PaperHeight * 10;
    end;

    if (not _SO.PortraitOrientation) then
    begin
      tmp := _w;
      _w := _h;
      _h := tmp;
    end;

    xml.Attributes.Add(ZETag_fo_page_width, ODFGetSizeToStr(_w * 0.01), false);
    xml.Attributes.Add(ZETag_fo_page_height, ODFGetSizeToStr(_h * 0.01), false);

    xml.Attributes.Add('style:writing-mode', 'lr-tb', false); //???

    xml.WriteTagNode(ZETag_style_page_layout_properties, true, true, false);
    //Possible child tags:
    //  style:background-image
    //  style:columns
    //  style:footnote-sep
    xml.WriteEndTagNode(); //style:page-layout-properties

    _AddHeaderFooter(ZETag_style_header_style, ZETag_fo_margin_bottom, _SO.HeaderMargins, _SO.HeaderBGColor);
    _AddHeaderFooter(ZETag_style_footer_style, ZETag_fo_margin_top, _SO.FooterMargins, _SO.FooterBGColor);
  end; //_AddPageLayoutProperties

begin
  for i := 0 to FUniquePageLayoutsCount - 1 do
  begin
    xml.Attributes.Clear();
    xml.Attributes.Add(ZETag_Attr_StyleName, 'Mpm' + IntToStr(i + 1), false);
    xml.WriteTagNode(ZETag_style_page_layout, true, true, false);
    _AddPageLayoutProperties(FUniquePageLayouts[i]);
    xml.WriteEndTagNode(); // style:page-layout
  end;
end; //WriteStylesPageLayouts

//Write to XML all master pages (<style:master-page> .. </style:master-page>)
//INPUT
//      xml: TZsspXMLWriterH
//  const _pages: TIntegerDynArray
procedure TZEODFWriteHelper.WriteStylesMasterPages(xml: TZsspXMLWriterH; const _pages: TIntegerDynArray);
var
  i: integer;
  _SO: TZSheetOptions;

  procedure _WriteRegion(const RegionName, Txt: string);
  begin
    xml.WriteTagNode(RegionName, true, true, false);
    xml.WriteTag(ZETag_text_p, Txt, true, false, true);
    xml.WriteEndTagNode();
  end; //_WriteRegion

  //Write header/footer item
  procedure _WriteHFItem(const TagName: string; const HFItem: TZSheetFooterHeader);
  begin
    xml.Attributes.Clear();
    if (not HFItem.IsDisplay) then
      xml.Attributes.Add(ZETag_style_display, 'false');
    xml.WriteTagNode(TagName, true, true, false);
    xml.Attributes.Clear();
    if ((HFItem.DataLeft = HFItem.DataRight) and (HFItem.DataLeft = '')) then
    begin
      xml.WriteTag(ZETag_text_p, HFItem.Data, true, false, true);
    end
    else
    begin
      _WriteRegion(ZETag_style_region_left, HFItem.DataLeft);
      _WriteRegion(ZETag_style_region_center, HFItem.Data);
      _WriteRegion(ZETag_style_region_right, HFItem.DataRight);
    end;

    xml.WriteEndTagNode(); //TagName
  end; //_WriteHFItem

  procedure _WriteMasterPage(Num: integer);
  begin
    if (num < 0) then
      _SO := FXMLSS.DefaultSheetOptions
    else
      _SO := FXMLSS.Sheets[_Pages[num]].SheetOptions;

    _WriteHFItem(ZETag_style_header, _SO.Header);
    if (not _SO.IsEvenHeaderEqual) then
      _WriteHFItem(ZETag_style_header_left, _SO.EvenHeader);
    _WriteHFItem(ZETag_style_footer, _SO.Footer);
    if (not _SO.IsEvenFooterEqual) then
      _WriteHFItem(ZETag_style_footer_left, _SO.EvenFooter);
  end; //_WriteMasterPage

  function _ifUniqueStyle(id: integer): boolean;
  var
    i: integer;

  begin
    result := true;
    for i := 0 to id - 1 do
      if (FMasterPages[i] = FMasterPages[id]) then
      begin
        result := false;
        break;
      end;
  end; //_ifUniqueStyle

begin
  for i := 0 to FMasterPagesCount - 1 do
    if (_ifUniqueStyle(i)) then
    begin
      xml.Attributes.Clear();
      xml.Attributes.Add(ZETag_Attr_StyleName, FMasterPagesNames[i], false);
      xml.Attributes.Add(ZETag_style_page_layout_name, 'Mpm' + IntToStr(FPageLayoutsIndexes[i] + 1));
      xml.WriteTagNode(ZETag_style_master_page, true, true, false);
      _WriteMasterPage(FMasterPages[i]);
      xml.WriteEndTagNode(); //style:master-page
    end;
end; //WriteStylesMasterPages

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

  Attr.Add('xmlns:drawooo', 'http://openoffice.org/2010/draw');
  Attr.Add('xmlns:calcext', 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0', false);

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
      ZEContinuous, ZEHair: s := ' solid';
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
    if (not(((ProcessedStyle.Border[0].LineStyle = ZEContinuous) or (ProcessedStyle.Border[0].LineStyle = ZEHair)) and 
            (ProcessedStyle.Border[0].Weight = 0))) then
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
    if (not(((ProcessedStyle.Border[j].LineStyle = ZEContinuous) or (ProcessedStyle.Border[j].LineStyle = ZEHair)) and 
            (ProcessedStyle.Border[j].Weight = 0))) then
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
      _xml.Attributes.Add(ZETag_fo_background_color, '#' + ColorToHTMLHex(ProcessedStyle.BGColor));

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
    _xml.Attributes.Add(ZETag_fo_color, '#' + ColorToHTMLHex(ProcessedFont.Color), false);

  if (fsItalic in ProcessedFont.Style) then
  begin
    s := 'italic';
    _xml.Attributes.Add('fo:font-style', s, false);
    _xml.Attributes.Add('style:font-style-asian', s, false);
    _xml.Attributes.Add('style:font-style-complex', s, false);
  end;

  _xml.WriteEmptyTag(ZETag_style_text_properties, true, true);
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
//  const WriteHelper: TZEODFWriteHelper- помошник для записи
//RETURN
//      integer
function ODFCreateStyles(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String;
                          BOM: ansistring; const WriteHelper: TZEODFWriteHelper): integer;
var
  _xml: TZsspXMLWriterH;
  i: integer;

  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  //Добавить стили условного форматирования
  procedure _AddConditionalStyles();
  var
    i, j, k: integer;
    _cf: TZConditionalFormatting;
    _cfHelper: TZODFConditionalWriteHelper;
    num: integer;

  begin
    _cfHelper := WriteHelper.ConditionWriter;
    for i := 0 to PageCount - 1 do
    begin
      _cf := XMLSS.Sheets[_pages[i]].ConditionalFormatting;
      for j := 0 to _cf.Count - 1 do
      for k := 0 to _cf[j].Count - 1 do
        if (_cfHelper.TryAddApplyCFStyle(_cf[j][k].ApplyStyleID, num)) then
        begin
          _xml.Attributes.Clear();
          _xml.Attributes.Add(ZETag_Attr_StyleName, const_ConditionalStylePrefix + IntToStr(num));
          _xml.Attributes.Add(ZETag_style_family, 'table-cell', false);
          _xml.WriteTagNode(ZETag_StyleStyle, true, true, true);
          ODFWriteTableStyle(XMLSS, _xml, _cf[j][k].ApplyStyleID, false);
          _xml.WriteEndTagNode();
        end; //for k
    end; //for i
  end; //_AddConditionalStyles
  {$ENDIF}

  //<office:automatic-styles>..</office:automatic-styles>
  //Contains page layout (paper size, orientation, etc)
  procedure _WriteAutomaticStyle();
  begin
    _xml.Attributes.Clear();
    _xml.WriteTagNode(ZETag_office_automatic_styles, true, true, true);
    WriteHelper.WriteStylesPageLayouts(_xml, _pages);
    _xml.WriteEndTagNode(); //office:automatic-styles
  end; //_WriteAutomaticStyle

  // <office:master-styles>..</office:master-styles>
  //Contains master-pages styles (footers/headers etc)
  procedure _WriteOfficeMasterStyles();
  begin
    _xml.Attributes.Clear();
    _xml.WriteTagNode(ZETag_office_master_styles, true, true, true);
    WriteHelper.WriteStylesMasterPages(_xml, _pages);
    _xml.WriteEndTagNode(); //office:master-styles
  end; //_WriteOfficeMasterStyles

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
    _xml.Attributes.Add(ZETag_style_family, 'table-cell', false);
    _xml.WriteTagNode(ZETag_StyleStyle, true, true, true);
    ODFWriteTableStyle(XMLSS, _xml, -1, true);
    _xml.WriteEndTagNode();

    //Number formats
    for i := 0 to XMLSS.Styles.Count - 1 do
      WriteHelper.NumberFormatWriter.TryWriteNumberFormat(_xml, i, XMLSS.Styles[i].NumberFormat);

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    _AddConditionalStyles();
    {$ENDIF}

    _xml.WriteEndTagNode(); //office:styles

    _WriteAutomaticStyle(); //<office:automatic-styles>..</office:automatic-styles>

    _WriteOfficeMasterStyles(); // <office:master-styles> .. </office:master-styles>

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
  _PageNum: integer;

  //<config:config-item config:name="ConfigName" config:type="ConfigType">ConfigValue</config:config-item>
  procedure _AddConfigItem(const ConfigName, ConfigType, ConfigValue: string);
  begin
    _xml.Attributes.Clear();
    _xml.Attributes.Add(ZETag_config_name, ConfigName);
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
    _SheetOptions: TZSheetOptions;
    b: boolean;

  begin
    _PageNum := _pages[num];
    _xml.Attributes.Clear();
    _xml.Attributes.Add(ZETag_config_name, _names[num]);
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

  procedure _WriteOtherSettings();
  var
    i: integer;

  begin
    //Выделенный лист (ActiveTable). В OO только 1 шт.
    for i := 0 to PageCount - 1 do
      if (XMLSS.Sheets[_pages[i]].Selected) then
      begin
        _AddConfigItem('ActiveTable', 'string', _names[i]);
        break;
      end;
  end; //_WriteOtherSettings

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

    _xml.Attributes.Add(ZETag_config_name, 'ooo:view-settings');
    _xml.WriteTagNode('config:config-item-set', true, true, false);

    _AddConfigItem('VisibleAreaTop', 'int', '0');
    _AddConfigItem('VisibleAreaLeft', 'int', '0');
    _AddConfigItem('VisibleAreaWidth', 'int', '6773');
    _AddConfigItem('VisibleAreaHeight', 'int', '1813');

    _xml.Attributes.Clear();
    _xml.Attributes.Add(ZETag_config_name, 'Views');
    _xml.WriteTagNode('config:config-item-map-indexed', true, true, false);

    _xml.Attributes.Clear();
    _xml.WriteTagNode('config:config-item-map-entry', true, true, false);

    _xml.Attributes.Add(ZETag_config_name, 'Tables');
    _xml.WriteTagNode(ZETag_config_config_item_map_named, true, true, false);

    for i := 0 to PageCount - 1 do
      _WritePageSettings(i);

    _xml.WriteEndTagNode(); //config:config-item-map-named

    _WriteOtherSettings();

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
//  const _pages: TIntegerDynArray      - массив страниц
//  const _names: TStringDynArray       - массив имён страниц
//    PageCount: integer                - кол-во страниц
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//  const WriteHelper: TZEODFWriteHelper- помошник для записи
//RETURN
//      integer
function ODFCreateContent(var XMLSS: TZEXMLSS; Stream: TStream; const _pages: TIntegerDynArray;
                          const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String;
                          BOM: ansistring; const WriteHelper: TZEODFWriteHelper): integer;
var
  _xml: TZsspXMLWriterH;
  ColumnStyle, RowStyle: array of array of integer;  //стили столбцов/строк
  i: integer;
  _dt: TDateTime;
  _currColumn: TZColOptions;
  _currRow: TZRowOptions;
  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  _cfwriter: TZODFConditionalWriteHelper;
  {$ENDIF}

  //Заголовок для content.xml
  procedure WriteHeader();
  var
    i, j: integer;
    kol: integer;
    n: integer;
    ColStyleNumber, RowStyleNumber: integer;
    s: string;

    //Стили для колонок
    procedure WriteColumnStyle(now_i, now_j, now_StyleNumber, count_i{, count_j}: integer);
    var
      i, j: integer;
      start_j: integer;
      b: boolean;

    begin
      if (ColumnStyle[now_i][now_j] > -1) then
        exit;

      _currColumn := XMLSS.Sheets[_pages[now_i]].Columns[now_j];

      _xml.Attributes.Clear();
      _xml.Attributes.Add(ZETag_Attr_StyleName, 'co' + IntToStr(now_StyleNumber));
      _xml.Attributes.Add(ZETag_style_family, 'table-column', false);
      _xml.WriteTagNode(ZETag_StyleStyle, true, true, false);

      _xml.Attributes.Clear();
      //разрыв страницы (fo:break-before = auto | column | page)
      s := 'auto';
      if (_currColumn.Breaked) then
        s := 'column';
      _xml.Attributes.Add('fo:break-before', s);
      //Ширина колонки style:column-width
      _xml.Attributes.Add('style:column-width', ODFGetSizeToStr(_currColumn.WidthMM * 0.1), false);

      if (_currColumn.AutoFitWidth) then
        _xml.Attributes.Add(ZETag_style_use_optimal_column_width, ODFBoolToStr(true), false);

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
            if (XMLSS.Sheets[_pages[i]].Columns[j].WidthPix <> _currColumn.WidthPix) then
              b := false;
            //fo:break-before
            if (XMLSS.Sheets[_pages[i]].Columns[j].Breaked <> _currColumn.Breaked) then
              b := false;
            //style:use-optimal-column-width
            if (XMLSS.Sheets[_pages[i]].Columns[j].AutoFitWidth <> _currColumn.AutoFitWidth) then
               b := false;

            if (b) then
              ColumnStyle[i][j] := now_StyleNumber;
          end;

        start_j := 0;
      end;
    end; //WriteColumnStyle

    //Стили для строк
    procedure WriteRowStyle(now_i, now_j, now_StyleNumber, count_i{, count_j}: integer);
    var
      i, j, k: integer;
      start_j: integer;
      b: boolean;

    begin
      if (RowStyle[now_i][now_j] > -1) then
        exit;
      _xml.Attributes.Clear();
      _xml.Attributes.Add(ZETag_Attr_StyleName, 'ro' + IntToStr(now_StyleNumber));
      _xml.Attributes.Add(ZETag_style_family, 'table-row', false);
      _xml.WriteTagNode(ZETag_StyleStyle, true, true, false);

      _currRow := XMLSS.Sheets[_pages[now_i]].Rows[now_j];

      _xml.Attributes.Clear();
      //разрыв страницы (fo:break-before = auto | column | page)
      s := 'auto';
      if (_currRow.Breaked) then
        s := 'page';
      _xml.Attributes.Add('fo:break-before', s);
      //высота строки style:row-height
      _xml.Attributes.Add('style:row-height', ODFGetSizeToStr(_currRow.HeightMM * 0.1), false);
       //?? style:min-row-height

      //style:use-optimal-row-height - пересчитывать ли высоту, если содержимое ячеек изменилось
      //if (abs(_currRow.Height - XMLSS.Sheets[_pages[now_i]].DefaultRowHeight) < 0.001) then
      if (_currRow.AutoFitHeight) then
        _xml.Attributes.Add(ZETag_style_use_optimal_row_height, ODFBoolToStr(true), false);
      //fo:background-color - цвет фона
      k := XMLSS.Sheets[_pages[now_i]].Rows[now_j].StyleID;
      if (k > -1) then
        if (XMLSS.Styles.Count - 1 >= k) then
          if (XMLSS.Styles[k].BGColor <> XMLSS.Styles[-1].BGColor) then
            _xml.Attributes.Add(ZETag_fo_background_color, '#' + ColorToHTMLHex(XMLSS.Styles[k].BGColor), false);

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
            if (XMLSS.Sheets[_pages[i]].Rows[j].HeightPix <> _currRow.HeightPix) then
              b := false;
            //fo:break-before
            if (XMLSS.Sheets[_pages[i]].Rows[j].Breaked <> _currRow.Breaked) then
              b := false;
            //style:use-optimal-row-height
            if (XMLSS.Sheets[_pages[i]].Rows[j].AutoFitHeight <> _currRow.AutoFitHeight) then
              b := false;

            if (b) then
              RowStyle[i][j] := now_StyleNumber;
          end;

        start_j := 0;
      end;
    end; //WriteRowStyle

    procedure WriteTableStyle(num: integer);
    begin
      _xml.Attributes.Clear();
      _xml.Attributes.Add(ZETag_Attr_StyleName, 'ta' + IntToStr(num + 1), false);
      _xml.Attributes.Add(ZETag_style_family, 'table', false);
      _xml.Attributes.Add(ZETag_style_master_page_name, WriteHelper.GetMasterPageName(num), false);
      _xml.WriteTagNode(ZETag_StyleStyle, true, true, false);

      _xml.Attributes.Clear();
      //Possible attributes for <style:table-properties>
      //    fo:background-color           ??
      //    fo:break-after                ??
      //    fo:break-before               ??
      //    fo:keep-with-next             ??
      //    fo:margin                     ??
      //    fo:margin-bottom              ??
      //    fo:margin-left                ??
      //    fo:margin-right               ??
      //    fo:margin-top                 ??
      //    style:may-break-between-rows  ??
      //    style:page-number             ??
      //    style:rel-width               ??
      //    style:shadow                  ??
      //    style:width                   ??
      //    style:writing-mode            ??
      //    table:align                   ??
      //    table:border-model            ??
      //    table:display                 ??
      //    tableooo:tab-color            ??
      _xml.Attributes.Add('table:display', ODFBoolToStr(true), false);
      if (XMLSS.Sheets[_pages[num]].TabColor <> clWindow) then
        _xml.Attributes.Add(ZETag_tableooo_tab_color, '#' + ColorToHTMLHex(XMLSS.Sheets[_pages[num]].TabColor));

      _xml.WriteEmptyTag(ZETag_style_table_properties, true, false);

      _xml.WriteEndTagNode();
    end; //WriteTableStyle

  begin
    ZEWriteHeaderCommon(_xml, CodePageName, BOM);
    GenODContentAttr(_xml.Attributes);
    _xml.WriteTagNode('office:document-content', true, true, false);
    _xml.Attributes.Clear();
    _xml.WriteEmptyTag('office:scripts', true, false);  //потом на досуге можно подумать
    ZEWriteFontFaceDecls(XMLSS, _xml);

    ///********   Automatic Styles   ********///
    _xml.WriteTagNode(ZETag_office_automatic_styles, true, true, false);
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
        WriteColumnStyle(i, j, ColStyleNumber, kol{, n - 1});
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
        WriteRowStyle(i, j, RowStyleNumber, kol{, n});
        inc(RowStyleNumber);
      end;
    end;

    //******* остальные стили
    for i := 0 to XMLSS.Styles.Count - 1 do
    begin
      _xml.Attributes.Clear();
      _xml.Attributes.Add(ZETag_Attr_StyleName, 'ce' + IntToStr(i));
      _xml.Attributes.Add(ZETag_style_family, 'table-cell', false);

      if (WriteHelper.NumberFormatWriter.TryGetNumberFormatName(i, s)) then
        _xml.Attributes.Add(ZETag_style_data_style_name, s);

        //??style:parent-style-name = Default
      _xml.WriteTagNode(ZETag_StyleStyle, true, true, false);
      ODFWriteTableStyle(XMLSS, _xml, i, false);
      _xml.WriteEndTagNode(); //style:style
    end;

    //Стили для условного форматирования
    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    _cfwriter.WriteCFStyles(_xml);
    {$ENDIF}

    //****** Styles for tables
    for i := 0 to PageCount - 1 do
      WriteTableStyle(i);

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
    _t: Double;

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
        xml.WriteTagNode(ZETag_text_p, true, false, true);
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
            xml.WriteTag(ZETag_text_p, s, true, false, true);
            s := '';
          end else
            if (CellData[i] <> AnsiChar(#13)) then
              s := s + CellData[i];
        end;
        if (s > '') then
          xml.WriteTag(ZETag_text_p, s, true, false, true);
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
    _xml.Attributes.Add(ZETag_table_style_name, 'ta' + IntToStr(PageIndex + 1), false);
    b := ProcessedSheet.Protect;
    if (b) then
      _xml.Attributes.Add('table:protected', ODFBoolToStr(b), false);
    _xml.WriteTagNode('table:table', true, true, true);

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
      _xml.Attributes.Add(ZETag_table_style_name, 'co' + IntToStr(ColumnStyle[PageIndex][i]));
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
      _xml.Attributes.Add(ZETag_table_style_name, 'ro' + IntToStr(RowStyle[PageIndex][i]));
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
        {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
        _StyleID := _cfwriter.GetStyleNum(PageIndex, j, i);
        if ((_StyleID >= 0) and (_StyleID < XMLSS.Styles.Count + _cfwriter.StylesCount)) then
        {$ELSE}
        _StyleID := ProcessedSheet.Cell[j, i].CellStyle;
        if ((_StyleID >= 0) and (_StyleID < XMLSS.Styles.Count)) then
        {$ENDIF}
          _xml.Attributes.Add(ZETag_table_style_name, 'ce' + IntToStr(_StyleID), false);

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
              WriteHelper.NumberFormatWriter.TryGetNumberFormatAddProp(ProcessedSheet.Cell[j, i].CellStyle, t);

              if (t and ZE_NUMFORMAT_NUM_IS_PERCENTAGE = ZE_NUMFORMAT_NUM_IS_PERCENTAGE) then
                ss := 'percentage'
              else
              if (t and ZE_NUMFORMAT_NUM_IS_CURRENCY = ZE_NUMFORMAT_NUM_IS_CURRENCY) then
              begin
                //TODO: add attribute office:currency
                //_xml.Attributes.Add('office:currency', 'BYR', false);
                ss := 'currency';
              end
              else
                ss := 'float';

              _xml.Attributes.Add('office:value', ZEFloatSeparator(FormatFloat('0.#######', ZETryStrToFloat(_CellData))), false);
            end;
          ZEBoolean:
            begin
              ss := 'boolean';
              _xml.Attributes.Add('office:boolean-value', ODFBoolToStr(ZETryStrToBoolean(_CellData)), false);
            end;
          ZEDateTime:
            begin
              b := TryZEStrToDateTime(_CellData, _dt);
              if (not b) then
                if (ZEIsTryStrToFloat(_CellData, _t)) then
                begin
                  b := true;
                  _CellData := ZEDateTimeToStr(_t);
                  _dt := _t;
                end;

              if (b) then
              begin
                WriteHelper.NumberFormatWriter.TryGetNumberFormatAddProp(ProcessedSheet.Cell[j, i].CellStyle, t);
                if (t and ZE_NUMFORMAT_DATE_IS_ONLY_TIME = ZE_NUMFORMAT_DATE_IS_ONLY_TIME) then
                begin
                  ss := 'time';
                  _xml.Attributes.Add('office:time-value', ZEDateTimeToPTDurationStr(_dt), false);
                end
                else
                begin
                  ss := 'date';
                  _xml.Attributes.Add('office:date-value', _CellData, false);
                end;
              end; //if
            end; //ZEDateTime
          else
            // всё остальное считаем строкой (потом подправить, возможно, добавить новые типы)
            {ZEansistring ZEError ZEDateTime}
        end; //case

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
      end; //for j
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

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    _cfwriter.WriteCalcextCF(_xml, PageIndex);
    {$ENDIF}

    _xml.WriteEndTagNode(); //table:table

    //автофильтр
    if Trim(ProcessedSheet.AutoFilter)<>'' then
    begin
      _xml.Attributes.Clear;
      _xml.WriteTagNode('table:database-ranges', true, true, false);
      _xml.Attributes.Add('table:display-filter-buttons','true');
      ss:=ProcessedSheet.AutoFilter;
      s:=#39+ProcessedSheet.Title+#39+'.'+Copy(ss,1,pos(':',ss))          
        +#39+ProcessedSheet.Title+#39+'.'+Copy(ss,pos(':',ss)+1,Length(ss));
      _xml.Attributes.Add('table:target-range-address',s);
      _xml.WriteEmptyTag('table:database-range', true, true);
      _xml.WriteEndTagNode();
    end;
    
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

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    _cfwriter := WriteHelper.ConditionWriter;
    {$ENDIF}

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
    s := ZEDateTimeToStr(XMLSS.DocumentProperties.Created);
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
  _WriteHelper: TZEODFWriteHelper;

begin
  result := 0;
  Stream := nil;
  _WriteHelper := nil;
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

    _WriteHelper := TZEODFWriteHelper.Create(XMLSS, _pages, _names, kol);

    Stream := TFileStream.Create(PathName + 'styles.xml', fmCreate);
    ODFCreateStyles(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM, _WriteHelper);
    FreeAndNil(Stream);

    Stream := TFileStream.Create(PathName + 'content.xml', fmCreate);
    ODFCreateContent(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM, _WriteHelper);
    FreeAndNil(Stream);

    Stream := TFileStream.Create(PathName + 'meta.xml', fmCreate);
    ODFCreateMeta(XMLSS, Stream, TextConverter, CodePageName, BOM);
    FreeAndNil(Stream);

    Stream := TFileStream.Create(PathName + 'settings.xml', fmCreate);
    ODFCreateSettings(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
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
    if (Assigned(_WriteHelper)) then
      FreeAndNil(_WriteHelper);
  end;
end; //SaveXmlssToODFSPath

//Сохраняет незапакованный документ в формате Open Document
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//RETURN
//      integer
function SaveXmlssToODFSPath(var XMLSS: TZEXMLSS; PathName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string): integer; overload;
begin
  result := SaveXmlssToODFSPath(XMLSS, PathName, SheetsNumbers, SheetsNames, ZEGetDefaultUTF8Converter(), 'UTF-8', '');
end; //SaveXmlssToODFSPath

//Сохраняет незапакованный документ в формате Open Document
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//RETURN
//      integer
function SaveXmlssToODFSPath(var XMLSS: TZEXMLSS; PathName: string): integer; overload;
begin
  result := SaveXmlssToODFSPath(XMLSS, PathName, [], []);
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
  _WriteHelper: TZEODFWriteHelper;

begin
  zip := nil;
  StreamC := nil;
  StreamME := nil;
  StreamS := nil;
  StreamST := nil;
  StreamMA := nil;
  _WriteHelper := nil;
  result := 0;
  try

    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      result := 2;
      exit;
    end;

    _WriteHelper := TZEODFWriteHelper.Create(XMLSS, _pages, _names, kol);

    zip := TZipper.Create();

    StreamST := TMemoryStream.Create();
    ODFCreateStyles(XMLSS, StreamST, _pages, _names, kol, TextConverter, CodePageName, BOM, _WriteHelper);

    StreamC := TMemoryStream.Create();
    ODFCreateContent(XMLSS, StreamC, _pages, _names, kol, TextConverter, CodePageName, BOM, _WriteHelper);

    StreamME := TMemoryStream.Create();
    ODFCreateMeta(XMLSS, StreamME, TextConverter, CodePageName, BOM);

    StreamS := TMemoryStream.Create();
    ODFCreateSettings(XMLSS, StreamS, _pages, _names, kol, TextConverter, CodePageName, BOM);

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
    if (Assigned(_WriteHelper)) then
      FreeAndNil(_WriteHelper);

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

//Сохраняет документ в формате Open Document
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      FileName: string                  - имя файла для сохранения
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//RETURN
//      integer
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string): integer; overload;
begin
  result := SaveXmlssToODFS(XMLSS, FileName, SheetsNumbers, SheetsNames, ZEGetDefaultUTF8Converter(), 'UTF-8', '');
end; //SaveXmlssToODFS

//Сохраняет документ в формате Open Document
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      FileName: string                  - имя файла для сохранения
//RETURN
//      integer
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
begin
  result := SaveXmlssToODFS(XMLSS, FileName, [], []);
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
  _WriteHelper: TZEODFWriteHelper;

begin
  azg := nil;
  _WriteHelper := nil;
  try

    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      result := 2;
      exit;
    end;

    _WriteHelper := TZEODFWriteHelper.Create(XMLSS, _pages, _names, kol);

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

    Stream := azg.NewStream('styles.xml');
    ODFCreateStyles(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM, _WriteHelper);
    azg.SealStream(Stream);

    Stream := azg.NewStream('content.xml');
    ODFCreateContent(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM, _WriteHelper);
    azg.SealStream(Stream);

    Stream := azg.NewStream('meta.xml');
    ODFCreateMeta(XMLSS, Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);

    Stream := azg.NewStream('settings.xml');
    ODFCreateSettings(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);

    Stream := azg.NewStream('META-INF/manifest.xml');
    ODFCreateManifest(Stream, TextConverter, CodePageName, BOM);
    azg.SealStream(Stream);

    azg.SaveAndSeal;
  finally
    ZESClearArrays(_pages, _names);
    azg.Free;

    if (Assigned(_WriteHelper)) then
      FreeAndNil(_WriteHelper);
  end;
  Result := 0;
end; //ExportXmlssToODFS


/////////////////// чтение

//Возвращает размер измерения в ММ
//INPUT
//  const value: string     - строка со значением
//  out RetSize: real       - возвращаемое значение
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
            {$IFDEF Z_USE_FORMAT_SETTINGS}
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
  if (su = 'pt') then
    RetSize := r * _PointToMM
  else
  if (su = 'in') then
    RetSize := r * 25.4
  else
    result := false;
end; //ODFGetValueSizeMM

//Читает стиль ячейки
//INPUT
//  var xml: TZsspXMLReaderH  - тэг-парсер
//  var _style: TZSTyle       - прочитанный стиль
//  var StyleProperties: TZEODFStyleProperties - условия для условного форматирования
procedure ZEReadODFCellStyleItem(var xml: TZsspXMLReaderH; var _style: TZSTyle
                                {$IFDEF ZUSE_CONDITIONAL_FORMATTING}; var StyleProperties: TZEODFStyleProperties{$ENDIF});
var
  t: integer;
  HAutoForced: boolean;
  r: real;
  s: string;

  function ODF12AngleUnit(const un: string; const scale: double): boolean;
  var
    err: integer;
    d: double;

  begin
    Result := AnsiEndsStr(un, s);
    if Result then
    begin
      Val( Trim(Copy( s, 1, length(s) - length(un))), d, err);
      Result := err = 0;
      if Result then
         t := round(d * scale);
    end;
  end;

begin
  HAutoForced := false; // pre-clean: paragraph properties may come before cell properties

  while ((xml.TagType <> 6) or (xml.TagName <> ZETag_StyleStyle)) do
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
      s := xml.Attributes.ItemsByName[ZETag_fo_background_color];
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
      if not HAutoForced then
      begin
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
      end; //if
    end else //if

    if ((xml.TagName = ZETag_style_text_properties) and (xml.TagType in [4, 5])) then
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
      s := xml.Attributes.ItemsByName[ZETag_fo_color];
      if (s > '') then
        _style.Font.Color := HTMLHexToColor(s);
    end; //if

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    if ((xml.TagType = 5) and (xml.TagName = ZETag_style_map)) then
    begin
      t := StyleProperties.ConditionsCount;
      inc(StyleProperties.ConditionsCount);
      SetLength(StyleProperties.Conditions, StyleProperties.ConditionsCount);
      StyleProperties.Conditions[t].ConditionValue := xml.Attributes[ZETag_style_condition];
      StyleProperties.Conditions[t].ApplyStyleName := xml.Attributes[ZETag_style_apply_style_name];
      StyleProperties.Conditions[t].ApplyBaseCellAddres := xml.Attributes['style:base-cell-address'];
      StyleProperties.Conditions[t].ApplyStyleIDX := -1;
    end;
    {$ENDIF}
  end; //while
end; //ZEReadODFCellStyleItem

//Чтение стилей документа и автоматических стилей (составная часть стилей)
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//      stream: TStream - поток для чтения
//  var ReadHelper: TZEODFReadHelper - для хранения доп. инфы
//RETURN
//      boolean - true - всё ок
function ReadODFStyles(var XMLSS: TZEXMLSS; stream: TStream; var ReadHelper: TZEODFReadHelper): boolean;
var
  xml: TZsspXMLReaderH;
  _Style: TZStyle;
  num: integer;
  s: string;

  //Прочитать один стиль
  procedure _ReadOneStyle();
  var
    i: integer;

  begin
    if (xml.Attributes[ZETag_style_family] = 'table-cell') then
    begin
      num := ReadHelper.StylesCount;
      ReadHelper.AddStyle();
      _Style := ReadHelper.Style[num];

      ReadHelper.StylesProperties[num].name := xml.Attributes[ZETag_Attr_StyleName];
      s := xml.Attributes['style:parent-style-name'];
      ReadHelper.StylesProperties[num].ParentName := s;
      if (length(s) > 0) then
      begin
        ReadHelper.StylesProperties[num].isHaveParent := true;
        for i := 0 to num - 1 do
          if (ReadHelper.StylesProperties[num].name = s) then
          begin
            _Style.Assign(ReadHelper.Style[i]);
            break;
          end;
      end; //if

      ZEReadODFCellStyleItem(xml, _style {$IFDEF ZUSE_CONDITIONAL_FORMATTING}, ReadHelper.StylesProperties[num] {$ENDIF});
    end;
  end; //_ReadOneStyle

  //Чтение стилей
  procedure _ReadStyles();
  begin
    while (not xml.Eof()) do
    begin
      if (not xml.ReadTag()) then
        break;

      if (xml.TagType = 4) then
      begin
        //style:style - стиль
        if (xml.TagName = ZETag_StyleStyle) then
          _ReadOneStyle()
        else
        //office:automatic-styles (Page Layouts etc)
        if (xml.TagName = ZETag_office_automatic_styles) then
          ReadHelper.ReadAutomaticStyles(xml)
        else
        //office:master-styles
        if (xml.TagName = ZETag_office_master_styles) then
          ReadHelper.ReadMasterStyles(xml)
        else
        //NumberStyles:
        //    number:date-style
        //    number:number-style
        //    number:currency-style
        //    number:percentage-style
          ReadHelper.NumberStylesHelper.ReadKnownNumberFormat(xml);
      end; //if

      //style:default-style - стиль по умолчанию
      //number:number-style - числовой стиль
      //number:currency-style - валютный стиль
      //office:automatic-styles - автоматические стили
      //office:master-styles - мастер страница
    end; //while
  end; //_ReadStyles

begin
  result := false;
  xml := nil;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(stream) <> 0) then
      exit;
    _ReadStyles();
    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
  end;
end; //ReadODFStyles

//Чтение содержимого документа ODS (content.xml)
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//      stream: TStream - поток для чтения
//  var ReadHelper: TZEODFReadHelper - для хранения доп. инфы
//RETURN
//      boolean - true - всё ок
function ReadODFContent(var XMLSS: TZEXMLSS; stream: TStream; var ReadHelper: TZEODFReadHelper): boolean;
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
  _celltext: string;
  _Sheet: TZSheet;                      //Current reading sheet
  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  _RowDefaultStyleCFNumber: integer;    //Является ли стиль по умолчанию для ячейки в строке условным стилем
  _CellStyleCFNumber: integer;          //Является ли стиль ячейки условным (номер в массиве)
  {$ENDIF}

  function IfTag(const TgName: string; const TgType: integer): boolean;
  begin
    result := (xml.TagType = TgType) and (xml.TagName = TgName);
  end;

  //Ищет номер стиля по названию
  //INPUT
  //  const st: string        - название стиля
  //  out retCFIndex: integer - возвращает номер условного стиля в массиве,
  //                            если < 0 - стиль не условный.
  //                            Если >= StyleCount, то стиль в массиве ReadHelper.StylesProperties.
  function _FindStyleID(const st: string {$IFDEF ZUSE_CONDITIONAL_FORMATTING};
                                         out retCFIndex: integer
                                         {$ENDIF}): integer;
  var
    i: integer;

  begin
    result := -1;
    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    retCFIndex := -1;
    {$ENDIF}
    for i := 0 to StyleCount - 1 do
    if (ODFStyles[i].name = st) then
    begin
      result := ODFStyles[i].index;
      {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
      if (ODFStyles[i].ConditionsCount > 0) then
        retCFIndex := i;
      {$ENDIF}
      break;
    end;
    if (result < 0) then
    begin
      for i := 0 to ReadHelper.StylesCount - 1 do
        if (ReadHelper.StylesProperties[i].name = st) then
        begin
          if (ReadHelper.StylesProperties[i].index = -2) then
            ReadHelper.StylesProperties[i].index := XMLSS.Styles.Add(ReadHelper.Style[i], true);
          result := ReadHelper.StylesProperties[i].index;

          {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
          if (ReadHelper.StylesProperties[i].ConditionsCount > 0) then
            retCFIndex := i + StyleCount;
          {$ENDIF}

          break;
        end;
    end;
  end; //_FindStyleID

  //Автоматические стили
  procedure _ReadAutomaticStyle();
  var
    _stylefamily: string;
    _stylename: string;
    s: string;
    _style: TZSTyle;
    _data_style_name: string;

    //Чтение стиля для ячейки
    procedure _ReadCellStyle();
    var
      i: integer;
      b: boolean;

    begin
      if (StyleCount >= MaxStyleCount) then
      begin
        MaxStyleCount := StyleCount + 20;
        SetLength(ODFStyles, MaxStyleCount);
      end;

      _data_style_name := xml.Attributes[ZETag_style_data_style_name];

      ODFClearStyleProperties(ODFStyles[StyleCount]);
      ODFStyles[StyleCount].name := _stylename;
      _style.Assign(XMLSS.Styles.DefaultStyle);

      s := xml.Attributes['style:parent-style-name'];
      ODFStyles[StyleCount].ParentName := s;
      if (length(s) > 0) then
      begin
        ODFStyles[StyleCount].isHaveParent := true;
        b := true;
        for i := 0 to StyleCount - 1 do
          if (ODFStyles[i].name = s) then
          begin
            b := false;
            _style.Assign(XMLSS.Styles[ODFStyles[i].index]);
            break;
          end;
        if (b) then
          for i := 0 to ReadHelper.StylesCount - 1 do
            if (ReadHelper.StylesProperties[i].name = s) then
            begin
              _style.Assign(ReadHelper.Style[i]);
              break;
            end;
      end; //if

      if (xml.TagType = 4) then
        ZEReadODFCellStyleItem(xml, _style {$IFDEF ZUSE_CONDITIONAL_FORMATTING}, ODFStyles[StyleCount] {$ENDIF});

      if (_data_style_name <> '') then
        if (ReadHelper.NumberStylesHelper.TryGetFormatStrByNum(_data_style_name, s)) then
          _style.NumberFormat := s;

      ODFStyles[StyleCount].index := XMLSS.Styles.Add(_style, true);
      inc(StyleCount);
    end; //_ReadCellStyle

    procedure _ReadTableColumnStyle();
    begin
      if (ColStyleCount >= MaxColStyleCount) then
      begin
        MaxColStyleCount := ColStyleCount + 20;
        SetLength(ODFColumnStyles, MaxColStyleCount);
      end;
      ODFColumnStyles[ColStyleCount].name := '';
      ODFColumnStyles[ColStyleCount].breaked := false;
      ODFColumnStyles[ColStyleCount].width := 25;

      if (xml.TagType = 4) then
        while (not IfTag(ZETag_StyleStyle, 6)) do
        begin
          if (xml.Eof()) then
            break;
          xml.ReadTag();
          if ((xml.TagName = 'style:table-column-properties') and (xml.TagType in [4, 5])) then
          begin
            ODFColumnStyles[ColStyleCount].name := _stylename;
            s := xml.Attributes.ItemsByName['fo:break-before'];
            if (s = 'column') then
              ODFColumnStyles[ColStyleCount].breaked := true;
            s := xml.Attributes.ItemsByName['style:column-width'];
            if (s > '') then
              ODFGetValueSizeMM(s, ODFColumnStyles[ColStyleCount].width);
            s := xml.Attributes.ItemsByName[ZETag_style_use_optimal_column_width];
            ODFColumnStyles[ColStyleCount].AutoWidth := ZETryStrToBoolean(s);
          end;
        end; //while
      inc(ColStyleCount);
    end; //_ReadTableColumnStyle

    procedure _ReadTableRowStyle();
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

      if (xml.TagType = 4) then
        while (not ifTag(ZETag_StyleStyle, 6)) do
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
            s := xml.Attributes.ItemsByName[ZETag_fo_background_color];
            if (s > '') then
              ODFRowStyles[RowStyleCount].color := HTMLHexToColor(s);

            s := xml.Attributes.ItemsByName[ZETag_style_use_optimal_row_height];
            ODFRowStyles[RowStyleCount].AutoHeight := ZETryStrToBoolean(s);
          end;
        end; //while
      inc(RowStyleCount);
    end; //_ReadTableRowStyle

    procedure _ReadTableStyle();
    begin
      if (TableStyleCount >= MaxTableStyleCount) then
      begin
        MaxTableStyleCount := TableStyleCount + 20;
        SetLength(ODFTableStyles, MaxTableStyleCount);
      end;
      ODFTableStyles[TableStyleCount].name := _stylename;
      ODFTableStyles[TableStyleCount].isColor := false;
      ODFTableStyles[TableStyleCount].MasterPageName := xml.Attributes.ItemsByName[ZETag_style_master_page_name];

      if (xml.TagType = 4) then
        while (not ifTag(ZETag_StyleStyle, 6)) do
        begin
          if (xml.Eof()) then
            break;
          xml.ReadTag();

          if ((xml.TagName = ZETag_style_table_properties) and (xml.TagType in [4, 5])) then
          begin
            s := xml.Attributes.ItemsByName[ZETag_tableooo_tab_color];
            if (s > '') then
            begin
              ODFTableStyles[TableStyleCount].isColor := true;
              ODFTableStyles[TableStyleCount].Color := HTMLHexToColor(s);
            end;
          end;

        end; //while
      inc(TableStyleCount);
    end; //_ReadTableStyle

  begin
    _style := nil;
    try
      _style := TZStyle.Create();
      while (not IfTag(ZETag_office_automatic_styles, 6)) do
      begin
        if (xml.Eof()) then
          break;
        xml.ReadTag();

        if ((xml.TagType in [4, 5]) and (xml.TagName = ZETag_StyleStyle)) then
        begin
          _stylefamily := xml.Attributes.ItemsByName[ZETag_style_family];
          _stylename := xml.Attributes.ItemsByName[ZETag_Attr_StyleName];

          if (_stylefamily = 'table-column') then //столбец
            _ReadTableColumnStyle()
          else
          if (_stylefamily = 'table-row') then //строка
            _ReadTableRowStyle()
          else
          if (_stylefamily = 'table') then //таблица
            _ReadTableStyle()
          else
          if (_stylefamily = 'table-cell') then //ячейка
            _ReadCellStyle();
        end
        else
          ReadHelper.NumberStylesHelper.ReadKnownNumberFormat(xml);
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
    _RowDefaultStyleID: integer;  //ID стиля по-умолчанию в строке
    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    _tmpNum: integer;
    {$ENDIF}

    //Повторить строку
    procedure _RepeatRow();
    var
      i, j, n, z: integer;

    begin
      {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
      if (isRepeatRow) then
      begin
        if (_RepeatRowCount <= 2000) then
        begin
          z := _RepeatRowCount;
          if ((not _IsHaveTextInRow) and (z > 512)) then
            z := 512;
        end else
          z := 1;
      end else
        z := 1;
      ReadHelper.ConditionReader.ProgressLine(_CurrentRow - 1, z);
      {$ENDIF}
      //Максимум строка может повторятся 2000 раз
      if ((isRepeatRow) and (_RepeatRowCount <= 2000)) then
      begin
        //Если в строке небыло текста и её нужно повторить более 512 раз - уменьшаем кол-во
        // до 512 раз
        if ((not _IsHaveTextInRow) and (_RepeatRowCount > 512)) then
          _RepeatRowCount := 512;
        n := _Sheet.ColCount - 1;
        z := _CurrentRow - 1;
        CheckRow(_CurrentPage, _CurrentRow + _RepeatRowCount);
        for i := 1 to _RepeatRowCount do
        begin
          for j := 0 to n do
            _Sheet.Cell[j, _CurrentRow].Assign(_Sheet.Cell[j, z]);
        end;
        inc(_CurrentRow, _RepeatRowCount - 1);
      end; //if
    end; //_RepeatRow

    //Чтение ячейки
    procedure _ReadCell();
    var
      i: integer;
      _isnf: boolean;
      _numX, _numY: integer;
      _isHaveTextCell: boolean;
      _kol: integer;
      _isStringValue: boolean;
      _stringValue: string;

    begin
      if (((xml.TagName = 'table:table-cell') or (xml.TagName = 'table:covered-table-cell')) and (xml.TagType in [4, 5])) then
      begin
        _stringValue := '';
        _isStringValue := false;
        CheckCol(_CurrentPage, _CurrentCol + 1);
        _CurrCell := _Sheet.Cell[_CurrentCol, _CurrentRow];
        s := xml.Attributes.ItemsByName['table:number-columns-repeated'];
        isRepeatCell := TryStrToInt(s, _RepeatCellCount);
        if (not isRepeatCell) then
          _RepeatCellCount := 1;

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
          _Sheet.MergeCells.AddRectXY(_CurrentCol, _CurrentRow, _CurrentCol + _numX, _CurrentRow + _numY);
        end;

        //стиль ячейки
        //TODO: Нужно разобраться с приоритетом. Если для ячейки не указан стиль,
        //      какой стиль нужно брать раньше: стиль столбца или строки?
        //      Пока пусть будет как для столбца. Или, может, если не указан стиль,
        //      то ставить дефолтный (-1)?
        s := xml.Attributes.ItemsByName[ZETag_table_style_name];
        _CurrCell.CellStyle := _Sheet.Columns[_CurrentCol].StyleID;
        {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
        _CellStyleCFNumber := ReadHelper.ConditionReader.GetColumnCF(_CurrentCol);
        {$ENDIF}
        if (s > '') then
          _CurrCell.CellStyle := _FindStyleID(s {$IFDEF ZUSE_CONDITIONAL_FORMATTING}, _CellStyleCFNumber{$ENDIF})
        else
        begin
          {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
          _CellStyleCFNumber := _RowDefaultStyleCFNumber;
          {$ENDIF}
        end;

        {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
        ReadHelper.ConditionReader.CheckCell(_CurrentCol, _CellStyleCFNumber, _RepeatCellCount);
        {$ENDIF}

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
          s := xml.Attributes.ItemsByName['office:value-type'];
        _CurrCell.CellType := ODFTypeToZCellType(s);

        case (_CurrCell.CellType) of
          ZENumber:
            _CurrCell.Data := xml.Attributes.ItemsByName['office:value'];
          ZEDateTime:
            begin
              if (UpperCase(s) = 'TIME') then
                _CurrCell.AsDateTime := ZEPTDateDurationToDateTime(xml.Attributes.ItemsByName['office:time-value'])
              else
                _CurrCell.Data := xml.Attributes.ItemsByName['office:date-value'];
            end;
          ZEString:
            begin
              _stringValue := xml.Attributes.ItemsByName['office:string-value'];
              _isStringValue := _stringValue <> '';
            end;
          ZEBoolean:
            begin
              s := '0';
              if (ZETryStrToBoolean(xml.Attributes.ItemsByName['office:boolean-value'])) then
                s := '1';
              _CurrCell.Data := s;
            end;
        end; //case

        //защищённость ячейки
        s := xml.Attributes.ItemsByName['table:protected']; //{tut} надо будет добавить ещё один стиль
        //table:number-matrix-rows-spanned ??
        //table:number-matrix-columns-spanned ??

        _celltext := '';
        _isHaveTextCell := false;
        if (xml.TagType = 4) then
        begin
          _isnf := false;
          while (not(((xml.TagType = 6) and (xml.TagName = 'table:table-cell') or (xml.TagName = 'table:covered-table-cell')))) do
          begin
            if (xml.Eof()) then
              break;
            xml.ReadTag();

            //Текст ячейки
            if (IfTag(ZETag_text_p, 4)) then
            begin
              _IsHaveTextInRow := true;
              _isHaveTextCell := true;
              if (_isnf) then
                _celltext := _celltext + {$IFDEF FPC} LineEnding {$ELSE} sLineBreak {$ENDIF};
              while (not IfTag(ZETag_text_p, 6)) do
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
                if ((xml.TagName <> ZETag_text_p) and (xml.TagName <> 'text:a') and (xml.TagName <> 'text:s') and
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
                if (IfTag(ZETag_text_p, 4)) then
                begin
                  while (not IfTag(ZETag_text_p, 6)) do
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

        if (not (_CurrCell.CellType in [ZENumber, ZEDateTime, ZEBoolean])) then
          _CurrCell.Data := ZEReplaceEntity(_celltext);

        if ((_CurrCell.CellType = ZEString) and _isStringValue) then
        begin
          _CurrCell.Data := ZEReplaceEntity(_stringValue);
          _isHaveTextCell := true;
        end;

        //Если ячейку нужно повторить
        if (isRepeatCell) then
          //Если в ячейке небыло текста и нужно повторить её более 255 раз - игнорируется повторение
          if (_isHaveTextCell) or (_RepeatCellCount < 255) then
          begin
            CheckCol(_CurrentPage, _CurrentCol + _RepeatCellCount + 1);
            for i := 1 to _RepeatCellCount do
              _Sheet.Cell[_CurrentCol + i, _CurrentRow].Assign(_CurrCell);
            //-1, т.к. нужно учитывать, что номер ячейки увеличивается на 1 каждый раз
            inc(_CurrentCol, _RepeatCellCount - 1);
          end;

        inc(_CurrentCol);
      end; //if
    end; //_ReadCell

  begin
    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    ReadHelper.ConditionReader.Clear();
    {$ENDIF}
    _CurrentRow := 0; 
    _CurrentCol := 0;
    _MaxCol := 0;
    _CurrentPage := XMLSS.Sheets.Count;
    XMLSS.Sheets.Count := _CurrentPage + 1;
    _Sheet := XMLSS.Sheets[_CurrentPage];
    _Sheet.RowCount := 1;
    _Sheet.Title := ZEReplaceEntity(xml.Attributes.ItemsByName['table:name']);
    _Sheet.Protect := ZEStrToBoolean(xml.Attributes.ItemsByName['table:protected']);

    s := xml.Attributes.ItemsByName[ZETag_table_style_name];
    if (s > '') then
      for i := 0 to TableStyleCount - 1 do
      if (ODFTableStyles[i].name = s) then
      begin
        if (ODFTableStyles[i].isColor) then
          _Sheet.TabColor := ODFTableStyles[i].Color;
        ReadHelper.ApplyMasterPageStyle(_Sheet.SheetOptions, ODFTableStyles[i].MasterPageName);
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
        {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
        ReadHelper.ConditionReader.ClearLine();
        {$ENDIF}

        //кол-во повторений строки
        s := xml.Attributes.ItemsByName['table:number-rows-repeated'];
        isRepeatRow := TryStrToInt(s, _RepeatRowCount);
        _IsHaveTextInRow := false;
        //стиль строки
        s := xml.Attributes.ItemsByName[ZETag_table_style_name];

        if (s > '') then
          for i := 0 to RowStyleCount - 1 do
            if (ODFRowStyles[i].name = s) then
            begin
              CheckRow(_CurrentPage, _CurrentRow + 1);
              _Sheet.Rows[_CurrentRow].Breaked := ODFRowStyles[i].breaked;
              if (ODFRowStyles[i].height >= 0) then
                _Sheet.Rows[_CurrentRow].HeightMM := ODFRowStyles[i].height;
              _Sheet.Rows[_CurrentRow].AutoFitHeight := ODFRowStyles[i].AutoHeight;
            end;

        //стиль ячейки по умолчанию
        s := xml.Attributes.ItemsByName['table:default-cell-style-name'];
        if (Length(s) > 0) then
        begin
          _RowDefaultStyleID := _FindStyleID(s {$IFDEF ZUSE_CONDITIONAL_FORMATTING}, _RowDefaultStyleCFNumber {$ENDIF});
          _Sheet.Rows[_CurrentRow].StyleID := _RowDefaultStyleID;
        end else
        begin
          //_RowDefaultStyleID := -1;
          {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
          _RowDefaultStyleCFNumber := -1;
          {$ENDIF}
        end;

        //Видимость: visible | collapse | filter
        s := xml.Attributes.ItemsByName['table:visibility'];
        if (s = 'collapse') then
          _Sheet.Rows[_CurrentRow].Hidden := true;

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
        s := xml.Attributes.ItemsByName[ZETag_table_style_name];
        for i := 0 to ColStyleCount - 1 do
          if (ODFColumnStyles[i].name = s) then
          begin
            _Sheet.Columns[_MaxCol].Breaked := ODFColumnStyles[i].breaked;
            _Sheet.Columns[_MaxCol].AutoFitWidth := ODFColumnStyles[i].AutoWidth;
            if (ODFColumnStyles[i].width >= 0) then
              _Sheet.Columns[_MaxCol].WidthMM := ODFColumnStyles[i].width;
            break;
          end;

        s := xml.Attributes.ItemsByName['table:default-cell-style-name'];
        if (s > '') then
        {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
        begin
          _Sheet.Columns[_MaxCol].StyleID := _FindStyleID(s, _tmpNum);
          if (_tmpNum >= 0) then
            ReadHelper.ConditionReader.AddColumnCF(_MaxCol, _tmpNum);
        end else
          _tmpNum := -1;
        {$ELSE}
          _Sheet.Columns[_MaxCol].StyleID := _FindStyleID(s);
        {$ENDIF}

        s := xml.Attributes.ItemsByName['table:number-columns-repeated'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            if (t < 255) then
            begin
              dec(t); //т.к. один столбец уже есть
              CheckCol(_CurrentPage, _MaxCol + t + 1);
              for i := 1 to t do
              begin
                _Sheet.Columns[_MaxCol + i].Assign(XMLSS.Sheets[_CurrentPage].Columns[_MaxCol]);
                {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
                if (_tmpNum >= 0) then
                  ReadHelper.ConditionReader.AddColumnCF(_MaxCol + i, _tmpNum);
                {$ENDIF}
              end;
              inc(_MaxCol, t);
            end;

        inc(_MaxCol);
      end; //if

      //ячейка
      _ReadCell();

      // для LibreOffice >= 4.0 условное форматирование
      {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
      if (ifTag(const_calcext_conditional_formats, 4)) then
        ReadHelper.ConditionReader.ReadCalcextTag(xml, _CurrentPage);
      {$ENDIF}
    end; //while

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    ReadHelper.ConditionReader.ApplyConditionStylesToSheet(_CurrentPage,
                                                           ReadHelper.StylesProperties,
                                                           ReadHelper.StylesCount,
                                                           ODFStyles,
                                                           StyleCount);
    {$ENDIF}

    for i := 0 to _Sheet.ColCount - 1 do
      _Sheet.Columns[i].StyleID := -1;
  end; //_ReadTable

  //считываем фильтр
  procedure _ReadAutoFilter();
  var st:string;
  begin
    //В ods область фильра задается через имя листа
    //при этом есть два варианта:
    // 1. "Лист1.A4:Лист1.C4" - если имя листа односложное
    // 2. "'Лист №1'.A4:'Лист №1'.C4" - если имя листа сложное
    // поэтому дважды вычищаем считанное значение
    st := xml.Attributes.ItemsByName['table:target-range-address'];
    {$IFDEF FPC_OR_DELPHI_UNICODE}
    st := ReplaceStr(st, #39 + _Sheet.Title + #39 + '.', '');
    st := ReplaceStr(st, _Sheet.Title + '.', '');
    {$ELSE}
    st := AnsiReplaceStr(st, #39 + _Sheet.Title + #39 + '.', '');
    st := AnsiReplaceStr(st, _Sheet.Title + '.', '');
    {$ENDIF}

    _Sheet.AutoFilter:=st;
  end;

  procedure _ReadDocument();
  begin
    while (not xml.Eof()) do
    begin
      xml.ReadTag();
      ErrorReadCode := ErrorReadCode or xml.ErrorCode;
      if (ifTag(ZETag_office_automatic_styles, 4)) then
        _ReadAutomaticStyle();
      //if (ifTag('office:styles', 4)) then
      //if (ifTag('office:master-styles', 4)) then

      if (ifTag('table:table', 4)) then
        _ReadTable();

      //считываем фильтр
      if (IfTag('table:database-range', 5)) then _ReadAutoFilter();
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

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    for RowStyleCount := 0 to StyleCount - 1 do
      SetLength(ODFStyles[RowStyleCount].Conditions, 0);
    {$ENDIF}

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
  _ConfigName: string;
  _ConfigType: string;
  _ConfigValue: string;
  _Sheet: TZSheet;

  function _GetSplitModeByNum(const num: integer): TZSplitMode;
  begin
    result := ZSplitNone;
    case (num) of
      1: result := ZSplitSplit;
      2: result := ZSplitFrozen
    end;
  end; //_GetSplitModeByNum

  procedure _GetConfigTypeAndValue();
  begin
    _ConfigName := xml.Attributes.ItemsByName[ZETag_config_name];
    _ConfigType := xml.Attributes.ItemsByName['config:type'];
  end; //_GetConfigTypeAndValue

  function _FindSheetByName(const ASheetName: string; out retSheet: TZSheet): boolean;
  var
    i: integer;

  begin
    result := false;
    for i := 0 to XMLSS.Sheets.Count - 1 do
    if (XMLSS.Sheets[i].Title = ASheetName) then
    begin
      retSheet := XMLSS.Sheets[i];
      result := true;
      break;
    end;
  end; //_FindSheetByName

  procedure _ReadSettingsPage();
  var
    s: string;
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
    s := ZEReplaceEntity(xml.Attributes.ItemsByName[ZETag_config_name]);
    if (s <> '') then
      if (_FindSheetByName(s, _Sheet)) then
        while not ((xml.TagType = 6) and (xml.TagName = 'config:config-item-map-entry')) do
        begin
          if (xml.Eof()) then
            break;

          if (xml.TagName = 'config:config-item') then
          begin
            if (xml.TagType = 4) then
              _GetConfigTypeAndValue()
            else
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
    while not ((xml.TagName = ZETag_config_config_item_map_named) and (xml.TagType = 6)) do
    begin
      if (xml.Eof) then
        break;

      xml.ReadTag();
      if ((xml.TagType = 4) and (xml.TagName = 'config:config-item-map-entry')) then
        _ReadSettingsPage();
    end; //while
  end; //_ReadSettings

  //Проверка итема (не для листа)
  procedure _CheckConfigValue();
  begin
    _ConfigValue := ZEReplaceEntity(xml.TextBeforeTag);
    if (_ConfigName = 'ActiveTable') then
    begin
      if (_FindSheetByName(_ConfigValue, _Sheet)) then
        _Sheet.Selected := true;
    end;
  end; //_CheckOtherParams()

  //<config:config-item>...</config:config-item> вне настроек листа
  procedure _ReadConfigItem();
  begin
    _GetConfigTypeAndValue();

    while not ((xml.TagName = 'config:config-item') and (xml.TagType = 6)) do
    begin
      if (xml.Eof) then
        break;

      xml.ReadTag();
      if ((xml.TagType = 6) and (xml.TagName = 'config:config-item')) then
        _CheckConfigValue();
    end; //while
  end; //_ReadConfigItem

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
      if ((xml.TagName = ZETag_config_config_item_map_named) and (xml.TagType = 4)) then
      begin
        if (xml.Attributes.ItemsByName[ZETag_config_name] = 'Tables') then
          _ReadSettings();
      end else
      if (xml.TagType = 4) and (xml.TagName = 'config:config-item') then
        _ReadConfigItem();
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
  ReadHelper: TZEODFReadHelper;

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
  ReadHelper := TZEODFReadHelper.Create(XMLSS);

  try
    //манифест (META_INF/manifest.xml)

    //стили (styles.xml)
    try
      stream := TFileStream.Create(DirName + 'styles.xml', fmOpenRead or fmShareDenyNone);
    except
      result := 2;
      exit;
    end;
    if (not ReadODFStyles(XMLSS, stream, ReadHelper)) then
      result := result or 2;
    FreeAndNil(stream);

    //содержимое (content.xml)
    try
      stream := TFileStream.Create(DirName + 'content.xml', fmOpenRead or fmShareDenyNone);
    except
      result := 2;
      exit;
    end;
    if (not ReadODFContent(XMLSS, stream, ReadHelper)) then
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
    if (Assigned(ReadHelper)) then
      FreeAndNil(ReadHelper);
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
    ZH := TODFZipHelper.Create();
    ZH.XMLSS := XMLSS;
    u_zip := TUnZipper.Create();
    u_zip.FileName := FileName;
    u_zip.OnCreateStream := @ZH.DoCreateOutZipStream;
    u_zip.OnDoneStream := @ZH.DoDoneOutZipStream;

    lst.Clear();
    lst.Add('styles.xml'); //стили (styles.xml)
    ZH.FileType := 2;
    u_zip.UnZipFiles(lst);
    result := result or ZH.RetCode;

    lst.Clear();
    lst.Add('content.xml'); //содержимое
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
{$ENDIF} //FPC

//Get MediaType by string (from manifest.xml)
//INPUT
//  const MediaType: string - string representation of media type
//RETURN
//      TODSManifestMediaType - media type
function GetODSMediaTypeByStr(const MediaType: string): TODSManifestMediaType;
var
  i: integer;

begin
  Result := ZEODSMediaTypeUnknown;
  for i := 0 to High(const_ODS_manifest_mediatypes_str) do
    if (MediaType = const_ODS_manifest_mediatypes_str[i]) then
    begin
      Result := const_ODS_manifest_mediatypes[i];
      break;
    end;
end;

//Get string representation for manifest.xml MediaType
//INPUT
//  const MediaType: TODSManifestMediaType - media type
//RETURN
//      string
function GetODSStrByMediaType(const MediaType: TODSManifestMediaType): string;
var
  i: integer;

begin
  Result := '';
  for i := 0 to High(const_ODS_manifest_mediatypes) - 1 do
    if (MediaType = const_ODS_manifest_mediatypes[i]) then
    begin
      Result := const_ODS_manifest_mediatypes_str[i];
      break;
    end;
end;

{$IFNDEF FPC}
{$I odszipfuncimpl.inc}
{$ENDIF}

end.
