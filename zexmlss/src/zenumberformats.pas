//****************************************************************
// Routines for translate number formats from ods/xlsx/excel xml
//  to internal (ods like) number format and back etc.
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2016.09.10
//----------------------------------------------------------------
{
 Copyright (C) 2015 Ruslan Neborak

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
//Sorry for my english.
unit zenumberformats;

{$I zexml.inc}
{$I compver.inc}

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

interface

uses
  SysUtils,
  zsspxml           //TZsspXMLReaderH
  ;

const
  //Main number formats
  ZE_NUMFORMAT_IS_UNKNOWN   = 0;
  ZE_NUMFORMAT_IS_NUMBER    = 1;
  ZE_NUMFORMAT_IS_DATETIME  = 2;
  ZE_NUMFORMAT_IS_STRING    = 4;

  //Additional properties for number styles
  ZE_NUMFORMAT_NUM_IS_PERCENTAGE  = 1 shl 10;
  ZE_NUMFORMAT_NUM_IS_SCIENTIFIC  = 1 shl 11;
  ZE_NUMFORMAT_NUM_IS_CURRENCY    = 1 shl 12;
  ZE_NUMFORMAT_NUM_IS_FRACTION    = 1 shl 13;

  ZE_NUMFORMAT_DATE_IS_ONLY_TIME  = 1 shl 14;

  //DateStyles
  ZETag_number_date_style     = 'number:date-style';
  ZETag_number_time_style     = 'number:time-style';

  ZETag_number_day            = 'number:day';
  ZETag_number_text           = 'number:text';
  ZETag_number_style          = 'number:style';
  ZETag_number_month          = 'number:month';
  ZETag_number_year           = 'number:year';
  ZETag_number_hours          = 'number:hours';
  ZETag_number_minutes        = 'number:minutes';
  ZETag_number_seconds        = 'number:seconds';
  ZETag_number_day_of_week    = 'number:day-of-week';
  ZETag_number_textual        = 'number:textual';
  ZETag_number_possessive_form = 'number:possessive-form';

  ZETag_number_am_pm          = 'number:am-pm';
  ZETag_number_quarter        = 'number:quarter';
  ZETag_number_week_of_year   = 'number:week-of-year';
  ZETag_number_era            = 'number:era';

  //NumberStyles:
  //WARNING: number style = currency style = percentage style!
  //TODO:
  //      Is need separate number/currency/percentage number styles?
  ZETag_number_number_style         = 'number:number-style';
  ZETag_number_currency_style       = 'number:currency-style';
  ZETag_number_percentage_style     = 'number:percentage-style';

  //for currency
  ZETag_number_currency_symbol      = 'number:currency-symbol';
  ZETag_number_language             = 'number:language';
  ZETag_number_country              = 'number:country';

  ZETag_number_fraction             = 'number:fraction';
  ZETag_number_scientific_number    = 'number:scientific-number';
  ZETag_number_embedded_text        = 'number:embedded-text';
  ZETag_number_number               = 'number:number';
  ZETag_number_decimal_places       = 'number:decimal-places';
  ZETag_number_decimal_replacement  = 'number:decimal-replacement';
  ZETag_number_display_factor       = 'number:display-factor';
  ZETag_number_grouping             = 'number:grouping';
  ZETag_number_min_integer_digits   = 'number:min-integer-digits';
  ZETag_number_position             = 'number:position';
  ZETag_number_min_exponent_digits  = 'number:min-exponent-digits';

  ZETag_number_min_numerator_digits = 'number:min-numerator-digits';
  ZETag_number_min_denominator_digits = 'number:min-denominator-digits';
  ZETag_number_denominator_value    = 'number:denominator-value';

  ZETag_number_text_style           = 'number:text-style';
  ZETag_number_text_content         = 'number:text-content';

  ZETag_style_text_properties = 'style:text-properties';
  ZETag_style_map             = 'style:map';
  ZETag_fo_color              = 'fo:color';

  ZETag_Attr_StyleName        = 'style:name';
  ZETag_style_condition       = 'style:condition';
  ZETag_style_apply_style_name = 'style:apply-style-name';
  ZETag_long                  = 'long';
  ZETag_short                 = 'short';
  ZETag_style_volatile        = 'style:volatile';

type
  TZODSNumberItemOptions = record
    isColor: boolean;
    ColorStr: string;
    StyleType: byte;
  end;

  TODSEmbeded_text_props = record
    Txt: string;
    NumberPosition: integer;
  end;

  //Date/Time item for processing date number style
  TZDateTimeProcessItem = record
    //Item type:
    //    -1 - error item (ignore)
    //     0 - text
    //     1 - year (Y/YY/YYYY)
    //     2 - month (M/MM/MMM/MMMM/MMMMM)
    //     3 - day (D/DD/DDD/DDDD/NN/NNN/NNNN)
    //     4 - hour (h/hh)
    //     5 - minute (m/mm)
    //     6 - second (s/ss)
    //     7 - week (WW)
    //     8 - quarterly (Q/QQ)
    //     9 - era jap (G/GG/GGG/RR/GGGEE)
    //    10 - number of the year in era (E/EE/R)
    //    11 - AM/PM (a/p AM/PM)
    ItemType: integer;
    //Text value (for ItemType = 0)
    TextValue: string;
    //Length for item
    Len: integer;
    //Additional settings for item
    Settings: integer;
  end;

  //Simple parser for number format
  TNumFormatParser = class
  private
    FStr: string;
    FLen: integer;
    FPos: integer;
    FReadedSymbol: string;
    FReadedSymbolType: integer;
    FIsError: integer;
    FFirstSymbol: char;
  protected
    procedure Clear();
  public
    constructor Create();
    procedure BeginRead(const AStr: string);
    function ReadSymbol(): boolean;
    procedure IncPos(ADelta: integer);
    property FirstSymbol: char read FFirstSymbol;
    property ReadedSymbol: string read FReadedSymbol;
    property ReadedSymbolType: integer read FReadedSymbolType;
    property StrLength: integer read FLen;
    property CurrentPos: integer read FPos;
    property IsError: integer read FIsError;
  end;

  //Parser for ODS datetime format
  TZDateTimeODSFormatParser = class
  private
    FCount: integer;
    FMaxCount: integer;
  protected
    procedure IncCount(ADelta: integer = 1);
    procedure CheckMonthMinute();
  public
    FItems: array of TZDateTimeProcessItem;
    constructor Create();
    destructor Destroy(); override;
    procedure DeleteRepeatedItems();
    function GetValidCount(): integer;
    function TryToParseDateFormat(const AFmtStr: string; const AFmtParser: TNumFormatParser = nil): integer;
    property Count: integer read FCount;
  end;

  //Number format item for write
  TODSNumberFormatMapItem = class
  private
    FCondition: string;
    FisCondition: boolean;
    FColorStr: string;
    FisColor: boolean;
    FNumberFormat: string;
    FConditionsArray: array[0..1] of array[0..1] of string;
    FConditionsCount: integer;
    FEmbededTextCount: integer;
    FEmbededMaxCount: integer;
    FEmbededTextArray: array of TODSEmbeded_text_props;
    FNumberFormatParser: TNumFormatParser;
    FDateTimeODSFormatParser: TZDateTimeODSFormatParser;
  protected
    procedure PrepareCommonStyleAttributes(const xml: TZsspXMLWriterH; const AStyleName: string; isVolatile: boolean = false);
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Clear();
    function TryToParse(const FNStr: string): boolean;
    //Add condition for this number format (max 2)
    //INPUT
    //     const ACondition: string
    //     const AStyleName: string
    function AddCondition(const ACondition, AStyleName: string): boolean;

    //Write number style item (<number:number-style> </number:number-style>)
    //INPUT
    //     const xml: TZsspXMLWriterH - xml
    //     const AStyleName: string   - style name
    //     const NumProperties: integer - additional number properties (currency/percentage etc)
    //           isVolatile: boolean  - is volatile?
    procedure WriteNumberStyle(const xml: TZsspXMLWriterH; const AStyleName: string; const NumProperties: integer; isVolatile: boolean = false);

    //Write number text style item (<number:text-style> </number:text-style>)
    //INPUT
    //     const xml: TZsspXMLWriterH - xml
    //     const AStyleName: string   - style name
    //           isVolatile: boolean  - is volatile? (for now - ignore)
    procedure WriteTextStyle(const xml: TZsspXMLWriterH; const AStyleName: string; isVolatile: boolean = false);

    //Write datetime style item (<number:date-style> </number:date-style>)
    //INPUT
    //     const xml: TZsspXMLWriterH - xml
    //     const AStyleName: string   - style name
    //           isVolatile: boolean  - is volatile? (for now - ignore)
    //RETURN
    //      integer - additional properties for datetime style
    function WriteDateTimeStyle(const xml: TZsspXMLWriterH;
                                 const AStyleName: string;
                                 isVolatile: boolean = false): integer;

    property Condition: string read FCondition write FCondition;
    property isCondition: boolean read FisCondition write FisCondition;
    property ColorStr: string read FColorStr write FColorStr;
    property isColor: boolean read FisColor write FisColor;
    property NumberFormat: string read FNumberFormat write FNumberFormat;
  end;

  //Reads and stores number formats for ODS
  TZEODSNumberFormatReader = class
  private
    FItems: array of array [0..1] of string;  //index 0 - format num
                                              //index 1 - format
    FItemsOptions: array of TZODSNumberItemOptions;
    FCount: integer;
    FCountMax: integer;

    FEmbededTextCount: integer;
    FEmbededMaxCount: integer;
    FEmbededTextArray: array of TODSEmbeded_text_props;
    procedure AddEmbededText(const AText: string; ANumberPosition: integer);
  protected
    procedure AddItem();
    procedure ReadNumberFormatCommon(const xml: TZsspXMLReaderH;
                                     const NumberFormatTag: string;
                                     sub_number_type: integer);
    function BeginReadFormat(const xml: TZsspXMLReaderH; out retStartString: string; const NumFormat: integer): integer;
  public
    constructor Create();
    destructor Destroy(); override;
    procedure ReadKnownNumberFormat(const xml: TZsspXMLReaderH);
    procedure ReadDateFormat(const xml: TZsspXMLReaderH; const ATagName: string);
    procedure ReadNumberFormat(const xml: TZsspXMLReaderH);
    procedure ReadCurrencyFormat(const xml: TZsspXMLReaderH);
    procedure ReadPercentageFormat(const xml: TZsspXMLReaderH);
    procedure ReadStringFormat(const xml: TZsspXMLReaderH);
    function TryGetFormatStrByNum(const DataStyleName: string; out retFormatStr: string): boolean;
    property Count: integer read FCount;
  end;

  TZEODSNumberFormatWriterItem = record
    StyleIndex: integer;
    NumberFormatName: string;
    NumberFormat: string;
  end;

  //Writes to ODS number formats and stores number formats names
  TZEODSNumberFormatWriter = class
  private
    FItems: array of TZEODSNumberFormatWriterItem;
    FCount: integer;
    FCountMax: integer;
    FCurrentNFIndex: integer;

    FNFItems: array of TODSNumberFormatMapItem;
    FNFItemsCount: integer;
                                               //Additional properties for number formats (currency, percentage etc)
    FNumberAdditionalProps: array of integer;

  protected
    function TryAddNFItem(const NFStr: string): boolean;
    function SeparateNFItems(const NFStr: string): integer;
  public
    constructor Create(const AMaxCount: integer);
    destructor Destroy(); override;
    function TryGetNumberFormatName(StyleID: integer; out NumberFormatName: string): boolean;
    //Try to find additional properties for number format
    //INPUT
    //      StyleID: integer          - style ID
    //  out NumberFormatProp: integer - finded number additional properties
    //RETURN
    //      boolean - true - additional properties is found
    function TryGetNumberFormatAddProp(StyleID: integer; out NumberFormatProp: integer): boolean;
    //Try to write number format to xml
    //INPUT
    //  const xml: TZsspXMLWriterH  - xml
    //      StyleID: integer        - Style ID
    //      ANumberFormat: string   - number format
    //RETURN
    //      boolean - true - NumberFormat was written ok
    function TryWriteNumberFormat(const xml: TZsspXMLWriterH; StyleID: integer; ANumberFormat: string): boolean;
    property Count: integer read FCount;
  end;

// Try to get xlsx number format type by string (very simplistic)
//INPUT
//  const FormatStr: string - format ("YYYY.MM.DD" etc)
//RETURN
//      integer - 0 - unknown
//                1 and 1 = 1 - number
//                2 and 2 = 2 - datetime
//                4 and 4 = 4 - string
function GetXlsxNumberFormatType(const FormatStr: string): integer;

// Try to get native number format type by string (very simplistic)
//INPUT
//  const FormatStr: string - format ("YYYY.MM.DD" etc)
//RETURN
//      integer - 0 - unknown
//                1 and 1 = 1 - number
//                2 and 2 = 2 - datetime
//                4 and 4 = 4 - string
function GetNativeNumberFormatType(const FormatStr: string): integer;

//Convert native number format to xlsx
//INPUT
//     const FormatNative: string                   - number format
//     const AFmtParser: TNumFormatParser           - format parser (not NIL!)
//     const ADateParser: TZDateTimeODSFormatParser - date parser (not NIL!)
//RETURN
//      string - number format fox xlsx and excel 2003 xml
function ConvertFormatNativeToXlsx(const FormatNative: string;
                                   const AFmtParser: TNumFormatParser;
                                   const ADateParser: TZDateTimeODSFormatParser): string; overload;

//Convert native number format to xlsx
//INPUT
//     const FormatNative: string                   - number format
//RETURN
//      string - number format fox xlsx and excel 2003 xml
function ConvertFormatNativeToXlsx(const FormatNative: string): string; overload;
function ConvertFormatXlsxToNative(const FormatXlsx: string): string;
function TryXlsxTimeToDateTime(const XlsxDateTime: string; out retDateTime: TDateTime; is1904: boolean = false): boolean;

implementation

uses
  zesavecommon,
  StrUtils            {IfThen}
  ;

const
  ZE_NUMBER_FORMAT_DECIMAL_SEPARATOR    = '.';

  ZE_MAX_NF_ITEMS_COUNT                 = 3;

  ZE_MAP_CONDITIONAL_COLORS_COUNT       = 8;

  ZE_MAP_CONDITIONAL_COLORS: array [0 .. ZE_MAP_CONDITIONAL_COLORS_COUNT - 1] of
                         array [0..1] of string =
                         (('#000000', 'BLACK'),
                          ('#FFFFFF', 'WHITE'),
                          ('#FF0000', 'RED'),
                          ('#00FF00', 'GREEN'),
                          ('#0000FF', 'BLUE'),
                          ('#FF00FF', 'MAGENTA'),
                          ('#00FFFF', 'CYAN'),
                          ('#FFFF00', 'YELLOW')
                         );

  ZE_VALID_CONDITIONS_STR: array [0..4] of string =
                        (
                          '>',
                          '<',
                          '>=',
                          '<=',
                          '='
                        );

  ZE_VALID_NAMED_FORMATS_COUNT = 15;

  ZE_VALID_NAMED_FORMATS: array[0..ZE_VALID_NAMED_FORMATS_COUNT - 1] of
                        array [0..1] of string =
                        (('GENERAL',      ''),
                         ('FIXED',        '0.00'),
                         ('CURRENCY',     '0.00'),
                         ('STANDARD',     ''),
                         ('PERCENT',      '0.00%'),
                         ('SCIENTIFIC',   '0,00E+00'),
                         ('GENERAL DATE', 'DD.MM.YYYY'),
                         ('DATE',         'DD.MM.YYYY'),
                         ('LONG DATE',    'DD.MM.YYYY'),
                         ('MEDIUM DATE',  'DD-MMM-YY'),
                         ('SHORT DATE',   'DD.MM.YY'),
                         ('LONG TIME',    'HH:MM:SS'),
                         ('MEDIUM TIME',  'HH:MM AM/PM'),
                         ('SHORT TIME',   'HH:MM'),
                         ('TIME',         'HH:MM')
                        );

  ZE_DATETIME_ITEM_ERROR    = -1;
  ZE_DATETIME_ITEM_TEXT     = 0;
  ZE_DATETIME_ITEM_YEAR     = 1;
  ZE_DATETIME_ITEM_MONTH    = 2;
  ZE_DATETIME_ITEM_DAY      = 3;
  ZE_DATETIME_ITEM_HOUR     = 4;
  ZE_DATETIME_ITEM_MINUTE   = 5;
  ZE_DATETIME_ITEM_SECOND   = 6;
  ZE_DATETIME_ITEM_WEEK     = 7;
  ZE_DATETIME_ITEM_QUARTER  = 8;
  ZE_DATETIME_ITEM_ERA_JAP  = 9;
  ZE_DATETIME_ITEM_ERA_YEAR = 10;
  ZE_DATETIME_ITEM_AMPM     = 11;

  ZE_DATETIME_AMPM_SHORT_LOW  = 0;
  ZE_DATETIME_AMPM_SHORT_UP   = 1;
  ZE_DATETIME_AMPM_LONG_LOW   = 2;
  ZE_DATETIME_AMPM_LONG_UP    = 3;

{

LO:

M             Month as 3.
MM            Month as 03.
MMM           Month as Jan-Dec
MMMM          Month as January-December 	MMMM
MMMMM         First letter of Name of Month 	MMMMM
D             Day as 2 	D
DD            Day as 02 	DD
NN or DDD     Day as Sun-Sat
NNN or DDDD   Day as Sunday to Saturday
NNNN          Day followed by comma, as in "Sunday," 	NNNN
YY            Year as 00-99 	YY
YYYY          Year as 1900-2078 	YYYY
WW            Calendar week
Q             Quarterly as Q1 to Q4 	Q
QQ            Quarterly as 1st quarter to 4th quarter 	QQ
G             Era on the Japanese Gengou calendar, single character (possible values are: M, T, S, H)
GG            Era, abbreviation
GGG           Era, full name
E             Number of the year within an era, without a leading zero for single-digit years
EE or R       Number of the year within an era, with a leading zero for single-digit years
RR or GGGEE   Era, full name and year

h             Hours as 0-23 	h
hh            Hours as 00-23
m             Minutes as 0-59
mm            Minutes as 00-59
s             Seconds as 0-59
ss            Seconds as 00-59

[~buddhist] 	  Thai Buddhist Calendar
[~gengou] 	    Japanese Gengou Calendar
[~gregorian] 	  Gregorian Calendar
[~hanja]        Korean Calendar
[~hanja_yoil] 	Korean Calendar
[~hijri] 	      Arabic Islamic Calendar, currently supported for the following locales: ar_EG, ar_LB, ar_SA, and ar_TN
[~jewish] 	    Jewish Calendar
[~ROC] 	        Republic Of China Calendar

  m$
m       Displays the month as a number without a leading zero.
mm      Displays the month as a number with a leading zero when appropriate.
mmm     Displays the month as an abbreviation (Jan to Dec).
mmmm    Displays the month as a full name (January to December).
mmmmm   Displays the month as a single letter (J to D).
d       Displays the day as a number without a leading zero.
dd      Displays the day as a number with a leading zero when appropriate.
ddd     Displays the day as an abbreviation (Sun to Sat).
dddd    Displays the day as a full name (Sunday to Saturday).
yy      Displays the year as a two-digit number.
yyyy    Displays the year as a four-digit number.



h       Displays the hour as a number without a leading zero.
[h]     Displays elapsed time in hours. If you are working with a formula that returns a time in which the number of hours exceeds 24, use a number format that resembles [h]:mm:ss.
hh      Displays the hour as a number with a leading zero when appropriate. If the format contains AM or PM, the hour is based on the 12-hour clock. Otherwise, the hour is based on the 24-hour clock.
m       Displays the minute as a number without a leading zero.
[m]     Displays elapsed time in minutes. If you are working with a formula that returns a time in which the number of minutes exceeds 60, use a number format that resembles [mm]:ss.
mm      Displays the minute as a number with a leading zero when appropriate.
        Note   The m or mm code must appear immediately after the h or hh code or immediately before the ss code; otherwise, Excel displays the month instead of minutes.
s       Displays the second as a number without a leading zero.
[s]     Displays elapsed time in seconds. If you are working with a formula that returns a time in which the number of seconds exceeds 60, use a number format that resembles [ss].
ss      Displays the second as a number with a leading zero when appropriate. If you want to display fractions of a second, use a number format that resembles h:mm:ss.00.

AM/PM, am/pm, A/P, a/p  Displays the hour using a 12-hour clock. Excel displays AM, am, A, or a for times from midnight unt

}

//Return true if in string AStr after position AStartPos have one of symbols SymbolsArr
//This function checks quotas and brackets. If desired symbol between the quotas - function return FALSE.
//INPUT
//      AStartPos: integer              - start position
//      ALen: integer                   - string length
//  out retPos: integer                 - returned position of symbol
//  const AStr: string                  - string
//  const SymbolsArr: array of string   - searching symbols
//RETURN
//      boolean - true - one of symbols was found in string after AStartPos (and not between quotas)
function IsHaveSymbolsAfterPosQuotas(AStartPos: integer; ALen: integer; out retPos: integer; const AStr: string; const SymbolsArr: array of string): boolean; overload;
var
  i, j: integer;
  _IsQuote: boolean;
  _IsBracket: boolean;
  ch: char;
  _max, _min: integer;

begin
  Result := false;
  _IsQuote := false;
  _IsBracket := false;
  retPos := -1;
  _min := Low(SymbolsArr);
  _max := High(SymbolsArr);
  i := AStartPos + 1;
  while (i <= ALen) do
  begin
    ch := AStr[i];

    if (not _IsQuote) then
    begin
      case (ch) of
        '[': _IsBracket := true;
        ']': _IsBracket := false;
        '\':
            begin
              inc(i, 2);
              if (i > ALen) then
                break;
              ch := AStr[i];
            end;
      end;
    end;

    if ((not _IsBracket) and (ch = '"')) then
      _IsQuote := not _IsQuote;

    if ((not _IsQuote) and (not _IsBracket)) then
      for j := _min to _max do
        if (ch = SymbolsArr[j]) then
        begin
          retPos := i;
          Result := true;
          exit;
        end;
    inc(i);
  end; //while i
end; //IsHaveSymbolsAfterPosQuotas

//Return true if in string AStr after position AStartPos have one of symbols SymbolsArr
//This function checks quotas and brackets. If desired symbol between the quotas - function return FALSE.
//INPUT
//      AStartPos: integer              - start position
//      ALen: integer                   - string length
//  const AStr: string                  - string
//  const SymbolsArr: array of string   - searching symbols
//RETURN
//      boolean - true - one of symbols was found in string after AStartPos (and not between quotas)
function IsHaveSymbolsAfterPosQuotas(AStartPos: integer; ALen: integer; const AStr: string; const SymbolsArr: array of string): boolean; overload;
var
  _retPos: integer;

begin
  Result := IsHaveSymbolsAfterPosQuotas(AStartPos, ALen, _retPos, AStr, SymbolsArr);
end; //IsHaveSymbolsAfterPosQuotas

// Try to get xlsx number format type by string (very simplistic)
//INPUT
//  const FormatStr: string - format ("YYYY.MM.DD" etc)
//RETURN
//      integer - 0 - unknown
//                1 and 1 = 1 - number
//                2 and 2 = 2 - datetime
//                4 and 4 = 4 - string
function GetXlsxNumberFormatType(const FormatStr: string): integer;
var
  i, l: integer;
  ch: char;
  _isQuote: boolean;
  _isBracket: boolean;
  _isFraction: boolean;

begin
  Result := ZE_NUMFORMAT_IS_UNKNOWN;
  _isFraction := false;

  //General is not for dates
  if ((UpperCase(FormatStr) = 'GENERAL') or (FormatStr = '')) then
  begin
    Result := ZE_NUMFORMAT_IS_NUMBER;
    exit;
  end;

  _isQuote := false;
  _isBracket := false;

  l := length(FormatStr);
  for i := 1 to l do
  begin
    ch := FormatStr[i];

    if ((ch = '"') and (not _isBracket)) then
      _isQuote := not _isQuote;

    if ((ch = '[') and (not _isQuote)) then
      _isBracket := true;

    if ((ch = ']') and (not _isQuote) and _isBracket) then
       _isBracket := false;

    //[$RUB] / [$UAH] etc
    //TODO: need check for valid country code
    if (_isBracket and (ch = '$')) then
       Result := Result or ZE_NUMFORMAT_NUM_IS_CURRENCY;

    if ((not _isQuote) and (not _isBracket)) then
      case (ch) of
        '0', '#', 'E', 'e', '%', '?':
          begin
            Result := ZE_NUMFORMAT_IS_NUMBER;

            if (IsHaveSymbolsAfterPosQuotas(i - 1, l, FormatStr, ['e', 'E'])) then
              Result := Result or ZE_NUMFORMAT_NUM_IS_SCIENTIFIC
            else
            if (IsHaveSymbolsAfterPosQuotas(i - 1, l, FormatStr, ['%'])) then
              Result := Result or ZE_NUMFORMAT_NUM_IS_PERCENTAGE
            else
            if (IsHaveSymbolsAfterPosQuotas(i - 1, l, FormatStr, ['/']) or _isFraction) then
              Result := Result or ZE_NUMFORMAT_NUM_IS_FRACTION;

            exit;
          end;
        '@':
          begin
            Result := ZE_NUMFORMAT_IS_STRING;
            exit;
          end;
        '/': _isFraction := true;
        'H', 'h', 'S', 's', 'm', 'M', 'd', 'D', 'Y', 'y', ':':
          begin
            Result := ZE_NUMFORMAT_IS_DATETIME;
            exit;
          end;
      end;
  end; //for i
end; //GetXlsxNumberFormatType

// Try to get native number format type by string (very simplistic)
//INPUT
//  const FormatStr: string - format ("YYYY.MM.DD" etc)
//RETURN
//      integer - 0 - unknown
//                1 and 1 = 1 - number
//                2 and 2 = 2 - datetime
//                4 and 4 = 4 - string
function GetNativeNumberFormatType(const FormatStr: string): integer;
var
  i, l: integer;
  ch, _prev: char;
  _isQuote: boolean;
  _isBracket: boolean;
  _isSemicolon: boolean;
  t: integer;

  function _CheckSemicolon(): boolean;
  begin
    if (not _isSemicolon) then
    begin
      t := i - 1;
      _isSemicolon := IsHaveSymbolsAfterPosQuotas(t, l, i, FormatStr, [';']);
    end;

    Result := not _isSemicolon;
  end;

begin
  Result := ZE_NUMFORMAT_IS_UNKNOWN;

  _isSemicolon := false;
  _isBracket := false;
  _isQuote := false;
  _prev := #0;

  l := length(FormatStr);
  i := 1;
  while (i <= l) do
  begin
    ch := FormatStr[i];

    if ((ch = '"') and (not _isBracket)) then
      _isQuote := not _isQuote;

    if ((ch = '[') and (not _isQuote)) then
      _isBracket := true;

    if ((ch = ']') and (not _isQuote) and _isBracket) then
       _isBracket := false;

    //[$RUB] / [$UAH] etc
    //TODO: need check for valid country code
    if (_isBracket and (ch = '$')) then
       Result := Result or ZE_NUMFORMAT_NUM_IS_CURRENCY;

    if ((not _isQuote) and (not _isBracket)) then
      case (ch) of
        '0', '#', '?':
          begin
            if (_isSemicolon) then
              Result := Result or ZE_NUMFORMAT_IS_NUMBER
            else
              Result := ZE_NUMFORMAT_IS_NUMBER;

            if (IsHaveSymbolsAfterPosQuotas(i - 1, l, t, FormatStr, ['e', 'E'])) then
            begin
              Result := Result or ZE_NUMFORMAT_NUM_IS_SCIENTIFIC;
              i := t;
            end
            else
            if (IsHaveSymbolsAfterPosQuotas(i - 1, l, t, FormatStr, ['%'])) then
            begin
              Result := Result or ZE_NUMFORMAT_NUM_IS_PERCENTAGE;
              i := t;
            end
            else
            if (IsHaveSymbolsAfterPosQuotas(i - 1, l, t, FormatStr, ['/'])) then
            begin
              Result := Result or ZE_NUMFORMAT_NUM_IS_FRACTION;
              i := t;
            end;

            if (_CheckSemicolon()) then
              exit;
          end;
        '%':
          begin
            if (_isSemicolon) then
              Result := Result or ZE_NUMFORMAT_IS_NUMBER or ZE_NUMFORMAT_NUM_IS_PERCENTAGE
            else
              Result := ZE_NUMFORMAT_IS_NUMBER or ZE_NUMFORMAT_NUM_IS_PERCENTAGE;

            if (_CheckSemicolon()) then
              exit;
          end;
        'E', 'e':
          begin
            if ((_prev = '0') or (_prev = '#')) then
              Result := ZE_NUMFORMAT_IS_NUMBER or ZE_NUMFORMAT_NUM_IS_SCIENTIFIC
            else
              Result := ZE_NUMFORMAT_IS_DATETIME;
            exit;
          end;
        '@':
          begin
            Result := ZE_NUMFORMAT_IS_STRING;
            exit;
          end;
        'H', 'h', 'S', 's', 'm', 'M', 'd', 'D', 'Y', 'y', ':', 'G', 'Q', 'R', 'W', 'N':
          begin
            Result := ZE_NUMFORMAT_IS_DATETIME;
            exit;
          end;
        ';': _isSemicolon := True;
      end;
    _prev := ch;
    inc(i);
  end; //while i
end; //GetNativeNumberFormatType

//Convert native number format to xlsx
//INPUT
//     const FormatNative: string                   - number format
//     const AFmtParser: TNumFormatParser           - format parser (not NIL!)
//     const ADateParser: TZDateTimeODSFormatParser - date parser (not NIL!)
//RETURN
//      string - number format fox xlsx and excel 2003 xml
function ConvertFormatNativeToXlsx(const FormatNative: string;
                                   const AFmtParser: TNumFormatParser;
                                   const ADateParser: TZDateTimeODSFormatParser): string; overload;
var
  _FmtParser: TNumFormatParser;
  _DateParser: TZDateTimeODSFormatParser;
  _delP: boolean;
  _delD: boolean;
  _fmt: integer;
  _len: integer;

  function _AddText(const AStr: string): string;
  begin
    Result := '';
    _len := Length(AStr);
    if (_len = 1) then
      case (AStr[1]) of
        ' ', '.', ',', ':', '-', '+', '/':
          Result := AStr;
        else
          Result := '\' + AStr;
      end //case
    else
      Result := '"' + AStr + '"';
  end; //_AddText

  function _RepeatSymbol(const ASymbol: string; ALen: integer; AMin, AMax: integer): string;
  var
    i: integer;

  begin
    Result := '';
    if (ALen < AMin) then
      ALen := AMin;
    if (ALen > AMax) then
      ALen := AMax;
    for i := 1 to ALen do
      Result := Result + ASymbol;
  end; //_RepeatSymbol

  function _AddMonth(var item: TZDateTimeProcessItem): string;
  begin
    Result := _RepeatSymbol('M', item.Len, 1, 5);
  end; //_AddMonth

  function _AddYear(var item: TZDateTimeProcessItem): string;
  begin
    if (item.Len <= 2) then
      Result := 'YY'
    else
      Result := 'YYYY';
  end; //_AddYear

  function _AddDay(var item: TZDateTimeProcessItem): string;
  begin
    Result := _RepeatSymbol('D', item.Len, 1, 4);
  end; //_AddDay

  function _AddHour(var item: TZDateTimeProcessItem): string;
  begin
    Result := _RepeatSymbol('H', item.Len, 1, 2);
  end; //_AddHour

  function _AddMinute(var item: TZDateTimeProcessItem): string;
  begin
    Result := _RepeatSymbol('M', item.Len, 1, 2);
  end; //_AddMinute

  function _AddSecond(var item: TZDateTimeProcessItem): string;
  begin
    Result := _RepeatSymbol('S', item.Len, 1, 2);

    if (item.Settings > 0) then
      Result := Result + '.' + _RepeatSymbol('0', item.Settings, 1, item.Settings);
  end; //_AddSecond

  function _AddAMPM(var item: TZDateTimeProcessItem): string;
  begin
    case (item.Settings) of
      ZE_DATETIME_AMPM_SHORT_LOW:
        Result := 'a/p';
      ZE_DATETIME_AMPM_SHORT_UP:
        Result := 'A/P';
      ZE_DATETIME_AMPM_LONG_LOW:
        Result := 'am/pm';
      else
        Result := 'AM/PM';
    end;
  end; //_AddAMPM

  function _GetXlsxDateFormat(): string;
  var
    i: integer;

  begin
    Result := '';
    for i := 0 to _DateParser.Count - 1 do
      case (_DateParser.FItems[i].ItemType) of
        ZE_DATETIME_ITEM_TEXT:
                              Result := Result + _AddText(_DateParser.FItems[i].TextValue);

        ZE_DATETIME_ITEM_YEAR:
                              Result := Result + _AddYear(_DateParser.FItems[i]);

        ZE_DATETIME_ITEM_MONTH:
                              Result := Result + _AddMonth(_DateParser.FItems[i]);

        ZE_DATETIME_ITEM_DAY:
                              Result := Result + _AddDay(_DateParser.FItems[i]);

        ZE_DATETIME_ITEM_HOUR:
                              Result := Result + _AddHour(_DateParser.FItems[i]);

        ZE_DATETIME_ITEM_MINUTE:
                              Result := Result + _AddMinute(_DateParser.FItems[i]);

        ZE_DATETIME_ITEM_SECOND:
                              Result := Result + _AddSecond(_DateParser.FItems[i]);

        ZE_DATETIME_ITEM_WEEK:;     //??
        ZE_DATETIME_ITEM_QUARTER:;  //??
        ZE_DATETIME_ITEM_ERA_JAP:;  //??
        ZE_DATETIME_ITEM_ERA_YEAR:; //??
        ZE_DATETIME_ITEM_AMPM:
                              Result := Result + _AddAMPM(_DateParser.FItems[i]);
      end; //case

  end; //_GetXlsxDateFormat

begin
  _fmt := GetNativeNumberFormatType(FormatNative);

  //For now difference only for datetime
  if (_fmt and ZE_NUMFORMAT_IS_DATETIME = ZE_NUMFORMAT_IS_DATETIME) then
  begin
    Result := '';
    _FmtParser := AFmtParser;
    _DateParser := ADateParser;
    _delP := AFmtParser = Nil;
    _delD := ADateParser = Nil;
    try
      if (_delD) then
        _DateParser := TZDateTimeODSFormatParser.Create();
      if (_delP) then
        _FmtParser := TNumFormatParser.Create();

      if (_DateParser.TryToParseDateFormat(FormatNative, _FmtParser) > 0) then
        Result := _GetXlsxDateFormat();

    finally
      if (_delP) then
        FreeAndNil(_FmtParser);
      if (_delD) then
        FreeAndNil(_DateParser);
    end;
  end
  else
    Result := FormatNative;
end; //ConvertFormatNativeToXlsx

function ConvertFormatNativeToXlsx(const FormatNative: string): string; overload;
begin
  Result := ConvertFormatNativeToXlsx(FormatNative, nil, nil);
end; //ConvertFormatNativeToXlsx

function ConvertFormatXlsxToNative(const FormatXlsx: string): string;
var
  i, l: integer;
  _isQuote: boolean;
  _isBracket: boolean;
  s: string;
  ch: char;
  _semicolonCount: integer;
  _prevCh: char;
  _strDateList: string;
  z: string;
  b: boolean;
  _isSlash: boolean;
  t: integer;

  procedure _AddToResult(const strItem: string; charDate: char);
  begin
    Result := Result + strItem;
    _strDateList := _strDateList + charDate;
    s := '';
  end;

  procedure _CheckStringItem(currCh: char);
  begin
    z := UpperCase(s);

    if ((z = 'YY') or (z= 'YYYY')) then
      _AddToResult(z, 'Y')
    else
    if ((z = 'D') or (z = 'DD')) then
      _AddToResult(z, 'D')
    else
    if (z = 'DDD') then
      _AddToResult('NN', 'D')
    else
    if (z = 'DDDD') then
      _AddToResult('NNN', 'D')
    else
    if (z = 'H') then
      _AddToResult('h', 'H')
    else
    if (z = 'HH') then
      _AddToResult('hh', 'H')
    else
    if (z = 'S') then
      _AddToResult('s', 'S')
    else
    if (z = 'SS') then
      _AddToResult('ss', 'S')
    else
    if ((z = 'MMM') or (z = 'MMMM') or (z = 'MMMMM')) then
      _AddToResult(z, 'M')
    else
    //Minute or Month?
    //  If M/MM between 'H/S' - minutes
    if (z = 'M') or (z = 'MM') then
    begin
      //Is it minute?
      b := (_prevCh = ':') or (_prevCh = 'H') or
           (currCh = 'S') or (currCh = ':');
      if (not b) then
      begin
        t := Length(_strDateList);
        //if some spaces (or some other symbols) between date symbols
        if (t > 0) then
        begin
          if (_strDateList[t] = 'H') or (_strDateList[t] = 'S') then
            b := true;
          if (not b) then
            //If previous date symbol was "month" then for now - "minute"
            b := pos('M', _strDateList) <> 0;
        end;
      end;

      //If previous date symbal was "minute" then for now - "month"
      if (b) then
        b := pos('N', _strDateList) = 0;

      //minutes
      if (b) then
      begin
        if (Z = 'M') then
          _AddToResult('m', 'N')
        else
          _AddToResult('mm', 'N')
      end
      else
        _AddToResult(z, 'M'); //months
    end
    else
      Result := Result + s;

    _prevCh := currCh;
    s := '';
  end; //_CheckStringItem

  procedure _ProcessOpenBracket();
  begin
    if (_isQuote) then
      s := s + ch
    else
      if (not _isBracket) then
      begin
        _CheckStringItem(ch);
        _isBracket := true;
      end;
  end; //_ProcessOpenBracket

  //[some data]
  procedure _ProcessCloseBracket();
  var
    z: string;

  begin
    _isBracket := not _isBracket;

    z := UpperCase(s);
    if (z = 'COLOR1') then
      s := 'BLACK'
    else
    if (z = 'COLOR2') then
      s := 'WHITE'
    else
    if (z = 'COLOR3') then
      s := 'RED';
    //TODO: need add all possible colorXX (1..64??)

    Result := Result + '[' + s + ']';
    _prevCh := ch;
    s := '';
  end; //_ProcessCloseBracket

  procedure _ProcessQuote(addCloseQuote: boolean = true);
  begin
    if (addCloseQuote) then
      _isQuote := not _isQuote;

    if (_isQuote) then
    begin
      if (addCloseQuote) then
        Result := Result + '"';

      Result := Result + s;

      if (addCloseQuote) then
        Result := Result + '"';
    end;

    s := '';
    _prevCh := ch;
  end; //_ProcessQuote

  procedure _ProcessSemicolon();
  begin
    inc(_semicolonCount);
    _CheckStringItem(ch);
  end; //_ProcessSemicolon

  procedure _ProcessDateTimeSymbol(DTSymbol: char);
  begin
    if (_prevCh = #0) then
      _prevCh := DTSymbol;

    if (_prevCh <> DTSymbol) then
      _CheckStringItem(DTSymbol);

    s := s + ch;
  end;

begin
  Result := '';
  _isQuote := false;
  _isBracket := false;
  s := '';
  _semicolonCount := 0;
  _prevCh := #0;

  _strDateList := '';

  l := length(FormatXlsx);

  _isSlash := false;

  for i := 1 to l do
  begin
    ch := FormatXlsx[i];

    if (_isSlash) then
    begin
      Result := Result + ch;
      _isSlash := false;
    end
    else
    if (((_isQuote and (ch <> '"')) or (_isBracket and (ch <> ']')))) then
      s := s + ch
    else
      case (ch) of
        '[': _ProcessOpenBracket();
        ']': _ProcessCloseBracket();
        '"': _ProcessQuote();
        ';':
          begin
            //only 3 sections maximum available!
            _ProcessSemicolon();
            if (_semicolonCount >= 3) then
              break;
            Result := Result + ch;
          end;
        'y', 'Y': _ProcessDateTimeSymbol('Y');
        'm', 'M': _ProcessDateTimeSymbol('M');
        'd', 'D': _ProcessDateTimeSymbol('D');
        's', 'S': _ProcessDateTimeSymbol('S');
        'h', 'H': _ProcessDateTimeSymbol('H');
        '\':
          begin
            _CheckStringItem(ch);
            Result := Result + ch;
            _isSlash := true;
          end;
        else
        begin
          _CheckStringItem(ch);
          s := s + ch;
        end;
      end; //case ch
  end; //for i

  _CheckStringItem(#0);
end; //TryConvertXlsxToNative

//Try to convert xlsx datetime as number to DateTime
//INPUT
//  const XlsxDateTime: string  - datetime string from xlsx cell value
//  out retDateTime: TDateTime  - output datetime (no sense if function returns false!)
//      is1904: boolean         - if true than calc dates from 1904 and from 1900 otherwise
//RETURN
//      boolean - true - ok
function TryXlsxTimeToDateTime(const XlsxDateTime: string; out retDateTime: TDateTime; is1904: boolean = false): boolean;
var
  t: Double;
  s1, s2: string;
  i: integer;
  b: boolean;
  ch: char;

begin
  b := false;
  Result := false;
  s1 := '';
  s2 := '';

  for i := 1 to length(XlsxDateTime) do
  begin
    ch := XlsxDateTime[i];
    if ((ch = '.') or (ch = ',')) then
    begin
      if (b) then
        exit;
      b := true;
    end
    else
      if (b) then
        s2 := s2 + ch
      else
        s1 := s1 + ch;
  end;

  if (s1 = '') then
    s1 := '0';

  if (TryStrToInt(s1, i)) then
  begin
    retDateTime := i;
    if (is1904) then
      retDateTime := IncMonth(retDateTime, 12 * 4);

    if (s2 <> '') then
    {$IFDEF Z_USE_FORMAT_SETTINGS}
      if (TryStrToFloat('0' + FormatSettings.DecimalSeparator + s2, t)) then
    {$ELSE}
      if (TryStrToFloat('0' + DecimalSeparator + s2, t)) then
    {$ENDIF}
        retDateTime := retDateTime + t;
    Result := true;
  end;
end; //TryXlsxTimeToDateTime

function TryGetNumFormatByName(ANamedFormat: string; out retNumFormat: string): boolean;
var
  i: integer;

begin
  Result := False;
  ANamedFormat := UpperCase(ANamedFormat);
  for i := 0 to ZE_VALID_NAMED_FORMATS_COUNT - 1 do
    if (ZE_VALID_NAMED_FORMATS[i][0] = ANamedFormat) then
    begin
      Result := True;
      retNumFormat := ZE_VALID_NAMED_FORMATS[i][1];
      break;
    end;
end;

function TryGetMapColorName(AColor: string; out retColorName: string): boolean;
var
  i: integer;

begin
  Result := false;
  for i := 0 to ZE_MAP_CONDITIONAL_COLORS_COUNT - 1 do
    if (ZE_MAP_CONDITIONAL_COLORS[i][0] = AColor) then
    begin
      Result := true;
      retColorName := ZE_MAP_CONDITIONAL_COLORS[i][1];
      break;
    end;
end; //TryGetMapColor

function TryGetMapColorColor(AColorName: string; out retColor: string): boolean;
var
  i: integer;

begin
  Result := false;
  for i := 0 to ZE_MAP_CONDITIONAL_COLORS_COUNT - 1 do
    if (ZE_MAP_CONDITIONAL_COLORS[i][1] = AColorName) then
    begin
      Result := true;
      retColor := ZE_MAP_CONDITIONAL_COLORS[i][0];
      break;
    end;
end; //TryGetMapColorColor

function TryGetMapCondition(AConditionStr: string; out retODSCondution: string): boolean;
var
  i: integer;
  s: string;
  a: array [0..3] of string;
  kol: integer;
  ch: char;
  _isNumber: boolean;
  _isCond: boolean;

  procedure _AddItem();
  begin
    if (kol < 3) then
    begin
      a[kol] := s;
      s := '';
      inc(kol);
    end;
  end; //_AddItem

  procedure _ProcessSymbol(var isPrevTypeSymbol: boolean; var newTypeSymbol: boolean);
  begin
    if (isPrevTypeSymbol and (s <> '')) then
      _AddItem();

    s := s + ch;

    isPrevTypeSymbol := false;
    newTypeSymbol := true;
  end; //_ProcessSymbol

  function _CheckCondition(): boolean;
  var
    i: integer;
    d: double;

  begin
    Result := false;
    for i := Low(ZE_VALID_CONDITIONS_STR) to High(ZE_VALID_CONDITIONS_STR) do
      if (a[0] = ZE_VALID_CONDITIONS_STR[i]) then
      begin
        Result := true;
        break;
      end;

    if (Result) then
      Result := ZEIsTryStrToFloat(a[1], d);

    if (Result) then
      retODSCondution := 'value()' + a[0] + a[1];
  end; //_CheckCondition

begin
  Result := false;
  retODSCondution := '';

  kol := 0;

  _isNumber := false;
  _isCond := false;

  for i := 1 to length(AConditionStr) do
  begin
    ch := AConditionStr[i];
    case (ch) of
      '0'..'9', '.', ',':
        begin
          if (ch = ',') then
            ch := '.';
          _ProcessSymbol(_isCond, _isNumber);
        end;
      '>', '<', '=':  _ProcessSymbol(_isNumber, _isCond);
      ' ':
        if (s <> '') then
           _AddItem();
    end;
  end; //for i

  if (s <> '') then
    _AddItem();

  if (kol >= 2) then
    Result := _CheckCondition();
end; //TryGetMapCondition


////::::::::::::: TZDateTimeODSFormatParser :::::::::::::::::////

constructor TZDateTimeODSFormatParser.Create();
begin
  FCount := 0;
  FMaxCount := 16;
  SetLength(FItems, FMaxCount);
end;

destructor TZDateTimeODSFormatParser.Destroy();
begin
  SetLength(FItems, 0);
  inherited Destroy;
end;

procedure TZDateTimeODSFormatParser.IncCount(ADelta: integer = 1);
begin
  if (ADelta > 0) then
  begin
    inc(FCount, ADelta);
    if (FCount >= FMaxCount) then
    begin
      FMaxCount := FCount + 10;
      SetLength(FItems, FMaxCount);
    end;
  end;
end; //IncCount

procedure TZDateTimeODSFormatParser.CheckMonthMinute();
var
  i: integer;
  _left: boolean;
  _right: boolean;

  //Return FALSE if date and TRUE if time
  function _CheckNeighbors(ADateType: integer): boolean;
  begin
    Result := false;
    case (ADateType) of
      (*
      ZE_DATETIME_ITEM_DAY,
      ZE_DATETIME_ITEM_YEAR,
      ZE_DATETIME_ITEM_WEEK,
      ZE_DATETIME_ITEM_MONTH,
      ZE_DATETIME_ITEM_QUARTER:
        Result := false;
      *)
      ZE_DATETIME_ITEM_AMPM,
      ZE_DATETIME_ITEM_HOUR,
      ZE_DATETIME_ITEM_SECOND:
        Result := true;
    end;
  end; //_CheckNeighbors

  procedure _TryToCheckMonth(AIndex: integer);
  var
    i: integer;

  begin
    _right := false;
    _left := false;
    FItems[AIndex].Settings := 0;

    for i := AIndex - 1 downto 0 do
      if (FItems[i].ItemType > 0) then
      begin
        _left := _CheckNeighbors(FItems[i].ItemType);
        break;
      end;

    for i := AIndex + 1 to FCount - 1 do
      if (FItems[i].ItemType > 0) then
      begin
        _right := _CheckNeighbors(FItems[i].ItemType);
        break;
      end;

    if (_left or _right) then
      FItems[AIndex].ItemType := ZE_DATETIME_ITEM_MINUTE;
  end; //_TryToCheckMonth

begin
  for i := 0 to FCount - 1 do
    if ((FItems[i].ItemType = ZE_DATETIME_ITEM_MONTH) and (FItems[i].Settings = 1)) then
      _TryToCheckMonth(i);
end; //CheckMonthMinute

function TZDateTimeODSFormatParser.TryToParseDateFormat(const AFmtStr: string;
                                                        const AFmtParser: TNumFormatParser): integer;
var
  _parser: TNumFormatParser;
  _isFree: boolean;
  s: string;
  _ch, _prevCh: char;
  _tmp: string;
  _len: integer;
  t: integer;
  _pos: integer;

  procedure _ProcessDateTimeItem();
  begin
    if (s <> '') then
    begin
      _tmp := UpperCase(s);
      _len := Length(_tmp);

      FItems[FCount].Settings := 0;
      FItems[FCount].Len := _len;
      FItems[FCount].TextValue := s;

      case (_tmp[1]) of
        'Y', 'J', 'V':
            FItems[FCount].ItemType := ZE_DATETIME_ITEM_YEAR;
        'M':
          begin
            FItems[FCount].ItemType := ZE_DATETIME_ITEM_MONTH;
            //If can't recognize month / minute
            if (_len <= 2) then
              FItems[FCount].Settings := 1;
          end;
        'D':
            FItems[FCount].ItemType := ZE_DATETIME_ITEM_DAY;
        'N':
          begin
            FItems[FCount].ItemType := ZE_DATETIME_ITEM_DAY;
            inc(FItems[FCount].Len);
          end;
        'H':
            FItems[FCount].ItemType := ZE_DATETIME_ITEM_HOUR;
        'W':
            FItems[FCount].ItemType := ZE_DATETIME_ITEM_WEEK;
        'Q':
            FItems[FCount].ItemType := ZE_DATETIME_ITEM_QUARTER;
        'R':
            if (_len = 1) then
            begin
              FItems[FCount].ItemType := ZE_DATETIME_ITEM_ERA_YEAR;
              FItems[FCount].Len := 2;
            end
            else
            begin
              FItems[FCount].ItemType := ZE_DATETIME_ITEM_ERA_JAP;
              FItems[FCount].Len := 4;
            end;
        'E':
            FItems[FCount].ItemType := ZE_DATETIME_ITEM_ERA_YEAR;
        else
            FItems[FCount].ItemType := ZE_DATETIME_ITEM_TEXT;
      end; //case

      IncCount();

      s := '';
    end;
  end; //_ProcessDateTimeItem()

  procedure _AddItemCommon(const AStr: string; ALen: integer; AItemType: integer; ASettings: integer = 0);
  begin
    FItems[FCount].ItemType := AItemType;
    FItems[FCount].Len := ALen;
    FItems[FCount].TextValue := AStr;
    FItems[FCount].Settings := ASettings;
    IncCount();
  end;

  procedure _ProcessAMPM();
  begin
    //check for a/p A/P
    if (_parser.CurrentPos + 1 <= _parser.StrLength) then
    begin
      _tmp := Copy(AFmtStr, _parser.CurrentPos - 1, 3);

      if (UpperCase(_tmp) = 'A/P') then
      begin
        if (_parser.FirstSymbol = 'a') then
          t := ZE_DATETIME_AMPM_SHORT_LOW
        else
          t := ZE_DATETIME_AMPM_SHORT_UP;
        _AddItemCommon(_tmp, 3, ZE_DATETIME_ITEM_AMPM, t);
        _parser.IncPos(2);

        exit;
      end;
    end;

    //check for am/pm AM/PM
    if (_parser.CurrentPos + 3 <= _parser.StrLength) then
    begin
      _tmp := Copy(AFmtStr, _parser.CurrentPos - 1, 5);

      if (UpperCase(_tmp) = 'AM/PM') then
      begin
        if (_parser.FirstSymbol = 'a') then
          t := ZE_DATETIME_AMPM_LONG_LOW
        else
          t := ZE_DATETIME_AMPM_LONG_UP;
        _AddItemCommon(_tmp, 5, ZE_DATETIME_ITEM_AMPM, t);
        _parser.IncPos(4);

        exit;
      end;
    end;

    //It is not AM/PM. May be a year?
    s := s + 'Y';
  end; //_ProcessAMPM

  procedure _ProcessSeconds();
  begin
    s := _ch;
    _len := 1;

    t := 0;
    while (_parser.CurrentPos <= _parser.StrLength) do
    begin
      _ch := AFmtStr[_parser.CurrentPos];
      case (_ch) of
        's', 'S':
          begin
            s := s + _ch;
            inc(_len);
          end
        else
          begin
            if ((_ch = '.') or (_ch = ',')) then
            begin
              s := s + '.';
              while (_parser.CurrentPos <= _parser.StrLength) do
              begin
                _parser.IncPos(1);
                if (AFmtStr[_parser.CurrentPos] = '0') then
                begin
                  s := s + '0';
                  inc(t);
                end
                else
                  break;
              end; //while
            end; //if

            break;
          end;
      end;
      _parser.IncPos(1);
    end; //while
    _AddItemCommon(s, _len, ZE_DATETIME_ITEM_SECOND, t);
    s := '';
  end; //_ProcessSeconds

  procedure _TryToAddEraJap();
  begin
    _tmp := UpperCase(s);

    t := -1;
    if (_tmp = 'G') then
      t := 1
    else
    if (_tmp = 'GG') then
      t := 2
    else
    if (_tmp = 'GGG') then
      t := 3
    else
    if (_tmp = 'GGGEE') then
      t := 4;

    if (t > 0) then
      _AddItemCommon(s, t, ZE_DATETIME_ITEM_ERA_JAP, t);

    s := '';
  end;

  procedure _ProcessEraJap();
  begin
    s := _ch;
    _pos := _parser.CurrentPos;
    while (_pos <= _parser.StrLength) do
    begin
      _ch := AFmtStr[_pos];
      case (_ch) of
        'g', 'G', 'e', 'E':
          s := s + _ch;
        else
          begin
            _TryToAddEraJap();
            exit;
          end;
      end;
      inc(_pos);
      _parser.IncPos(1);
    end; //while

    if (s <> '') then
      _TryToAddEraJap();
  end; //_ProcessEraJap

  procedure _ProcessSymbol();
  begin
    _ch := _parser.FirstSymbol;

    if (UpperCase(_prevCh) = UpperCase(_ch)) then
      s := s + _ch
    else
    begin
      _ProcessDateTimeItem();

      //TODO:
      //     Need check all symbols for other locales (A/J/V - as year etc)
      case (_ch) of
        'a', 'A':
                  _ProcessAMPM();
        'y', 'Y',
        'j', 'J', //German year  ??
        'v', 'V', //Finnish year ??
        'm', 'M',
        'd', 'D',
        'n', 'N',
        'h', 'H',
        'w', 'W',
        'r', 'R',
        'q', 'Q',
        'e', 'E':
                  s := s + _ch;
        's', 'S':
                  _ProcessSeconds();
        'g', 'G':
                  _ProcessEraJap();
        else
          s := s + _ch;
      end; //case
    end;

    _prevCh := _ch;
  end; //_ProcessSymbol

  procedure _ProcessText();
  begin
    _ProcessDateTimeItem();
    FItems[FCount].ItemType := ZE_DATETIME_ITEM_TEXT;
    FItems[FCount].TextValue := _parser.ReadedSymbol;
    IncCount();
  end; //_ProcessText

begin
  FCount := 0;
  _parser := AFmtParser;
  _isFree := AFmtParser = nil;
  if (_isFree) then
    _parser := TNumFormatParser.Create();

  s := '';
  _prevCh := #0;

  try
    _parser.BeginRead(AFmtStr);

    while (_parser.ReadSymbol()) do
    begin
      case (_parser.ReadedSymbolType) of
        0:
          _ProcessSymbol();
        1:; //brackets - modiefier: color, calendar or conditions ??
        2, 3:
          _ProcessText();
      end;
    end; //while

    _ProcessDateTimeItem();

    CheckMonthMinute();

  finally
    if (_isFree) then
      FreeAndNil(_parser);
  end;

  Result := FCount;
end; //TryToParseDateFormat

procedure TZDateTimeODSFormatParser.DeleteRepeatedItems();
var
  i, j: integer;

begin
  for i := 1 to FCount - 2 do
    if (FItems[i].ItemType > 0) then
      for j := FCount - 1 downto i + 1 do
        if (FItems[i].ItemType = FItems[j].ItemType) then
        begin
          if (FItems[i].ItemType = ZE_DATETIME_ITEM_DAY) then
          begin
            if (FItems[i].Len = FItems[j].Len) then
              FItems[j].ItemType := ZE_DATETIME_ITEM_ERROR;
          end
          else
            FItems[j].ItemType := ZE_DATETIME_ITEM_ERROR;
        end;
end; //DeleteRepeatedItems

function TZDateTimeODSFormatParser.GetValidCount(): integer;
var
  i: integer;

begin
  Result := 0;
  for i := 1 to FCount - 1 do
    if (FItems[i].ItemType >= 0) then
      inc(Result);
end; //GetValidCount

////::::::::::::: TNumFormatParser :::::::::::::::::////

procedure TNumFormatParser.Clear();
begin
  FLen := -1;
  FPos := 1;
  FStr := '';
  FIsError := 0;
  FReadedSymbolType := 0;
  FReadedSymbol := '';
  FFirstSymbol := #0;
end;

constructor TNumFormatParser.Create();
begin
  Clear();
end;

procedure TNumFormatParser.BeginRead(const AStr: string);
begin
  Clear();
  FStr := AStr;
  FLen := Length(AStr);
end;

function TNumFormatParser.ReadSymbol(): boolean;
var
  ch: char;

  procedure _ReadBeforeSymbol(Symbol: char);
  begin
    if (FPos <= FLen) then
      FFirstSymbol := FStr[FPos];

    while (FPos <= FLen) do
    begin
      ch := FStr[FPos];
      inc(FPos);

      if (ch = Symbol) then
        exit;

      FReadedSymbol := FReadedSymbol + ch;
    end; //while

    FIsError := FIsError or 2;
  end; //_ReadBeforeSymbol

begin
  FFirstSymbol := #0;
  if (FPos > FLen) then
  begin
    Result := false;
    exit;
  end;

  FReadedSymbol := '';

  Result := true;
  while (FPos <= FLen) do
  begin
    ch := FStr[FPos];
    inc(FPos);

    case ch of
      '[':
        begin
          FReadedSymbolType := 1;
          _ReadBeforeSymbol(']');
          exit;
        end;
      '"':
        begin
          FReadedSymbolType := 2;
          _ReadBeforeSymbol('"');
          exit;
        end;
      '\':
        begin
          if (FPos <= FLen) then
          begin
            FFirstSymbol := FStr[FPos];
            FReadedSymbol := FFirstSymbol;
          end
          else
          begin
            FIsError := FIsError or 4;
            FReadedSymbol := '';
          end;
          inc(FPos);
          FReadedSymbolType := 3;
          exit;
        end;
      else
        begin
          FReadedSymbol := ch;
          FFirstSymbol := ch;
          FReadedSymbolType := 0;
          exit;
        end;
    end; //case
  end; //while

  FIsError := FIsError or 1;
end; //ReadSymbol

procedure TNumFormatParser.IncPos(ADelta: integer);
begin
  inc(FPos, ADelta);
  if (FPos < 1) then
    FPos := 1;
end; //IncPos

////::::::::::::: TZEODSNumberFormatReader :::::::::::::::::////

procedure TZEODSNumberFormatReader.AddEmbededText(const AText: string;
                                                  ANumberPosition: integer);
var
  i: integer;
  _pos: integer;

begin
  if (FEmbededTextCount >= FEmbededMaxCount) then
  begin
    inc(FEmbededMaxCount, 10);
    SetLength(FEmbededTextArray, FEmbededMaxCount);
  end;

  _pos := -1;

  for i := 0 to FEmbededTextCount - 1 do
    if (ANumberPosition < FEmbededTextArray[i].NumberPosition) then
    begin
      _pos := i;
      break;
    end;

  if (_pos >= 0) then
  begin
    for i := FEmbededTextCount + 1 downto _pos + 1 do
      FEmbededTextArray[i] := FEmbededTextArray[i - 1];
  end
  else
    _pos := FEmbededTextCount;

  FEmbededTextArray[_pos].Txt := AText;
  FEmbededTextArray[_pos].NumberPosition := ANumberPosition;

  inc(FEmbededTextCount);
end;

procedure TZEODSNumberFormatReader.AddItem();
var
  i: integer;

begin
  inc(FCount);
  if (FCount >= FCountMax) then
  begin
    inc(FCountMax, 20);
    SetLength(FItems, FCountMax);
    SetLength(FItemsOptions, FCountMax);
    for i := FCount to FCount - 1 do
    begin
      FItemsOptions[i].isColor := false;
      FItemsOptions[i].ColorStr := '';
    end;
  end;
end;

constructor TZEODSNumberFormatReader.Create();
var
  i: integer;

begin
  FCount := 0;
  FCountMax := 20;
  FEmbededMaxCount := 10;
  SetLength(FItems, FCountMax);
  SetLength(FItemsOptions, FCountMax);
  SetLength(FEmbededTextArray, FEmbededMaxCount);
  for i := 0 to FCountMax - 1 do
  begin
    FItemsOptions[i].isColor := false;
    FItemsOptions[i].ColorStr := '';
    FItemsOptions[i].StyleType := 0;
  end;
end;

destructor TZEODSNumberFormatReader.Destroy();
begin
  SetLength(FItems, 0);
  SetLength(FItemsOptions, 0);
  SetLength(FEmbededTextArray, 0);
  inherited;
end;

function TZEODSNumberFormatReader.BeginReadFormat(const xml: TZsspXMLReaderH; out retStartString: string; const NumFormat: integer): integer;
begin
  Result := FCount;
  AddItem();
  FItems[Result][0] := xml.Attributes[ZETag_Attr_StyleName];
  FItemsOptions[Result].StyleType := NumFormat;
  retStartString := '';
end; //BeginReadFormat

//Read date format: <number:date-style>.. </number:date-style>
procedure TZEODSNumberFormatReader.ReadDateFormat(const xml: TZsspXMLReaderH; const ATagName: string);
var
  num: integer;
  s, _result: string;
  _isLong: boolean;
  t: integer;
  i: integer;

  function CheckIsLong(const isTrue, isFalse: string): string;
  begin
    if (xml.Attributes[ZETag_number_style] = ZETag_long) then
      Result := isTrue
    else
      Result := isFalse;
  end;

begin
  num := BeginReadFormat(xml, _result, ZE_NUMFORMAT_IS_DATETIME);

  while ((xml.TagType <> 6) or (xml.TagName <> ATagName)) do
  begin
    xml.ReadTag();

    //Day
    if ((xml.TagName = ZETag_number_day) and (xml.TagType and 4 = 4)) then
      _result := _result + CheckIsLong('DD', 'D')
    else
    //Text
    if ((xml.TagName = ZETag_number_text) and (xml.TagType = 6)) then
    begin
      s := xml.TextBeforeTag;
      t := length(s);
      if (t = 1) then
      begin
        case (s[1]) of
          ' ', '.', ':', '-', '/', '*':;
          else
            s := '\' + s;
        end; //case
      end
      else
        s := '"' + s + '"';

      _result := _result + s;
    end
    else
    //Month
    if (xml.TagName = ZETag_number_month) then
    begin
      _isLong := xml.Attributes[ZETag_number_style] = ZETag_long;
      s := xml.Attributes[ZETag_number_textual];
      if (ZEStrToBoolean(s)) then
        _result := _result + IfThen(_isLong, 'MMMM', 'MMM')
      else
        _result := _result + IfThen(_isLong, 'MM', 'M')
    end
    else
    //Year
    if (xml.TagName = ZETag_number_year) then
      _result := _result + CheckIsLong('YYYY', 'YY')
    else
    //Hours
    if (xml.TagName = ZETag_number_hours) then
      _result := _result + CheckIsLong('HH', 'H')
    else
    //Minutes
    if (xml.TagName = ZETag_number_minutes) then
      _result := _result + CheckIsLong('mm', 'm')
    else
    //Seconds
    if (xml.TagName = ZETag_number_seconds) then
    begin
      _result := _result + CheckIsLong('ss', 's');
      s := xml.Attributes[ZETag_number_decimal_places];
      if (s <> '') then
        if (TryStrToInt(s, t)) then
          if (t > 0) then
          begin
            _result := _result + '.';
            for i := 1 to t do
              _result := _result + '0';
          end;
    end
    else
    //AM/PM
    if (xml.TagName = ZETag_number_am_pm) then
    begin
      _result := _result + 'AM/PM';
    end
    else
    //Era
    if (xml.TagName = ZETag_number_era) then
    begin
      //Attr: number:calendar
      //      number:style
      _result := _result + IfThen(_isLong, 'GG', 'G')
    end
    else
    //Quarter
    if (xml.TagName = ZETag_number_quarter) then
    begin
      //Attr: number:calendar
      //      number:style
      _result := _result + CheckIsLong('QQ', 'Q')
    end
    else
    //Day of week
    if (xml.TagName = ZETag_number_day_of_week) then
    begin
      //Attr: number:calendar
      //      number:style
      _result := _result + CheckIsLong('NNN', 'NN')
    end
    else
    //Week of year
    if (xml.TagName = ZETag_number_week_of_year) then
    begin
      //Attr: number:calendar
      _result := _result + 'WW';
    end;

    if (xml.Eof()) then
      break;
  end; //while
  FItems[num][1] := _result;
end; //ReadDateFormat

//Read string format <number:text-style> .. </number:text-style>
procedure TZEODSNumberFormatReader.ReadStringFormat(const xml: TZsspXMLReaderH);
var
  num: integer;
  _result: string;

begin
  num := BeginReadFormat(xml, _result, ZE_NUMFORMAT_IS_STRING);

  //Possible attributes for tag "number:text-style":
  //         number:country
  //         number:language
  //         number:rfc-language-tag
  //         number:script
  //         number:title
  //         number:transliteration-country
  //         number:transliteration-format
  //         number:transliteration-language
  //         number:transliteration-style
  //         style:display-name
  //         style:name
  //         style:volatile

  //Possible child elements:
  //         number:text           *
  //         number:text-content   *
  //         style:map
  //         style:textproperties
  while ((xml.TagType <> 6) or (xml.TagName <> ZETag_number_text_style)) do
  begin
    xml.ReadTag();

    //number:text-content
    if ((xml.TagName = ZETag_number_text_content)) then
      _result := _result + '@'
    else
    //Text
    if ((xml.TagName = ZETag_number_text) and (xml.TagType = 6)) then
      _result := _result + '"' + xml.TextBeforeTag + '"';

    if (xml.Eof()) then
      break;
  end; //while
  FItems[num][1] := _result;
end; //ReadStringFormat

//Read known numbers formats (date/number/percentage etc)
procedure TZEODSNumberFormatReader.ReadKnownNumberFormat(const xml: TZsspXMLReaderH);
begin
  if (xml.TagName = ZETag_number_number_style) then
    ReadNumberFormat(xml)
  else
  if (xml.TagName = ZETag_number_currency_style) then
    ReadCurrencyFormat(xml)
  else
  if (xml.TagName = ZETag_number_percentage_style) then
    ReadPercentageFormat(xml)
  else
  if (xml.TagName = ZETag_number_date_style) then
    ReadDateFormat(xml, ZETag_number_date_style)
  else
  if (xml.TagName = ZETag_number_time_style) then
    ReadDateFormat(xml, ZETag_number_time_style)
  else
  if (xml.TagName = ZETag_number_text_style) then
    ReadStringFormat(xml);
end;

procedure TZEODSNumberFormatReader.ReadCurrencyFormat(const xml: TZsspXMLReaderH);
begin
  ReadNumberFormatCommon(xml, ZETag_number_currency_style, ZE_NUMFORMAT_NUM_IS_CURRENCY);
end;

//Read number style: <number:number-style> .. </number:number-style>
procedure TZEODSNumberFormatReader.ReadNumberFormat(const xml: TZsspXMLReaderH);
begin
  ReadNumberFormatCommon(xml, ZETag_number_number_style, 0);
end;

//Read number style: <number:percentage-style> .. </number:percentage-style>
procedure TZEODSNumberFormatReader.ReadPercentageFormat(const xml: TZsspXMLReaderH);
begin
  ReadNumberFormatCommon(xml, ZETag_number_percentage_style, ZE_NUMFORMAT_NUM_IS_PERCENTAGE);
end;

//Read Number/currency/percentage number format style
//INPUT
//  const xml: TZsspXMLReaderH    - xml
//  const NumberFormatTag: string - tag name
//      sub_number_type: integer  - additional flag for number (percentage/scientific etc)
procedure TZEODSNumberFormatReader.ReadNumberFormatCommon(const xml: TZsspXMLReaderH;
                                                          const NumberFormatTag: string;
                                                          sub_number_type: integer);
var
  num: integer;
  s, _result, _txt, _style_name: string;
  _cond_text: string;
  _cond: string;
  _decimalPlaces: integer;
  _min_int_digits: integer;
  _display_factor: integer;
  _number_grouping: boolean;
  _number_position: integer;
  _is_number_decimal_replacement: boolean;
  ch: char;

  _min_numerator_digits: integer;
  _min_denominator_digits: integer;
  _denominator_value: integer;

  procedure _TryGetIntValue(const ATagName: string; out retIntValue: integer; const ADefValue: integer = 0);
  begin
    s := xml.Attributes[ATagName];
    if (not TryStrToInt(s, retIntValue)) then
      retIntValue := ADefValue;
  end;

  procedure _ReadNumber_NumberPrepare();
  begin
    FEmbededTextCount := 0;

    _TryGetIntValue(ZETag_number_decimal_places, _decimalPlaces);
    _TryGetIntValue(ZETag_number_min_integer_digits, _min_int_digits);

    s := xml.Attributes[ZETag_number_display_factor];
    if (s <> '') then
    begin
      if (not TryStrToInt(s, _display_factor)) then
        _display_factor := 1;
    end
    else
      _display_factor := 1;

    _number_grouping := false;
    s := xml.Attributes[ZETag_number_grouping];
    if (s <> '') then
      _number_grouping := ZEStrToBoolean(s);

    _is_number_decimal_replacement := xml.Attributes.IsContainsAttribute(ZETag_number_decimal_replacement);
  end; //_ReadNumber_NumberPrepare

  procedure _ReadEmbededText();
  begin
    _number_position := -100;
    while ((xml.TagType <> 6) or (xml.TagName <> ZETag_number_number)) do
    begin
      xml.ReadTag();

      //<number:embedded-text number:position="1">..</number:embedded-text>
      if (xml.TagName = ZETag_number_embedded_text) then
      begin
        if (xml.TagType = 4) then
          _TryGetIntValue(ZETag_number_position, _number_position, -100)
        else
        if ((xml.TagType = 6) and (_number_position >= 0)) then
          if (xml.TextBeforeTag <> '') then
            AddEmbededText(ZEReplaceEntity(xml.TextBeforeTag), _number_position);
      end;

      if (xml.Eof()) then
        break;
    end; //while
  end; //_ReadEmbededText

  function _GetRepeatedString(ACount: integer; const AStr: string): string;
  var
    i: integer;

  begin
    Result := '';
    for i := 1 to ACount do
      Result := Result + AStr;
  end;

  // <number:number>..</number:number>
  procedure _ReadNumber_Number();
  var
    i, j: integer;
    _pos: integer;
    _currentpos: integer;   // current position for embeded text

  begin
    _ReadNumber_NumberPrepare();

    if (xml.TagType = 4) then
      _ReadEmbededText();

    if (FEmbededTextCount > 0) then
    begin
      s := '';
      _pos := 0;

      for i := 0 to FEmbededTextCount - 1 do
        if (FEmbededTextArray[i].NumberPosition >= 0) then
        begin
          _currentpos := FEmbededTextArray[i].NumberPosition;
          //TODO: need test. For example: if symbol "%" not one? (%0%0.00% or "%"0.0)
          (*
          if (FEmbededTextArray[i].Txt = '%') then
            _txt := '%'
          else
          *)
          _txt := '"' + ZEReplaceEntity(FEmbededTextArray[i].Txt) + '"';

          if (_currentpos <= _min_int_digits) then
            ch := '0'
          else
            ch := '#';

          for j := _pos to _currentpos - 1 do
            s := ch + s;
          s := _txt + s ;
          _pos := _currentpos;
        end;

      if (_currentpos < _min_int_digits) then
        for j := _pos to _min_int_digits - 1 do
          s := '0' + s;

      _result := _result + s;
    end
    else
    begin
      if (_min_int_digits = 0) then
        _result := _result + '#'
      else
        for i := 0 to _min_int_digits - 1 do
          _result := _result + '0';
    end;

    if (_decimalPlaces > 0) then
    begin
      if (_is_number_decimal_replacement) then
        ch := '#'
      else
        ch := '0';

      _result := _result + ZE_NUMBER_FORMAT_DECIMAL_SEPARATOR;
      for i := 0 to _decimalPlaces - 1 do
        _result := _result + ch;
    end;
  end; //_ReadNumber_Number

  // <number:fraction />
  procedure _ReadNumber_Fraction();
  begin
    _ReadNumber_NumberPrepare();

    _TryGetIntValue(ZETag_number_min_numerator_digits, _min_numerator_digits);
    _TryGetIntValue(ZETag_number_min_denominator_digits, _min_denominator_digits);
    //TODO: do not forget about denominator_value!
    _TryGetIntValue(ZETag_number_denominator_value, _denominator_value);

    if (_min_int_digits <= 0) then
      s := '#'
    else
      s := _GetRepeatedString(_min_int_digits, '0');

    _result := _result + s;

    if ((_min_numerator_digits > 0) and (_min_denominator_digits > 0)) then
      _result := _result + ' ' +
                _GetRepeatedString(_min_numerator_digits, '?') +
                '/' +
                _GetRepeatedString(_min_denominator_digits, '?');

    FItemsOptions[num].StyleType := FItemsOptions[num].StyleType or ZE_NUMFORMAT_NUM_IS_FRACTION;
  end; //_ReadNumber_Fraction

  // <number:scientific-number .. />
  procedure _ReadNumber_Scientific();
  var
    _min_exponent_digits: integer;

  begin
    _ReadNumber_NumberPrepare();

    _TryGetIntValue(ZETag_number_min_exponent_digits, _min_exponent_digits);

    if (_min_exponent_digits > 0) then
    begin
      _result := _result + _GetRepeatedString(_min_int_digits, '0') +
                ZE_NUMBER_FORMAT_DECIMAL_SEPARATOR +
                _GetRepeatedString(_decimalPlaces, '0') +
                'E+' +
                _GetRepeatedString(_min_exponent_digits, '0');

      FItemsOptions[num].StyleType := FItemsOptions[num].StyleType or ZE_NUMFORMAT_NUM_IS_SCIENTIFIC;
    end;
  end; //_ReadNumber_Scientific

  //<number:currency-symbol> .. </<number:currency-symbol>
  procedure _ReadCurrecny_Symbol();
  begin
    //TODO: need get currency symbol by attributes:
    //      ZETag_number_language = 'number:language'
    //      ZETag_number_country  = 'number:country'
    if (xml.TagType = 4) then
      while ((xml.TagType <> 6) or (xml.TagName <> ZETag_number_currency_symbol)) do
      begin
        xml.ReadTag();

        (* //if need currency name in future
        if (xml.TagName = ZETag_number_currency_symbol) then
           s := xml.TextBeforeTag;
        *)

        if (xml.Eof()) then
          break;
      end; //while
  end; //_ReadCurrecny_Symbol

  // <style:map />
  procedure _ReadStyleMap();
  var
    i: integer;

  begin
    _cond := ZEReplaceEntity(xml.Attributes[ZETag_style_condition]);
    if (_cond <> '') then
    begin
      i := pos('value()', _cond);
      if (i = 1) then
        delete(_cond, 1, 7)
      else
        exit;

      if (_cond <> '') then
      begin
        _style_name := xml.Attributes[ZETag_style_apply_style_name];
        for i := 0 to FCount - 1 do
          if (FItems[i][0] = _style_name) then
          begin
            _txt := FItems[i][1];

            _cond_text := _cond_text + '[' + _cond + ']' + _txt + ';';

            break;
          end;
      end; //if
    end; //if
  end; //_ReadStyleMap

  // <style:text-properties .. /> (color by condition)
  procedure _ReadStyleTextProperties();
  begin
    //for now only colors
    s := UpperCase(xml.Attributes[ZETag_fo_color]);
    if (TryGetMapColorName(s, _txt)) then
    begin
      FItemsOptions[num].isColor := true;
      FItemsOptions[num].ColorStr := _txt;
      _result := _result + '[' + _txt + ']';
    end;
  end; //_ReadStyleTextProperties

begin
  num := BeginReadFormat(xml, _result, ZE_NUMFORMAT_IS_NUMBER or sub_number_type);
  _cond_text := '';

  while ((xml.TagType <> 6) or (xml.TagName <> NumberFormatTag)) do
  begin
    xml.ReadTag();

    if (xml.TagName = ZETag_number_number) then
      _ReadNumber_Number()
    else
    if (xml.TagName = ZETag_number_fraction) then
      _ReadNumber_Fraction()
    else
    if (xml.TagName = ZETag_number_scientific_number) then
      _ReadNumber_Scientific()
    else
    if (xml.TagName = ZETag_number_currency_symbol) then
      _ReadCurrecny_Symbol()
    else
    if ((xml.TagName = ZETag_number_text) and (xml.TagType = 6)) then
    begin
      s := ZEReplaceEntity(xml.TextBeforeTag);
      if (s = '"') then
        _result := _result + '\' + s
      else
      if (s = '%') then
        _result := _result + s
      else
        _result := _result + '"' + s + '"';
    end
    else
    if (xml.TagName = ZETag_style_map) then
      _ReadStyleMap()
    else
    if (xml.TagName = ZETag_style_text_properties) then
      _ReadStyleTextProperties();

    if (xml.Eof()) then
      break;
  end; //while

  //first - map number styles, after - current readed number format
  FItems[num][1] := _cond_text + _result;
end; //ReadNumberFormatCommon

function TZEODSNumberFormatReader.TryGetFormatStrByNum(const DataStyleName: string;
                                                       out retFormatStr: string): boolean;
var
  i: integer;

begin
  Result := false;
  for i := 0 to FCount - 1 do
    if (FItems[i][0] = DataStyleName) then
    begin
      Result := true;
      retFormatStr := FItems[i][1];
      break;
    end;
end; //TryGetFormatStrByNum


////::::::::::::: TZEODSNumberFormatWriter :::::::::::::::::////

constructor TZEODSNumberFormatWriter.Create(const AMaxCount: integer);
var
  i: integer;

begin
  FCount := 0;
  FCountMax := AMaxCount;
  if (FCountMax < 10) then
    FCountMax := 10;
  SetLength(FItems, FCountMax);

  SetLength(FNumberAdditionalProps, FCountMax);
  for i := 0 to FCountMax - 1 do
    FNumberAdditionalProps[i] := 0;

  FCurrentNFIndex := 100;

  FNFItemsCount := 0;
  SetLength(FNFItems, ZE_MAX_NF_ITEMS_COUNT);
  for i := 0 to ZE_MAX_NF_ITEMS_COUNT - 1 do
    FNFItems[i] := TODSNumberFormatMapItem.Create();
end;

destructor TZEODSNumberFormatWriter.Destroy();
var
  i: integer;

begin
  SetLength(FItems, 0);
  for i := 0 to ZE_MAX_NF_ITEMS_COUNT - 1 do
    FreeAndNil(FNFItems[i]);
  SetLength(FNFItems, 0);

  SetLength(FNumberAdditionalProps, 0);

  inherited;
end;

//Try to find number format name for style num StyleID
//INPUT
//      StyleID: integer          - style ID
//  out NumberFormatName: string  - finded number format name
//RETURN
//      boolean - true - number format finded
function TZEODSNumberFormatWriter.TryGetNumberFormatName(StyleID: integer;
                                                         out NumberFormatName: string): boolean;
var
  i: integer;

begin
  Result := false;
  for i := 0 to FCount - 1 do
    if (FItems[i].StyleIndex = StyleID) then
    begin
      NumberFormatName := FItems[i].NumberFormatName;
      Result := true;
      break;
    end;
end; //TryGetNumberFormatName

//Try to find additional properties for number format
//INPUT
//      StyleID: integer          - style ID
//  out NumberFormatProp: integer - finded number additional properties
//RETURN
//      boolean - true - additional properties is found
function TZEODSNumberFormatWriter.TryGetNumberFormatAddProp(StyleID: integer;
                                                            out NumberFormatProp: integer): boolean;
begin
  Result := (StyleID >= 0) and (StyleID < FCountMax);

  if (Result) then
    NumberFormatProp := FNumberAdditionalProps[StyleID]
  else
    NumberFormatProp := 0;
end; //TryGetNumberFormatAddProp

//Separate number format string  by ";".
//INPUT
//  const NFStr: string - Number format string (like "nf1;nf2;nf3")
//RETURN
//      integer - count of number format items
function TZEODSNumberFormatWriter.SeparateNFItems(const NFStr: string): integer;
var
  i, l: integer;
  b: boolean;
  s: string;
  ch: char;

begin
  b := true;
  s := '';

  l := length(NFStr);

  for i := 1 to l do
  begin
    ch := NFStr[i];

    if (ch = '"') then
      b := not b;

    if (b) then
    begin
      if (ch = ';') then
      begin
        TryAddNFItem(s);
        s := '';
      end
      else
        s := s + ch;
    end
    else
      s := s + ch;
  end; //for i

  if (s <> '') then
    TryAddNFItem(s);

  Result := FNFItemsCount;
end; //PrepareNFItems

//Try to add NumberFormat item (from string "NF1;NF2;NF3")
//Checks:
//  1. Count of NF items (max = 3)
//  2. is NF item valid
//INPUT
//  const NFStr: string - NF item
//RETURN
//      boolean - true - item added
function TZEODSNumberFormatWriter.TryAddNFItem(const NFStr: string): boolean;
begin
  Result := false;

  if (FNFItemsCount < ZE_MAX_NF_ITEMS_COUNT) then
  begin
    Result := FNFItems[FNFItemsCount].TryToParse(NFStr);
    if (Result) then
      inc(FNFItemsCount);
  end;
end; //TryAddNFItem

//Try to write number format to xml
//INPUT
//  const xml: TZsspXMLWriterH  - xml
//      StyleID: integer        - Style ID
//      ANumberFormat: string   - number format
//RETURN
//      boolean - true - NumberFormat was written ok
function TZEODSNumberFormatWriter.TryWriteNumberFormat(const xml: TZsspXMLWriterH;
                                                       StyleID: integer;
                                                       ANumberFormat: string): boolean;
var
  s: string;
  _nfType: integer;
  _nfName: string;

  function _WriteNumberNumber(NumProperties: integer = 0): boolean;
  var
    i: integer;
    num: integer;
    _item_name: string;

  begin
    Result := false;
    //  NumberFormat = "part1;part2;part3"
    //    part1 - for numbers > 0 (or for condition1)
    //    part2 - for numbers < 0 (or for condition2)
    //    part3 - for 0 (or for other numbers if used condition1 and condition2)
    //  partX = [condition][color]number_format

    if (SeparateNFItems(ANumberFormat) > 0) then
    begin
      if (FNFItemsCount = 1) then
      begin
        FNFItems[0].WriteNumberStyle(xml, _nfName, NumProperties);
        Result := true;
      end
      else
      begin
        num := 0;
        for i := 0 to FNFItemsCount - 2 do
        begin
          if (FNFItems[i].isCondition) then
            s := FNFItems[i].Condition
          else
            case i of
              0: s := '>0';
              1: s := '< 0';
              else
                s := '';
            end;

          if (s <> '') then
          begin
            _item_name := _nfName + 'P' + IntToStr(num);
            FNFItems[FNFItemsCount - 1].AddCondition(s, _item_name);
            FNFItems[i].WriteNumberStyle(xml, _item_name, NumProperties, true);
            inc(num);
          end;
        end; //for i

        FNFItems[FNFItemsCount - 1].WriteNumberStyle(xml, _nfName, NumProperties);
        Result := true;
      end;
    end; //if
  end; //_WriteNumberNumber

  function _WriteCurrency(): boolean;
  begin
    Result := _WriteNumberNumber(ZE_NUMFORMAT_NUM_IS_CURRENCY);
  end;

  function _WritePercentage(): boolean;
  begin
    Result := _WriteNumberNumber(ZE_NUMFORMAT_NUM_IS_PERCENTAGE);
  end;

  function _WriteDateTime(): boolean;
  var
    _addProp: integer;

  begin
    Result := false;

    if (SeparateNFItems(ANumberFormat) > 0) then
    begin
      //For now use only first NF item
      //TODO:
      //     Are conditions implements for date styles?
      _addProp := FNFItems[0].WriteDateTimeStyle(xml, _nfName);
      _nfType := _nfType or _addProp;
      Result := true;
    end; //if
  end; //_WriteDateTime

  function _WriteStringFormat(): boolean;
  begin
    Result := false;
    if (SeparateNFItems(ANumberFormat) > 0) then
    begin
      //For now use only first NF item
      //TODO:
      //     Are conditions implements for text styles?
      FNFItems[0].WriteTextStyle(xml, _nfName);
      Result := true;
    end; //if
  end; //_WriteStringFormat

  function _WriteNumberStyle(): boolean;
  begin
    FNFItemsCount := 0;
    _nfName := 'N' + IntToStr(FCurrentNFIndex);
    Result := false;

    _nfType := GetNativeNumberFormatType(ANumberFormat);

    case (_nfType and $FF) of
      ZE_NUMFORMAT_IS_NUMBER:
        begin
          if (_nfType and ZE_NUMFORMAT_NUM_IS_CURRENCY = ZE_NUMFORMAT_NUM_IS_CURRENCY) then
            Result := _WriteCurrency()
          else
          if (_nfType and ZE_NUMFORMAT_NUM_IS_PERCENTAGE = ZE_NUMFORMAT_NUM_IS_PERCENTAGE) then
            Result := _WritePercentage()
          else
            Result := _WriteNumberNumber();
        end;
      ZE_NUMFORMAT_IS_DATETIME:
        Result := _WriteDateTime();
      ZE_NUMFORMAT_IS_STRING:
        Result := _WriteStringFormat();
    end;
  end; //_WriteNumberStyle

  procedure _AddItem(const NFName: string);
  begin
    if (FCount >= FCountMax) then
    begin
      inc(FCountMax, 10);
      SetLength(FItems, FCountMax);
      SetLength(FNumberAdditionalProps, FCountMax);
    end;

    FItems[FCount].StyleIndex := StyleID;
    FItems[FCount].NumberFormatName := NFName;
    FItems[FCount].NumberFormat := ANumberFormat;

    FNumberAdditionalProps[StyleID] := _nfType;

    inc(FCount);
  end;

  function _CheckIsDuplicate(): boolean;
  var
    i: integer;

  begin
    Result := false;
    for i := 0 to FCount - 1 do
      if (FItems[i].NumberFormat = ANumberFormat) then
      begin
        _nfType := FNumberAdditionalProps[StyleID];
        _AddItem(FItems[i].NumberFormatName);
        Result := true;
        break;
      end;
  end; //_CheckIsDuplicate


begin
  Result := false;
  ANumberFormat := Trim(ANumberFormat);

  if ((ANumberFormat = '@') or (ANumberFormat = '')) then
    exit;

  s := UpperCase(ANumberFormat);

  if ((s = 'GENERAL') or (s = 'STANDART')) then
    exit;

  if (TryGetNumFormatByName(ANumberFormat, s)) then
    if (s = '') then
      exit
    else
      ANumberFormat := s;

  if (_CheckIsDuplicate()) then
    Result := true
  else
  begin
    Result := _WriteNumberStyle();

    if (Result) then
    begin
      _AddItem(_nfName);
      inc(FCurrentNFIndex);
    end;
  end;
end; //TryWriteNumberFormat


////::::::::::::: TODSNumberFormatMapItem :::::::::::::::::////

constructor TODSNumberFormatMapItem.Create();
begin
  FEmbededMaxCount := 10;
  SetLength(FEmbededTextArray, FEmbededMaxCount);
  FNumberFormatParser := TNumFormatParser.Create();
  FDateTimeODSFormatParser := TZDateTimeODSFormatParser.Create();
end;

destructor TODSNumberFormatMapItem.Destroy();
begin
  SetLength(FEmbededTextArray, 0);
  FreeAndNil(FNumberFormatParser);
  FreeAndNil(FDateTimeODSFormatParser);
  inherited Destroy;
end;

procedure TODSNumberFormatMapItem.Clear();
begin
  FCondition := '';
  FisCondition := false;
  FColorStr := '';
  FisColor := false;
  FNumberFormat := '';
  FConditionsCount := 0;
end; //Clear

function TODSNumberFormatMapItem.TryToParse(const FNStr: string): boolean;
var
  s: string;
  i: integer;
  _isQuote: boolean;
  _isBracket: boolean;
  ch: char;  
  _isError: boolean;
  _raw: string; //raw string without brackets
  _tmp: string;

  procedure _ProcessOpenBracket();
  begin
    if (_isBracket) then
      _isError := true
    else
    begin  
      _raw := _raw + s;
      s := '';
      _isBracket := true;
    end;  
  end; //_ProcessOpenBracket
  
  procedure _ProcessCloseBracket();
  begin
    if (_isBracket) then
    begin
      //is it color?
      if (TryGetMapColorColor(Trim(UpperCase(s)), _tmp)) then
      begin
        FisColor := true;
        FColorStr :=  _tmp;      
      end
      else
      //is it condition?
      if (TryGetMapCondition(s, _tmp)) then
      begin
        FisCondition := true;
        FCondition := _tmp;
      end;

      //TODO: need add:
      //    calendar:
      //            [~buddhist]
      //            [~gengou]
      //            [~gregorian])
      //            [~hanja] [~hanja_yoil]
      //            [~hijri]
      //            [~jewish]
      //            [~ROC]
      //    NatNumX / DBNumX transliteration
      //    currency
      
      _isBracket := false;
      s := '';
    end
    else
      _isError := true  
  end; //_ProcessCloseBracket
  
  procedure _ProcessQuote();
  begin            
    if (not _isBracket) then
      _isQuote := not _isQuote;
  end; //_ProcessQuote

  function _FinalCheck(): boolean;
  begin
    Result := true;

    if (not _isQuote) then
    begin
      _raw := _raw + s;
      s := '';
    end;
    
    if (_isQuote and (not _isError) and (s <> '')) then
    begin
      _raw := _raw + s + '"';
      s := '';
      _ProcessQuote();
    end;

    //TODO: add checking for valid NF here
    
  end; //_FinalCheck
  
begin
  Result := false;
  Clear();
  _raw := '';

  _isError := false;
  _isQuote := false;
  _isBracket := false;
  s := '';

  for i := 1 to length(FNStr) do
  begin  
    ch := FNStr[i];

    if (ch = '"') then
      _ProcessQuote();
    
    if (_isQuote) then
      s := s + ch
    else  
      case (ch) of
        '[': _ProcessOpenBracket();
        ']': _ProcessCloseBracket();
        else
          s := s + ch;
      end; //case
  end; //for i

  if (_FinalCheck()) then    
    Result := not (_isError or _isQuote or _isBracket or (_raw = ''));  

  if (Result) then
    FNumberFormat := _raw;
end; //TryToParse

function TODSNumberFormatMapItem.AddCondition(const ACondition, AStyleName: string): boolean;
begin
  Result := FConditionsCount < 2;
  if (Result) then
  begin
    FConditionsArray[FConditionsCount][0] := ACondition;
    FConditionsArray[FConditionsCount][1] := AStyleName;
    inc(FConditionsCount);
  end;
end; //AddCondition

procedure TODSNumberFormatMapItem.PrepareCommonStyleAttributes(const xml: TZsspXMLWriterH; const AStyleName: string; isVolatile: boolean = false);
begin
  xml.Attributes.Clear();
  xml.Attributes.Add(ZETag_Attr_StyleName, AStyleName);
  if (isVolatile) then
    xml.Attributes.Add(ZETag_style_volatile, 'true');
end; //PrepareCommonStyleAttributes

procedure TODSNumberFormatMapItem.WriteNumberStyle(const xml: TZsspXMLWriterH;
                                                   const AStyleName: string;
                                                   const NumProperties: integer;
                                                   isVolatile: boolean = false);
var
  i: integer;
  _DecimalCount: integer;
  _IntDigitsCount: integer;
  _TotalDigitsCount: integer;
  _MinIntDigitsCount: integer;
  _CurrentPos: integer;
  _isFirstText: boolean;
  _firstText: string;
  _txt_len: integer;
  _numeratorDigitsCount: integer;
  _denomenatorDigitsCount: integer;
  _isFraction: boolean;
  _isSci: boolean;
  _exponentDigitsCount: integer;
  s: string;

  // <style:map style:condition="" style:apply-style-name=""/>
  procedure _WriteStyleMap(num: integer);
  begin
    xml.Attributes.Clear();
    xml.Attributes.Add(ZETag_style_condition, FConditionsArray[num][0]);
    xml.Attributes.Add(ZETag_style_apply_style_name, FConditionsArray[num][1]);
    xml.WriteEmptyTag(ZETag_style_map, true, true);
  end; //_WriteConditionItem

  //<style:text-properties />
  procedure _WriteTextProperties();
  begin
    if (isColor) then
    begin
      xml.Attributes.Clear();
      xml.Attributes.Add(ZETag_fo_color, ColorStr);
      xml.WriteEmptyTag(ZETag_style_text_properties, true, true);
    end;
  end; //_WriteTextProperties

  procedure _ParseFormat();
  var
    i: integer;
    _isQuote: boolean;
    ch: char;
    _isDecimal: boolean;

    //Check digit
    //INPUT
    //     isExtrazero: boolean - true = 0, false = #
    procedure _CheckDigit(isExtrazero: boolean);
    begin
      if (_isSci) then
      begin
        inc(_exponentDigitsCount);
        Exit;
      end;

      inc(_TotalDigitsCount);
      inc(_CurrentPos);

      if (_isDecimal) then
      begin
        inc(_DecimalCount);
      end
      else
      begin
        inc(_IntDigitsCount);
        if (isExtrazero) then
          inc(_MinIntDigitsCount);
      end;
    end; //_CheckDigit

    //Calc symbols "?" for fraction numerator and denominator
    //TODO: is it possible to use 0 or # in fraction as numerator or/and denominator?
    procedure _CheckFractionDigit();
    begin
      if (_isFraction) then
        inc(_denomenatorDigitsCount)
      else
        inc(_numeratorDigitsCount);
    end; //_CheckFractionDigit

    procedure _AddEmbebedText(isAdd: boolean);
    begin
      if (isAdd) then
      begin
        if ((_TotalDigitsCount > 0) or _isFirstText) then
        begin
          if (FEmbededTextCount >= FEmbededMaxCount) then
          begin
            inc(FEmbededMaxCount, 10);
            SetLength(FEmbededTextArray, FEmbededMaxCount);
          end;
          FEmbededTextArray[FEmbededTextCount].Txt := s;
          FEmbededTextArray[FEmbededTextCount].NumberPosition := _CurrentPos;
          inc(FEmbededTextCount);
        end
        else
        begin
          _isFirstText := true;
          _firstText := s;
        end;
        s := '';
      end;
    end; //_AddEmbebedText

    procedure _ProgressPercent();
    begin
      s := s + ch;
      if ((not _isFirstText) and (_TotalDigitsCount = 0) and (FEmbededTextCount = 0)) then
      begin
        _isFirstText := true;
        _firstText := s;
        s := '';
      end
      else
        if (i <> _txt_len) then
          _AddEmbebedText(true);
    end; //_ProgressPercent

  begin
    s := '';
    _isQuote := false;
    _isDecimal := false;
    i := 1;
    _txt_len := Length(FNumberFormat);
    while (i <= _txt_len) do
    begin
      ch := FNumberFormat[i];

      if ((ch = '\') and (not _isQuote)) then
      begin
        inc(i);
        if (i > _txt_len) then
          break;

        s := FNumberFormat[i];
        _AddEmbebedText(true);
        ch := #0;
      end;

      if (ch = '"') then
      begin
        _AddEmbebedText(_isQuote and (not _isDecimal));
        _isQuote := not _isQuote;
      end;

      if (_isQuote) then
      begin
        if (ch <> '"') then
          s := s + ch
      end
      else
        case (ch) of
          '0': _CheckDigit(true);
          '#': _CheckDigit(false);
          '.': _isDecimal := true;
          '?': _CheckFractionDigit();
          '/': _isFraction := true;
          'E', 'e': _isSci := true;
          '+':; //??
          '-':; //??
          ' ':;
          '%': _ProgressPercent();
        end;
      inc(i);
    end; //while i
  end; //_ParseFormat

  //<number:number > </number:number>
  procedure _WriteNumberMain();
  var
    i: integer;

    procedure _FillMainAttrib();
    begin
      xml.Attributes.Clear();

      if (_isFraction) then
        if ((_numeratorDigitsCount <= 0) or (_denomenatorDigitsCount <= 0)) then
          _isFraction := False;

      //TODO: it is trouble. For now ignore fraction and sci.
      if (_isFraction and _isSci) then
      begin
        _isFraction := false;
        _isSci := false;
      end;

      if (_isFraction) then
      begin
        xml.Attributes.Add(ZETag_number_min_numerator_digits, IntToStr(_numeratorDigitsCount));
        xml.Attributes.Add(ZETag_number_min_denominator_digits, IntToStr(_denomenatorDigitsCount));
      end
      else
      begin
        if (_DecimalCount > 0) then
          xml.Attributes.Add(ZETag_number_decimal_places, IntToStr(_DecimalCount));
        if (_isSci) then
          xml.Attributes.Add(ZETag_number_min_exponent_digits, IntToStr(_exponentDigitsCount));
      end;
      xml.Attributes.Add(ZETag_number_min_integer_digits, IntToStr(_MinIntDigitsCount));
    end; //_FillMainAttrib

    procedure _StartEmbededTextTag();
    begin
      //Empty tag for:
      //      number:fraction
      //      number:scientific-number

      if (_isFraction) then
        xml.WriteEmptyTag(ZETag_number_fraction, true, true)
      else
      if (_isSci) then
        xml.WriteEmptyTag(ZETag_number_scientific_number, true, true)
      else
        xml.WriteTagNode(ZETag_number_number, true, true, false);
    end; //_StartEmbededTextTag

    procedure _EndEmbededTextTag();
    begin
      if ((not _isFraction) and (not _isSci)) then
        xml.WriteEndTagNode(); //number:number
    end; //_EndEmbededTextTag

    //<number:number/> or
    //<number:fraction/> or
    //<number:scientific-number/>
    procedure _WriteEmptyNumberTag();
    begin
      if (_isFraction) then
        xml.WriteEmptyTag(ZETag_number_fraction, true, true)
      else
      if (_isSci) then
        xml.WriteEmptyTag(ZETag_number_scientific_number, true, true)
      else
        xml.WriteEmptyTag(ZETag_number_number, true, true);
    end; //_WriteEmptyNumberTag

  begin
    FEmbededTextCount := 0;
    _DecimalCount := 0;
    _CurrentPos := 0;
    _IntDigitsCount := 0;
    _TotalDigitsCount := 0;
    _MinIntDigitsCount := 0;
    _isFirstText := false;
    _firstText := '';
    _numeratorDigitsCount := 0 ;
    _denomenatorDigitsCount := 0;
    _isFraction := False;
    _isSci := False;
    _exponentDigitsCount := 0;

    _ParseFormat();

    if (_isFirstText) then
    begin
      xml.Attributes.Clear();
      xml.WriteTag(ZETag_number_text, _firstText, true, false, true);
    end;

    _FillMainAttrib();

    if (FEmbededTextCount > 0) then
    begin
      _StartEmbededTextTag();

      //TODO: Is it possible to use embeded text for fraction and scientific formats?
      if (not _isFraction) then
        for i := 0 to FEmbededTextCount - 1 do
        begin
          xml.Attributes.Clear();
          xml.Attributes.Add(ZETag_number_position, IntToStr(_IntDigitsCount - FEmbededTextArray[i].NumberPosition));
          xml.WriteTag(ZETag_number_embedded_text, FEmbededTextArray[i].Txt, true, false, true);
        end;

      _EndEmbededTextTag();
    end
    else
      _WriteEmptyNumberTag();

    if (s <> '') then
    begin
      xml.Attributes.Clear();
      xml.WriteTag(ZETag_number_text, s, true, false, true);
    end;
  end; //_WriteNumberMain

begin
  PrepareCommonStyleAttributes(xml, AStyleName, isVolatile);

  if (NumProperties = ZE_NUMFORMAT_NUM_IS_PERCENTAGE) then
    xml.WriteTagNode(ZETag_number_percentage_style, true, true, false)
  else
  if (NumProperties = ZE_NUMFORMAT_NUM_IS_CURRENCY) then
    xml.WriteTagNode(ZETag_number_currency_style, true, true, false)
  else
    xml.WriteTagNode(ZETag_number_number_style, true, true, false);

  _WriteTextProperties();

  _WriteNumberMain();

  for i := 0 to FConditionsCount - 1 do
    _WriteStyleMap(i);

  xml.WriteEndTagNode(); //number:number-style
end; //WriteNumberStyle

//Write number text style item (<number:text-style> </number:text-style>)
//INPUT
//     const xml: TZsspXMLWriterH - xml
//     const AStyleName: string   - style name
//           isVolatile: boolean  - is volatile? (for now - ignore)
procedure TODSNumberFormatMapItem.WriteTextStyle(const xml: TZsspXMLWriterH; const AStyleName: string; isVolatile: boolean = false);
var
  _isText: boolean;

begin
  _isText := false;
  PrepareCommonStyleAttributes(xml, AStyleName, isVolatile);

  xml.WriteTagNode(ZETag_number_text_style, true, true, false);

  xml.Attributes.Clear();
  FNumberFormatParser.BeginRead(FNumberFormat);

  while (FNumberFormatParser.ReadSymbol()) do
  begin
    case (FNumberFormatParser.ReadedSymbolType) of
      0:
        begin
          if (FNumberFormatParser.ReadedSymbol = '@') then
          begin
            _isText := true;
            xml.WriteEmptyTag(ZETag_number_text_content, true, false);
          end;
        end;
      2, 3:
        begin
          xml.WriteTag(ZETag_number_text, FNumberFormatParser.ReadedSymbol, true, false, true);
        end;
    end; //case
  end; //while

  if (not _isText) then
    xml.WriteEmptyTag(ZETag_number_text_content, true, false);

  xml.WriteEndTagNode(); //number:text-style
end; //WriteTextStyle

function TODSNumberFormatMapItem.WriteDateTimeStyle(const xml: TZsspXMLWriterH;
                                                    const AStyleName: string;
                                                    isVolatile: boolean = false): integer;

var
  s: string;
  _tagName: string;

  procedure _WriteYear(var item: TZDateTimeProcessItem);
  begin
    if (item.Len > 2) then
      xml.Attributes.Add(ZETag_number_style, ZETag_long);

    xml.WriteEmptyTag(ZETag_number_year, true, false);
  end; //_WriteYear

  procedure _WriteMonth(var item: TZDateTimeProcessItem);
  begin
    if ((item.Len >= 4) or (item.Len = 2)) then
      xml.Attributes.Add(ZETag_number_style, ZETag_long);

    if (item.Len >= 3) then
       xml.Attributes.Add(ZETag_number_textual, 'true');

    xml.WriteEmptyTag(ZETag_number_month, true, false);
  end; //_WriteMonth

  procedure _WriteDay(var item: TZDateTimeProcessItem);
  begin
    if (item.Len > 2) then
      s := ZETag_number_day_of_week
    else
      s := ZETag_number_day;

    if ((item.Len >= 4) or (item.Len = 2)) then
      xml.Attributes.Add(ZETag_number_style, ZETag_long);

    xml.WriteEmptyTag(s, true, false);
  end; //_WriteDay

  procedure _WriteHour(var item: TZDateTimeProcessItem);
  begin
    if (item.Len >= 2) then
      xml.Attributes.Add(ZETag_number_style, ZETag_long);

    xml.WriteEmptyTag(ZETag_number_hours, true, false);
  end; //_WriteHour

  procedure _WriteMinute(var item: TZDateTimeProcessItem);
  begin
    if (item.Len >= 2) then
      xml.Attributes.Add(ZETag_number_style, ZETag_long);

    xml.WriteEmptyTag(ZETag_number_minutes, true, false);
  end; //_WriteMinute

  procedure _WriteSecond(var item: TZDateTimeProcessItem);
  begin
    if (item.Len >= 2) then
      xml.Attributes.Add(ZETag_number_style, ZETag_long);

    if (item.Settings > 0) then
      xml.Attributes.Add(ZETag_number_decimal_places, IntToStr(item.Settings));

    xml.WriteEmptyTag(ZETag_number_seconds, true, false);
  end; //_WriteSecond

  procedure _WriteWeek(var item: TZDateTimeProcessItem);
  begin
    xml.WriteEmptyTag(ZETag_number_week_of_year, true, false);
  end; //_WriteWeek

  procedure _WriteQuarter(var item: TZDateTimeProcessItem);
  begin
    if (item.Len >= 2) then
      xml.Attributes.Add(ZETag_number_style, ZETag_long);

    xml.WriteEmptyTag(ZETag_number_quarter, true, false);
  end; //_WriteQuarter

  procedure _WriteEraYear(var item: TZDateTimeProcessItem);
  begin
    //TODO
    (*
    if (item.Len >= 2) then
      xml.Attributes.Add(ZETag_number_style, ZETag_long);

    xml.WriteEmptyTag(ZETag_number_quarter, true, false);
    *)
  end; //_WriteEraYear

  procedure _WriteEraJap(var item: TZDateTimeProcessItem);
  begin
    if (item.Len >= 2) then
      xml.Attributes.Add(ZETag_number_style, ZETag_long);

    xml.WriteEmptyTag(ZETag_number_era, true, false);
  end; //_WriteEraJap

  procedure _WriteItems();
  var
    i: integer;

  begin
    for i := 0 to FDateTimeODSFormatParser.FCount - 1 do
    begin
      xml.Attributes.Clear();
      case (FDateTimeODSFormatParser.FItems[i].ItemType) of
        ZE_DATETIME_ITEM_TEXT:
                              xml.WriteTag(ZETag_number_text, FDateTimeODSFormatParser.FItems[i].TextValue, true, false, true);

        ZE_DATETIME_ITEM_YEAR:
                              _WriteYear(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_MONTH:
                              _WriteMonth(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_DAY:
                              _WriteDay(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_HOUR:
                              _WriteHour(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_MINUTE:
                              _WriteMinute(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_SECOND:
                              _WriteSecond(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_WEEK:
                              _WriteWeek(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_QUARTER:
                              _WriteQuarter(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_ERA_JAP:
                              _WriteEraJap(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_ERA_YEAR:
                              _WriteEraYear(FDateTimeODSFormatParser.FItems[i]);

        ZE_DATETIME_ITEM_AMPM:
                              xml.WriteEmptyTag(ZETag_number_am_pm, true, false);
      end; //case
    end; //for i
  end; //_WriteItems

  function _GetAdditionalProperties(): integer;
  var
    i: integer;

  begin
    for i := 0 to FDateTimeODSFormatParser.FCount - 1 do
      case (FDateTimeODSFormatParser.FItems[i].ItemType) of
        ZE_DATETIME_ITEM_YEAR,
        ZE_DATETIME_ITEM_MONTH,
        ZE_DATETIME_ITEM_DAY,
        ZE_DATETIME_ITEM_WEEK,
        ZE_DATETIME_ITEM_QUARTER,
        ZE_DATETIME_ITEM_ERA_JAP,
        ZE_DATETIME_ITEM_ERA_YEAR:
          begin
            Result := 0;
            exit;
          end;
      end; //case
    Result := ZE_NUMFORMAT_DATE_IS_ONLY_TIME;
  end; //_GetAdditionalProperties

begin
  Result := 0;
  FDateTimeODSFormatParser.TryToParseDateFormat(FNumberFormat, FNumberFormatParser);
  FDateTimeODSFormatParser.DeleteRepeatedItems();

  if (FDateTimeODSFormatParser.GetValidCount() > 0) then
  begin
    Result := _GetAdditionalProperties();

    if (Result = 0) then
      _tagName := ZETag_number_date_style
    else
      _tagName := ZETag_number_time_style;

    PrepareCommonStyleAttributes(xml, AStyleName, isVolatile);
    xml.WriteTagNode(_tagName, true, true, false);

    _WriteItems();

    xml.WriteEndTagNode(); //number:date-style / number:time-style
  end;
end; //WriteDateTimeStyle

end.
