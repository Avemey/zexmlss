//****************************************************************
// Routines for translate number formats from ods/xlsx/excel xml
//  to internal (ods like) number format and back etc.
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2015.11.21
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

  //DateStyles
  ZETag_number_date_style     = 'number:date-style';
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

  ZETag_style_text_properties = 'style:text-properties';
  ZETag_style_map             = 'style:map';
  ZETag_fo_color              = 'fo:color';

  ZETag_Attr_StyleName        = 'style:name';
  ZETag_style_condition       = 'style:condition';
  ZETag_style_apply_style_name = 'style:apply-style-name';
  ZETag_long                  = 'long';
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
  protected
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
    //           isVolatile: boolean  - is volatile?
    procedure WriteNumberStyle(const xml: TZsspXMLWriterH; const AStyleName: string; isVolatile: boolean = false);

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
  public
    constructor Create();
    destructor Destroy(); override;
    procedure ReadKnownNumberFormat(const xml: TZsspXMLReaderH);
    procedure ReadDateFormat(const xml: TZsspXMLReaderH);
    procedure ReadNumberFormat(const xml: TZsspXMLReaderH);
    procedure ReadCurrencyFormat(const xml: TZsspXMLReaderH);
    function TryGetFormatStrByNum(const DataStyleName: string; out retFormatStr: string): boolean;
    property Count: integer read FCount;
  end;

  TZEODSNumberFormatWriterItem = record
    StyleIndex: integer;
    NumberFormatName: string;
    NUmberFormat: string;
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

  protected
    function TryAddNFItem(const NFStr: string): boolean;
    function SeparateNFItems(const NFStr: string): integer;
  public
    constructor Create(const AMaxCount: integer);
    destructor Destroy(); override;
    function TryGetNumberFormatName(StyleID: integer; out NumberFormatName: string): boolean;
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
function ConvertFormatNativeToXlsx(const FormatNative: string): string;
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
//INPUT
//      AStartPos: integer              - start position
//      ALen: integer                   - string length
//  const AStr: string                  - string
//  const SymbolsArr: array of string   - searching symbols
//RETURN
//      boolean - true - one of symbols was found in string after AStartPos
function IsHaveSymbolsAfterPos(AStartPos: integer; ALen: integer; const AStr: string; const SymbolsArr: array of string): boolean;
var
  i, j: integer;
  _IsQuote: boolean;
  ch: char;
  _max, _min: integer;

begin
  Result := false;
  _IsQuote := false;
  _min := Low(SymbolsArr);
  _max := High(SymbolsArr);
  for i := AStartPos + 1 to Alen do
  begin
    ch := AStr[i];
    if (ch = '"') then
      _IsQuote := not _IsQuote;

    if (not _IsQuote) then
      for j := _min to _max do
        if (ch = SymbolsArr[j]) then
        begin
          Result := true;
          exit;
        end;
  end; //for i
end; //IsHaveSymbolsAfterPos

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
  s: string;
  ch: char;
  _isQuote: boolean;
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

  s := '';
  l := length(FormatStr);
  for i := 1 to l do
  begin
    ch := FormatStr[i];

    if (ch = '"') then
      _isQuote := not _isQuote;

    if (not _isQuote) then
      case (ch) of
        '0', '#', 'E', 'e', '%', '?':
          begin
            Result := ZE_NUMFORMAT_IS_NUMBER;

            if (IsHaveSymbolsAfterPos(i - 1, l, FormatStr, ['e', 'E'])) then
              Result := Result or ZE_NUMFORMAT_NUM_IS_SCIENTIFIC
            else
            if (IsHaveSymbolsAfterPos(i - 1, l, FormatStr, ['%'])) then
              Result := Result or ZE_NUMFORMAT_NUM_IS_PERCENTAGE
            else
            if (IsHaveSymbolsAfterPos(i - 1, l, FormatStr, ['/']) or _isFraction) then
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
  s: string;
  ch, _prev: char;
  _isQuote: boolean;

begin
  Result := ZE_NUMFORMAT_IS_UNKNOWN;

  _isQuote := false;
  _prev := #0;

  s := '';
  l := length(FormatStr);
  for i := 1 to l do
  begin
    ch := FormatStr[i];

    if (ch = '"') then
      _isQuote := not _isQuote;

    if (not _isQuote) then
      case (ch) of
        '0', '#', '%', '?':
          begin
            Result := ZE_NUMFORMAT_IS_NUMBER;

            if (IsHaveSymbolsAfterPos(i - 1, l, FormatStr, ['e', 'E'])) then
              Result := Result or ZE_NUMFORMAT_NUM_IS_SCIENTIFIC
            else
            if (IsHaveSymbolsAfterPos(i - 1, l, FormatStr, ['%'])) then
              Result := Result or ZE_NUMFORMAT_NUM_IS_PERCENTAGE
            else
            if (IsHaveSymbolsAfterPos(i - 1, l, FormatStr, ['/'])) then
              Result := Result or ZE_NUMFORMAT_NUM_IS_FRACTION;

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
      end;
    _prev := ch;
  end; //for i
end; //GetNativeNumberFormatType

function ConvertFormatNativeToXlsx(const FormatNative: string): string;
begin
  //TODO:
  Result := FormatNative;
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
      if (TryStrToFloat('0' + FormatSettings.DecimalSeparator + s2, t)) then
        retDateTime := retDateTime + t;

    Result := true;
  end;
end; //TryXlsxTimeToDateTime

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
      retColor := ZE_MAP_CONDITIONAL_COLORS[i][1];
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

//Read date format: <number:date-style>.. </number:date-style>
procedure TZEODSNumberFormatReader.ReadDateFormat(const xml: TZsspXMLReaderH);
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
  num := FCount;
  AddItem();
  FItems[num][0] := xml.Attributes[ZETag_Attr_StyleName];
  FItemsOptions[num].StyleType := ZE_NUMFORMAT_IS_DATETIME;
  _result := '';
  while ((xml.TagType <> 6) or (xml.TagName <> ZETag_number_date_style)) do
  begin
    xml.ReadTag();

    //Day
    if ((xml.TagName = ZETag_number_day) and (xml.TagType and 4 = 4)) then
      _result := _result + CheckIsLong('NN', 'DD')
    else
    //Text
    if ((xml.TagName = ZETag_number_text) and (xml.TagType = 6)) then
      _result := _result + xml.TextBeforeTag
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
    end
    else
    //Era
    if (xml.TagName = ZETag_number_era) then
    begin
      //Attr: number:calendar
      //      number:style
    end
    else
    //Quarter
    if (xml.TagName = ZETag_number_quarter) then
    begin
      //Attr: number:calendar
      //      number:style
    end
    else
    //Day of week
    if (xml.TagName = ZETag_number_day_of_week) then
    begin
      //Attr: number:calendar
      //      number:style
    end
    else
    //Week of year
    if (xml.TagName = ZETag_number_week_of_year) then
    begin
      //Attr: number:calendar
    end;

    if (xml.Eof()) then
      break;
  end; //while
  FItems[num][1] := _result;
end; //ReadDateFormat

//Read known numbers formats (date/number/percentage etc)
procedure TZEODSNumberFormatReader.ReadKnownNumberFormat(const xml: TZsspXMLReaderH);
begin
  if (xml.TagName = ZETag_number_number_style) then
    ReadNumberFormat(xml)
  else
  if (xml.TagName = ZETag_number_date_style) then
    ReadDateFormat(xml)
  else
  if (xml.TagName = ZETag_number_currency_style) then
    ReadCurrencyFormat(xml);
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
  num := FCount;
  AddItem();
  FItems[num][0] := xml.Attributes[ZETag_Attr_StyleName];
  FItemsOptions[num].StyleType := ZE_NUMFORMAT_IS_NUMBER or sub_number_type;
  _result := '';
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
    if ((xml.TagName = ZETag_number_text) and (xml.TagType = 6)) then
      _result := _result + '"' + ZEReplaceEntity(xml.TextBeforeTag) + '"'
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
  _tmp: string;
  ch: char;

begin
  b := true;
  s := '';

  _tmp := NFStr;

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
  _tmp: string;
  _nfType: integer;
  _nfName: string;
  l: integer;

  function _WriteCurrency(): boolean;
  begin
    Result := false;
  end;

  function _WritePercentage(): boolean;
  begin
    Result := false;
  end;

  function _WriteNumberNumber(): boolean;
  var
    b: boolean;

  begin
    Result := false;
    //  NumberFormat = "part1;part2;part3"
    //    part1 - for numbers > 0 (or for condition1)
    //    part2 - for numbers < 0 (or for condition2)
    //    part3 - for 0 (or for other numbers if used condition1 and condition2)
    //  partX = [condition][color]number_format

    if (SeparateNFItems(ANumberFormat) > 0) then
    begin
      //if (FNFItemsCount)
    end;
  end; //_WriteNumberNumber

  function _WriteDateTime(): boolean;
  begin
    Result := false;
  end;

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
    end;
  end; //_WriteNumberStyle

  procedure _AddItem(const NFName: string);
  begin
    if (FCount >= FCountMax) then
    begin
      inc(FCountMax, 10);
      SetLength(FItems, FCountMax);
    end;

    FItems[FCount].StyleIndex := StyleID;
    FItems[FCount].NumberFormatName := NFName;
    FItems[FCount].NUmberFormat := ANumberFormat;

    inc(FCount);
  end;

  function _CheckIsDuplicate(): boolean;
  var
    i: integer;

  begin
    Result := false;
    for i := 0 to FCount - 1 do
      if (FItems[i].NUmberFormat = ANumberFormat) then
      begin
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
end;

destructor TODSNumberFormatMapItem.Destroy();
begin
  SetLength(FEmbededTextArray, 0);
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
      
      _isBracket := true;
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

procedure TODSNumberFormatMapItem.WriteNumberStyle(const xml: TZsspXMLWriterH;
                                                   const AStyleName: string;
                                                   isVolatile: boolean = false);
var
  i: integer;
  _DecimalCount: integer;
  _CurrentPos: integer;

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
    s: string;
    _isQuote: boolean;
    ch: char;

  begin
    s := '';
    _isQuote := false;
    for i := 1 to Length(FNumberFormat) do
    begin
      ch := FNumberFormat[i];

      if (ch = '"') then
      begin
        if (_isQuote) then
        begin
          if (FEmbededTextCount + 1 >= FEmbededMaxCount) then
          begin
            //inc();
          end;
        end;
        _isQuote := not _isQuote;
      end;

      if (_isQuote) then
        s := s + ch
      else
        case (ch) of
          '0':;
          '#':;
          '.':;
          '?':;
          '/':;
          ' ':;
        end;
    end; //for i
  end; //_ParseFormat

  //<number:number > </number:number>
  procedure _WriteNumberMain();
  begin
    FEmbededTextCount := 0;
    _ParseFormat();
  end; //_WriteNumberMain

begin
  xml.Attributes.Clear();
  xml.Attributes.Add(ZETag_Attr_StyleName, AStyleName);
  if (isVolatile) then
     xml.Attributes.Add(ZETag_style_volatile, 'true');

  xml.WriteTagNode(ZETag_number_number_style, true, true, false);

  _WriteTextProperties();

  _WriteNumberMain();

  for i := 0 to FConditionsCount - 1 do
    _WriteStyleMap(i);

  xml.WriteEndTagNode(); //number:number-style
end; //WriteNumberStyle

end.
