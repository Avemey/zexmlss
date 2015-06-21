//****************************************************************
// Routines for translate number formats from ods/xlsx/excel xml
//  to internal (ods like) number format and back etc.
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2015.06.06
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
  ZE_NUMFORMAT_IS_UNKNOWN   = 0;
  ZE_NUMFORMAT_IS_NUMBER    = 1;
  ZE_NUMFORMAT_IS_DATETIME  = 2;
  ZE_NUMFORMAT_IS_STRING    = 4;

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

  ZETag_style_text_properties = 'style:text-properties';

  ZETag_Attr_StyleName        = 'style:name';
  ZETag_long                  = 'long';

type
  TZODSNumberItemOptions = record
    isColor: boolean;
    ColorStr: string;
    StyleType: byte;
  end;

  TZEODSNumberFormatReader = class
  private
    FItems: array of array [0..1] of string;  //index 0 - format num
                                              //index 1 - format
    FItemsOptions: array of TZODSNumberItemOptions;
    FCount: integer;
    FCountMax: integer;
  protected
    procedure AddItem();
    procedure ReadNumberFormatCommon(const xml: TZsspXMLReaderH; const NumberFormatTag: string);
  public
    constructor Create();
    destructor Destroy();
    procedure ReadDateFormat(const xml: TZsspXMLReaderH);
    procedure ReadNumberFormat(const xml: TZsspXMLReaderH);
    function TryGetFormatStrByNum(const DataStyleName: string; out retFormatStr): boolean;
    property Count: integer read FCount;
  end;

function GetXlsxNumberFormatType(const FormatStr: string): integer;
function ConvertFormatNativeToXlsx(const FormatNative: string): string;
function ConvertFormatXlsxToNative(const FormatXlsx: string): string;
function TryXlsxTimeToDateTime(const XlsxDateTime: string; out retDateTime: TDateTime; is1904: boolean = false): boolean;

implementation

uses
  zesavecommon,
  StrUtils            {IfThen}
  ;

  const
    const_format_type_number    = 0;
    const_format_type_datetime  = 1;
    const_format_type_boolean   = 2;

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

// Try to get xlsx number format type by string (very simplistic)
//INPUT
//  const FormatStr: string - format ("YYYY.MM.DD" etc)
//RETURN
//      integer - 0 - unknown
//                1 - number
//                2 - datetime
//                4 - string
function GetXlsxNumberFormatType(const FormatStr: string): integer;
var
  i, l: integer;
  s: string;
  ch: char;
  _isQuote: boolean;

begin
  Result := ZE_NUMFORMAT_IS_UNKNOWN;

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
            exit;
          end;
        '@':
          begin
            Result := ZE_NUMFORMAT_IS_STRING;
            exit;
          end;
        'H', 'h', 'S', 's', 'm', 'M', 'd', 'D', 'Y', 'y', ':':
          begin
            Result := ZE_NUMFORMAT_IS_DATETIME;
            exit;
          end;
      end;
  end; //for i
end; //GetXlsxNumberFormatType

function ConvertFormatNativeToXlsx(const FormatNative: string): string;
begin
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


////::::::::::::: TZEODSNumberFormatReader :::::::::::::::::////

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
  SetLength(FItems, FCountMax);
  SetLength(FItemsOptions, FCountMax);
  for i := 0 to FCountMax - 1 do
  begin
    FItemsOptions[i].isColor := false;
    FItemsOptions[i].ColorStr := '';
  end;
end;

destructor TZEODSNumberFormatReader.Destroy();
begin
  SetLength(FItems, 0);
  SetLength(FItemsOptions, 0);
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
    end
    else
    //Quarter
    if (xml.TagName = ZETag_number_quarter) then
    begin
    end;

    if (xml.Eof()) then
      break;
  end; //while
  FItems[num][1] := _result;
end; //ReadDateFormat

//Read number style: <number:number-style> .. </number:number-style>
procedure TZEODSNumberFormatReader.ReadNumberFormat(const xml: TZsspXMLReaderH);
begin
  ReadNumberFormatCommon(xml, ZETag_number_number_style);
end;

//Read Number/currency/percentage number format style
//INPUT
//  const xml: TZsspXMLReaderH    - xml
//  const NumberFormatTag: string - tag name
procedure TZEODSNumberFormatReader.ReadNumberFormatCommon(const xml: TZsspXMLReaderH;
                                                          const NumberFormatTag: string);
var
  num: integer;
  s, _result: string;

begin
  num := FCount;
  AddItem();
  FItems[num][0] := xml.Attributes[ZETag_Attr_StyleName];
  FItemsOptions[num].StyleType := const_format_type_number;
  _result := '';
  while ((xml.TagType <> 6) or (xml.TagName <> NumberFormatTag)) do
  begin
    xml.ReadTag();

    if (xml.TagName = ZETag_number_number) then
    begin
      s := xml.Attributes[ZETag_number_decimal_places];

    {
    The <number:number> element specifies the display formatting properties for a decimal number.

The <number:number> element has the following attributes:
number:decimal-places
number:decimal-replacement
number:display-factor
number:grouping
number:min-integer-digits
The <number:number> element has the following child element:
number:embedded-text

  ZETag_number_fraction       = 'number:fraction';
  ZETag_number_scientific_number = 'number:scientific-number';
  ZETag_number_embedded_text  = 'number:embedded-text';
  ZETag_number_number         = 'number:number';
  ZETag_number_decimal_places = 'number:decimal-places';
  ZETag_number_decimal_replacement = 'number:decimal-replacement';
  ZETag_number_display_factor = 'number:display-factor';
  ZETag_number_grouping       = 'number:grouping';
  ZETag_number_min_integer_digits = 'number:min-integer-digits';
    }

    end
    else
    if ((xml.TagName = ZETag_number_text) and (xml.TagType = 6)) then
      _result := _result + xml.TextBeforeTag;



    {
    ZETag_number_fraction       = 'number:fraction';
    ZETag_number_scientific_number = 'number:scientific-number';
    ZETag_number_embedded_text  = 'number:embedded-text';
    ZETag_number_number         = 'number:number';
    }

    if (xml.Eof()) then
      break;
  end; //while

  {
  <number:fraction> 16.27.6, <number:number> 16.27.3, <number:scientific-number> 16.27.5,
  <number:text> 16.27.26, <style:map> 16.3 and <style:text-properties> 16.27.28.
  }

end;

function TZEODSNumberFormatReader.TryGetFormatStrByNum(const DataStyleName: string;
                                                       out retFormatStr): boolean;
begin
  Result := false;
end;

end.
