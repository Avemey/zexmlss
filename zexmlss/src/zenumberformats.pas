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
  SysUtils;

const
  ZE_NUMFORMAT_IS_UNKNOWN   = 0;
  ZE_NUMFORMAT_IS_NUMBER    = 1;
  ZE_NUMFORMAT_IS_DATETIME  = 2;
  ZE_NUMFORMAT_IS_STRING    = 4;

function GetXlsxNumberFormatType(const FormatStr: string): integer;
function ConvertFormatXlsxToNative(const FormatXlsx: string): string;
function TryXlsxTimeToDateTime(const XlsxDateTime: string; out retDateTime: TDateTime; is1904: boolean = false): boolean;

implementation

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

function ConvertFormatXlsxToNative(const FormatXlsx: string): string;
var
  i, l: integer;
  _isQuote: boolean;
  _isBracket: boolean;
  s: string;
  ch: char;
  _semicolonCount: integer;

  procedure _ProcessOpenBracket();
  begin
    if (_isQuote) then
      s := s + ch
    else
      if (not _isBracket) then
        _isBracket := true;
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
  end;

begin
  Result := '';
  _isQuote := false;
  _isBracket := false;
  s := '';
  _semicolonCount := 0;

  l := length(FormatXlsx);

  for i := 1 to l do
  begin
    ch := FormatXlsx[i];

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
          end;
      end; //case ch
  end; //for i

  _ProcessQuote(false);
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

end.
