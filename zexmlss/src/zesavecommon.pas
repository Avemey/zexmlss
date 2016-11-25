//****************************************************************
// Common routines for saving and loading data
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2016.09.10
//----------------------------------------------------------------
// This software is provided "as-is", without any express or implied warranty.
// In no event will the authors be held liable for any damages arising from the
// use of this software.

unit zesavecommon;

interface

{$I zexml.inc}
{$I compver.inc}

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

uses
  SysUtils, Types, zexmlss, zsspxml;

const
  ZE_MMinInch: real = 25.4;

//type
//  TZESaveIntArray = array of integer; // Since Delphi 7 and FPC 2005 - TIntegerDynArray
//  TZESaveStrArray = array of string;  // Since Delphi 7 and FPC 2005 - TStringDynArray

//ѕопытка преобразовать строку в число
function ZEIsTryStrToFloat(const st: string; out retValue: double): boolean;
function ZETryStrToFloat(const st: string; valueIfError: double = 0): double; overload;
function ZETryStrToFloat(const st: string; out isOk: boolean; valueIfError: double = 0): double; overload;

//ѕопытка преобразовать строку в boolean
function ZETryStrToBoolean(const st: string; valueIfError: boolean = false): boolean;

//замен€ет все зап€тые на точки
function ZEFloatSeparator(st: string): string;

//BOM<?xml version="1.0" encoding="CodePageName"?>
procedure ZEWriteHeaderCommon(xml: TZsspXMLWriterH; const CodePageName: string; const BOM: ansistring);

//ѕровер€ет заголовки страниц, при необходимости корректирует
function ZECheckTablesTitle(var XMLSS: TZEXMLSS; const SheetsNumbers:array of integer;
                            const SheetsNames: array of string; out _pages: TIntegerDynArray;
                            out _names: TStringDynArray; out retCount: integer): boolean;

//ќчищает массивы
procedure ZESClearArrays(var _pages: TIntegerDynArray;  var _names: TStringDynArray);

//ѕереводит строку в boolean
function ZEStrToBoolean(const val: string): boolean;

//ѕровер€ет, есть ли папка така€. ≈сли путь не заканчиваетс€ на разделитель директории - добавл€ет в конец.
function ZE_CheckDirExist(var DirName: string): boolean;

//«амен€ет в строке последовательности на спецсимволы
function ZEReplaceEntity(const st: string): string;

// despite formal angle datatype declaration in default "range check off" mode
//    it can be anywhere -32K to +32K
// This fn brings it back into -90 .. +90 range
function ZENormalizeAngle90(const value: TZCellTextRotate): integer;

// despite formal angle datatype declaration in default "range check off" mode
//    it can be anywhere -32K to +32K
// This fn brings it back into 0 .. +179 range
function ZENormalizeAngle180(const value: TZCellTextRotate): integer;

// single place to update version et all
function ZELibraryName: string;

//Trying to convert string like "n%" to integer
function TryStrToIntPercent(s: string; out Value: integer): boolean;

//Get DateTime from XML scheme 2 duration string
//see https://www.w3.org/TR/xmlschema-2/#duration
//  example: PT537199H55M05.123000219S = 1961.04.12 7:55:05.123
//INPUT
//     const APTTime: string - string in PT format
//RETURN
//      TDateTime - calculated datetime
function ZEPTDateDurationToDateTime(const APTTime: string): TDateTime;

//Get XML schema 2 duration string from TDateTime
//see https://www.w3.org/TR/xmlschema-2/#duration
//  example: PT537199H55M05.123000219S = 1961.04.12 7:55:05.123
//INPUT
//     const ADateTime: TDateTime - datetime
//RETURN
//      string
function ZEDateTimeToPTDurationStr(const ADateTime: TDateTime): string;

const ZELibraryVersion = '0.0.15';
      ZELibraryFork = '';//'Arioch';  // or empty str   // URL ?

implementation

uses
  dateutils //IncHour etc
  ;

function ZELibraryName: string;
begin   // todo: compiler version and date ?
  Result := 'ZEXMLSSlib/' + ZELibraryVersion;
  {$WARNINGS OFF}
  if ZELibraryFork <> '' then Result := Result + '@' + ZELibraryFork;
  {$WARNINGS ON}

  Result := Result + '$' +
    {$IFDEF FPC}
     'FPC';
    {$ELSE}
    'DELPHI_or_CBUILDER';
    {$ENDIF}
  Result := Result + ' ZEXMLSS/' + {$I git_hash.inc};
end;

// despite formal angle datatype declaration in default "range check off" mode
//    it can be anywhere -32K to +32K
// This fn brings it back into -90 .. +90 range for Excel XML
function ZENormalizeAngle90(const value: TZCellTextRotate): integer;
var Neg: boolean; A: integer;
begin
   if (value >= -90) and (value <= +90)
      then Result := value
   else begin                             (* Special values: 270; 450; -450; 180; -180; 135 *)
      Neg := Value < 0;                             (*  F, F, T, F, T, F *)
      A := Abs(value) mod 360;      // 0..359       (*  270, 90, 90, 180, 180, 135  *)
      if A > 180 then A := A - 360; // -179..+180   (*  -90, 90, 90, 180, 180, 135 *)
      if A < 0 then begin
         Neg := not Neg;                            (*  T,  -"- F, T, F, T, F  *)
         A := - A;                  // 0..180       (*  90, -"- 90, 90, 180, 180, 135 *)
      end;
      if A > 90 then A := A - 180; // 91..180 -> -89..0 (* 90, 90, 90, 0, 0, -45 *)
      Result := A;
      If Neg then Result := - Result;               (* -90, +90, -90, 0, 0, -45 *)
   end;
end;

// despite formal angle datatype declaration in default "range check off" mode
//    it can be anywhere -32K to +32K
// This fn brings it back into 0 .. +180 range
function ZENormalizeAngle180(const value: TZCellTextRotate): integer;
begin
  Result := ZENormalizeAngle90(value);
  If Result < 0 then Result := 90 - Result;
end;


//«амен€ет в строке последовательности на спецсимволы
//INPUT
//  const st: string - вход€ща€ строка
//RETURN
//      string - обработанна€ строка
function ZEReplaceEntity(const st: string): string;
var
  s, s1: string;
  i: integer;
  isAmp: boolean;
  ch: char;

  procedure _checkS();
  begin
    s1 := UpperCase(s);
    if (s1 = '&GT;') then
      s := '>'
    else
    if (s1 = '&LT;') then
      s := '<'
    else
    if (s1 = '&AMP;') then
      s := '&'
    else
    if (s1 = '&APOS;') then
      s := ''''
    else
    if (s1 = '&QUOT;') then
      s := '"';
  end; //_checkS

begin
  s := '';
  result := '';
//  n := length(st);
//  i := 1;
  isAmp := false;
  for i := 1 to length(st) do
  begin
    ch := st[i];
    case ch of
      '&':
          begin
            if (isAmp) then
            begin
              result := result + s;
              s := ch;
            end else
            begin
              isAmp := true;
              s := ch;
            end;
          end;
      ';':
          begin
            if (isAmp) then
            begin
              s := s + ch;
              _checkS();
              result := result + s;
              s := '';
              isAmp := false;
            end else
            begin
              result := result + s + ch;
              s := '';
            end;
          end;
      else
        if (isAmp) then
          s := s + ch
        else
          result := result + ch;
    end; //case
  end; //for
  if (s > '') then
  begin
    _checkS();
    result := result + s;
  end;
end; //ZEReplaceEntity


//ѕровер€ет, есть ли папка така€. ≈сли путь не заканчиваетс€ на разделитель директории - добавл€ет в конец.
//INPUT
//  var DirName: string - путь
//RETURN
//      boolean - true - OK
function ZE_CheckDirExist(var DirName: string): boolean;
var
  n: integer;

begin
  result := true;
  n := length(DirName);
  if (n > 0) then
  begin
    if (DirName[n] <> PathDelim) then
      DirName := DirName + PathDelim;
  end;

  if (not DirectoryExists(DirName)) then
    result := false;
end; //__CheckDirExist

//ѕереводит строку в boolean
//INPUT
//  const val: string - переводима€ строка
function ZEStrToBoolean(const val: string): boolean;
begin
  if (val = '1' ) or (UpperCase(val)='TRUE')  then
    result := true
  else
    result := false;
end;

//ѕопытка преобразовать строку в boolean
//  const st: string        - строка дл€ распознавани€
//    valueIfError: boolean - значение, которое подставл€етс€ при ошибке преобразовани€
function ZETryStrToBoolean(const st: string; valueIfError: boolean = false): boolean;
begin
  result := valueIfError;
  if (st > '') then
  begin
    {$IFDEF DELPHI_UNICODE}
    if (CharInSet(st[1], ['T', 't', '1', '-'])) then
      result := true
    else if (CharInSet(st[1], ['F', 'f', '0'])) then
    {$ELSE}
    if (st[1] in ['T', 't', '1', '-']) then
      result := true
    else if (st[1] in ['F', 'f', '0']) then
    {$ENDIF}
      result := false
    else
      result := valueIfError;
  end;
end; //ZETryStrToBoolean

function ZEIsTryStrToFloat(const st: string; out retValue: double): boolean;
begin
  retValue := ZETryStrToFloat(st, Result);
end;

//ѕопытка преобразовать строку в число
//INPUT
//  const st: string        - строка
//  out isOk: boolean       - если true - ошибки небыло
//    valueIfError: double  - значение, которое подставл€етс€ при ошибке преобразовани€
function ZETryStrToFloat(const st: string; out isOk: boolean; valueIfError: double = 0): double;
var
  s: string;
  i: integer;

begin
  result := 0;
  isOk := true;
  if (length(trim(st)) <> 0) then
  begin
    s := '';
    for i := 1 to length(st) do
      {$IFDEF DELPHI_UNICODE}
      if (CharInSet(st[i], ['.', ','])) then
      {$ELSE}
      if (st[i] in ['.', ',']) then
      {$ENDIF}
        {$IFDEF Z_USE_FORMAT_SETTINGS}
        s := s + FormatSettings.DecimalSeparator
        {$ELSE}
        s := s + DecimalSeparator
        {$ENDIF}
      else if (st[i] <> ' ') then
        s := s + st[i];

    isOk := TryStrToFloat(s, result);
    if (not isOk) then
      result := valueIfError;
  end;
end; //ZETryStrToFloat

//ѕопытка преобразовать строку в число
//INPUT
//  const st: string        - строка
//    valueIfError: double  - значение, которое подставл€етс€ при ошибке преобразовани€
function ZETryStrToFloat(const st: string; valueIfError: double = 0): double;
var
  s: string;
  i: integer;

begin
  result := 0;
  if (trim(st) <> '') then
  begin
    s := '';
    for i := 1 to length(st) do
      {$IFDEF DELPHI_UNICODE}
      if (CharInSet(st[i], ['.', ','])) then
      {$ELSE}
      if (st[i] in ['.', ',']) then
      {$ENDIF}
        {$IFDEF Z_USE_FORMAT_SETTINGS}
        s := s + FormatSettings.DecimalSeparator
        {$ELSE}
        s := s + DecimalSeparator
        {$ENDIF}
      else if (st[i] <> ' ') then
        s := s + st[i];
      try
        result := StrToFloat(s);
      except
        result := valueIfError;
      end;
  end;
end; //ZETryStrToFloat

//замен€ет все зап€тые на точки
function ZEFloatSeparator(st: string): string;
var
  k: integer;

begin
  result := '';
  for k := 1 to length(st) do
    if (st[k] = ',') then
      result := result + '.'
    else
      result := result + st[k];
  {
  result := st;
  k := pos(',', result);
  while k <> 0 do
  begin
    result[k] := '.';
    k := pos(',', result);
  end;
  }
end;

//BOM<?xml version="1.0" encoding="CodePageName"?>
//INPUT
//    xml: TZsspXMLWriterH      - писатель
//  const CodePageName: string  - им€ кодировки
//  const BOM: ansistring       - BOM
procedure ZEWriteHeaderCommon(xml: TZsspXMLWriterH; const CodePageName: string; const BOM: ansistring);
begin
  with (xml) do
  begin
    WriteRaw(BOM, false, false);
    Attributes.Add('version', '1.0');
    if (CodePageName > '') then
      Attributes.Add('encoding', CodePageName);
    WriteInstruction('xml', false);
    Attributes.Clear();
  end;
end;

//ќчищает массивы
procedure ZESClearArrays(var _pages: TIntegerDynArray;  var _names: TStringDynArray);
begin
  SetLength(_pages, 0);
  SetLength(_names, 0);
  _names := nil;
  _pages := nil;
end;

resourcestring
  DefaultSheetName = 'Sheet';

//делает уникальную строку, добавл€€ к строке '(num)'
//топорно, но работает
//INPUT
//  var st: string - строка
//      n: integer - номер
procedure ZECorrectStrForSave(var st: string; n: integer);
var
  l, i, m, num: integer;
  s: string;

begin
  if Trim(st) = '' then
     st := DefaultSheetName;  // behave uniformly with ZECheckTablesTitle

  l := length(st);
  if st[l] <> ')' then
    st := st + '(' + inttostr(n) + ')'
  else
  begin
    m := l;
    for i := l downto 1 do
    if st[i] = '(' then
    begin
      m := i;
      break;
    end;
    if m <> l then
    begin
      s := copy(st, m+1, l-m - 1);
      try
        num := StrToInt(s) + 1;
      except
        num := n;
      end;
      delete(st, m, l-m + 1);
      st := st + '(' + inttostr(num) + ')';
    end else
      st := st + '(' + inttostr(n) + ')';
  end;
end; //ZECorrectStrForSave

//делаем уникальные значени€ массивов
//INPUT
//  var mas: array of string - массив со значени€ми
procedure ZECorrectTitles(var mas: array of string);
var
  i, num, k, _kol: integer;
  s: string;

begin
  num := 0;
  _kol := High(mas);
  while (num < _kol) do
  begin
    s := UpperCase(mas[num]);
    k := 0;
    for i := num + 1 to _kol do
    begin
      if (s = UpperCase(mas[i])) then
      begin
        inc(k);
        ZECorrectStrForSave(mas[i], k);
      end;
    end;
    inc(num);
    if k > 0 then num := 0;
  end;
end; //CorrectTitles

//ѕровер€ет заголовки страниц, при необходимости корректирует
//INPUT
//  var XMLSS: TZEXMLSS
//  const SheetsNumbers:array of integer
//  const SheetsNames: array of string
//  var _pages: TIntegerDynArray
//  var _names: TStringDynArray
//  var retCount: integer
//RETURN
//      boolean - true - всЄ нормально, можно продолжать дальше
//                false - что-то не то подсунули, дальше продолжать нельз€
function ZECheckTablesTitle(var XMLSS: TZEXMLSS; const SheetsNumbers:array of integer;
                            const SheetsNames: array of string; out _pages: TIntegerDynArray;
                            out _names: TStringDynArray; out retCount: integer): boolean;
var
  t1, t2, i: integer;

  // '!' is allowed; ':' is not; whatever else ?
  procedure SanitizeTitle(var s: string);   var i: integer;
  begin
    s := Trim(s);
    for i := 1 to length(s) do
       if s[i] = ':' then s[i] := ';';
  end;
  function CoalesceTitle(const i: integer; const checkArray: boolean): string;
  begin
    if checkArray then begin
       Result := SheetsNames[i];
       SanitizeTitle(Result);
    end else
       Result := '';

    if Result = '' then begin
       Result := XMLSS.Sheets[_pages[i]].Title;
       SanitizeTitle(Result);
    end;

    if Result = '' then
       Result := DefaultSheetName + ' ' + IntToStr(_pages[i] + 1);
  end;

begin
  result := false;
  t1 :=  Low(SheetsNumbers);
  t2 := High(SheetsNumbers);
  retCount := 0;
  //если пришЄл пустой массив SheetsNumbers - берЄм все страницы из Sheets
  if t1 = t2 + 1 then
  begin
    retCount := XMLSS.Sheets.Count;
    setlength(_pages, retCount);
    for i := 0 to retCount - 1 do
      _pages[i] := i;
  end else
  begin
    //иначе берЄм страницы из массива SheetsNumbers
    for i := t1 to t2 do
    begin
      if (SheetsNumbers[i] >= 0) and (SheetsNumbers[i] < XMLSS.Sheets.Count) then
      begin
        inc(retCount);
        setlength(_pages, retCount);
        _pages[retCount-1] := SheetsNumbers[i];
      end;
    end;
  end;

  if (retCount <= 0) then
    exit;

  //названи€ страниц
//  t1 :=  Low(SheetsNames); // we anyway assume later that Low(_names) == t1 - then let us just skip this.
  t2 := High(SheetsNames);
  setlength(_names, retCount);
//  if t1 = t2 + 1 then
//  begin
//    for i := 0 to retCount - 1 do
//    begin
//      _names[i] := XMLSS.Sheets[_pages[i]].Title;
//      if trim(_names[i]) = '' then _names[i] := 'list';
//    end;
//  end else
//  begin
//    if (t2 > retCount) then
//      t2 := retCount - 1;
//    for i := t1 to t2 do
//      _names[i] := SheetsNames[i];
//    if (t2 < retCount) then
//    for i := t2 + 1 to retCount - 1 do
//    begin
//      _names[i] := XMLSS.Sheets[_pages[i]].Title;
//      if trim(_names[i]) = '' then _names[i] := 'list';
//    end;
//  end;
  for i := Low(_names) to High(_names) do begin
      _names[i] := CoalesceTitle(i, i <= t2);
  end;


  ZECorrectTitles(_names);
  result := true;
end; //ZECheckTablesTitle

//Trying to convert string like "n%" to integer
//INPUT
//      s: string     - input string
//  out Value: integer  - returned value
//RETURN
//      boolean - true - Ok
function TryStrToIntPercent(s: string; out Value: integer): boolean;
var
  l: integer;

begin
  l := length(s);
  if (l > 1) then
    if (s[l] = '%') then
      Delete(s, l, 1);
  result := TryStrToInt(s, Value);
end; //TryStrToIntPercent

//Get XML schema 2 duration string from TDateTime
//see https://www.w3.org/TR/xmlschema-2/#duration
//  example: PT537199H55M05.123000219S = 1961.04.12 7:55:05.123
//INPUT
//     const ADateTime: TDateTime - datetime
//RETURN
//      string
function ZEDateTimeToPTDurationStr(const ADateTime: TDateTime): string;
var
  _t: TDateTime;
  _h: integer;
  _min: integer;
  _sec: integer;
  _msec: integer;

begin
  _h := HoursBetween(ADateTime, 0);
  _t := IncHour(0, _h);

  _min := MinutesBetween(ADateTime, _t);
  _t := IncMinute(_t, _min);

  _sec := SecondsBetween(ADateTime, _t);
  _t := IncSecond(_t, _sec);

  _msec := MilliSecondsBetween(ADateTime, _t);

  Result := 'PT' + IntToStr(_h) + 'H' +
            IntToStr(_min) + 'M' +
            ZEFloatSeparator(FloatToStr(_sec + _msec / 1000)) + 'S'
            ;
end; //ZEDateTimeToPTDurationStr

//Get DateTime from XML schema 2 duration string
//see https://www.w3.org/TR/xmlschema-2/#duration
//  example: PT537199H55M05.123000219S = 1961.04.12 7:55:05.123
//INPUT
//     const APTTime: string - string in PT format
//RETURN
//      TDateTime - calculated datetime
function ZEPTDateDurationToDateTime(const APTTime: string): TDateTime;
var
  i: integer;
  l: integer;
  s: string;
  _y: integer;
  _m: integer;
  _d: integer;
  _h: integer;
  _min: integer;
  _sec: integer;
  _msec: integer;
  _t: double;
  _ch: char;
  _isTime: boolean;

  procedure _CheckSeconds();
  begin
    _t := ZETryStrToFloat(s, 0);
    _sec := trunc(_t);
    _msec := Round(frac(_t) * 1000);
    s := '';
  end; //_CheckSeconds

  function _GetIntPart(): integer;
  begin
    if (not TryStrToInt(s, Result)) then
      Result := 0;
    s := '';
  end; //_GetIntPart

begin
  Result := 0;
  _y := 0;
  _m := 0;
  _d := 0;
  _h := 0;
  _min := 0;
  _sec := 0;
  _msec := 0;
  s := '';
  _isTime := false;

  l := length(APTTime);
  for i := 1 to l do
  begin
    _ch := APTTime[i];
    case (_ch) of
      'P', 'p': s := '';
      'Y', 'y':
           _y := _GetIntPart();
      'D', 'd':
           _d := _GetIntPart();
      'T', 't':
          begin
            _isTime := true;
            s := '';
          end;
      'H', 'h':
           _h := _GetIntPart();
      'M', 'm':
          if (_isTime) then
            _min := _GetIntPart()
          else
            _m := _GetIntPart();
      'S', 's':
          _checkSeconds();
      else
        s := s + _ch;
    end; //case
  end; // for i

  if (_y <> 0) then
    Result := IncYear(Result, _y);
  if (_m <> 0) then
    Result := IncMonth(Result, _m);
  if (_d <> 0) then
    Result := IncDay(Result, _d);
  if (_h <> 0) then
    Result := IncHour(Result, _h);
  if (_min <> 0) then
    Result := IncMinute(Result, _min);
  if (_sec <> 0) then
    Result := IncSecond(Result, _sec);
  if (_msec <> 0) then
    Result := IncMilliSecond(Result, _msec);
end; //ZEPTDateDurationToDateTime

end.
