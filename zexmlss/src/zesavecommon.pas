//****************************************************************
// Common routines for saving and loading data
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2014.07.20
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

//������� ������������� ������ � �����
function ZETryStrToFloat(const st: string; valueIfError: double = 0): double; overload;
function ZETryStrToFloat(const st: string; out isOk: boolean; valueIfError: double = 0): double; overload;

//������� ������������� ������ � boolean
function ZETryStrToBoolean(const st: string; valueIfError: boolean = false): boolean;

//�������� ��� ������� �� �����
function ZEFloatSeparator(st: string): string;

//��������� ���� � ������ ��� XML (YYYY-MM-DDTHH:MM:SS)
function ZEDateToStr(ATime: TDateTime): string;

function IntToStrN(value: integer; NullCount: integer): string;

//BOM<?xml version="1.0" encoding="CodePageName"?>
procedure ZEWriteHeaderCommon(xml: TZsspXMLWriterH; const CodePageName: string; const BOM: ansistring);

//��������� ��������� �������, ��� ������������� ������������
function ZECheckTablesTitle(var XMLSS: TZEXMLSS; const SheetsNumbers:array of integer;
                            const SheetsNames: array of string; out _pages: TIntegerDynArray;
                            out _names: TStringDynArray; out retCount: integer): boolean;

//������� �������
procedure ZESClearArrays(var _pages: TIntegerDynArray;  var _names: TStringDynArray);

//��������� ������ � boolean
function ZEStrToBoolean(const val: string): boolean;

//���������, ���� �� ����� �����. ���� ���� �� ������������� �� ����������� ���������� - ��������� � �����.
function ZE_CheckDirExist(var DirName: string): boolean;

//�������� � ������ ������������������ �� �����������
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

const ZELibraryVersion = '0.0.7';
      ZELibraryFork = '';//'Arioch';  // or empty str   // URL ?

implementation

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


//�������� � ������ ������������������ �� �����������
//INPUT
//  const st: string - �������� ������
//RETURN
//      string - ������������ ������
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


//���������, ���� �� ����� �����. ���� ���� �� ������������� �� ����������� ���������� - ��������� � �����.
//INPUT
//  var DirName: string - ����
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

//��������� ������ � boolean
//INPUT
//  const val: string - ����������� ������
function ZEStrToBoolean(const val: string): boolean;
begin
  if (val = '1' ) or (UpperCase(val)='TRUE')  then
    result := true
  else
    result := false;
end;

//������� ������������� ������ � boolean
//  const st: string        - ������ ��� �������������
//    valueIfError: boolean - ��������, ������� ������������� ��� ������ ��������������
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

//������� ������������� ������ � �����
//INPUT
//  const st: string        - ������
//  out isOk: boolean       - ���� true - ������ ������
//    valueIfError: double  - ��������, ������� ������������� ��� ������ ��������������
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
        {$IFDEF DELPHI_UNICODE}
        s := s + FormatSettings.DecimalSeparator
        {$ELSE}
          {$IFDEF Z_FPC_USE_FORMATSETTINGS}
        s := s + FormatSettings.DecimalSeparator
          {$ELSE}
        s := s + DecimalSeparator
          {$ENDIF}
        {$ENDIF}
      else if (st[i] <> ' ') then
        s := s + st[i];

    isOk := TryStrToFloat(s, result);
    if (not isOk) then
      result := valueIfError;
  end;
end; //ZETryStrToFloat

//������� ������������� ������ � �����
//INPUT
//  const st: string        - ������
//    valueIfError: double  - ��������, ������� ������������� ��� ������ ��������������
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
        {$IFDEF DELPHI_UNICODE}
        s := s + FormatSettings.DecimalSeparator
        {$ELSE}
          {$IFDEF Z_FPC_USE_FORMATSETTINGS}
        s := s + FormatSettings.DecimalSeparator
          {$ELSE}
        s := s + DecimalSeparator
          {$ENDIF}
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

//�������� ��� ������� �� �����
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

//��������� ����� � ������ ����������� ����� NullCount
//TODO: ���� ������� ��� � ��������� � FlyLogReader-�
//INPUT
//      value: integer     - �����
//      NullCount: integer - ���-�� ������ � ������
//RETURN
//      string
function IntToStrN(value: integer; NullCount: integer): string;
var
  t: integer;
  k: integer;

begin
  t := value;
  k := 0;
  if (t = 0) then
    k := 1;
  while t > 0 do
  begin
    inc(k);
    t := t div 10;
  end;
  result := IntToStr(value);
  for t := 1 to (NullCount - k) do
    result := '0' + result;
end; //IntToStrN

//��������� ���� � ������ ��� XML (YYYY-MM-DDTHH:MM:SS)
//INPUT
//    ATime: TDateTime - ������ ����/�����
function ZEDateToStr(ATime: TDateTime): string;
var
  HH, MM, SS, MS: word;

begin
  DecodeDate(ATime, HH, MM, SS);
  result := IntToStrN(HH, 4) + '-' + IntToStrN(MM, 2) + '-' + IntToStrN(SS, 2) + 'T';
  DecodeTime(ATime, HH, MM, SS, MS);
  result := result + IntToStrN(HH, 2) + ':' + IntToStrN(MM, 2) + ':' + IntToStrN(SS, 2);
end;

//BOM<?xml version="1.0" encoding="CodePageName"?>
//INPUT
//    xml: TZsspXMLWriterH      - ��������
//  const CodePageName: string  - ��� ���������
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

//������� �������
procedure ZESClearArrays(var _pages: TIntegerDynArray;  var _names: TStringDynArray);
begin
  SetLength(_pages, 0);
  SetLength(_names, 0);
  _names := nil;
  _pages := nil;
end;

resourcestring
  DefaultSheetName = 'Sheet';

//������ ���������� ������, �������� � ������ '(num)'
//�������, �� ��������
//INPUT
//  var st: string - ������
//      n: integer - �����
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

//������ ���������� �������� ��������
//INPUT
//  var mas: array of string - ������ �� ����������
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

//��������� ��������� �������, ��� ������������� ������������
//INPUT
//  var XMLSS: TZEXMLSS
//  const SheetsNumbers:array of integer
//  const SheetsNames: array of string
//  var _pages: TIntegerDynArray
//  var _names: TStringDynArray
//  var retCount: integer
//RETURN
//      boolean - true - �� ���������, ����� ���������� ������
//                false - ���-�� �� �� ���������, ������ ���������� ������
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
  //���� ������ ������ ������ SheetsNumbers - ���� ��� �������� �� Sheets
  if t1 = t2 + 1 then
  begin
    retCount := XMLSS.Sheets.Count;
    setlength(_pages, retCount);
    for i := 0 to retCount - 1 do
      _pages[i] := i;
  end else
  begin
    //����� ���� �������� �� ������� SheetsNumbers
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

  //�������� �������
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

end.
