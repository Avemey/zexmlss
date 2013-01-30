//****************************************************************
// Common routines for saving and loading data
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2012.08.05
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
  sysutils, zexmlss, zsspxml;

const
  ZE_MMinInch: real = 25.4;

type
  TZESaveIntArray = array of integer;
  TZESaveStrArray = array of string;

//������� ������������� ������ � �����
function ZETryStrToFloat(const st: string; valueIfError: double = 0): double;

//������� ������������� ������ � boolean
function ZETryStrToBoolean(const st: string; valueIfError: boolean = false): boolean;

//�������� ��� ������� �� �����
function ZEFloatSeparator(st: string): string;

//��������� ���� � ������ ��� XML (YYYY-MM-DDTHH:MM:SS)
function ZEDateToStr(ATime: TDateTime): string;

function IntToStrN(value: integer; NullCount: integer): string;

//BOM<?xml version="1.0" encoding="CodePageName"?>
procedure ZEWriteHeaderCommon(xml: TZsspXMLWriterH; const CodePageName: string; const BOM: ansistring);
{$IfDef DELPHI_UNICODE}
   overload;

procedure ZEWriteHeaderCommon(xml: TZsspXMLWriterH; const CodePageName: AnsiString; const BOM: ansistring);
   overload; inline;
{$EndIf}

//��������� ��������� �������, ��� ������������� ������������
function ZECheckTablesTitle(var XMLSS: TZEXMLSS; const SheetsNumbers:array of integer;
                            const SheetsNames: array of string; var _pages: TZESaveIntArray;
                            var _names: TZESaveStrArray; var retCount: integer): boolean;

//������� �������
procedure ZESClearArrays(var _pages: TZESaveIntArray;  var _names: TZESaveStrArray);

//��������� ������ � boolean
function ZEStrToBoolean(const val: string): boolean;

//���������, ���� �� ����� �����. ���� ���� �� ������������� �� ����������� ���������� - ��������� � �����.
function ZE_CheckDirExist(var DirName: string): boolean;

//�������� � ������ ������������������ �� �����������
function ZEReplaceEntity(const st: string): string;

implementation

//�������� � ������ ������������������ �� �����������
//INPUT
//  const st: string - �������� ������
//RETURN
//      string - ������������ ������
function ZEReplaceEntity(const st: string): string;
var
  s, s1: string;
  i, n: integer;
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
  n := length(st);
  i := 1;
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
  if (length(s) > 0) then
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
  if (length(st) > 0) then
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
//    valueIfError: double  - ��������, ������� ������������� ��� ������ ��������������
function ZETryStrToFloat(const st: string; valueIfError: double = 0): double;
var
  s: string;
  i: integer;

begin
  result := 0;
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
    if length(CodePageName) > 0 then
      Attributes.Add('encoding', CodePageName);
    WriteInstruction('xml', false);
    Attributes.Clear();
  end;
end;

{$IfDef DELPHI_UNICODE}
procedure ZEWriteHeaderCommon(xml: TZsspXMLWriterH; const CodePageName: AnsiString; const BOM: ansistring);
begin
  ZEWriteHeaderCommon(xml, UnicodeString(CodePageName), BOM);
  // intended typecast - supress warnings
end;
{$EndIf}



//������� �������
procedure ZESClearArrays(var _pages: TZESaveIntArray;  var _names: TZESaveStrArray);
begin
  SetLength(_pages, 0);
  SetLength(_names, 0);
  _names := nil;
  _pages := nil;
end;

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
  l := length(st);
  if (l = 0) then
   begin
     l := 4;
     st := 'List';
   end;
  if st[l] <>')' then
    st := st + '(' + inttostr(n) + ')'
  else
  begin
    m :=  l;
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
//  var _pages: TZESaveIntArray
//  var _names: TZESaveStrArray
//  var retCount: integer
//RETURN
//      boolean - true - �� ���������, ����� ���������� ������
//                false - ���-�� �� �� ���������, ������ ���������� ������
function ZECheckTablesTitle(var XMLSS: TZEXMLSS; const SheetsNumbers:array of integer;
                            const SheetsNames: array of string; var _pages: TZESaveIntArray;
                            var _names: TZESaveStrArray; var retCount: integer): boolean;
var
  t1, t2, i: integer;

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
  t1 :=  Low(SheetsNames);
  t2 := High(SheetsNames);
  setlength(_names, retCount);
  if t1 = t2 + 1 then
  begin
    for i := 0 to retCount - 1 do
    begin
      _names[i] := XMLSS.Sheets[_pages[i]].Title;
      if trim(_names[i]) = '' then _names[i] := 'list';
    end;
  end else
  begin
    if (t2 > retCount) then
      t2 := retCount - 1;
    for i := t1 to t2 do
      _names[i] := SheetsNames[i];
    if (t2 < retCount) then
    for i := t2 + 1 to retCount - 1 do
    begin
      _names[i] := XMLSS.Sheets[_pages[i]].Title;
      if trim(_names[i]) = '' then _names[i] := 'list';
    end;
  end;

  ZECorrectTitles(_names);
  result := true;
end; //ZECheckTablesTitle

end.
