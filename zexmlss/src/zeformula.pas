//****************************************************************
// Routines for formulas.
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2013.11.05
//----------------------------------------------------------------
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

unit zeformula;

{$I zexml.inc}
{$I compver.inc}

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

interface

uses
  sysutils;

const
  //ZE_RTA = ZE R1C1 to A1
  ZE_RTA_ODF            =   1;  //преобразовывать для ODF (=[.A1] + [.B1])
  ZE_RTA_ODF_PREFIX     =   2;  //добавлять префикс для ODF, если первый символ в формуле '=' (of:=[.A1] + [.B1])
  ZE_RTA_NO_ABSOLUTE    =   4;  //все абсолютные ссылки заменять на относительные (R1C1 => A1) (относительные не меняет)
  ZE_RTA_ONLY_ABSOLUTE  =   8;  //все относительные ссылки заменять на абсолютные (R[1]C[1] => $C$3) (абсолютные не меняет)
  ZE_RTA_ODF_NO_BRACKET = $10;  //Для ODF, но не добавлять квадратные скобки, разделитель лист/ячейка - точка ".".

  //ZE_ATR = ZE A1 to R1C1
  ZE_ATR_DEL_PREFIX     =   1;  //Удалять все символы до первого '='

function ZEGetA1byCol(ColNum: integer; StartZero: boolean = true): string;
function ZEGetColByA1(AA: string; StartZero: boolean = true): integer;

function ZER1C1ToA1(const formula: string; CurCol, CurRow: integer; options: integer; StartZero: boolean = true): string;
function ZEA1ToR1C1(const formula: string; CurCol, CurRow: integer; options: integer; StartZero: boolean = true): string;

function ZEGetCellCoords(const cell: string; out column, row: integer; StartZero: boolean = true): boolean;

implementation

const
  ZE_STR_ARRAY: array [0..25] of char = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');

//Получает номер строки и столбца по строковому значению (для A1 стилей)
//INPUT
//  const cell: string      - номер ячейки в A1 стиле
//  out column: integer     - возвращаемый номер столбца
//  out row: integer        - возвращаемый номер строки
//      StartZero: boolean  - признак нумерации с нуля
//RETURN
//      boolean - true - координаты успешно определены
function ZEGetCellCoords(const cell: string; out column, row: integer; StartZero: boolean = true): boolean;
var
  i: integer;
  s1, s2: string;
  _isOk: boolean;
  b: boolean;

begin
  _isOk := true;
  s1 := '';
  s2 := '';
  b := false;
  for i := 1 to length(cell) do
  case cell[i] of
    'A'..'Z', 'a'..'z':
      begin
        s1 := s1 + cell[i];
        b := true;
      end;
    '0'..'9':
      begin
        if (not b) then
        begin
          _isOk := false;
          break;
        end;
        s2 := s2 + cell[i];
      end;
    else
      begin
        _isOk := false;
        break;
      end;
  end;
  if (_isOk) then
  begin
    if (not TryStrToInt(s2, row)) then
      _isOk := false
    else
    begin
      if (StartZero) then
        dec(row);
      column := ZEGetColByA1(s1, StartZero);
      if (column < 0) then
        _isOk := false;
    end;
  end;
  result := _isOk;
end; //ZEGetCellCoords

//Попытка преобразовать номер ячейки из R1C1 в A1 стиль
//если не удалось распознать номер ячейки, то возвратит обратно тот же текст
//INPUT
//  const st: string        - предположительно номер ячеки (диапазон)
//      CurCol: integer     - номер столбца ячейки
//      CurRow: integer     - номер строки ячейки
//      options: integer    - параметры преобразования
//      StartZero: boolean  - признак нумерации с нуля
//RETURN
//      string - номер ячейки в стиле A1
function ReturnA1(const st: string; CurCol, CurRow: integer; options: integer; StartZero: boolean = true): string;
var
  s: string;
  i, kol: integer;
  retTxt: string;
  isApos: boolean;
  t: integer;
  isList: boolean;
  isODF: boolean;
  isSq: boolean;
  isOk: boolean;
  isNumber: boolean;
  isR, isC: boolean;
  isNotLast: boolean;
  _c, _r: string;
  is_only_absolute: boolean;
  is_no_absolute: boolean;
  isDelim: boolean;
  _num: integer;
  _use_bracket: boolean;

  //Возвращает строку
  procedure _getR(num: integer);
  begin
    if (isSq or (num = 0)) then
      num := CurRow + num;
    if (is_only_absolute and isSq) then
      isSq := false
    else
      if (is_no_absolute and (not isSq)) then
        isSq := true;
    _r := IntToStr(num);
    if (not isSq) then
      _r := '$' + _r;
    isNumber := false;
    inc(_num);
  end; //_getR

  //Возвращает столбец
  procedure _getC(num: integer);
  begin
    if (isSq or (num = 0)) then
      num := CurCol + num;
    if (is_only_absolute and isSq) then
      isSq := false
    else
      if (is_no_absolute and (not isSq)) then
        isSq := true;
    _c := ZEGetA1byCol(num, false);
    if (not isSq) then
      _c := '$' + _c;
    isNumber := false;
    inc(_num);
  end; //_getС

  //Проверяет символ
  procedure _checksymbol(ch: char);
  begin
    if (isApos) then
    begin
      if (ch <> '''') then
      begin
        s := s + ch;
        exit;
      end;
    end else
    begin
      if (isNumber) then
      begin
        {$IFDEF DELPHI_UNICODE}
        if (not CharInSet(ch, ['-', '0'..'9', ']', '[', ''''])) then
        {$ELSE}
        if (not (ch in ['-', '0'..'9', ']', '[', ''''])) then
        {$ENDIF}
        begin
          if (not (isC xor isR)) then
          begin
            isOk := false;
            exit;
          end;
          if (not TryStrToInt(s, t)) then
          begin
            isOk := false;
            exit;
          end;

          if (isC) then
            _getC(t);
          if (isR) then
            _getR(t);
          isSq := false;
          s := '';
        end;
      end else //if (isNumber)
      begin
        //если адрес: RC (без чисел - нули)
        {$IFDEF DELPHI_UNICODE}
        if (isR and CharInSet(ch, ['C', 'c'])) then
        {$ELSE}
        if (isR and (ch in ['C', 'c'])) then
        {$ENDIF}
        begin
          _getR(0);
          s := '';
          isSq := false;
        end else
        if (isC and (not isNotLast)) then
        begin
          _getC(0);
          s := '';
          isSq := false;
        end;
      end;
    end;
    case ch of
      '''':
        begin
          s := s + ch;
          isApos := not isApos;
        end;
      '[': {хм..};
      ']': isSq := true;
      'R', 'r':
        begin
          //R - ok, CR - что-то не то
          if (isR or isC) then
            isOk := false;
          isR := true;
          s := '';
          isDelim := false;
        end;
      'C', 'c':
        begin
          if (isC or (not isR)) then
            isOk := false
          else
          begin
            isC := true;
            isR := false;
          end;
          s := '';
          isDelim := false;
        end;
      '-', '0'..'9':
        begin
          s := s + ch;
          if (isC or isR) then
            if (not isNumber) then
              isNumber := true;
        end;
      '!':  //разделитель страницы
        begin
          retTxt := retTxt + s;
          if (isODF) then //ODF
            retTxt := retTxt + '.'
          else
            retTxt := retTxt + ch;
          s := '';
          isList := true;
          isDelim := false;
        end;
      else
        if (isDelim and isNotLast) then
          s := s + ch
        else
          if (isNotLast) then
            isOk := false; //O_o - вроде как не ячейка, выходим и возвращаем всё как есть
    end; //case
  end; //_checksymbol

begin
  result := '';
  if (TryStrToInt(st, t)) then
  begin
    result := st;
    exit;
  end;
  kol := length(st);
  s := '';
  retTxt := '';
  isApos := false;
  isList := false;
  isSq := false;
  isOk := true;
  isNumber := false;
  isR := false;
  isC := false;
  isNotLast := true;
  isDelim := true;
  _c := '';
  _r := '';
  _num := 0;

  is_no_absolute := (options and ZE_RTA_NO_ABSOLUTE = ZE_RTA_NO_ABSOLUTE);
  is_only_absolute := (options and ZE_RTA_ONLY_ABSOLUTE = ZE_RTA_ONLY_ABSOLUTE);
  isODF := (options and ZE_RTA_ODF = ZE_RTA_ODF);
  for i := 1 to kol do
  begin
    _checksymbol(st[i]);
    if (not isOk) then
      break;
  end;
  isNotLast := false;
  //нужно подумать, что делать, если было не 2 преобразования
  if ((kol <= 0) or (_num = 0)) then
    isOk := false;
  _checksymbol(';');
  if (not isOk) then
  begin
    result := st;
    exit;
  end;
  result := retTxt + _c + _r + s;
  _use_bracket := not (options and ZE_RTA_ODF_NO_BRACKET = ZE_RTA_ODF_NO_BRACKET);
  if (isODF and _use_bracket) then
  begin
    if (not isList) then
      result := '.' + result;
    result := '[' + result + ']';  
  end;
end; //ReturnA1

//Переводит формулу из стиля R1C1 в стиль A1
//INPUT
//  const formula: string - формула в стиле R1C1
//      CurRow: integer   - номер строки ячейки
//      CurCol: integer   - номер столбца ячейки
//      options: integer  - настройки преобразования (ZE_RTA_ODF и ZE_RTA_ODF_PREFIX)
//                              options and ZE_RTA_ODF = ZE_RTA_ODF - преобразовывать для ODF (=[.A1] + [.B1])
//                              options and ZE_RTA_ODF_PREFIX = ZE_RTA_ODF_PREFIX - добавлять префикс для ODF, если первый символ в формуле '=' (of:=[.A1] + [.B1])
//      StartZero: boolean- при true счёт строки/ячейки начинается с 0.
//RETURN
//      string  - текст формулы в стиле R1C1
function ZER1C1ToA1(const formula: string; CurCol, CurRow: integer; options: integer; StartZero: boolean = true): string;
var
  kol: integer;
  i: integer;
  retFormula: string;
  s: string;
  isQuote: boolean; // " ... "
  isApos: boolean;  // ' ... '
  isNotLast: boolean;
  isSq: boolean;

  procedure _checksymbol(ch: char);
  begin
    case ch of
      '"':
        begin
          if (isApos) then
            s := s + ch
          else
          begin
            if (isQuote) then
            begin
              retFormula := retFormula + s + ch;
              s := '';
            end else
            begin
              if (s > '') then
              begin
                //O_o Странно
                retFormula := retFormula + ReturnA1(s, CurCol, CurRow, options, StartZero);
                s := '';
              end;
              s := ch
            end;
            isQuote := not isQuote;
          end;
        end;
      '''':
        begin
          s := s + ch;
          if (not isQuote) then
            isApos := not isApos;
        end;
      '[':
        begin
          s := s + ch;
          if (not (isQuote or isApos)) then
            isSq := true;
        end;
      ']':
        begin
          s := s + ch;
          if (not (isQuote or isApos)) then
            isSq := false;
        end;
      ':', ';', ' ', '-', '%', '^', '*', '/', '+', '&', '<', '>', '(', ')', '=': //разделители
        begin
          if (isApos or isQuote or isSq) then
            s := s + ch
          else
          begin
            retFormula := retFormula + ReturnA1(s, CurCol, CurRow, options, StartZero);
            if (isNotLast) then
              retFormula := retFormula + ch;
            s := '';
          end;
        end;
      else
        s := s + ch;
    end;
  end; //_checksymbol

begin
  result := '';
  kol := length(formula);
  retFormula := '';
  s := '';
  if (StartZero) then
  begin
    inc(CurRow);
    inc(CurCol);
  end;
  isApos := false;
  isQuote := false;
  isNotLast := true;
  isSq := false;
  for i := 1 to kol do
    _checksymbol(formula[i]);
  isNotLast := false;  
  _checksymbol(';');
  result := retFormula;
  if (options and ZE_RTA_ODF = ZE_RTA_ODF) and (options and ZE_RTA_ODF_PREFIX = ZE_RTA_ODF_PREFIX) then
    if (kol > 0) then
      if (formula[1] = '=') then
        result := 'of:' + result;
end; //ZER1C1ToA1

//Попытка преобразовать номер ячейки из A1 в R1C1 стиль
//если не удалось распознать номер ячейки, то возвратит обратно тот же текст
//INPUT
//  const st: string        - предположительно номер ячеки (диапазон)
//      CurCol: integer     - номер столбца ячейки
//      CurRow: integer     - номер строки ячейки
//      options: integer    - настройки
//      StartZero: boolean  - признак нумерации с нуля
//RETURN
//      string - номер ячейки в стиле R1C1
function ReturnR1C1(const st: string; CurCol, CurRow: integer; StartZero: boolean = true): string;
var
  i: integer;
  s: string;
  isApos: boolean;
  _startNumber: boolean;
  kol: integer;
  num: integer;
  t: integer;
  isAbsolute: byte;
  sa: string;
  isNotLast: boolean;
  column: string;
  isC: boolean;

  procedure _GetColumn();
  begin
    //попробовать преобразовать
    num := ZEGetColByA1(s, false);
    if (num >= 0) then //распознался вроде нормально
    begin
      if (num > 25000) then //сколько там колонок возможно?
        result := result + sa + s
      else
      begin
        column := '';
        if (isAbsolute > 0) then
          column := 'C' + IntToStr(num)
        else
        begin
          t := num - CurCol;
          if (t <> 0) then
            column := 'C[' + IntToStr(t) + ']'
          else
            column := 'C';
        end;
      end;
    end else //что-то не то
      result := result + sa + s;
    if (isAbsolute > 0) then
      dec(isAbsolute);
    sa := '';
    s := '';
    isC := true;
  end; //_GetColumn

  procedure _CheckSymbol(ch: char);
  begin
    {$IFDEF DELPHI_UNICODE}
    if (not CharInSet(ch, ['0'..'9'])) then
    {$ELSE}
    if (not (ch in ['0'..'9'])) then
    {$ENDIF}
      if (not isApos) then
      begin
        if (_startNumber) then
        begin
          if (TryStrToInt(s, t)) then //удалось получить число
          begin
            if (isAbsolute > 0) then
              result := result + 'R' + s + column
            else
            begin
              t := t - CurRow;
              if (t <> 0) then
                result := result + 'R[' + IntToStr(t) + ']' + column
              else
                result := result + 'R' + column;  
            end;
            if (isAbsolute > 0) then
              dec(isAbsolute);
            isC := false;  
          end else
            result := result + sa + s;
          s := '';
          sa := '';
        end;
        _startNumber := false;
      end;
    case ch of
      '''':
        begin
          s := s + ch;
          if (isApos) then
          begin
            result := result + s;
            s := '';
          end;
          isApos := not isApos;
        end;
      '.': //разделитель для листа (OpenOffice/LibreOffice)
        begin
           if (isApos) then
            s := s + ch
          else
          begin
            if (s > '') then
              result := result + s + '!';
            s := '';
          end;
        end;
      '!': //разделитель для листа (excel)
        begin
          if (isApos) then
            s := s + ch
          else
          begin
            result := result + s + ch;
            s := '';
          end;
        end;
      '$':
        begin
          if (isApos) then
            s := s + ch
          else
          begin
            if (not _startNumber) and (s > '') then
              _GetColumn();
            inc(isAbsolute);
            sa := ch;
          end;  
        end;
      '[':
        begin
          if (isApos) then
            s := s + ch
          else
          begin
          end;
        end;
      ']':
        begin
          if (isApos) then
            s := s + ch
          else
          begin
          end;
        end;
      '0'..'9':
         begin
           if (isApos) then
             s := s + ch
           else
           begin
             if ((not _startNumber) and (not isC)) then //первая цифирка
             begin
               _GetColumn();
               s := '';
             end;
             s := s+ ch;
             _startNumber := true;
           end;
         end;
      else
        if (isNotLast) then
          s := s + ch;
    end; //case
  end; // _CheckSymbol

  //Проверяет, с какого символа в строке начать
  procedure FindStartNumber(out num: integer);
  var
    i: integer;
    z: boolean;

  begin
    num := 1;
    z := false;
    for i := 1 to kol do
      case st[i] of
        '''':
          begin
            s := s + st[i];
            z := not z;
          end;
        '!', '.':
           if (not z) then
           begin
             num := i;
             exit;
           end;
         else
           s := s + st[i];  
      end; //case
    s := '';  
  end; //FindStartNumber

begin
  result := '';
  s := '';
  isApos := false;
  kol := length(st);
  if (kol >= 1) then
    if (st[1] <> '$') then
      if (TryStrToInt(st, t)) then
      begin
        result := st;
        exit;
      end;
  FindStartNumber(i);
  _startNumber := false;
  isAbsolute := 0;
  sa := '';
  column := '';
  isNotlast := true;
  isC := false;
  while (i <= kol) do
  begin
    _CheckSymbol(st[i]);
    inc(i);
  end; //while
  isNotLast := false;
  _CheckSymbol(';');
  if (s > '') then
    result := result + s; 
end; //ReturnR1C1

//Переводит формулу из стиля A1 в стиль R1C1
//INPUT
//  const formula: string - формула в стиле A1
//      CurRow: integer   - номер строки ячейки
//      CurCol: integer   - номер столбца ячейки
//      options: integer  - настройки преобразования
//      StartZero: boolean- при true счёт строки/ячейки начинается с 0.
//RETURN
//      string  - текст формулы в стиле R1C1
function ZEA1ToR1C1(const formula: string; CurCol, CurRow: integer; options: integer; StartZero: boolean = true): string;
var
  i, l: integer;
  s: string;
  retFormula: string;
  isQuote: boolean; // " ... "
  isApos: boolean;  // ' ... '
  isNotLast: boolean;
  start_num: integer;

  //Проверить символ
  //INPUT
  //  const ch: char - символ для проверки
  procedure _CheckSymbol(const ch: char);
  begin
    case ch of
      '"':
        begin;
          if (isApos) then
            s := s + ch
          else
          begin
            if (isQuote) then
            begin
              retFormula := retFormula + s + ch;
              s := '';
            end else
            begin
              if (s > '') then
              begin
                //O_o Странно
                retFormula := retFormula + ReturnR1C1(s, CurCol, CurRow, StartZero);
                s := '';
              end;
              s := ch
            end;
            isQuote := not isQuote;
          end;  
        end;
      '''':
        begin
          s := s + ch;
          if (not isQuote) then
            isApos := not isApos;
        end; 
      ':', ';', ' ', '-', '%', '^', '*', '/', '+', '&', '<', '>', '(', ')', ']', '[', '=': //разделители
        begin
          if (isQuote or isApos) then
            s := s + ch
          else
          begin
            retFormula := retFormula + ReturnR1C1(s, CurCol, CurRow, StartZero);
            if (isNotLast) then
              {$IFDEF DELPHI_UNICODE}
              if (not CharInSet(ch, ['[',']'])) then
              {$ELSE}
              if (not (ch in ['[', ']'])) then
              {$ENDIF}
                retFormula := retFormula + ch;
            s := '';
          end;
        end;
      else
        s := s + ch;
    end;
  end; //_CheckSymbol

  procedure FindStartNum(var start_num: integer);
  var
    i: integer;

  begin
    for i := 1 to l do
      if (formula[i] = '=') then
      begin
        start_num := i;
        exit;
      end;
  end; //FindStartNum

begin
  result := '';
  l := length(formula);
  s := '';
  retFormula := '';
  isQuote := false;
  isApos := false;
  isNotLast := true;
  if (StartZero) then
  begin
    inc(CurRow);
    inc(CurCol);
  end;

  start_num := 1;
  if (options and ZE_ATR_DEL_PREFIX = ZE_ATR_DEL_PREFIX) then
    FindStartNum(start_num);

  for i := start_num to l do
    _CheckSymbol(formula[i]);
  isNotLast := false;  
  _CheckSymbol(';');
  if (isQuote or isApos) then
    retFormula := retFormula + s;
  result := retFormula;
end; //ZEA1ToR1C1

//Возвращает номер столбца по буквенному обозначению
//INPUT
//  const AA: string      - буквенное обозначение столбца
//      StartZero: boolean  - если true, то счёт начинает с нуля (т.е. A = 0), в противном случае с 1.
//RETURN
//      integer -   -1 - не удалось преобразовать
function ZEGetColByA1(AA: string; StartZero: boolean = true): integer;
var
  i: integer;
  num, t, kol, s: integer;

begin
  result := -1;
  num := 0;
  t := 0;
  AA := UpperCase(AA);
  kol := length(AA);
  s := 1;
  for i := kol downto 1 do
  begin
    case AA[i] of
      'A': t :=  0;
      'B': t :=  1;
      'C': t :=  2;
      'D': t :=  3;
      'E': t :=  4;
      'F': t :=  5;
      'G': t :=  6;
      'H': t :=  7;
      'I': t :=  8;
      'J': t :=  9;
      'K': t := 10;
      'L': t := 11;
      'M': t := 12;
      'N': t := 13;
      'O': t := 14;
      'P': t := 15;
      'Q': t := 16;
      'R': t := 17;
      'S': t := 18;
      'T': t := 19;
      'U': t := 20;
      'V': t := 21;
      'W': t := 22;
      'X': t := 23;
      'Y': t := 24;
      'Z': t := 25;
      else
        exit;
    end;
    num := num + (t + 1) * s;
    s := s * 26;
    if (s < 0) or (num < 0) then
      exit;
  end;
  result := num;
  if (StartZero) then
    result := result - 1;
end; //ZEGetColByAA

//Возвращает буквенное обозначение столбца для АА стиля
//INPUT
//      ColNum: integer     - номер столбца
//      StartZero: boolean  - если true, то счёт начинается с 0, в противном случае - с 1.
function ZEGetA1byCol(ColNum: integer; StartZero: boolean = true): string;
var
  t, n: integer;
  s: string;

begin
  t := ColNum;
  if (not StartZero) then
    dec(t);
  result := '';
  s := '';
  while t >= 0 do
  begin
    n := t mod 26;
    t := (t div 26) - 1;
    //ХЗ как там с кодировками будет
    s := s + ZE_STR_ARRAY[n];
  end;
  for t := length(s) downto 1 do
    result := result + s[t];
end; //ZEGetAAbyCol

end.
