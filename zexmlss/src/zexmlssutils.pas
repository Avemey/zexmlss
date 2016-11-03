//****************************************************************
// zexmlssutils  (Z Excel XML SpreadSheet Utils)
// Различные дополнительные утилитки для ZEXMLSS
// Накалякано в Мозыре в 2009 году
// Автор:  Неборак Руслан Владимирович (Ruslan V. Neborak)
// e-mail: avemey(мяу)tut(точка)by
// URL:    http://avemey.com
// Ver:    0.0.11
// Лицензия: zlib
// Last update: 2016.09.10
//----------------------------------------------------------------
// This software is provided "as-is", without any express or implied warranty.
// In no event will the authors be held liable for any damages arising from the
// use of this software.
//****************************************************************
unit zexmlssutils;

interface

{$I zexml.inc}
{$I compver.inc}

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

uses
  {$IFNDEF FPC}
  Windows,
  {$ENDIF}
  SysUtils, Types, Classes, Grids,
  {$IFNDEF NOZCOLORSTRINGGRID}
  ZColorStringGrid,
  {$ENDIF}
  zexmlss, zsspxml, zesavecommon, graphics
//  { $IFDEF VER130},
//  sysd7
//  { $ENDIF}
  ;

{$IFNDEF NOZCOLORSTRINGGRID}

//Копирует стиль из TCellStyle в TZCellStyle
procedure CellStyleTo(const CellStyle: TCellStyle; var XStyle: TZStyle; ignorebgcolor: boolean; _border: byte);

//Копирует стиль из TZStyle в TCellStyle
procedure ZStyleToCellStyle(const ZZStyle: TZStyle; CellStyle: TCellStyle; StyleCopy: integer);

//Копирует данные из TZColorStringGrid-а на страницу TZEXMLSS
function GridToXmlSS(var XMLSS: TZEXMLSS; const PageNum: integer;
                     var Grid: TZColorStringGrid; ToCol: integer; ToRow: integer;
                         BCol, BRow, ECol, ERow: integer; ignorebgcolor: boolean; _border: byte): boolean; overload;

//Копирует данные из страницы TZEXMLSS в TZColorStringGrid
function XmlSSToGrid(var Grid: TZColorStringGrid; var XMLSS: TZEXMLSS; const PageNum: integer;
                         ToCol: integer; ToRow: integer; BCol, BRow, ECol, ERow: integer; InsertMode: byte; StyleCopy: integer = 1023): boolean; overload;
{$ENDIF}

//Копирует данные из TStringGrid-а на страницу TZEXMLSS
function GridToXmlSS(var XMLSS: TZEXMLSS; const PageNum: integer;
                     var Grid: TStringGrid; ToCol: integer; ToRow: integer;
                         BCol, BRow, ECol, ERow: integer; ignorebgcolor: boolean; _border: byte): boolean; overload;

//Копирует данные из страницы TZEXMLSS в TStringGrid
function XmlSSToGrid(var Grid: TStringGrid; var XMLSS: TZEXMLSS; const PageNum: integer;
                         ToCol: integer; ToRow: integer; BCol, BRow, ECol, ERow: integer; InsertMode: byte; StyleCopy: integer = 511): boolean; overload;

function SaveXmlssToHtml(var XMLSS: TZEXMLSS; const PageNum: integer; Title: string; Stream: TStream;
                         TextConverter: TAnsiToCPConverter; CodePageName: string): integer; overload;

function SaveXmlssToHtml(var XMLSS: TZEXMLSS; const PageNum: integer; Title: string; FileName: string;
                         TextConverter: TAnsiToCPConverter; CodePageName: string): integer; overload;

//Сохраняет в поток в формате Excel XML SpreadSheet
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; Stream: TStream; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;

//Сохраняет в поток в формате Excel XML SpreadSheet
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; Stream: TStream; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string): integer; overload;

//Сохраняет в поток в формате Excel XML SpreadSheet
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; Stream: TStream): integer; overload;

//Сохраняет в файл в формате Excel XML SpreadSheet
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;

//Сохраняет в файл в формате Excel XML SpreadSheet
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string): integer; overload;

//Сохраняет в файл в формате Excel XML SpreadSheet
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; FileName: string): integer; overload;

//Читает из потока Excel XML SpreadSheet (EXMLSS)
//      XMLSS: TZEXMLSS                 - хранилище
//      Stream: TStream                 - поток
function ReadEXMLSS(var XMLSS: TZEXMLSS; Stream: TStream): integer; overload;

//Читает из файла Excel XML SpreadSheet (EXMLSS)
//      XMLSS: TZEXMLSS                 - хранилище
//      FileName: string                - имя файла
function ReadEXMLSS(var XMLSS: TZEXMLSS; FileName: string): integer; overload;

// needed for uniform save routines: zeSave*
// missed in pre-Unicode Dlephi and in FPC
function SplitString(const buffer: string; const delimeter: char): TStringDynArray;
{$IfDef DELPHI_UNICODE} overload; {$EndIf}

implementation
uses
  zenumberformats  //ConvertFormatNativeToXlsx / ConvertFormatXlsxToNative
{$IfDef DELPHI_UNICODE}
  ,
  StrUtils,  // stock SplitString(string, string) implementation
  AnsiStrings  // AnsiString targeted overloaded versions of Pos, Trim, etc
{$EndIf}
  ;

{$IFDEF DELPHI_UNICODE}
  {$DEFINE USE_STRUTILS_SPLIT_STRING}
{$ENDIF}

{$IFDEF VER200} // RAD Studio 2009
  {$UNDEF USE_STRUTILS_SPLIT_STRING} //There are no StrUtils.SplitString in D2009!!!!
{$ENDIF VER200}

function SplitString(const buffer: string; const delimeter: char): TStringDynArray;
{$IfDef USE_STRUTILS_SPLIT_STRING}
begin
   Result := StrUtils.SplitString(buffer, delimeter); // implicit typecast
end;
{$Else}
var
  i, from, till: integer;
  L: array of integer;
  _count, _maxcount: integer;

  procedure _add(num: integer);
  begin
    if (_count + 1 >= _maxcount) then
    begin
      inc(_maxcount, 10);
      setlength(L, _maxcount)
    end;
    L[_count] := num;
    inc(_count);
  end;

begin
  _maxcount := 20;
  _count := 0;
  SetLength(L, _maxcount);
  try
    for i := 1 to length(buffer) do
      if delimeter = buffer[i] then
        _add(i);
    _add(length(buffer) + 1);

    SetLength(Result, _count);

    from := 1;
    for i := 0 to _count - 1 do
    begin
      till := L[i];
      Result[i] := Copy(buffer, from, till - from);
      from := till + 1;
    end;
  finally
    setlength(L, 0);
  end;
end;
(* //FPC: warnings on pointer(i)
var i, from, till: integer; L: TList;
begin
  L := TList.Create;
  try
    for i := 1 to length(buffer) do
        if delimeter = buffer[i] then L.Add(pointer(i));
    L.Add(pointer(length(buffer) + 1));

    SetLength(Result, L.Count);

    from := 1;
    for i := 0 to L.Count - 1 do begin
        till := Integer(L[i]);
        Result[i] := Copy(buffer, from, till - from);
        from := till + 1;
    end;
  finally
    L.Free;
  end;
end;
*)
{$EndIf}


{$IFNDEF NOZCOLORSTRINGGRID}

// Переводит TCellStyle в TZCellStyle
//
//      CellStyle: TCellStyle   - стиль ячейки из TZColorStringGrid-а
//      XStyle: TZStyle         - итоговый стиль
//      ignorebgcolor: boolean  - игнорировать ли цвет фона ячейки
//                                true - игнорировать
//      _border: integer        - обработка рамки ячейки
//              1 - все в рамках
//              0 - все без рамок
//              2 - все с рамками, кроме стиля sgNone
procedure CellStyleTo(const CellStyle: TCellStyle; var XStyle: TZStyle; ignorebgcolor: boolean; _border: byte);
var
  i: integer;

begin
  XStyle.Font.Assign(CellStyle.Font);
  if not ignorebgcolor then
  begin
    XStyle.BGColor := CellStyle.BGColor;
    XStyle.CellPattern := ZPSolid;
  end;

  if _border > 0 then
  begin
    for i := 0 to 3 do
      if not((CellStyle.BorderCellStyle = sgNone) and (_border = 2)) then
      begin
        XStyle.Border[i].Weight := 1;
        XStyle.Border[i].LineStyle := ZEContinuous;
      end else
      begin
        XStyle.Border[i].Weight := 0;
        XStyle.Border[i].LineStyle := ZENone;
      end;
  end else
  for i := 0 to 3 do
  begin
    XStyle.Border[i].Weight := 0;
    XStyle.Border[i].LineStyle := ZENone;
  end;

  XStyle.Alignment.WrapText := CellStyle.WordWrap;
  Case CellStyle.VerticalAlignment of
    vaTop    : XStyle.Alignment.Vertical := ZVTop;
    vaCenter : XStyle.Alignment.Vertical := ZVCenter;
    vaBottom : XStyle.Alignment.Vertical := ZVBottom;
  end;
  case CellStyle.HorizontalAlignment of
    taLeftJustify  : XStyle.Alignment.Horizontal := ZHLeft;
    taRightJustify : XStyle.Alignment.Horizontal := ZHRight;
    taCenter       : XStyle.Alignment.Horizontal := ZHCenter;
  end;
  XStyle.Alignment.Rotate := CellStyle.Rotate;
end;
{$ENDIF}

//Проверка на возможность получения данных из грида
function GridCheck(var XMLSS: TZEXMLSS; const PageNum: integer;
                       GridIsNil: boolean; ToCol: integer; ToRow: integer;
                       BCol, BRow, ECol, ERow: integer): boolean;
begin
  result := true;
  if XMLSS = nil then
  begin
    result := false;
    exit;
  end;
  if GridIsNil then
  begin
    result := false;
    exit;
  end;
  if (BCol > ECol) or (BCol < 0) then
  begin
    result := false;
    exit;
  end;
  if (BRow > ERow) or (BRow < 0) then
  begin
    result := false;
    exit;
  end;
  if (ToCol < 0) or (ToRow < 0) then
  begin
    result := false;
    exit;
  end;
  if (PageNum >= XMLSS.Sheets.Count) or (PageNum < 0) then
  begin
    result := false;
    exit;
  end;
end;

//Копирует данные из TZColorStringGrid-а на страницу TZEXMLSS
//Input:
//      XMLSS: TZEXMLSS - Хранилище
//      PageNum: integer - Номер страници
//      Grid: TObject - грид, из которого нужно брать данные
//      ToCol: integer - номер столбца для вставки
//      ToRow: integer - номер строки для вставки
//      BCol: integer - верхняя левая колонка грида
//      BRow: integer - верхняя левая строка грида
//      ECol: integer - нижняя правая колонка грида
//      ERow: integer - нижняя правая строка грида
//      ignorebgcolor: boolean - игнорировать ли цвет фона ячейки
//      _border: byte - обработка рамки ячейки
//              1 - все в рамках
//              0 - все без рамок
//              2 - все с рамками, кроме стиля sgNone (только для TZColorStringGrid)
//Result:
//      true - копирование удалось
function MainGridToXmlSS(var XMLSS: TZEXMLSS; const PageNum: integer;
                     var Grid: TObject; ToCol: integer; ToRow: integer;
                         BCol, BRow, ECol, ERow: integer; ignorebgcolor: boolean; _border: byte): boolean;
var
  i, j, kolrow, kolcol, x, y: integer;
  tmpStyle: TZStyle;
  t : TRect;
  {$IFNDEF NOZCOLORSTRINGGRID}
  zGrid: TZColorStringGrid;
  {$ENDIF}
  sGrid: TStringGrid;

begin
  result := GridCheck(XMLSS, PageNum, Grid = nil, ToCol, ToRow, BCol, BRow, ECol, Erow);
  if not result then
    exit;

  tmpstyle := TZStyle.Create();
  kolrow := ERow - BRow;
  kolcol := ECol - BCol;
  if XMLSS.Sheets[PageNum].RowCount <= ToRow + kolrow then
    XMLSS.Sheets[PageNum].RowCount := ToRow + kolrow + 1;

  if XMLSS.Sheets[PageNum].ColCount <= ToCol + kolcol then
    XMLSS.Sheets[PageNum].ColCount := ToCol + kolcol + 1;
  //col, row
  //TZColorStringGrid
  {$IFNDEF NOZCOLORSTRINGGRID}
  if Grid is TZColorStringGrid then
  begin
    zGrid := Grid as TZColorStringGrid;
    if (BCol >= zGrid.ColCount) or (ECol >= zGrid.ColCount) or
       (BRow >= zGrid.RowCount) or (ERow >= zGrid.RowCount)  then
      result := false
    else
    begin
      for i := BCol to ECol do
      begin
        x := ToCol + i - BCol;
        XMLSS.Sheets[PageNum].Columns[x].WidthPix := zGrid.ColWidths[i];
        for j := BRow to ERow do
        begin
          y := ToRow + j - BRow;
          XMLSS.Sheets[PageNum].Rows[y].HeightPix := zGrid.RowHeights[j];
          CellStyleTo(zGrid.CellStyle[i, j], tmpstyle, ignorebgcolor, _border);
          XMLSS.Sheets[PageNum].Cell[x, y].Clear();
          XMLSS.Sheets[PageNum].Cell[x, y].Data := zGrid.Cells[i, j];
          XMLSS.Sheets[PageNum].Cell[x, y].CellStyle := XMLSS.Styles.Add(tmpstyle, true);
        end;
      end;
      for i := 0 to zGrid.MergeCells.Count - 1 do
      if (zGrid.MergeCells.Items[i].Left >= BCol) and (zGrid.MergeCells.Items[i].Left <= ECol) and
         (zGrid.MergeCells.Items[i].Top >= BRow) and (zGrid.MergeCells.Items[i].Top <= ERow) then
      begin
        t := zGrid.MergeCells.Items[i];
        XMLSS.Sheets[PageNum].MergeCells.AddRectXY(ToCOl + t.Left - BCol, ToRow + t.Top - BRow, ToCol + t.Right - BCol, ToRow + t.Bottom - BRow);
      end;
    end;
  end else
  {$ENDIF}
  //TStringGrid
  if Grid is TStringGrid then
  begin
    sGrid := Grid as TStringGrid;
    if (BCol >= sGrid.ColCount) or (ECol >= sGrid.ColCount) or
       (BRow >= sGrid.RowCount) or (ERow >= sGrid.RowCount)  then
      result := false
    else
    begin
      tmpstyle.Assign(XMLSS.Styles.DefaultStyle);
      if _border > 0 then
      begin
        tmpstyle.Border[0].LineStyle := ZEContinuous;
        tmpstyle.Border[0].Weight := 1;
      end else
      begin
        tmpstyle.Border[0].LineStyle := ZENone;
        tmpstyle.Border[0].Weight := 0;
      end;
      for i := 1 to 3 do
        tmpstyle.Border[i].Assign(tmpstyle.Border[0]);
      //номер стиля
      t.Left := XMLSS.Styles.Add(tmpstyle, true);
      for i := BCol to ECol do
      begin
        x := ToCol + i - BCol;
        XMLSS.Sheets[PageNum].Columns[x].WidthPix := sGrid.ColWidths[i];
        for j := BRow to ERow do
        begin
          y := ToRow + j - BRow;
          XMLSS.Sheets[PageNum].Rows[y].HeightPix := sGrid.RowHeights[j];
          XMLSS.Sheets[PageNum].Cell[x, y].Clear();
          XMLSS.Sheets[PageNum].Cell[x, y].Data := sGrid.Cells[i, j];
          XMLSS.Sheets[PageNum].Cell[x, y].CellStyle := t.Left;
        end;
      end;
    end;
  end else
    result := false;
  FreeAndNil(tmpstyle);
end;

//Копирует данные из TStringGrid-а на страницу TZEXMLSS
//Input:
//      XMLSS: TZEXMLSS - Хранилище
//      PageNum: integer - Номер страници
//      Grid: TStringGrid - грид, из которого нужно брать данные
//      ToCol: integer - номер столбца для вставки
//      ToRow: integer - номер строки для вставки
//      BCol: integer - верхняя левая колонка грида
//      BRow: integer - верхняя левая строка грида
//      ECol: integer - нижняя правая колонка грида
//      ERow: integer - нижняя правая строка грида
//      ignorebgcolor: boolean - игнорировать ли цвет фона ячейки
//      _border: byte - обработка рамки ячейки
//              1 - все в рамках
//              0 - все без рамок
//Result:
//      true - копирование удалось
function GridToXmlSS(var XMLSS: TZEXMLSS; const PageNum: integer;
                     var Grid: TStringGrid; ToCol: integer; ToRow: integer;
                         BCol, BRow, ECol, ERow: integer; ignorebgcolor: boolean; _border: byte): boolean; overload;
begin
  result := MainGridToXmlSS(XMLSS,PageNum, TObject(Grid), ToCol, ToRow, BCol, BRow,
                        ECol, ERow, ignorebgcolor, _border);
end;

{$IFNDEF NOZCOLORSTRINGGRID}

//Копирует данные из TZColorStringGrid-а на страницу TZEXMLSS
//Input:
//      XMLSS: TZEXMLSS - Хранилище
//      PageNum: integer - Номер страници
//      Grid: TZColorStringGrid - грид, из которого нужно брать данные
//      ToCol: integer - номер столбца для вставки
//      ToRow: integer - номер строки для вставки
//      BCol: integer - верхняя левая колонка грида
//      BRow: integer - верхняя левая строка грида
//      ECol: integer - нижняя правая колонка грида
//      ERow: integer - нижняя правая строка грида
//      ignorebgcolor: boolean - игнорировать ли цвет фона ячейки
//      _border: byte - обработка рамки ячейки
//              1 - все в рамках
//              0 - все без рамок
//              2 - все с рамками, кроме стиля sgNone
//Result:
//      true - копирование удалось
function GridToXmlSS(var XMLSS: TZEXMLSS; const PageNum: integer;
                     var Grid: TZColorStringGrid; ToCol: integer; ToRow: integer;
                         BCol, BRow, ECol, ERow: integer; ignorebgcolor: boolean; _border: byte): boolean; overload;
begin
  result := MainGridToXmlSS(XMLSS,PageNum, TObject(Grid), ToCol, ToRow, BCol, BRow,
                         ECol, ERow, ignorebgcolor, _border);
end;

//Копирует стиль из TZStyle в TCellStyle
//      ZZStyle: TZStyle - стиль XMLSS
//      var CellStyle: TCellStyle - стиль TZColorStringGrid-а
//      StyleCopy: integer - что копировать
//              0 - ничего из стиля не копируется
//              StyleCopy and  1 =  1 - копируется BGColor
//              StyleCopy and  2 =  2 - копируется Vertical Alignment
//              StyleCopy and  4 =  4 - копируется Horizontal Alignment
//              StyleCopy and  8 =  8 - копируется Font
//              StyleCopy and 32 = 32 - WordWrap
procedure ZStyleToCellStyle(const ZZStyle: TZStyle; CellStyle: TCellStyle; StyleCopy: integer);
begin
  if StyleCopy and 1  =  1 then
    CellStyle.BGColor := ZZStyle.BGColor;
  if StyleCopy and 2  =  2 then
  case ZZStyle.Alignment.Vertical of
    ZVTop: CellStyle.VerticalAlignment := vaTop;
    ZVCenter, ZVAutomatic, ZVJustify,
    ZVDistributed, ZVJustifyDistributed: CellStyle.VerticalAlignment := vaCenter;
    ZVBottom: CellStyle.VerticalAlignment := vaBottom;
  end;
  if StyleCopy and 4  =  4 then
  case ZZStyle.Alignment.Horizontal of
    ZHAutomatic, ZHLeft: CellStyle.HorizontalAlignment := taLeftJustify;
    ZHCenter, ZHFill, ZHJustify, ZHCenterAcrossSelection,
    ZHDistributed, ZHJustifyDistributed: CellStyle.HorizontalAlignment := taCenter;
    ZHRight: CellStyle.HorizontalAlignment := taRightJustify;
  end;
  if StyleCopy and 8  =  8 then
    CellStyle.Font.Assign(ZZStyle.Font);
  if StyleCopy and 32  =  32 then
    CellStyle.WordWrap := ZZStyle.Alignment.WrapText;
end;

//Копирует данные из страницы TZEXMLSS в TZColorStringGrid
//Input:
//      Grid: TZColorStringGrid - грид
//      XMLSS: TZEXMLSS - Хранилище
//      PageNum: integer - Номер страници
//      ToCol: integer - номер столбца для вставки
//      ToRow: integer - номер строки для вставки
//      BCol: integer - верхняя левая колонка хранилища
//      BRow: integer - верхняя левая строка хранилища
//      ECol: integer - нижняя правая колонка хранилища
//      ERow: integer - нижняя правая строка хранилища
//      InsertMode: byte - добавлять или заменять ячейки
//                      0 - ячейки в гриде не смещаются
//                      1 - ячейки в гриде смещаются вправо на размер добавляемой области
//                      2 - ячейки в гриде смещаются вниз на размер добавляемой области
//                      3 - ячейки в гриде смещаются вправо и вниз на размер добавляемой области
//      StyleCopy: integer - копирует стиль
//              0 - ничего из стиля не копируется
//              StyleCopy and   1  =    1 - копируется BGColor
//              StyleCopy and   2  =    2 - копируется Vertical Alignment
//              StyleCopy and   4  =    4 - копируется Horizontal Alignment
//              StyleCopy and   8  =    8 - копируется Font
//              StyleCopy and  16  =   16 - копируется Размер ячейки
//              StyleCopy and  32  =   32 - WordWrap
//              StyleCopy and  64  =   64 - копируется Merge Area
//              StyleCopy and  128 =  128 - удалять объединённые ячейки на месте вставки
//              StyleCopy and  256 =  256 - установить на месте вставки BorderCellStyle = sqNone
//              StyleCopy and  512 =  512 - Увеличивать высоту ячейки, если текст не помещается
//              StyleCopy and 1024 = 1024 - Увеличивать длину ячейки, если текст не помещается
//              StyleCopy and 2048 = 2048 - устанавливать default-стиль из XMLSS для
//                                          добавленных на место сдвига ячеек вместо
//                                          default-стиля из грида
//Result:
//      true - копирование удалось
function XmlSSToGrid(var Grid: TZColorStringGrid; var XMLSS: TZEXMLSS; const PageNum: integer;
                         ToCol: integer; ToRow: integer; BCol, BRow, ECol, ERow: integer; InsertMode: byte; StyleCopy: integer = 1023): boolean; overload;
var
  i, j, kolrow, kolcol, x, y, acount: integer;
  t : TRect;
  a: array of array[0..4] of integer;
  b1: boolean;
  bb: array [0..3] of boolean;

  procedure ProcessShiftMerge(xx, yy: integer);
  var
    i: integer;
    t : TRect;

  begin
    for i := 0 to Grid.MergeCells.Count - 1 do
    begin
      t := Grid.MergeCells.Items[i];
      if (t.Top >= ToRow) and (t.Left >= ToCol) then
      begin
        setlength(a, acount + 1);
        a[acount][0] := t.Left;
        a[acount][1] := t.Top;
        a[acount][2] := t.Right;
        a[acount][3] := t.Bottom;
        a[acount][4] := i;
        inc(acount);
      end;
    end;
    for i := acount - 1  downto 0 do
      Grid.MergeCells.DeleteItem(a[i][4]);
//      Grid.MergeCells.DeleteItem(Grid.MergeCells.InLeftTopCorner(a[j][0], a[j][1]));
    for i := 0 to acount - 1 do
      Grid.MergeCells.AddRectXY(a[i][0] + xx, a[i][1] + yy, a[i][2] + xx, a[i][3] + yy);
  end;

  {tut} //Просмотреть в модуле zexmlss и сделать одну нормальную функцию
  function usl(rct1, rct2: TRect): boolean;
  begin
    result := (((rct1.Left >= rct2.Left) and (rct1.Left <= rct2.Right)) or
        ((rct1.right >= rct2.Left) and (rct1.right <= rct2.Right))) and
       (((rct1.Top >= rct2.Top) and (rct1.Top <= rct2.Bottom)) or
        ((rct1.Bottom >= rct2.Top) and (rct1.Bottom <= rct2.Bottom)));
  end;

  procedure _ClearCells_(bx, ex, by, ey: integer);
  var
    i, j: integer;

  begin
    for i := by to ey do
    for j := bx to ex do
    begin
      Grid.Cells[j, i] := '';
      if StyleCopy and 2048 = 2048 then
        ZStyleToCellStyle(XMLSS.Styles.DefaultStyle, Grid.CellStyle[j, i], StyleCopy)
      else
        Grid.CellStyle[j, i].Assign(Grid.DefaultCellStyle);
        //Grid.CellStyle[j, i] := Grid.DefaultCellStyle;
    end;
  end;

begin
  result := GridCheck(XMLSS, PageNum, Grid = nil, ToCol, ToRow, BCol, BRow, ECol, ERow);
  if not result then
    exit;
  if StyleCopy < 0 then StyleCopy := 2047; //всё
  if ToCol >= Grid.ColCount then
    Grid.ColCount := ToCol;
  if ToRow >= Grid.RowCount then
    Grid.RowCount := ToRow;
  kolrow := ERow - BRow + 1;
  kolcol := ECol - BCol + 1;
  acount := 0;
  b1 := false;

  bb[0] := Grid.SizingHeight;
  bb[1] := Grid.SizingWidth;
  bb[2] := Grid.UseCellSizingHeight;
  bb[3] := Grid.UseCellSizingWidth;
  if StyleCopy and 512 = 512 then
    Grid.SizingHeight := true
  else
    Grid.SizingHeight := false;
  if StyleCopy and 1024 = 1024 then
    grid.SizingWidth := true
  else
    grid.SizingWidth := false;
  Grid.UseCellSizingHeight := false;
  Grid.UseCellSizingWidth := false;

  if StyleCopy and 128 = 128 then
    b1 := true;

  x := Grid.ColCount - 1;
  y := Grid.RowCount - 1;
  if InsertMode = 1 then
  begin
    Grid.ColCount := Grid.ColCount + kolcol;
    if Grid.RowCount <= ToRow + kolrow then
      Grid.RowCount := ToRow + kolrow;
    for i := ToRow to y do
    for j := x downto ToCol do
    begin
      Grid.Cells[j + kolcol, i] :=  Grid.Cells[j, i];
      Grid.CellStyle[j + kolcol , i].Assign(Grid.CellStyle[j, i]);
      if StyleCopy and 256 = 256 then
        Grid.CellStyle[j, i].BorderCellStyle := sgNone;
    end;
    for j := x downto ToCol do
      Grid.ColWidths[j + kolcol] := Grid.ColWidths[j];
    _ClearCells_(ToCol, ToCol + kolcol - 1, ToRow + kolrow, y);
    ProcessShiftMerge(kolcol, 0);
  end else
  if InsertMode = 2 then
  begin
    Grid.RowCount := Grid.RowCount + kolrow;
    if Grid.ColCount <= ToCol + kolcol then
      Grid.ColCount := ToCol + kolcol;
    for i := y downto ToRow do
    begin
      for j := ToCol to x do
      begin
        Grid.Cells[j, i + kolrow] :=  Grid.Cells[j, i];
        Grid.CellStyle[j, i + kolrow].Assign(Grid.CellStyle[j, i]);
        if StyleCopy and 256 = 256 then
          Grid.CellStyle[j, i].BorderCellStyle := sgNone;
      end;
      Grid.RowHeights[i + kolrow] := Grid.RowHeights[i];
    end;
    _ClearCells_(ToCol + kolcol, x, ToRow, ToRow + kolrow - 1);
    ProcessShiftMerge(0, kolrow);
  end else
  if InsertMode = 3 then
  begin
    Grid.RowCount := Grid.RowCount + kolrow;
    Grid.ColCount := Grid.ColCount + kolcol;
    for i := y downto ToRow do
    begin
      for j := x downto ToCol do
      begin
        Grid.Cells[j + kolcol, i + kolrow] :=  Grid.Cells[j, i];
        Grid.CellStyle[j + kolcol, i + kolrow].Assign(Grid.CellStyle[j, i]);
        if StyleCopy and 256 = 256 then
          Grid.CellStyle[j, i].BorderCellStyle := sgNone;
      end;
      Grid.RowHeights[i + kolrow] := Grid.RowHeights[i];
    end;
    for j := x downto ToCol do
      Grid.ColWidths[j + kolcol] := Grid.ColWidths[j];
    _ClearCells_(ToCol + kolcol, Grid.ColCount - 1, ToRow, ToRow + kolrow - 1);
    _ClearCells_(ToCol, ToCol + kolcol - 1, ToRow + kolrow, Grid.RowCount - 1);
    ProcessShiftMerge(kolcol, kolrow);
  end else
  begin
    if Grid.RowCount < ToRow + kolrow then
      Grid.RowCount := ToRow + kolrow;
    if Grid.ColCount < ToCol + kolcol then
      Grid.ColCount := ToCol + kolcol;
    if b1 then
    for i := 0 to Grid.MergeCells.Count - 1 do
    begin
      t.Left := ToCol;
      t.Top := ToRow;
      t.Right := ToCol + kolcol;
      t.Bottom := ToRow + kolrow;
      if usl(t, Grid.MergeCells.Items[i]) or usl(Grid.MergeCells.Items[i], t) then
      begin
        setlength(a, acount + 1);
        a[acount][0] := Grid.MergeCells.Items[i].Left;
        a[acount][1] := Grid.MergeCells.Items[i].Top;
        a[acount][4] := i;
        inc(acount);
      end;
    end;
  end;

    //удаление объединённых ячеек
  if b1 and ((InsertMode <= 0) or (InsertMode > 3)) then
    for i := 0 to acount - 1 do
    begin
      j := Grid.MergeCells.InMergeRange(a[i][0], a[i][1]);
      Grid.MergeCells.DeleteItem(j);
    end;
  setlength(a, 0);
  a := nil;

  //Merge
  if StyleCopy and 64 = 64 then
  for i := 0 to XMLSS.Sheets[PageNum].MergeCells.Count - 1 do
  begin
    t := XMLSS.Sheets[PageNum].MergeCells.Items[i];
    if (t.Left >= BCol) and (t.Right <= ECol) and
       (t.Top >= BRow) and (t.Bottom <= ERow) then
    begin
      t.Left := t.Left + ToCol - BCol;
      t.Right := t.Right + ToCol - BCol;
      t.Top := t.Top + ToRow - BRow;
      t.Bottom := t.Bottom + ToRow - BRow;
      Grid.MergeCells.AddRect(t);
    end;
  end;

  {tut!}
  if StyleCopy and 16 = 16 then
  for j := BRow to ERow do
    Grid.RowHeights[ToRow + j] := XMLSS.Sheets[PageNum].Rows[j].HeightPix;

  for i := BCol to ECol do
  begin
    x := ToCol + i - BCol;
    if StyleCopy and 16 = 16 then
      Grid.ColWidths[x] := XMLSS.Sheets[PageNum].Columns[i].WidthPix;
    for j := BRow to ERow do
    begin
      y := ToRow + j - BRow;
      ZStyleToCellStyle(XMLSS.Styles[XMLSS.Sheets[PageNum].Cell[i, j].CellStyle],
                        Grid.CellStyle[x, y], StyleCopy);
      Grid.Cells[x, y] := ''; //Сильно не пинать!                  
      Grid.Cells[x, y] := XMLSS.Sheets[PageNum].Cell[i, j].Data;
    end;
  end;

  Grid.SizingHeight := bb[0];
  grid.SizingWidth := bb[1];
  Grid.UseCellSizingHeight := bb[2];
  Grid.UseCellSizingWidth := bb[3];
end;
{$ENDIF}

//Копирует данные из страницы TZEXMLSS в TStringGrid
//Input:
//      Grid: TStringGrid - грид
//      XMLSS: TZEXMLSS - Хранилище
//      PageNum: integer - Номер страници
//      ToCol: integer - номер столбца для вставки
//      ToRow: integer - номер строки для вставки
//      BCol: integer - верхняя левая колонка хранилища
//      BRow: integer - верхняя левая строка хранилища
//      ECol: integer - нижняя правая колонка хранилища
//      ERow: integer - нижняя правая строка хранилища
//      InsertMode: byte - добавлять или заменять ячейки
//                      0 - ячейки в гриде не смещаются
//                      1 - ячейки в гриде смещаются вправо на размер добавляемой области
//                      2 - ячейки в гриде смещаются вниз на размер добавляемой области
//                      3 - ячейки в гриде смещаются вправо и вниз на размер добавляемой области
//      StyleCopy: integer - копирует стиль
//              StyleCopy and 16  =  16 - копируется Размер ячейки
//Result:
//      true - копирование удалось
function XmlSSToGrid(var Grid: TStringGrid; var XMLSS: TZEXMLSS; const PageNum: integer;
                         ToCol: integer; ToRow: integer; BCol, BRow, ECol, ERow: integer; InsertMode: byte; StyleCopy: integer = 511): boolean; overload;
var
  i, j, kolrow, kolcol, x, y: integer;

  procedure _ClearCells_(bx, ex, by, ey: integer);
  var
    i, j: integer;

  begin
    for i := by to ey do
    for j := bx to ex do
      Grid.Cells[j, i] := '';
  end;

begin
  result := GridCheck(XMLSS, PageNum, Grid = nil, ToCol, ToRow, BCol, BRow, ECol, ERow);
  if not result then
    exit;
  if StyleCopy < 0 then StyleCopy := 511; //всё
  if ToCol >= Grid.ColCount then
    Grid.ColCount := ToCol;
  if ToRow >= Grid.RowCount then
    Grid.RowCount := ToRow;
  kolrow := ERow - BRow + 1;
  kolcol := ECol - BCol + 1;

  x := Grid.ColCount - 1;
  y := Grid.RowCount - 1;
  if InsertMode = 1 then
  begin
    Grid.ColCount := Grid.ColCount + kolcol;
    if Grid.RowCount <= ToRow + kolrow then
      Grid.RowCount := ToRow + kolrow;
    for i := ToRow to y do
    for j := x downto ToCol do
      Grid.Cells[j + kolcol, i] :=  Grid.Cells[j, i];
    for j := x downto ToCol do
      Grid.ColWidths[j + kolcol] := Grid.ColWidths[j];
    _ClearCells_(ToCol, ToCol + kolcol - 1, ToRow + kolrow, y);
  end else
  if InsertMode = 2 then
  begin
    Grid.RowCount := Grid.RowCount + kolrow;
    if Grid.ColCount <= ToCol + kolcol then
      Grid.ColCount := ToCol + kolcol;
    for i := y downto ToRow do
    begin
      for j := ToCol to x do
        Grid.Cells[j, i + kolrow] :=  Grid.Cells[j, i];
      Grid.RowHeights[i + kolrow] := Grid.RowHeights[i];
    end;
    _ClearCells_(ToCol + kolcol, x, ToRow, ToRow + kolrow - 1);
  end else
  if InsertMode = 3 then
  begin
    Grid.RowCount := Grid.RowCount + kolrow;
    Grid.ColCount := Grid.ColCount + kolcol;
    for i := y downto ToRow do
    begin
      for j := x downto ToCol do
        Grid.Cells[j + kolcol, i + kolrow] :=  Grid.Cells[j, i];
      Grid.RowHeights[i + kolrow] := Grid.RowHeights[i];
    end;
    for j := x downto ToCol do
      Grid.ColWidths[j + kolcol] := Grid.ColWidths[j];
    _ClearCells_(ToCol + kolcol, Grid.ColCount - 1, ToRow, ToRow + kolrow - 1);
    _ClearCells_(ToCol, ToCol + kolcol - 1, ToRow + kolrow, Grid.RowCount - 1);
  end else
  begin
    if Grid.RowCount < ToRow + kolrow then
      Grid.RowCount := ToRow + kolrow;
    if Grid.ColCount < ToCol + kolcol then
      Grid.ColCount := ToCol + kolcol;
  end;

  if StyleCopy and 16 = 16 then
  for j := BRow to ERow do
    Grid.RowHeights[ToRow + j] := XMLSS.Sheets[PageNum].Rows[j].HeightPix;

  for i := BCol to ECol do
  begin
    x := ToCol + i - BCol;
    if StyleCopy and 16 = 16 then
      Grid.ColWidths[x] := XMLSS.Sheets[PageNum].Columns[i].WidthPix;
    for j := BRow to ERow do
    begin
      y := ToRow + j - BRow;
      Grid.Cells[x, y] := XMLSS.Sheets[PageNum].Cell[i, j].Data;
    end;
  end;
end;

//Сохраняет страницу TZEXMLSS в поток в формате HTML
//Input:
//      XMLSS: TZEXMLSS - Хранилище
//      PageNum: integer - Номер страницы
//      Title: string - Заголовок
//      Stream: TStream - поток
//      TextConverter: TAnsiToCodePageConverter - конвертер
//      CodePageName: string - имя кодировки
//Output:
//      0 - сохранение удалось
function SaveXmlssToHtml(var XMLSS: TZEXMLSS; const PageNum: integer; Title: string; Stream: TStream;
                         TextConverter: TAnsiToCPConverter; CodePageName: string): integer; overload;
var
  _xml: TZsspXMLWriterH;
  i, j, t, l: integer;
  NumTopLeft, NumArea: integer;
  s: string;
  Att: TZAttributesH;

function HTMLStyleTable(name: string; const Style: TZStyle): string;
var
  s: string;
  i, l: integer;

begin
  result := #13#10 + ' .' + name + '{'#13#10     ;
  for i := 0 to 3 do
  begin
    s := 'border-';
    l := 0;
    case i of
      0: s := s + 'left:';
      1: s := s + 'top:';
      2: s := s + 'right:';
      3: s := s + 'bottom:';
    end;
    s := s + '#' + ColorToHTMLHex(Style.Border[i].Color);
    if Style.Border[i].Weight <> 0 then
      s := s + ' ' + IntToStr(Style.Border[i].Weight) + 'px'
    else inc(l);
    case Style.Border[i].LineStyle of
      ZEContinuous, ZEHair: s := s + ' ' + 'solid';
      ZEDot, ZEDashDotDot: s := s + ' ' + 'dotted';
      ZEDash, ZEDashDot, ZESlantDashDot: s := s + ' ' + 'dashed';
      ZEDouble: s := s + ' ' + 'double';
      else inc(l);
    end;
    s := s + ';';
    if l <> 2 then
      result := result + s+#13#10;
  end;
  result := result + 'background:#' +  ColorToHTMLHex(Style.BGColor) + ';}';
end;

function HTMLStyleFont(name: string; const Style: TZStyle): string;
begin
  result := #13#10 + ' .' + name + '{'#13#10;
  result := result + 'color:#' + ColorToHTMLHex(Style.Font.Color) + ';';
  result := result + 'font-size:' + inttostr(Style.Font.Size) + 'px;';
  result := result + 'font-family:' + Style.Font.Name + ';}';
end;

begin
  result := 0;
  try
    _xml := TZsspXMLWriterH.Create();
    if @TextConverter <> nil then
      _xml.TextConverter := TextConverter;
    with _xml do
    begin
      TabLength := 1;
      BeginSaveToStream(Stream);
      //
      Attributes.Clear();
      WriteRaw('<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">', true, false);
      WriteTagNode('HTML',true, true, false);
      WriteTagNode('HEAD',true, true, false);
      WriteTag('TITLE',Title, true, false, false);

      //styles
      s := 'body {';
      s := s + 'background:#' + ColorToHTMLHex(XMLSS.Styles.DefaultStyle.BGColor) + ';';
      s := s + 'color:#' + ColorToHTMLHex(XMLSS.Styles.DefaultStyle.Font.Color) + ';';
      s := s + 'font-size:' + inttostr(XMLSS.Styles.DefaultStyle.Font.Size) + 'px;';
      s := s + 'font-family:' + XMLSS.Styles.DefaultStyle.Font.Name + ';}';
      s := s + HTMLStyleTable('T19', XMLSS.Styles.DefaultStyle);
      s := s + HTMLStyleFont('F19', XMLSS.Styles.DefaultStyle);
      for i := 0 to XMLSS.Styles.Count - 1 do
      begin
        s := s + HTMLStyleTable('T' + IntToStr(i + 20), XMLSS.Styles[i]);
        s := s +  HTMLStyleFont('F' + IntToStr(i + 20), XMLSS.Styles[i]);
      end;

      WriteTag('STYLE', s, true, true, false);
      Attributes.Add('HTTP-EQUIV', 'CONTENT-TYPE');

      s := '';
      if trim(CodePageName) > '' then
        s := '; CHARSET='+ CodePageName;
      Attributes.Add('CONTENT', 'TEXT/HTML'+s);
      WriteTag('META','', true, false, false);
      WriteEndTagNode(); // HEAD

      //BODY
      Attributes.Clear();
      WriteTagNode('BODY',true, true, false);

      //Table
      Attributes.Add('cellSpacing', '0');
      Attributes.Add('border','0');
      Attributes.Add('width','100%');
      WriteTagNode('TABLE',true, true, false);

      Att := TZAttributesH.Create();
      Att.Clear();
      for i := 0 to XMLSS.Sheets[PageNum].RowCount - 1 do
      begin
        Attributes.Clear();
        WriteTagNode('TR',true, true, true);
        for j := 0 to XMLSS.Sheets[PageNum].ColCount - 1 do
        begin
          NumTopLeft := XMLSS.Sheets[PageNum].MergeCells.InLeftTopCorner(j, i);
          NumArea := XMLSS.Sheets[PageNum].MergeCells.InMergeRange(j, i);
          // если ячейка входит в объединённые области и не является
          // верхней левой ячейков в этой области - пропускаем её
          if not ((NumArea >= 0) and (NumTopLeft = -1)) then
          begin
            Attributes.Clear();
            if NumTopLeft >= 0 then
            begin
              {tut}
              t := XMLSS.Sheets[PageNum].MergeCells.Items[NumTopLeft].Right -
                   XMLSS.Sheets[PageNum].MergeCells.Items[NumTopLeft].Left;
              if t > 0 then
                Attributes.Add('colspan', InttOstr(t+1));
              t := XMLSS.Sheets[PageNum].MergeCells.Items[NumTopLeft].Bottom -
                   XMLSS.Sheets[PageNum].MergeCells.Items[NumTopLeft].Top;
              if t > 0 then
                Attributes.Add('rowspan', InttOstr(t+1));
            end;
            t := XMLSS.Sheets[PageNum].Cell[j,i].CellStyle;
            if XMLSS.Styles[t].Alignment.Horizontal = ZHCenter then
              Attributes.Add('align', 'center') else
            if XMLSS.Styles[t].Alignment.Horizontal = ZHRight then
              Attributes.Add('align', 'right') else
            if XMLSS.Styles[t].Alignment.Horizontal = ZHJustify then
              Attributes.Add('align', 'justify');
            Attributes.Add('class', 'T'+IntToStr(t + 20));
            Attributes.Add('width',inttostr(XMLSS.Sheets[PageNum].Columns[j].WidthPix)+'px');

            WriteTagNode('TD', true, false, false);
            Attributes.Clear();
            Att.Clear();
            Att.Add('class', 'F' + IntToStr(t + 20));
            if fsbold in XMLSS.Styles[t].Font.Style then
              WriteTagNode('B', false, false, false);
            if fsItalic in XMLSS.Styles[t].Font.Style then
              WriteTagNode('I', false, false, false);
            if fsUnderline in XMLSS.Styles[t].Font.Style then
              WriteTagNode('U', false, false, false);
            if fsStrikeOut in XMLSS.Styles[t].Font.Style then
              WriteTagNode('S', false, false, false);

            l := Length(XMLSS.Sheets[PageNum].Cell[j, i].Href);
            if l > 0 then
            begin
              Attributes.Add('href', XMLSS.Sheets[PageNum].Cell[j, i].Href);
              //target?
              WriteTagNode('A', false, false, false);
              Attributes.Clear();
            end;

            WriteTag('FONT', XMLSS.Sheets[PageNum].Cell[j, i].Data, Att, false, false, true);
            if l > 0 then
              WriteEndTagNode(); // A

            if fsbold in XMLSS.Styles[t].Font.Style then
              WriteEndTagNode(); // B
            if fsItalic in XMLSS.Styles[t].Font.Style then
              WriteEndTagNode(); // I
            if fsUnderline in XMLSS.Styles[t].Font.Style then
              WriteEndTagNode(); // U
            if fsStrikeOut in XMLSS.Styles[t].Font.Style then
              WriteEndTagNode(); // S
            WriteEndTagNode(); // TD
          end;

        end;
        WriteEndTagNode(); // TR
      end;

      WriteEndTagNode(); // BODY
      WriteEndTagNode(); // HTML
      EndSaveTo();
    end;
    FreeAndNil(Att);
  finally
    FreeAndNil(_xml);
  end;
end;

//Сохраняет страницу TZEXMLSS в файл FileName
//Input:
//      XMLSS: TZEXMLSS - Хранилище
//      PageNum: integer - Номер страницы
//      Title: string - Заголовок
//      FileName: String - имя файла
//      TextConverter: TAnsiToCodePageConverter - конвертер
//      CodePageName: string - имя кодировки
//Output:
//      0 - сохранение удалось
function SaveXmlssToHtml(var XMLSS: TZEXMLSS; const PageNum: integer; Title: string; FileName: string;
                         TextConverter: TAnsiToCPConverter; CodePageName: string): integer; overload;
var
  Stream: TStream;

begin
  result := 0;
  Stream := nil; // это чтобы варнинги не кричали
  try
    try
      Stream := TFileStream.Create(FileName, fmCreate);
    except
      result := 1;
    end;
    if result = 0 then
      result := SaveXmlssToHtml(XMLSS, PageNum, Title, Stream, TextConverter, CodePageName);
  finally
    if Stream <> nil then
      Stream.Free;
  end;
end;

const RepeatablePrintedHeadersName = 'Print_Titles';

//Сохраняет в поток в формате Excel XML SpreadSheet
//      XMLSS: TZEXMLSS                 - хранилище
//      Stream: TStream                 - поток
//      SheetsNumbers: array of integer - массив номеров страниц в нужной последовательности
//      SheetsNames: array of string    - массив названий страниц
//              количество элементов в двух массивах должны совпадать
//      TextConverter: TAnsiToCodePageConverter - конвертер
//      CodePageName: string - имя кодировки
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; Stream: TStream; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
var
  _xml: TZsspXMLWriterH;
  _pages: TIntegerDynArray;    //номера страниц
  _names: TStringDynArray;    //названия страниц
  kol: integer;               //количество страниц
  i: integer;
  _FmtParser: TNumFormatParser;
  _DateParser: TZDateTimeODSFormatParser;


  //заголовок xml-ины
  procedure WriteHeader();
  begin
    with _xml do
    begin
      ZEWriteHeaderCommon(_xml, CodePageName, BOM);
      Attributes.Add('progid','Excel.Sheet');
      WriteInstruction('mso-application',true);
      Attributes.Clear();
      Attributes.Add('xmlns','urn:schemas-microsoft-com:office:spreadsheet');
      Attributes.Add('xmlns:o','urn:schemas-microsoft-com:office:office', false);
      Attributes.Add('xmlns:x','urn:schemas-microsoft-com:office:excel', false);
      Attributes.Add('xmlns:ss','urn:schemas-microsoft-com:office:spreadsheet', false);
      Attributes.Add('xmlns:html','http://www.w3.org/TR/REC-html40', false);
      WriteTagNode('Workbook', true, true, false);
      Attributes.Clear();
      WriteTagNode('DocumentProperties',[ToAttribute('xmlns','urn:schemas-microsoft-com:office:office')], true, true, false);
      WriteTag('Author', XMLSS.DocumentProperties.Author);
      WriteTag('LastAuthor', XMLSS.DocumentProperties.LastAuthor);
      WriteTag('Created', ZEDateTimeToStr(XMLSS.DocumentProperties.Created) + 'Z'{s});
      WriteTag('Company', XMLSS.DocumentProperties.Company);

      WriteTag('Version', XMLSS.DocumentProperties.Version); // зачем ???
//      WriteTag('Version', ZELibraryVersion); // Excel 2010 writes "14.00" here despite specifications say Integer
      WriteTag('NameOfApplication', ZELibraryName);

      WriteEndTagNode();
      WriteTagNode('ExcelWorkbook',[ToAttribute('xmlns','urn:schemas-microsoft-com:office:excel')], true, true, false);
      WriteTag('WindowHeight', inttostr(XMLSS.DocumentProperties.WindowHeight), true, false, false);
      WriteTag('WindowWidth', inttostr(XMLSS.DocumentProperties.WindowWidth), true, false, false);
      WriteTag('WindowTopX', inttostr(XMLSS.DocumentProperties.WindowTopX), true, false, false);
      WriteTag('WindowTopY', inttostr(XMLSS.DocumentProperties.WindowTopY), true, false, false);
      if XMLSS.DocumentProperties.ModeR1C1 then
        _xml.WriteEmptyTag('RefModeR1C1');
      WriteTag('ProtectStructure', 'False');
      WriteTag('ProtectWindows', 'False');
      WriteEndTagNode();
    end;
  end;

  //добавить целый атрибут
  // Input:
  //    _name: string           - имя атрибута
  //    Value: integer          - Значение атрибута
  //    _defvalue: integer      - Значение по-умолчанию
  //    _defStyle: integer      - Дефолтный стиль
  //    _def: boolean           - Является ли рассматриваемый стиль дефолтным
  procedure AddAttribute(const _name: string; const Value: integer; const _defvalue: integer; const _defStyle: integer; _def: boolean); overload;
  begin
    if _def then
    begin
      if value <> _defvalue then
        _xml.Attributes.Add(_name, inttostr(Value), false);
    end else
    begin
      if Value <> _defStyle then
        _xml.Attributes.Add(_name, inttostr(Value), false);
    end;
  end;

  //добавить логический атрибут
  // Input:
  //    _name: string           - имя атрибута
  //    Value: boolean          - Значение атрибута
  //    _defvalue: boolean      - Значение по-умолчанию
  //    _defStyle: boolean      - Дефолтный стиль
  //    _def: boolean           - Является ли рассматриваемый стиль дефолтным
  procedure AddAttribute(const _name: string; Value: boolean; _defvalue: boolean; _defStyle: boolean; _def: boolean); overload;
  var
    t: string;

  begin
    if value then t := '1' else t := '0';
    if _def then
    begin
      if value <> _defvalue then
        _xml.Attributes.Add(_name, t, false);
    end else
    begin
      if Value <> _defStyle then
        _xml.Attributes.Add(_name, t, false);
    end;
  end;

  //добавить string атрибут
  // Input:
  //    _name: string           - имя атрибута
  //    Value: string          - Значение атрибута
  //    _defvalue: string      - Значение по-умолчанию
  //    _defStyle: string      - Дефолтный стиль
  //    _def: boolean           - Является ли рассматриваемый стиль дефолтным
  procedure AddAttribute(const _name: string; const Value: string; const _defvalue: string; const _defStyle: string; _def: boolean); overload;
  begin
    if _def then
    begin
      if value <> _defvalue then
        _xml.Attributes.Add(_name, Value, false);
    end else
    begin
      if Value <> _defStyle then
        _xml.Attributes.Add(_name, Value, false);
    end;
  end;

  //добавить TColor атрибут
  // Input:
  //    _name: string           - имя атрибута
  //    Value: integer          - Значение атрибута
  //    _defvalue: integer      - Значение по-умолчанию
  //    _defStyle: integer      - Дефолтный стиль
  //    _def: boolean           - Является ли рассматриваемый стиль дефолтным
  procedure AddAttributeColor(const _name: string; const Value: TColor; const _defvalue: TColor; const _defStyle: TColor; _def: boolean);
  begin
    if _def then
    begin
      if value <> _defvalue then
        _xml.Attributes.Add(_name, '#'+ColorToHTMLHex(Value), false);
    end else
    begin
      if Value <> _defStyle then
        _xml.Attributes.Add(_name, '#'+ColorToHTMLHex(Value), false);
    end;
  end;

  function AddBorders(const _border: TZborder; _def: boolean): integer;
  var
    i: integer;
    t: TZAttributesH;

    //необязательные атрибуты
    procedure NonReqBorderAttr(num: integer);
    begin
      //Color
      AddAttributeColor('ss:Color', _border[num].Color, clBlack, XMLSS.Styles.DefaultStyle.Border[num].Color, _def);
      //LineStyle
      if _def then
      begin
        if _border[num].LineStyle <> ZENone then
          _xml.Attributes.Add('ss:LineStyle', ZBorderTypeToStr(_border[num].LineStyle), false);
      end else
      begin
        if _border[num].LineStyle <> XMLSS.Styles.DefaultStyle.Border[num].LineStyle then
          _xml.Attributes.Add('ss:LineStyle', ZBorderTypeToStr(_border[num].LineStyle), false);
      end;
      //Weight
      AddAttribute('ss:Weight', _border[num].Weight, 0, 0, _def);
    end;

  begin
     result := 0;
     for i := 0 to 5 do
     begin
       _xml.Attributes.Clear();
       if not(((_border[i].LineStyle = ZEContinuous) or (_border[i].LineStyle = ZEHair)) and 
              (_border[i].Weight = 0)) then
       begin
         case i of
           0: _xml.Attributes.Add('ss:Position','Left', false);
           1: _xml.Attributes.Add('ss:Position','Top', false);
           2: _xml.Attributes.Add('ss:Position','Right', false);
           3: _xml.Attributes.Add('ss:Position','Bottom', false);
           4: _xml.Attributes.Add('ss:Position','DiagonalLeft', false);
           5: _xml.Attributes.Add('ss:Position','DiagonalRight', false);
         end;
         NonReqBorderAttr(i);
       end;
       if _xml.Attributes.Count > 1 then
       begin
         result := result + 1;
         if result = 1 then
         begin
           t := TZAttributesH.Create();
           _xml.WriteTagNode('Borders', t, true, true, true);
           FreeAndNil(t);
         end;
         _xml.WriteEmptyTag('Border');
       end;
     end;
  end;

  procedure AddFont(const _font: TFont; _def: boolean);
  var
    b: boolean;
    s1, s2: string;

  begin
    _xml.Attributes.Clear();
    //FontName
    _xml.Attributes.Add('ss:FontName', _font.Name, false);
    //AddAttribute('ss:FontName', _font.Name, 'Arial', Styles.DefaultStyle.Font.Name, _def);
    //x:CharSet
    _xml.Attributes.Add('x:CharSet', inttostr(_font.Charset), false);
    //AddAttribute('x:CharSet', _font.Charset, 0, Styles.DefaultStyle.Font.Charset, _def);
    //bold
    if fsBold in XMLSS.Styles.DefaultStyle.Font.Style then
      b := true
    else
      b := false;
    if fsBold in _font.Style then
      AddAttribute('ss:Bold', true, false, b, _def);
    //Color
    AddAttributeColor('ss:Color', _font.Color, clBlack, XMLSS.Styles.DefaultStyle.Font.Color, _def);
    //Italic
    if fsItalic in XMLSS.Styles.DefaultStyle.Font.Style then
      b := true
    else
      b := false;
    if fsItalic in _font.Style then
      AddAttribute('ss:Italic', true, false, b, _def);
    //Size
    AddAttribute('ss:Size', _font.Size, 10, XMLSS.Styles.DefaultStyle.Font.Size, _def);
    //StrikeThrough
    if fsStrikeOut in XMLSS.Styles.DefaultStyle.Font.Style then
      b := true
    else
      b := false;
    if fsStrikeOut in _font.Style then
      AddAttribute('ss:StrikeThrough', true, false, b, _def);
    //Underline
    if fsUnderLine in XMLSS.Styles.DefaultStyle.Font.Style then
      s1 := 'Single'
    else
      s1 := 'None';
    if fsUnderLine in _font.Style then
      s2 := 'Single'
    else
      s2 := 'None';
    AddAttribute('ss:Underline', s2, 'None', s1, _def);
    if _xml.Attributes.Count > 0 then
      _xml.WriteEmptyTag('Font', true, true);
  end;

  //записать стиль
  procedure WriteStyle(const _id: string; const _name: string; const _style: TZStyle; _def: boolean);
  begin
    with _xml do
    begin
      Attributes.Clear();
      Attributes.Add('ss:ID', _id, false);
      if (_name > '') then
        Attributes.Add('ss:Name', _name, false);
      WriteTagNode('Style', true, true);
      Attributes.Clear();
      //=====aligment======
      // Horizontal / Vertical
      if _def then
      begin
        if (_style.Alignment.Horizontal <> ZHAutomatic) then
          Attributes.Add('ss:Horizontal', HAlToStr(_Style.Alignment.Horizontal), false);
        if (_style.Alignment.Vertical <> ZVAutomatic) then
          Attributes.Add('ss:Vertical', VAlToStr(_Style.Alignment.Vertical), false);
      end else
      begin
        if _style.Alignment.Horizontal <> XMLSS.Styles.DefaultStyle.Alignment.Horizontal then
          Attributes.Add('ss:Horizontal', HAlToStr(_Style.Alignment.Horizontal), false);
        if _style.Alignment.Vertical <> XMLSS.Styles.DefaultStyle.Alignment.Vertical then
          Attributes.Add('ss:Vertical', VAlToStr(_Style.Alignment.Vertical), false);
      end;

      // Indent
      AddAttribute('ss:Indent', _style.Alignment.Indent, 0, XMLSS.Styles.DefaultStyle.Alignment.Indent, _def);
      // Rotate
      AddAttribute('ss:Rotate', zeNormalizeAngle90(_style.Alignment.Rotate), 0,
                                zeNormalizeAngle90(XMLSS.Styles.DefaultStyle.Alignment.Rotate), _def);
      // ShrinkToFit
      AddAttribute('ss:ShrinkToFit', _style.Alignment.ShrinkToFit, false, XMLSS.Styles.DefaultStyle.Alignment.ShrinkToFit, _def);
      // VerticalText
      AddAttribute('ss:VerticalText', _style.Alignment.VerticalText, false, XMLSS.Styles.DefaultStyle.Alignment.VerticalText, _def);
      // WrapText
      AddAttribute('ss:WrapText', _style.Alignment.WrapText, false, XMLSS.Styles.DefaultStyle.Alignment.WrapText, true);
      if Attributes.Count > 0 then
        WriteEmptyTag('Alignment', true, true);
      //=====Borders======
      if AddBorders(_style.Border, _def) > 0 then
        WriteEndTagNode();
      //=====Font======
      AddFont(_style.Font, _def);
      //=====Interior======
      Attributes.Clear();
      AddAttributeColor('ss:Color', _style.BGColor, ClWindow, XMLSS.styles.DefaultStyle.BGColor, _def);
      if _def then
      begin
        if (_style.CellPattern <> ZPNone) then
          Attributes.Add('ss:Pattern', ZCellPatternToStr(_style.CellPattern), false);
      end else
      begin
        if _style.CellPattern <> XMLSS.Styles.DefaultStyle.CellPattern then
          Attributes.Add('ss:Pattern', ZCellPatternToStr(_style.CellPattern), false);
      end;
      //ss:PatternColor
      AddAttributeColor('ss:PatternColor', _style.PatternColor, ClWindow, XMLSS.styles.DefaultStyle.PatternColor, _def);
      if Attributes.Count > 0 then
        WriteEmptyTag('Interior', true, true);
      //=====NumberFormat======
      if (_style.NumberFormat <> '') then
      begin
        Attributes.Clear();
        AddAttribute('ss:Format', ConvertFormatNativeToXlsx(_style.NumberFormat, _FmtParser, _DateParser), 'General', ConvertFormatNativeToXlsx(XMLSS.Styles.DefaultStyle.NumberFormat, _FmtParser, _DateParser), _def);
        if Attributes.Count > 0 then
          WriteEmptyTag('NumberFormat', true, true);
      end;
      //=====Protection========
      Attributes.Clear();
      AddAttribute('x:HideFormula', _style.HideFormula, false, XMLSS.Styles.DefaultStyle.HideFormula, _def);
      AddAttribute('ss:Protected', _style.Protect, true, XMLSS.Styles.DefaultStyle.Protect, _def);
      if Attributes.Count > 0 then
        WriteEmptyTag('Protection', true, true);
      WriteEndTagNode();
    end;
  end;

  //пишет все стили
  //    TODO:  может, не сохранять те стили, которые не используются
  //           на страницах? (или сделать на выбор)
  procedure WriteStyles();
  var
    i: integer;

  begin
    _xml.Attributes.Clear();
    _xml.WriteTagNode('Styles', true, true);
    WriteStyle('Default', 'Normal', XMLSS.Styles.DefaultStyle, true);

    _FmtParser := TNumFormatParser.Create();
    _DateParser := TZDateTimeODSFormatParser.Create();

    try
      for i := 0 to XMLSS.Styles.Count - 1 do
        WriteStyle('s'+inttostr(i+20), '', XMLSS.Styles[i], false);
    finally
      FreeAndNil(_FmtParser);
      FreeAndNil(_DateParser);
    end;

    _xml.WriteEndTagNode();
  end;

  //Сохраняет страницу Sheets[PageNum] с подписью PageName
  procedure WritePage(const PageNum: integer; const PageName: string);
  var
    i, j, t: integer;
    b, CellIndex: boolean;
    CountCells: integer;
    AttrCell, AttrData, AttrComment, AttrCommentData, AttrPrintTitles: TZAttributesH;
    NumTopLeft, NumArea: integer;
    kol: integer;
    s: string;
    isFormula, isRepeatablePrint: boolean;
    ProcessedSheet: TZSheet; ProcessedCell: TZCell;
    _hMode, _vMode: TZSplitMode;

    procedure AddCellInRow(var CountCells: integer; var CellIndex: boolean);
    begin
      if CountCells = 0 then
      begin
        if _xml.Attributes.Count = 1 then
          _xml.WriteTagNode('Row', true, true, true);
      end;
      inc(CountCells);
      CellIndex := false;
    end;

    //Разрывы страницы
    procedure WritePageBreaks();
    var
      i: integer;
      b, b1: boolean;

      procedure WriteBreak(Breaked: boolean; s1, s2, s3: string; var b: boolean; var b1: boolean; num: integer);
      begin
        if Breaked then
        begin
          if not b then
          begin
           _xml.WriteTagNode('PageBreaks', AttrCell, true, true, false);
            b := true;
          end;
          if not b1 then
          begin
            _xml.WriteTagNode(s1, true, true, false);
            b1 := true;
          end;
          _xml.WriteTagNode(s2, true, true, false);
          _xml.WriteTag(s3, inttostr(num), true, false, false);
          _xml.WriteEndTagNode(); //s2
        end;
      end;

    begin
      with _xml do
      begin
        AttrCell.Clear();
        AttrCell.Add('xmlns', 'urn:schemas-microsoft-com:office:excel');
        Attributes.Clear();

        b := false;
        b1 := false;
        for i := 1 to ProcessedSheet.ColCount - 1 do
          WriteBreak(ProcessedSheet.Columns[i].Breaked, 'ColBreaks', 'ColBreak', 'Column', b, b1, i);
        if b1 then
          WriteEndTagNode(); //ColBreaks
        b1 := false;
        for i := 1 to ProcessedSheet.RowCount - 1 do
          WriteBreak(ProcessedSheet.Rows[i].Breaked, 'RowBreaks', 'RowBreak', 'Row', b, b1, i);
        if b1 then
          WriteEndTagNode(); //RowBreaks
        if b then
          WriteEndTagNode(); //PageBreaks
      end;
    end; //WritePageBreaks (Разрывы страницы)

    procedure _WriteSplitFrozen(_xml: TZsspXMLWriterH; SplitMode: TZSplitMode; const SplitValue: integer; const SplitName, SplitPane: string);
    var
      s: string;

    begin
      if ((SplitValue <> 0) and (SplitMode <> ZSplitNone)) then
      begin
        if (SplitMode = ZSplitSplit) then
          s := IntToStr(round(PixelToPoint(SplitValue) * 20))
        else
          s := IntToStr(SplitValue);
        _xml.WriteTag(SplitName, s, true, false, false);
        if (SplitMode = ZSplitFrozen) then
          _xml.WriteTag(SplitPane, s, true, false, false);
      end;
    end; //_WriteSplitFrozen

  begin
    ProcessedSheet := XMLSS.Sheets[PageNum];

    with _xml do
    begin
      Attributes.Clear();
      Attributes.Add('ss:Name',PageName);
      AddAttribute('ss:Protected', ProcessedSheet.Protect, false, false, true);
      AddAttribute('ss:RightToLeft', ProcessedSheet.RightToLeft, false, false, true);
      WriteTagNode('Worksheet', true, true, true);

      // Repeatable on print rows/columns
      if ProcessedSheet.RowsToRepeat.Active or ProcessedSheet.ColsToRepeat.Active then
      begin
        Attributes.Clear();
        WriteTagNode('Names', []);
        Attributes.Add( 'ss:Name', RepeatablePrintedHeadersName);

        s := ProcessedSheet.ColsToRepeat.ToString + ',' + ProcessedSheet.RowsToRepeat.ToString;
        //at least one of those calls should return non-empty street

        t := Length(s);
        if s[t] = ',' then SetLength(s, t-1); // no rows if ends with comma

        if s[1] = ','  // no columns if starts with comma
           then s[1] := '='
           else s := '=' + s;

        Attributes.Add( 'ss:RefersTo', s);
        WriteEmptyTag('NamedRange');
        WriteEndTagNode(); // Names
      end;

      //Table
      Attributes.Clear();
      Attributes.Add('ss:ExpandedColumnCount', inttostr(ProcessedSheet.ColCount), false);
      Attributes.Add('ss:ExpandedRowCount', inttostr(ProcessedSheet.RowCount), false);
      if round(ProcessedSheet.DefaultColWidth*100) <> 4800 then
        Attributes.Add('ss:DefaultColumnWidth', ZEFloatSeparator(FormatFloat('0.#####',ProcessedSheet.DefaultColWidth)), false);
      if round(ProcessedSheet.DefaultRowHeight*100) <> 1275 then
        Attributes.Add('ss:DefaultRowHeight', ZEFloatSeparator(FormatFloat('0.#####',ProcessedSheet.DefaultRowHeight)), false);
      Attributes.Add('x:FullRows', '1', false);
      Attributes.Add('x:FullColumns', '1', false);
      WriteTagNode('Table', true, true, true);
      //Columns
      b := false;
      for i := 0 to ProcessedSheet.ColCount - 1 do
      begin
        Attributes.Clear();
        if round(ProcessedSheet.DefaultColWidth*100) <> round(ProcessedSheet.Columns[i].Width*100) then
          Attributes.Add('ss:Width', ZEFloatSeparator(FormatFloat('0.#####',ProcessedSheet.Columns[i].Width)), false);
        if (ProcessedSheet.Columns[i].StyleID <> -1) and (ProcessedSheet.Columns[i].StyleID < XMLSS.Styles.Count) then
          AddAttribute('ss:StyleID', 's' + IntToStr(ProcessedSheet.Columns[i].StyleID + 20), '', '', true);
        AddAttribute('ss:AutoFitWidth', ProcessedSheet.Columns[i].AutoFitWidth, true, true, true);
        AddAttribute('ss:Hidden', ProcessedSheet.Columns[i].Hidden, false, false, true);
        if (Attributes.Count > 0) and (b) then
          Attributes.Insert(0, 'ss:Index', inttostr(i+1));
        if Attributes.Count > 0 then
        begin
          WriteEmptyTag('Column', true, false);
          b := false;
        end else
          b := true;
      end;

      //rows
      AttrCell := TZAttributesH.Create();
      AttrData := TZAttributesH.Create();
      AttrComment := TZAttributesH.Create();
      AttrCommentData := TZAttributesH.Create();

      AttrPrintTitles := TZAttributesH.Create();
        AttrPrintTitles.Add('ss:Name', RepeatablePrintedHeadersName);

      for i := 0 to ProcessedSheet.RowCount - 1 do
      begin
        Attributes.Clear();
        Attributes.Add('ss:Index', inttostr(i+1), false);
        if round(ProcessedSheet.DefaultRowHeight*100) <> round(ProcessedSheet.Rows[i].Height*100) then
          Attributes.Add('ss:Height', ZEFloatSeparator(FormatFloat('0.#####',ProcessedSheet.Rows[i].Height)), false);
        if (ProcessedSheet.Rows[i].StyleID <> -1) and (ProcessedSheet.Rows[i].StyleID < XMLSS.Styles.Count) then
          AddAttribute('ss:StyleID', 's' + IntToStr(ProcessedSheet.Rows[i].StyleID + 20), '', '', true);
        AddAttribute('ss:AutoFitHeight', ProcessedSheet.Rows[i].AutoFitHeight, true, true, true);
        AddAttribute('ss:Hidden', ProcessedSheet.Rows[i].Hidden, false, false, true);

        if (Attributes.Count > 1)  then
          WriteTagNode('Row', true, true, true);

        //Пробегаем по всем ячейкам
        CountCells := 0;
        CellIndex := false;
        for j := 0 to ProcessedSheet.ColCount - 1 do
        begin
          ProcessedCell := ProcessedSheet.Cell[j, i];
          AttrCell.Clear();
          if CellIndex then
            AttrCell.Add('ss:Index', IntToStr(j+1), false);

          NumTopLeft := ProcessedSheet.MergeCells.InLeftTopCorner(j, i);
          NumArea := ProcessedSheet.MergeCells.InMergeRange(j, i);

          // если ячейка входит в объединённые области и не является
          // верхней левой ячейков в этой области - пропускаем её
          if not ((NumArea >= 0) and (NumTopLeft = -1)) then
          begin
            if NumTopLeft >= 0 then
            begin
              //ss:MergeAcross - влево
              t := ProcessedSheet.MergeCells.Items[NumTopLeft].Right -
                   ProcessedSheet.MergeCells.Items[NumTopLeft].Left;
              if t > 0 then
                AttrCell.Add('ss:MergeAcross', InttOstr(t), false);
              //ss:MergeDown - вниз
              t := ProcessedSheet.MergeCells.Items[NumTopLeft].Bottom -
                   ProcessedSheet.MergeCells.Items[NumTopLeft].Top;
              if t > 0 then
                AttrCell.Add('ss:MergeDown', InttOstr(t), false);
            end;
            {tut}
            //Сделать проверку на формулу?
            isFormula := false;
            if (ProcessedCell.Formula > '') then
            begin
              AttrCell.Add('ss:Formula', ProcessedCell.Formula, false);
              isFormula := true;
            end;
            //Href
            if (ProcessedCell.Href > '') then
            begin
              AttrCell.Add('ss:HRef', ProcessedCell.Href, false);
              if (ProcessedCell.HRefScreenTip > '') then
                AttrCell.Add('x:HRefScreenTip', ProcessedCell.HRefScreenTip, false);
            end;
            //StyleID
            if (ProcessedCell.CellStyle <> -1) and
               (ProcessedCell.CellStyle < XMLSS.Styles.Count) then
               AttrCell.Add('ss:StyleID', 's'+inttostr(20 + ProcessedCell.CellStyle), false);

            kol := 0;

            /// Repeatable printed Rows/Columns
            isRepeatablePrint := false;
            if ( ProcessedSheet.RowsToRepeat.Active and
                 (ProcessedSheet.RowsToRepeat.From <= i) and
                 (ProcessedSheet.RowsToRepeat.Till >= i) )
            or ( ProcessedSheet.ColsToRepeat.Active and
                 (ProcessedSheet.ColsToRepeat.From <= j) and
                 (ProcessedSheet.ColsToRepeat.Till >= j) )
            then begin
                 inc(kol);
                 isRepeatablePrint := true;
            end;

            /// DATA
            AttrData.Clear();
            // для OpenOffice Calc добавлен текст "or (j = ProcessedSheet.ColCount - 1)"
            // т.к. в ОО таблицы отображались в некоторых случаях некорректно
            if (ProcessedCell.Data > '') or (j = ProcessedSheet.ColCount - 1) or
               (isFormula) then
            begin
              AttrData.Add('ss:Type', ZCellTypeToStr(ProcessedCell.CellType), false);
              inc(kol);
            end;

            /// Comment
            {tut}  //Добавить проверки
            AttrComment.Clear();
            AttrCommentData.Clear();
            if ProcessedCell.ShowComment then
            begin
              if (ProcessedCell.CommentAuthor > '') then
                AttrComment.Add('ss:Author', ProcessedCell.CommentAuthor, false);
              if ProcessedCell.AlwaysShowComment then
                AttrComment.Add('ss:ShowAlways','1', false);
              inc(kol);
            end;

            {tut}
            // подумай насчёт
            // xmlns="http://www.w3.org/TR/REC-html40" для текста ячеек и комментария

            //если kol = 0 то пустой тег рисуем
            if kol > 0 then
            begin
              AddCellInRow(CountCells, CellIndex);
              //Cell
              WriteTagNode('Cell', AttrCell,true, true, false);

              if isRepeatablePrint then
                  WriteEmptyTag('NamedCell', AttrPrintTitles);
              //Data
              if (ProcessedCell.Data > '') or
                 (isFormula) then
              begin
                if (ProcessedCell.CellType = ZEBoolean) then
                begin
                  if (ZETryStrToBoolean(ProcessedCell.Data)) then
                    s := '1'
                  else
                    s := '0';
                end
                else
                  CorrectStrForXML(ProcessedCell.Data, s, b);

                if b then
                  AttrData.Add('xmlns','http://www.w3.org/TR/REC-html40', false);
                WriteTag('ss:Data', s, AttrData, true, false, false);
              end;
              //Comment
              if ProcessedCell.ShowComment then
              begin
                WriteTagNode('Comment', AttrComment, true, true, true);
                CorrectStrForXML(ProcessedCell.Comment, s, b);
                if b then
                  AttrCommentData.Add('xmlns','http://www.w3.org/TR/REC-html40');
                WriteTag('ss:Data', s, AttrCommentData, true, false, false);
                WriteEndTagNode(); //Comment
              end;
              WriteEndTagNode(); //Cell
            end else
            begin
              if AttrCell.Count >= 1 then
              begin
                if not((AttrCell.Count = 1) and (AttrCell.Items[0][0] = 'ss:Index')) then
                begin
                  AddCellInRow(CountCells, CellIndex);
                  WriteEmptyTag('Cell', AttrCell, true, false);
                end else
                  CellIndex := true;  
              end else
                CellIndex := true;
            end;
          end else
            CellIndex := true;
        end;

        if (Attributes.Count > 1) or (CountCells > 0) then
        begin
          WriteEndTagNode(); //Row
        end;
      end; //for i

      WriteEndTagNode(); //Table

      //WorksheetOptions
      Attributes.Clear();
      Attributes.Add('xmlns','urn:schemas-microsoft-com:office:excel');
      WriteTagNode('WorksheetOptions',true, true, true);
      Attributes.Clear();
      WriteTagNode('PageSetup',true, true, true);

      //Layout
      if not ProcessedSheet.SheetOptions.PortraitOrientation then
        Attributes.Add('x:Orientation', 'Landscape', false);
      if ProcessedSheet.SheetOptions.CenterHorizontal then
        Attributes.Add('x:CenterHorizontal', '1', false);
      if ProcessedSheet.SheetOptions.CenterVertical then
        Attributes.Add('x:CenterVertical', '1', false);
      if ProcessedSheet.SheetOptions.StartPageNumber <> 1 then
        Attributes.Add('x:StartPageNumber', inttostr(ProcessedSheet.SheetOptions.StartPageNumber), false);
      if Attributes.Count > 0 then
        WriteEmptyTag('Layout', true, false);

      //PageMargins
      Attributes.Clear();
      Attributes.Add('x:Bottom', ZEFloatSeparator(FormatFloat('0.###############',ProcessedSheet.SheetOptions.MarginBottom / ZE_MMinInch)));
      Attributes.Add('x:Left', ZEFloatSeparator(FormatFloat('0.###############',ProcessedSheet.SheetOptions.MarginLeft / ZE_MMinInch)), false);
      Attributes.Add('x:Right', ZEFloatSeparator(FormatFloat('0.###############',ProcessedSheet.SheetOptions.MarginRight / ZE_MMinInch)), false);
      Attributes.Add('x:Top', ZEFloatSeparator(FormatFloat('0.###############',ProcessedSheet.SheetOptions.MarginTop / ZE_MMinInch)), false);
      WriteEmptyTag('PageMargins', true, false);

      //Header
      Attributes.Clear();
      if (ProcessedSheet.SheetOptions.HeaderData > '') then
      begin
        if (ProcessedSheet.SheetOptions.HeaderMargins.Height) <> 13 then
          Attributes.Add('x:Margin', ZEFloatSeparator(FormatFloat('0.##', ProcessedSheet.SheetOptions.HeaderMargins.Height / ZE_MMinInch)));
        Attributes.Add('x:Data', ProcessedSheet.SheetOptions.HeaderData, false);
        WriteEmptyTag('Header', true, true);
      end;

      //Footer
      Attributes.Clear();
      if (ProcessedSheet.SheetOptions.FooterData > '') then
      begin
        if ProcessedSheet.SheetOptions.FooterMargins.Height <> 13 then
          Attributes.Add('x:Margin', ZEFloatSeparator(FormatFloat('0.##', ProcessedSheet.SheetOptions.FooterMargins.Height / ZE_MMinInch)));
        Attributes.Add('x:Data', ProcessedSheet.SheetOptions.FooterData, false);
        WriteEmptyTag('Footer', true, true);
      end;

      WriteEndTagNode(); //PageSetup

      //Print
      Attributes.Clear();
      if ProcessedSheet.SheetOptions.PaperSize <> 0 then
      begin
        WriteTagNode('Print',true, true, true);
        WriteEmptyTag('ValidPrinterInfo');
        WriteTag('PaperSizeIndex',inttostr(ProcessedSheet.SheetOptions.PaperSize), true, false, false);
        WriteTag('HorizontalResolution', '600', true, false, false);  //может, тоже добавить?
        WriteTag('VerticalResolution', '600', true, false, false);
        WriteEndTagNode(); //Print
      end;

      if ProcessedSheet.Selected then
        WriteEmptyTag('Selected', true, false);

      //Фиксирование столбцов/строк
      _hMode := ProcessedSheet.SheetOptions.SplitHorizontalMode;
      _vMode := ProcessedSheet.SheetOptions.SplitVerticalMode;
      b := (_vMode <> ZSplitNone) or
           (_hMode <> ZSplitNone);
      if (b) then
        b := (ProcessedSheet.SheetOptions.SplitVerticalValue <> 0) or
             (ProcessedSheet.SheetOptions.SplitHorizontalValue <> 0);

      if (b) then
      begin
        b := (_vMode = ZSplitFrozen) or
             (_hMode = ZSplitFrozen);
        if (b) then
        begin
//          WriteEmptyTag('Selected', true, false);
          WriteEmptyTag('FreezePanes', true, false);
          WriteEmptyTag('FrozenNoSplit', true, false);
        end;

        _WriteSplitFrozen(_xml, _hMode,
                                ProcessedSheet.SheetOptions.SplitHorizontalValue,
                                'SplitHorizontal',
                                'TopRowBottomPane');

        _WriteSplitFrozen(_xml, _vMode,
                                ProcessedSheet.SheetOptions.SplitVerticalValue,
                                'SplitVertical',
                                'LeftColumnRightPane');

        s := '';
        if ((_hMode = ZSplitFrozen) and (_vMode = ZSplitFrozen)) then
          s := '0'
        else
        if (_hMode = ZSplitFrozen) then
          s := '2'
        else
        if (_vMode = ZSplitFrozen) then
          s := '1';

        if (s > '') then
          WriteTag('ActivePane', s, true, false, false);
      end; //fix or split

      if not((ProcessedSheet.SheetOptions.ActiveCol = 0) and
         (ProcessedSheet.SheetOptions.ActiveRow = 0)) then
      begin
        WriteTagNode('Panes',true, true, false);
        WriteTagNode('Pane',true, true, false);
        WriteTag('Number', '3', true, false, false);
        if ProcessedSheet.SheetOptions.ActiveRow <> 0 then
          WriteTag('ActiveRow', inttostr(ProcessedSheet.SheetOptions.ActiveRow), true, false, false);
        if ProcessedSheet.SheetOptions.ActiveCol <> 0 then
          WriteTag('ActiveCol', inttostr(ProcessedSheet.SheetOptions.ActiveCol), true, false, false);
        WriteEndTagNode(); //Pane
        WriteEndTagNode(); //Panes
      end;

      {tut}//Узнать, как шифруются индексы цветов
      if ProcessedSheet.TabColor <> ClWindow then
        WriteTag('TabColorIndex', inttostr(ProcessedSheet.TabColor), true, false, false);

      WriteEndTagNode(); //WorksheetOptions

      //PageBreaks
      WritePageBreaks();

      //Todo: consider moving this FANs upwards after "WriteEndTagNode(); //Table" and try-finally it
      FreeAndNil(AttrPrintTitles);
      FreeAndNil(AttrCell);
      FreeAndNil(AttrData);
      FreeAndNil(AttrComment);
      FreeAndNil(AttrCommentData);

      WriteEndTagNode(); //Worksheet
    end;
  end; //WritePage

begin
  result := 0;
  _xml := nil;

  try
    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      result := 2;
      exit;
    end;

    _xml := TZsspXMLWriterH.Create();
    _xml.TabLength := 1;
    _xml.TextConverter := TextConverter;
    _xml.BeginSaveToStream(Stream);
    WriteHeader();
    WriteStyles();
    for i := 0 to kol - 1 do
      WritePage(_pages[i], _names[i]);
    _xml.EndSaveTo();
  finally
    if (Assigned(_xml)) then
      FreeAndNil(_xml);
    ZESClearArrays(_pages, _names);
  end;
end; //SaveXmlssToEXML

//Сохраняет в поток в формате Excel XML SpreadSheet
//      XMLSS: TZEXMLSS                 - хранилище
//      Stream: TStream                 - поток
//      SheetsNumbers: array of integer - массив номеров страниц в нужной последовательности
//      SheetsNames: array of string    - массив названий страниц
//              количество элементов в двух массивах должны совпадать
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; Stream: TStream; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string): integer; overload;
begin
  result := SaveXmlssToEXML(XMLSS, Stream, SheetsNumbers, SheetsNames, ZEGetDefaultUTF8Converter(), 'UTF-8', '');
end;

//Сохраняет в поток в формате Excel XML SpreadSheet
//      XMLSS: TZEXMLSS                 - хранилище
//      Stream: TStream                 - поток
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; Stream: TStream): integer; overload;
begin
  result := SaveXmlssToEXML(XMLSS, Stream, [], []);
end;

//Сохраняет в файл в формате Excel XML SpreadSheet
//      XMLSS: TZEXMLSS                 - хранилище
//      FileName: string            - Имя файла
//      SheetsNumbers: array of integer - массив номеров страниц в нужной последовательности
//      SheetsNames: array of string    - массив названий страниц
//              количество элементов в двух массивах должны совпадать
//      TextConverter: TAnsiToCodePageConverter - конвертер
//      CodePageName: string - имя кодировки
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
var
  Stream: TStream;

begin
  result := 0;
  Stream := nil; // так надо!
  try
    try
      Stream := TFileStream.Create(FileName, fmCreate);
    except
      result := 1;
    end;
    if result = 0 then
      result := SaveXmlssToEXML(XMLSS, Stream, SheetsNumbers, SheetsNames, TextConverter, CodePageName, BOM);
  finally
    if Stream <> nil then
      Stream.Free;
  end;
end; //SaveXmlssToEXML

//Сохраняет в файл в формате Excel XML SpreadSheet
//      XMLSS: TZEXMLSS                 - хранилище
//      FileName: string            - Имя файла
//      SheetsNumbers: array of integer - массив номеров страниц в нужной последовательности
//      SheetsNames: array of string    - массив названий страниц
//              количество элементов в двух массивах должны совпадать
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string): integer; overload;
begin
  result := SaveXmlssToEXML(XMLSS, FileName, SheetsNumbers, SheetsNames, ZEGetDefaultUTF8Converter(), 'UTF-8', '');
end; //SaveXmlssToEXML

//Сохраняет в файл в формате Excel XML SpreadSheet
//      XMLSS: TZEXMLSS                 - хранилище
//      FileName: string            - Имя файла
function SaveXmlssToEXML(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
begin
  result := SaveXmlssToEXML(XMLSS, FileName, [], []);
end; //SaveXmlssToEXML

//Читает из потока Excel XML SpreadSheet (EXMLSS)
//      XMLSS: TZEXMLSS                 - хранилище
//      Stream: TStream                 - поток
function ReadEXMLSS(var XMLSS: TZEXMLSS; Stream: TStream): integer; overload;
var
  _xml: TZsspXMLReaderH;
  StyleNames: array of string;
  MaxStyles: integer;
  //count_styles: integer;

  function _StrToInt(val: string): integer;
  begin
    result := 0;
    if (val > '') then
    try
      result := StrToInt(val);
    except
      result := 0;
    end;
  end;

  {tut} //нужно ускорить, сделать без pos в один проход
  function ReplaceAll(const src: string; const substr: string; const value: string): string;
  var
    k: integer;
    ll: integer;

  begin
    result := src;
    ll := length(substr);
    k := pos(substr, result);
    while k > 0 do
    begin
      delete(result, k, ll);
      insert(value, result, k);
      k := pos(substr, result);
    end;
  end;

  function IDByStyleName(const StyleName: string): integer;
  var
    i: integer;

  begin
    result := -1;
    for i := 0 to XMLSS.Styles.Count - 1 do
      if StyleNames[i] = StyleName then
      begin
        result := i;
        break;
      end;
  end;

  function IfTag(const TgName: string; const TgType: integer): boolean;
  begin
    result := (_xml.TagName = TgName) and (_xml.TagType = TgType);
  end;

  //<Style> ... </Style>
  procedure ReadEXMLOneStyle(const IDStyle: integer);
  var
    s: string;
    i: integer;

  begin
    if IDStyle <> -1 then
        XMLSS.Styles[IDStyle].Assign(XMLSS.Styles[-1]);
    while not IfTag('Style', 6) do
    begin
      if _xml.Eof() then break;
      _xml.ReadTag();
      with XMLSS.Styles[IDStyle] do
      begin
        //Alignment
        if _xml.TagName = 'Alignment' then
        begin
          Alignment.Horizontal := StrToHal(_xml.Attributes.ItemsByName['ss:Horizontal']);
          Alignment.Vertical := StrToVal(_xml.Attributes.ItemsByName['ss:Vertical']);
          Alignment.Indent := _StrToInt(_xml.Attributes.ItemsByName['ss:Indent']);
          Alignment.Rotate := _StrToInt(_xml.Attributes.ItemsByName['ss:Rotate']);
          Alignment.ShrinkToFit := ZEStrToBoolean(_xml.Attributes.ItemsByName['ss:ShrinkToFit']);
          Alignment.VerticalText := ZEStrToBoolean(_xml.Attributes.ItemsByName['ss:VerticalText']);
          Alignment.WrapText := ZEStrToBoolean(_xml.Attributes.ItemsByName['ss:WrapText']);
        end else
        //Font
        if _xml.TagName = 'Font' then
        begin
          s := _xml.Attributes.ItemsByName['ss:FontName'];
          if (s > '') then
            font.Name := s;
          s := _xml.Attributes.ItemsByName['x:CharSet'];
          if (s > '') then
            font.Charset := _StrToInt(s);
          if ZEStrToBoolean(_xml.Attributes.ItemsByName['ss:Italic']) then
            font.Style := font.Style + [fsItalic];
          if ZEStrToBoolean(_xml.Attributes.ItemsByName['ss:Bold']) then
            font.Style := font.Style + [fsBold];
          if ZEStrToBoolean(_xml.Attributes.ItemsByName['ss:StrikeThrough']) then
            font.Style := font.Style + [fsStrikeout];
          s := _xml.Attributes.ItemsByName['ss:Underline'];
          if (s > '') then
            if UpperCase(s) <> 'NONE' then
              font.Style := font.Style + [fsUnderline];
          s := _xml.Attributes.ItemsByName['ss:Size'];
          if (s > '') then
            font.Size := _StrToInt(s);
          s := _xml.Attributes.ItemsByName['ss:Color'];
          if (s > '') then
            font.Color := HTMLHexToColor(s);
        end else
        //Borders
        if IfTag('Borders', 4) then
        while not IFTag('Borders', 6) do
        begin
          _xml.ReadTag();
          if _xml.Eof() then break;
          if _xml.TagName = 'Border' then
          begin
            s := _xml.Attributes.ItemsByName['ss:Position'];
            if (s > '') then
            begin
              if s = 'Left' then
                i := 0
              else if s = 'Top' then
                i := 1
              else if s = 'Right' then
                i := 2
              else if s = 'Bottom' then
                i := 3
              else if s = 'DiagonalLeft' then
                i := 4
              else
                i := 5;
              s := _xml.Attributes.ItemsByName['ss:LineStyle'];
              if (s > '') then
                Border[i].LineStyle := StrToZBorderType(s);
              s := _xml.Attributes.ItemsByName['ss:Color'];
              if (s > '') then
                Border[i].Color := HTMLHexToColor(s);
              s := _xml.Attributes.ItemsByName['ss:Weight'];
              if (s > '') then
                Border[i].Weight := _StrToInt(s);
            end
          end;
        end else
        //Interior
        if _xml.TagName = 'Interior' then
        begin
          s := _xml.Attributes.ItemsByName['ss:Color'];
          if (s > '') then
            BGColor := HTMLHexToColor(s);
          s := _xml.Attributes.ItemsByName['ss:Pattern'];
          if (s > '') then
            CellPattern := StrToZCellPattern(s);
          s := _xml.Attributes.ItemsByName['ss:PatternColor'];
          if (s > '') then
            PatternColor := HTMLHexToColor(s);
        end else
        //NumberFormat
        if _xml.TagName = 'NumberFormat' then
          NumberFormat := ConvertFormatXlsxToNative(_xml.Attributes.ItemsByName['ss:Format'])
        else
        //Protection
        if _xml.TagName = 'Protection' then
        begin
          s := _xml.Attributes.ItemsByName['x:HideFormula'];
          if (s > '') then
            HideFormula := ZEStrToBoolean(s);
          s := _xml.Attributes.ItemsByName['ss:Protected'];
          if (s > '') then
            Protect := ZEStrToBoolean(s);
        end;
      end;
    end;
  end;

  //<Styles> ... </Styles>
  procedure ReadEXMLStyles();
  var
    s: string;
    SN: integer;

  begin
    while not IfTag('Styles', 6) do
    begin
      if _xml.Eof() then break;
      _xml.ReadTag();
      if IfTag('Style', 4) then
      begin
        s := _xml.Attributes.ItemsByName['ss:ID'];
        if s = 'Default' then SN := -1 else
        begin
          SN := XMLSS.Styles.Count;
          XMLSS.Styles.Count := XMLSS.Styles.Count + 1;
          if (XMLSS.Styles.Count >= MaxStyles) then
          begin
            MaxStyles := XMLSS.Styles.Count + 50;
            SetLength(StyleNames, MaxStyles);
          end;
          StyleNames[SN] := s;
        end;
        ReadEXMLOneStyle(SN);
      end;
    end;
  end;

  // <DocumentProperties> ... </DocumentProperties>
  procedure ReadEXMLDP();
  begin
    while not IfTag('DocumentProperties', 4) do
    begin
      if _xml.Eof() then break;
      _xml.ReadTag();
      if _xml.TagType = 6 then
      begin
        if _xml.TagName = 'Author' then
          XMLSS.DocumentProperties.Author := trim(_xml.TextBeforeTag) else
        if _xml.TagName = 'LastAuthor' then
          XMLSS.DocumentProperties.LastAuthor := trim(_xml.TextBeforeTag) else
        if _xml.TagName = 'Created' then
        begin
          {tut}
          //XMLSS.DocumentProperties.Created :=
        end else
        if _xml.TagName = 'Company' then
          XMLSS.DocumentProperties.Company := trim(_xml.TextBeforeTag) else
        if _xml.TagName = 'Version' then
          XMLSS.DocumentProperties.Version := trim(_xml.TextBeforeTag);
      end;
    end;
  end;

  // <ExcelWorkbook> ... </ExcelWorkbook>
  procedure ReadEXMLExcelWorkbook();
  begin
    while not IfTag('ExcelWorkbook', 6) do
    begin
      if _xml.Eof() then break;
      _xml.ReadTag();
      if _xml.TagType = 6 then
      begin
        if _xml.TagName = 'WindowHeight' then
          XMLSS.DocumentProperties.WindowHeight := _StrToInt(trim(_xml.TextBeforeTag)) else
        if _xml.TagName = 'WindowWidth' then
          XMLSS.DocumentProperties.WindowWidth := _StrToInt(trim(_xml.TextBeforeTag)) else
        if _xml.TagName = 'WindowTopX' then
          XMLSS.DocumentProperties.WindowTopX := _StrToInt(trim(_xml.TextBeforeTag)) else
        if _xml.TagName = 'WindowTopY' then
          XMLSS.DocumentProperties.WindowTopY := _StrToInt(trim(_xml.TextBeforeTag));
      end else
      if _xml.TagType = 5 then
        if _xml.TagName = 'RefModeR1C1' then
          XMLSS.DocumentProperties.ModeR1C1 := true;
    end;
  end;

  procedure CheckCol(const PageNum: integer; const ColCount: integer);
  begin
    if XMLSS.Sheets[PageNum].ColCount < ColCount then
      XMLSS.Sheets[PageNum].ColCount := ColCount
  end;

  procedure CheckRow(const PageNum: integer; const RowCount: integer);
  begin
    if XMLSS.Sheets[PageNum].RowCount < RowCount then
      XMLSS.Sheets[PageNum].RowCount := RowCount
  end;

  // <Table> .. </Table>
  procedure ReadXMLTable(const PageNum: integer);
  var
    s: string;
    idC, idR, idColumn, cntColumn, cntRow: integer;
    t1, t2, i: integer;
    _isComment: boolean;
    ProcessedSheet: TZSheet;
    ProcessedColumn: TZColOptions;
    ProcessedRow: TZRowOptions;

  begin
    ProcessedSheet := XMLSS.Sheets[PageNum];
    s := _xml.Attributes.ItemsByName['ss:DefaultColumnWidth'];
    if (s > '') then
      ProcessedSheet.DefaultColWidth := ZETryStrToFloat(s)
    else
      ProcessedSheet.DefaultColWidth := 48;
    s := _xml.Attributes.ItemsByName['ss:DefaultRowHeight'];
    if (s > '') then
      ProcessedSheet.DefaultRowHeight := ZETryStrToFloat(s)
    else
      ProcessedSheet.DefaultRowHeight := 12.75;
    s := _xml.Attributes.ItemsByName['ss:ExpandedColumnCount'];
    if (s > '') then
      ProcessedSheet.ColCount := _StrToInt(s);
    s := _xml.Attributes.ItemsByName['ss:ExpandedRowCount'];
    if (s > '') then
      ProcessedSheet.RowCount := _StrToInt(s);
    idR := -1;
    idColumn := -1;
    idC := -1;
    while not IfTag('Table', 6) do
    begin
      if _xml.Eof() then break;
      _xml.ReadTag();
      //Cell
      if _xml.TagName = 'Cell' then
      begin
        if _xml.TagType in [4, 5] then
        begin
          s := _xml.Attributes.ItemsByName['ss:Index'];
          if (s > '') then
            idC := _StrToInt(s) - 1
          else
            inc(idC);
          t1 := 0; //влево
          t2 := 0; //вниз 
          CheckCol(PageNum, idC + 1);
          s := _xml.Attributes.ItemsByName['ss:MergeAcross'];
          if (s > '') then
          begin
            t1 := _StrToInt(s);
            CheckCol(PageNum, idC + 1 + t1);
          end;
          s := _xml.Attributes.ItemsByName['ss:MergeDown'];
          if (s > '') then
          begin
            t2 := _StrToInt(s);
            CheckRow(PageNum, idR + 1 + t2);
          end;
          //Объединённая ячейка
          if t1 + t2 > 0 then
            ProcessedSheet.MergeCells.AddRectXY(idC, idR, idC + t1, idR + t2);
          s := _xml.Attributes.ItemsByName['ss:Formula'];
          if (s > '') then
            ProcessedSheet.Cell[idC, idR].Formula := s;
          s := _xml.Attributes.ItemsByName['ss:HRef'];
          if (s > '') then
            ProcessedSheet.Cell[idC, idR].Href := s;
          s := _xml.Attributes.ItemsByName['x:HRefScreenTip'];
          if (s > '') then
            ProcessedSheet.Cell[idC, idR].HRefScreenTip := s;
          s := _xml.Attributes.ItemsByName['ss:StyleID'];
          if (s > '') then
            ProcessedSheet.Cell[idC, idR].CellStyle := IDByStyleName(s);

          if _xml.TagType = 4 then
          begin
            _isComment := false;
            while not IfTag('Cell', 6) do
            begin
              if _xml.Eof() then break;
              _xml.ReadTag();
              //Comment
              if _xml.TagName = 'Comment' then
              begin
                if _xml.TagType = 6 then
                  _isComment := false
                else
                begin
                  _isComment := true;
                  ProcessedSheet.Cell[idC, idR].ShowComment := true;
                  s := _xml.Attributes.ItemsByName['ss:Author'];
                  if (s > '') then
                    ProcessedSheet.Cell[idC, idR].CommentAuthor := s;
                  s := _xml.Attributes.ItemsByName['ss:ShowAlways'];
                  if (s > '') then
                    ProcessedSheet.Cell[idC, idR].AlwaysShowComment := ZEStrToBoolean(s);
                end;
              end else
              //ss:Data
              if (_xml.TagName = 'ss:Data') or (_xml.TagName = 'Data') then
              begin
                if _xml.TagType = 4 then
                begin
                  if not _isComment then
                  begin
                    s := _xml.Attributes.ItemsByName['ss:Type'];
                    if (s > '') then
                      ProcessedSheet.Cell[idC, idR].CellType := StrToZCellType(s);
                  end;
                  s := '';
                  while not (ifTag('ss:Data', 6) or ifTag('Data', 6)) do
                  begin
                    if _xml.Eof() then break;
                    _xml.ReadTag();
                    s := s + _xml.TextBeforeTag;
                    if (_xml.TagName <> 'ss:Data') and (_xml.TagName <> 'Data') then
                      s := s +  _xml.RawTextTag;
                  end;
                  s := ReplaceAll(s, '&#10;', {$IFDEF FPC} LineEnding {$ELSE} sLineBreak {$ENDIF});
                  if _isComment then
                    ProcessedSheet.Cell[idC, idR].Comment := s
                  else
                    ProcessedSheet.Cell[idC, idR].Data := s;
                end;
              end;
             end;
          end;
          idC := idC + t1; //{tut}
        end;
      end else
      //Row
      if _xml.TagName = 'Row' then
      begin
        if _xml.TagType in [4, 5] then
        begin
          idC := -1;
          s := _xml.Attributes.ItemsByName['ss:Index'];
          if (s > '') then
            idR := _StrToInt(s) - 1
          else
            inc(idR);

          s := _xml.Attributes.ItemsByName['ss:Span'];
          cntRow := 1;
          if (s > '') then
            if (not TryStrToInt(s, cntRow)) then
              cntRow := 1;

          CheckRow(PageNum, idR + cntRow);
          ProcessedRow := ProcessedSheet.Rows[idR];
          s := _xml.Attributes.ItemsByName['ss:Height'];
          if (s > '') then
            ProcessedRow.Height := ZETryStrToFloat(s);
          s := _xml.Attributes.ItemsByName['ss:StyleID'];
          if (s > '') then
            ProcessedRow.StyleID := IDByStyleName(s);
          s := _xml.Attributes.ItemsByName['ss:AutoFitHeight'];
          if (s > '') then
            ProcessedRow.AutoFitHeight := ZEStrToBoolean(s);
          s := _xml.Attributes.ItemsByName['ss:Hidden'];
          if (s > '') then
            ProcessedRow.Hidden := ZEStrToBoolean(s);

          dec(cntRow);
          for i := 1 to cntRow do
            ProcessedSheet.Rows[idr + i].Assign(ProcessedRow);
          inc(idR, cntRow);
        end;
      end else
      //Column
      if _xml.TagName = 'Column' then
      begin
        s := _xml.Attributes.ItemsByName['ss:Index'];
        if (s > '') then
          idColumn := _StrToInt(s) - 1
        else
          inc(idColumn);

        s := _xml.Attributes.ItemsByName['ss:Span'];
        cntColumn := 1;
        if (s > '') then
          if (not TryStrToInt(s, cntColumn)) then
            cntColumn := 1;

        CheckCol(PageNum, idColumn + cntColumn);
        ProcessedColumn := ProcessedSheet.Columns[idColumn];
        s := _xml.Attributes.ItemsByName['ss:Width'];
        if (s > '') then
          ProcessedColumn.Width := ZETryStrToFloat(s);
        s := _xml.Attributes.ItemsByName['ss:StyleID'];
        if (s > '') then
          ProcessedColumn.StyleID := IDByStyleName(s);
        s := _xml.Attributes.ItemsByName['ss:AutoFitWidth'];
        if (s > '') then
          ProcessedColumn.AutoFitWidth := ZEStrToBoolean(s);
        s := _xml.Attributes.ItemsByName['ss:Hidden'];
        if (s > '') then
          ProcessedColumn.Hidden := ZEStrToBoolean(s);

        dec(cntColumn);
        for i := 1 to cntColumn do
          ProcessedSheet.Columns[idColumn + i].Assign(ProcessedColumn);
        inc(idColumn, cntColumn);
      end; //if Column
    end;
  end;

  // <WorksheetOptions> .. </WorksheetOptions>
  procedure ReadXMLWorkSheetOptions(const PageNum: integer);
  var
    s: string;
    _isFreezePanes: boolean;
    _isFrozeNoSplit: boolean;
    _SheetOptions: TZSheetOptions;
    _SplitMode: TZSplitMode;
    _isH, _isV: boolean;

    function _GetSplitValue(const SplitMode: TZSplitMode; const SplitValue: integer): integer;
    begin
      result := SplitValue;
      if (SplitMode = ZSplitSplit) then
        result := round(PointToPixel(SplitValue/20));
    end; //_GetSplitValue

  begin
    _isFreezePanes := false;
    _isFrozeNoSplit := false;
    _isH := false;
    _isV := false;
    _SheetOptions := XMLSS.Sheets[PageNum].SheetOptions;
    _SheetOptions.SplitHorizontalMode := ZSplitNone;
    _SheetOptions.SplitVerticalMode := ZSplitNone;

    while not IfTag('WorksheetOptions', 6) do
    begin
      if _xml.Eof() then break;
      _xml.ReadTag();
      //PageSetup
      if IfTag('PageSetup', 4) then
      while not IfTag('PageSetup', 6) do
      begin
        if _xml.Eof() then break;
        _xml.ReadTag();
        if IfTag('Layout', 5) then
        begin
          if _xml.Attributes.ItemsByName['x:Orientation'] = 'Landscape' then
            _SheetOptions.PortraitOrientation := false
          else
            _SheetOptions.PortraitOrientation := true;
          if _xml.Attributes.ItemsByName['x:CenterHorizontal'] = '1' then
            _SheetOptions.CenterHorizontal := true else
          if _xml.Attributes.ItemsByName['x:CenterVertical'] = '1' then
            _SheetOptions.CenterVertical := true else
          begin
            s := _xml.Attributes.ItemsByName['x:StartPageNumber'];
            if (s > '') then
              _SheetOptions.StartPageNumber := _StrToInt(s);
          end;
        end else
        if IfTag('PageMargins', 5) then
        begin
          s := _xml.Attributes.ItemsByName['x:Bottom'];
          if (s > '') then
            _SheetOptions.MarginBottom := round(ZETryStrToFloat(s)*25.4);
          s := _xml.Attributes.ItemsByName['x:Left'];
          if (s > '') then
            _SheetOptions.MarginLeft := round(ZETryStrToFloat(s)*25.4);
          s := _xml.Attributes.ItemsByName['x:Right'];
          if (s > '') then
            _SheetOptions.MarginRight := round(ZETryStrToFloat(s)*25.4);
          s := _xml.Attributes.ItemsByName['x:Top'];
          if (s > '') then
            _SheetOptions.MarginTop := round(ZETryStrToFloat(s)*25.4);
        end else
        if IfTag('Header', 5) then
        begin
          s := _xml.Attributes.ItemsByName['x:Margin'];
          if (s > '') then
            _SheetOptions.HeaderMargins.Height := abs(round(ZETryStrToFloat(s)*25.4));
          _SheetOptions.HeaderData := _xml.Attributes.ItemsByName['x:Data'];
        end else
        if IfTag('Footer', 5) then
        begin
          s := _xml.Attributes.ItemsByName['x:Margin'];
          if (s > '') then
            _SheetOptions.FooterMargins.Height := abs(round(ZETryStrToFloat(s)*25.4));
          _SheetOptions.FooterData := _xml.Attributes.ItemsByName['x:Data'];
        end;
      end else
      //Print
      if IfTag('Print', 4) then
      while not IfTag('Print', 6) do
      begin
        _xml.ReadTag();
        if _xml.Eof() then break;
        if IfTag('PaperSizeIndex', 6) then
          _SheetOptions.PaperSize := _StrToInt(trim(_xml.TextBeforeTag));
      end else
      if IfTag('Selected', 5) then
        XMLSS.Sheets[PageNum].Selected := true else
      //Panes
      if IfTag('Panes', 4) then
      while not IfTag('Panes', 6) do
      begin
        _xml.ReadTag();
        if _xml.Eof() then break;
        if IfTag('Pane', 4) then
        while not IfTag('Pane', 6) do
        begin
          _xml.ReadTag();
          if _xml.Eof() then break;
          if IfTag('ActiveRow', 6) then
            _SheetOptions.ActiveRow := _StrToInt(trim(_xml.TextBeforeTag)) else
          if IfTag('ActiveCol', 6) then
            _SheetOptions.ActiveCol := _StrToInt(trim(_xml.TextBeforeTag));
        end;
      end else
      if IfTag('TabColorIndex', 6) then
        XMLSS.Sheets[PageNum].TabColor := _StrToInt(trim(_xml.TextBeforeTag))
      else
      //Spli/Frozen
      if (IfTag('FreezePanes', 5)) then
        _isFreezePanes := true
      else
      if (IfTag('FrozenNoSplit', 5)) then
        _isFrozeNoSplit := true
      else
      if (IfTag('SplitHorizontal', 6)) then
      begin
        _isH := true;
        s := _xml.TextBeforeTag;
        _SheetOptions.SplitHorizontalMode := ZSplitSplit;
        _SheetOptions.SplitHorizontalValue := round(ZETryStrToFloat(s));
      end else
      if (IfTag('SplitVertical', 6)) then
      begin
        _isV := true;
        s := _xml.TextBeforeTag;
        _SheetOptions.SplitVerticalMode := ZSplitSplit;
        _SheetOptions.SplitVerticalValue := round(ZETryStrToFloat(s));
      end;
    end; //while

    if (_isH or _isV) then
    begin
      _SplitMode := ZSplitSplit;
      if (_isFreezePanes or _isFrozeNoSplit) then
        _SplitMode := ZSplitFrozen;
    end else
      _SplitMode := ZSplitNone;

    if (_isH) then
    begin
      _SheetOptions.SplitHorizontalMode := _SplitMode;
      _SheetOptions.SplitHorizontalValue := _GetSplitValue(_SplitMode, _SheetOptions.SplitHorizontalValue);
    end;

    if (_isV) then
    begin
      _SheetOptions.SplitVerticalMode := _SplitMode;
      _SheetOptions.SplitVerticalValue := _GetSplitValue(_SplitMode, _SheetOptions.SplitVerticalValue);
    end;

  end; //ReadXMLWorkSheetOptions

  // <PageBreaks> ... </PageBreaks>
  procedure ReadPageBreaks(const PageNum: integer);
  var
    t: integer;

  begin
    while not IfTag('PageBreaks', 6) do
    begin
      if _xml.Eof() then break;
      _xml.ReadTag();
      if IfTag('Column', 6) then
      begin
        t := _StrToInt(trim(_xml.TextBeforeTag));
        if (t > 0) and (t <= XMLSS.Sheets[PageNum].ColCount - 1) then
          XMLSS.Sheets[PageNum].Columns[t].Breaked := true;
      end else
      if IfTag('Row', 6) then
      begin
        t := _StrToInt(trim(_xml.TextBeforeTag));
        if (t > 0) and (t <= XMLSS.Sheets[PageNum].RowCount - 1) then
          XMLSS.Sheets[PageNum].Rows[t].Breaked := true;
      end
    end;
  end; //ReadPageBreaks

  //Читает секцию <Worksheet> ... </Worksheet>
  procedure ReadEXMLWorkSheet();
  var
    num: integer;

  begin
    num := XMLSS.Sheets.Count;
    XMLSS.Sheets.Count := num + 1;
    XMLSS.Sheets[num].Title := _xml.Attributes.ItemsByName['ss:Name'];
    while not IfTag('Worksheet', 6) do
    begin
      if _xml.Eof() then break;
      _xml.ReadTag();
      if IfTag('Table', 4) then
        ReadXMLTable(num) else
      if IfTag('WorksheetOptions', 4) then
        ReadXMLWorkSheetOptions(num) else
      if IfTag('PageBreaks', 4) then
        ReadPageBreaks(num);
    end;
  end;

  procedure ProcessReadXML();
  begin
    while not _xml.Eof() do
    begin
      _xml.ReadTag();
      result := result or _xml.ErrorCode;
      if IfTag('DocumentProperties', 4) then
        ReadEXMLDP();
      if IfTag('ExcelWorkbook', 4) then
        ReadEXMLExcelWorkbook();
      if IfTag('Styles', 4) then
        ReadEXMLStyles();
      if IfTag('Worksheet', 4) then
        ReadEXMLWorkSheet();
    end;
  end;

begin
  result := 0;
  if Stream = nil then
  begin
    result := -1;
    exit;
  end;
  try
    _xml := TZsspXMLReaderH.Create();
    _xml.AttributesMatch := false;
    _xml.BeginReadStream(Stream);
    MaxStyles := -1;
    XMLSS.Styles.Clear();
    XMLSS.Sheets.Count := 0;
    ProcessReadXML();
  finally
    FreeAndNil(_xml);
    SetLength(StyleNames, 0);
    StyleNames := nil;
  end;
end; //ReadEXMLSS

//Читает из файла Excel XML SpreadSheet (EXMLSS)
//      XMLSS: TZEXMLSS             - хранилище
//      FileName: string            - имя файла
function ReadEXMLSS(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
var
  Stream: TStream;

begin
  result := 0;
  Stream := nil; 
  try
    try
      Stream := TFileStream.Create(FileName, fmOpenRead or fmShareDenyNone {fmShareDenyWrite});
    except
      result := -1;
    end;
    if result = 0 then
      result := ReadEXMLSS(XMLSS, stream);
  finally
    if Stream <> nil then
      Stream.Free;
  end;
end; //ReadEXMLSS

end.
