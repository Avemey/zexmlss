{
 Copyright (C) 2011 Ruslan Neborak

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

unit zcftext;

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

interface

uses
  {$IFDEF FPC}
  lcltype, types, FPCanvas
  {$ELSE}
  Windows
  {$ENDIF},
  graphics, SysUtils, controls
  ;

{$I 'compver.inc'}

const
  ZCF_SIZINGW         = 1;
  ZCF_SIZINGH         = 1 shl 1;
  ZCF_CALCONLY        = 1 shl 2;
  ZCF_SYMBOL_WRAP     = 1 shl 3;
  ZCF_NO_TRANSPARENT  = 1 shl 4;
  ZCF_NO_CLIP         = 1 shl 5;
  ZCF_DOTTED          = 1 shl 6;  //пока не работает
  ZCF_DOTTED_LINE     = 1 shl 7;  //пока не работает

function ZCWriteTextFormatted(canvas: TCanvas; text: string; fnt: TFont; fntColor: TColor; HA: integer; VA: integer; WordWrap: boolean; var Rct: TRect; IH, IV: byte; Params: integer; LineSpacing: integer; Rotate: integer = 0): boolean; overload;
function ZCWriteTextFormatted(canvas: TCanvas; text: string; fnt: TFont; HA: integer; VA: integer; WordWrap: boolean; var Rct: TRect; IH, IV: byte; Params: integer; LineSpacing: integer; Rotate: integer = 0): boolean; overload;

{$IFNDEF FPC}
function IsFontTrueType(Fnt: TFont): boolean;
{$ENDIF}

implementation

{$IFNDEF FPC}
function IsFontTrueType(Fnt: TFont): boolean;
var
  DC: HDC;
  OldFont: HFont;
  Metr: TTextMetric;

begin
  DC := GetDC(0);
  try
    OldFont := SelectObject(DC, Fnt.Handle);
    GetTextMetrics(DC, Metr);
    SelectObject(DC, OldFont);
  finally
    ReleaseDC(0, DC);
  end;
  result := (Metr.tmPitchAndFamily and tmpf_TrueType) <> 0;
end;
{$ENDIF}

//Выводит в canvas текст
//INPUT
//      canvas: TCanvas       - холст для рисования
//      text: string          - текст
//      fnt: TFont            - шрифт
//      fntColor: TColor      - цвет текста
//      HA: integer           - горизонтальное выравнивание (0 - по левому краю, 1 - по центру, 2 - по правому краю, 3 - распределённый)
//      VA: integer           - вертикальное выравнивание (0 - сверху, 1 - по центру, 2 - по низу)
//      WordWrap: boolean     - переносить-ли текст по словам, если не помещается в область
//  var Rct: TRect            - область для рисования
//      IH: byte              - отступ по горизонтали (учитывает HA)
//      IV: byte              - отступ по вертикали (учитывает VA)
//      Params: integer       - дополнительные параметры
//      LineSpacing: integer  - межстрочный интервал в пикселях
//      Rotate: integer       - угол поворота текста (игнорируется, если шрифт не TrueType)
//RETURN
//      true  - ok
//      false - что-то не так
function ZCWriteTextFormatted(canvas: TCanvas; text: string; fnt: TFont; fntColor: TColor; HA: integer; VA: integer; WordWrap: boolean; var Rct: TRect; IH, IV: byte; Params: integer; LineSpacing: integer; Rotate: integer = 0): boolean;
type
  TStrSize = record   //строки
    value: string;      //текст строки
    Size: TSize;        //размер строки
  end;

  TStrWordSize =  record  //слова в строке
    value: string;
    Size: TSize;            //размер БЕЗ пробелов
    SpacesBefore: integer;  //кол-во пробелов до слова
  end;

  TWordSize = array [0..1] of integer;
                    //0 элемент - кол-во слов
                    //1 элемент - кол-во пробелов после _последнего_ слова

var
  tmpcanvas: TCanvas;           //временный канвас
  {$IFNDEF FPC}
  LF: TLogFont;
  {$ENDIF}
  i: integer;
  strl: array of TStrSize;      //массив со строками
  strIndex: array of integer;   //массив с номерами строк
  kol: integer;                 //кол-во строк
  l: integer;
  s: string;
  Angle: real;                  //Угол поворота в радианах
  cosA: real;
  sinA: real;
  calc_height: integer;         //рассчётная высота прямоугольника
  calc_width: integer;          //рассчётная длина прямоугольника
  calc_heightN: integer;        //рассчётная высота прямоугольника (без поворота)
  calc_widthN: integer;         //рассчётная длина прямоугольника (без поворота)
  _height: integer;             //высота прямоугольника
  _width: integer;              //длина прямоугольника
  _spaceSize, _dotSize: TSize;  //размеры пробела и точки
  WordsArray: array of array of TStrWordSize;   //слова в предложении
  WordsSize: array of TWordSize;                //кол-во слов в предложении
  isSizingW: boolean;           //можно ли увеличивать ширину прямоугольника, если текст не помещается
  isSizingH: boolean;           //можно ли увеличивать высоту, если текст не помещается
  isCalcOnly: boolean;          //при true только расчёт без рисования текста
  isDotted: boolean;            //при true оканчивать непомещающийся текст "..." (игнорируется, если isSizingH = true) (не реализовано)
  isDottedLine: boolean;        //при true оканчивать непомещающуюся строку "..." (игнорируется, если isSizingW = true) (не реализовано)
  isSymbolWrap: boolean;        //при true разрыв слов происходит на любом символе (работает только при WordWrap = true)
  isTransparent: boolean;       //при true устанавливает прозрачный фон для вывода текста (пусть будет по умолчанию)
  isClip: boolean;              //При true обрезает текст, не помещающийся в прямоугольник

  ArraySizeRes: integer;

  //При необходимости увеличивает размеры массивов
  procedure CheckArraySize();
  var
    i: integer;
    num: integer;

  begin
    if (kol + 1 < ArraySizeRes) then
      exit;
    num := ArraySizeRes;
    inc(ArraySizeRes, 5);
    SetLength(WordsSize, ArraySizeRes);
    SetLength(WordsArray, ArraySizeRes);
    SetLength(strl, ArraySizeRes);
    SetLength(strIndex, ArraySizeRes);
    for i := num to ArraySizeRes - 1 do
    begin
      WordsSize[i][0] := 0;
      WordsSize[i][1] := 0;
    end;
  end;

  //Добавить новую строку в массивы str и strIndex
  //INPUT
  //    var NewStr: string      - новая строка
  //        LineIndex: integer  - номер строки, из которой копируют
  //    var StrSize: TSize      - размер новой строки
  procedure AddNewLine(var NewStr: string; LineIndex: integer; var StrSize: TSize);
  begin
    //потом посмотреть, равен ли новый размер старому за минусом части перенесённого текста
    strl[strIndex[LineIndex]].Size := tmpcanvas.TextExtent(strl[strIndex[LineIndex]].value);
    CheckArraySize();
    strl[kol].value := NewStr;
    strl[kol].Size := StrSize;
    //сдвиг массива
    move(strIndex[LineIndex], strIndex[LineIndex + 1], SizeOf(integer) * (kol - LineIndex + 1));
    strIndex[LineIndex + 1] := kol;
    inc(kol);
   end;

  //Делает перенос строк
  procedure DoWordWrap();
  var
    i, j, n: integer;
    s: string;
    tmpSize: TSize;
    ind: integer;

  begin
    i := 0;
    ind := 0;
    case (HA) of
      0, 2: ind := IH;
      3, 1: ind := 2 * IH;
    end;
    while (i < kol) do
    begin
      if (strl[strIndex[i]].Size.cx + ind > _width) then
      begin
        n := length(strl[strIndex[i]].value);
        if (isSymbolWrap) then
        begin
          //TODO: Подумать насчёт ускорения поиска потом (дихотомией, к примеру)
          for j := n downto 1 do
          begin
            s := copy(strl[strIndex[i]].value, j, n - j + 1);
            tmpSize := tmpcanvas.TextExtent(s);
            if (strl[strIndex[i]].Size.cx - tmpSize.cx + ind <= _width) then
            begin
              delete(strl[strIndex[i]].value, j, n - j + 1);
              AddNewLine(s, i, tmpSize);
              break;
            end;
          end; //for j
        end //if isSymbolWrap
        else
        begin //start WordWrap
          s := '';
          j := length(strl[strIndex[i]].value);
          while (j >= 1) do
          begin
            {$IFDEF DELPHI_UNICODE}
            if (CharInSet(strl[strIndex[i]].value[j], [' ', #9])) then
            {$ELSE}
            if (strl[strIndex[i]].value[j] in [' ', #9]) then
            {$ENDIF}
            begin
              //tut
              //тут нужно подумать как правильно переносить пробелы.
              //Возможно, не все пробелы нужно переносить в новую строку, а только те,
              //которые не помещаются на старой строке
              {$IFDEF DELPHI_UNICODE}
              while (CharInSet(strl[strIndex[i]].value[j], [' ', #9]) and (j >= 1)) do
              {$ELSE}
              while ((strl[strIndex[i]].value[j] in [' ', #9]) and (j >= 1)) do
              {$ENDIF}
              begin
                dec(j);
              end;
              s := copy(strl[strIndex[i]].value, j + 1, n - j);
              tmpSize := tmpcanvas.TextExtent(s);
              if (strl[strIndex[i]].Size.cx - tmpSize.cx + ind <= _width) and (j > 1) then
              begin
                delete(strl[strIndex[i]].value, j + 1, n - j);
                s := trim(s);
                tmpSize := tmpcanvas.TextExtent(s);
                AddNewLine(s, i, tmpSize);
                break;
              end;
            end;
            dec(j);
          end; //while
        end; //WordWrap end
      end; //if length
      inc(i);
    end; //while
  end;

  //Добавить слово в массив WordsArray
  //INPUT
  //    var NWord: string       - слово
  //        num: integer        - номер строки, из которой берётся слово
  //    var SpaceCount: integer - кол-во пробелов перед словом
  procedure AddNewWord(var NWord: string; num: integer; var SpaceCount: integer);
  var
    t: integer;

  begin
    if (length(NWord) > 0) then
    begin
      SetLength(WordsArray[num], WordsSize[num][0] + 1);
      t := WordsSize[num][0];
      WordsArray[num][t].value := s;
      WordsArray[num][t].Size := tmpcanvas.TextExtent(s);
      WordsArray[num][t].SpacesBefore := SpaceCount;
      { //этот кусочек не работает в XE, хз почему
      With (WordsArray[num][WordsSize[num][0]]) do
      begin
        value := s;
        Size := tmpcanvas.TextExtent(s);
        SpacesBefore := SpaceCount;
      end;
      }
      inc(WordsSize[num][0]);
      NWord := '';
      SpaceCount := 0;
    end;
  end;

  //Подсчитать размеры слов
  procedure CalcWords();
  var
    i, j: integer;
    s: string;
    SpaceCount: integer;

  begin
    for i := 0 to kol - 1 do
    begin
      s := '';
      SpaceCount := 0;
      for j := 1 to length(strl[i].value) do
      begin
        {$IFDEF DELPHI_UNICODE}
        if (CharInSet(strl[i].value[j], [' ', #9])) then
        {$ELSE}
        if (strl[i].value[j] in [' ', #9]) then
        {$ENDIF}
        begin
          AddNewWord(s, i, SpaceCount);
          inc(SpaceCount);
        end else
          s := s + strl[i].value[j];
      end; //for j
      AddNewWord(s, i, SpaceCount);
      WordsSize[i][1] := SpaceCount;
    end; //for i
  end;

  //Подсчитывает размер прямоугольника для текста
  procedure CalcRectSize();
  var
    i, j, tmp, _hh, _ww: integer;
    _NewH, _NewW: integer;

  begin
    _NewH := 0;
    for i := 0 to kol - 1 do
      inc(_NewH, strl[i].Size.cy);
    if (kol > 1) then
      _NewH := _NewH + (kol - 1) * LineSpacing;

    _NewW := 0;
    for i := 0 to kol - 1 do
    begin
      tmp := 0;
      if (HA = 3) then
      begin
        //пробелы игнорируются
        for j := 0 to WordsSize[i][0] - 1 do
          tmp := tmp + WordsArray[i][j].Size.cx;
        if (WordsSize[i][0] > 1) then
          tmp := tmp + (WordsSize[i][0] - 1) * _spaceSize.cx;
      end else
      begin
        //все пробелы считаются
        tmp := strl[i].Size.cx;
      end;

      if (tmp > _NewW) then
        _NewW := tmp;
    end; //for i

    _hh := _NewH;
    _ww := _NewW;

    calc_heightN := _NewH;
    calc_widthN := _NewW;

    //Может, если текст по центру, то отступы не добавлять?
    _NewH := IV + round(_ww * abs(sinA) + _hh * abs(cosA));
    _NewW := IH + round(_ww * abs(cosA) + _hh * abs(sinA));

    //NewW
    if (calc_width < _NewW) then
      calc_width := _NewW;
    if ((_width < _NewW) and (isSizingW)) then
    begin
      Rct.Right := Rct.Right + _NewW - _width;
      _width := _NewW;
    end;

    //NewH
    if (calc_height < _NewH) then
      calc_height := _NewH;
    if ((_height < _NewH) and (isSizingH)) then
    begin
      Rct.bottom := Rct.bottom + _NewH - _height;
      _height := _NewH;
    end;
  end;

  procedure DoDrawTxt();
  var
    i, j: integer;
    xx, yy, deltaX, deltaY: integer;
    idx: integer;
    {$IFDEF FPC}
    CanvasRgn: TRect;
    oldOpaque: boolean;
    oldTextStyle: TTextStyle;
    oldBrushStyle: TFPBrushStyle;
    oldColor: TColor;
    {$ELSE}
    CanvasRgn: HRGN;
    bkMode: integer;
    {$ENDIF}

  begin
    {$IFNDEF FPC}
    CanvasRgn := 0;
    {$ENDIF}

    if (isClip) then
    begin
      {$IFDEF FPC}
      CanvasRgn := canvas.ClipRect;
      canvas.ClipRect := Rct;
      {$ELSE}
      CanvasRgn := CreateRectRgn(Rct.Left, Rct.Top, Rct.Right, Rct.Bottom);
      SelectClipRgn(canvas.Handle, CanvasRgn);
      {$ENDIF}
    end;

    deltaX := 0;
    deltaY := 0;
    yy := 0;
    {$IFNDEF FPC}
    bkMode := 0;
    {$ENDIF}
    if (isTransparent) then
    begin
      {$IFDEF FPC}
        {$IFDEF LCLGtk2}
        //Hmm
        {$ELSE}
      //что-то у lazarus-а какие-то непонятки с прозрачностью и TextOut

      oldOpaque := canvas.TextStyle.Opaque;
      oldTextStyle := canvas.TextStyle;
      oldTextStyle.Opaque := true;
      canvas.TextStyle := oldTextStyle;
      oldColor := canvas.Brush.Color;
      canvas.Brush.Color := clNone;
      oldBrushStyle := Canvas.Brush.Style;
      Canvas.Brush.Style := bsClear;

        {$ENDIF}
      {$ELSE}
      bkMode := GetBkMode(canvas.Handle);
      SetBkMode(canvas.Handle, TRANSPARENT);
      {$ENDIF}

    end;

    if (VA in [0, 2]) then
    begin
      if ((Rotate >= 0) and (Rotate < 180)) then
      begin
        if (Rotate <= 90) then
          yy := round(calc_heightN*cosA + calc_widthN*sinA)
        else
          yy := round(calc_widthN*sinA);
      end else
      begin
        if (Rotate <= 270) then
          yy := 0
        else
          yy := round(calc_heightN*cosA);
      end;
    end;

    case (VA) of
      0:
        begin
          yy := Rct.Bottom - (_height - calc_height) - yy;
        end;
      1:
        begin
          yy := round(calc_heightN*cosA + calc_widthN*sinA);
          yy := Rct.Top + ((_height - yy) shr 1);
        end;
      2:
        begin
          yy := Rct.Bottom - yy - IV;
        end;
    end;

    if (HA = 0) then
      deltaY := round((calc_widthN) * sinA);
      
    for i := 0 to kol - 1 do
    begin
      idx := strIndex[i];
      case (HA) of
        0:  //left
          begin
            if ((Rotate >= 0) and (Rotate <= 90)) then
              xx := Rct.Left + IH
            else
            if ((Rotate > 90) and (Rotate <= 180)) then
              xx := Rct.Left + IH - round(calc_widthN*cosA)
            else
            if ((Rotate > 180) and (Rotate <= 270)) then
              xx := Rct.Left + IH - round(calc_widthN*cosA + calc_heightN*sinA)
            else
              xx := Rct.Left + IH - round(calc_heightN*sinA);
            canvas.TextOut(xx + deltaX, yy + deltaY, strl[idx].value);
          end;
        1:  //center
          begin
            deltaY := round(sinA * (calc_widthN + (strl[idx].Size.cx))/2);
            xx := Rct.Left + round((_width - calc_heightN*sinA - cosA*strl[idx].Size.cx)/2);
            canvas.TextOut(xx + deltaX, yy + deltaY, strl[idx].value);
          end;
        2:  //right
          begin
            deltaY := round(sinA*(strl[idx].Size.cx));
            if ((Rotate >= 0) and (Rotate <= 90)) then
              xx := Rct.Right - round(IH + strl[idx].Size.cx*cosA + calc_heightN*sinA)
            else
            if ((Rotate > 90) and (Rotate <= 180)) then
              xx := Rct.Right - round(IH + strl[idx].Size.cx*cosA + calc_heightN*sinA - calc_widthN*cosA)
            else
            if ((Rotate > 180) and (Rotate <= 270)) then
              xx := Rct.Right - round(IH + strl[idx].Size.cx*cosA - calc_widthN*cosA)
            else
              xx := Rct.Right - round(IH + strl[idx].Size.cx*cosA);
            canvas.TextOut(xx + deltaX, yy + deltaY, strl[idx].value);
          end;
        3:  //distributed
          begin
            for j := 0 to WordsSize[idx][0] - 1 do
            begin
            end;
          end;
      end; //case HA
      yy := yy + round(cosA * (strl[idx].Size.cy + LineSpacing));
      deltaX := deltaX + round(sinA * (strl[idx].Size.cy + LineSpacing));
    end; //for i
    if (isTransparent) then
    begin
      {$IFDEF FPC}
        {$IFDEF LCLGtk2}
        //Hmm
        {$ELSE}
      oldTextStyle.Opaque := oldOpaque;
      canvas.TextStyle := oldTextStyle;
      canvas.Brush.Color := oldColor;
      Canvas.Brush.Style := oldBrushStyle;
        {$ENDIF}
      {$ELSE}
      SetBkMode(canvas.Handle, bkMode);
      {$ENDIF}
    end;
    
    if (isClip) then
    begin
      {$IFDEF FPC}
      canvas.ClipRect := CanvasRgn;
      {$ELSE}
      SelectClipRgn(canvas.Handle, HRGN(nil));
      DeleteObject(CanvasRgn);
      {$ENDIF}
    end;
  end;

begin
  result := false;
  if (not (Assigned(canvas) and Assigned(fnt))) then
    exit;
  kol := 0;
  cosA := 1;
  sinA := 0;
  canvas.Font.Assign(fnt);
  ArraySizeRes := 0;
  try
    tmpcanvas := TCanvas.Create();//TControlCanvas.Create();
    //(tmpcanvas as TControlCanvas).Control := (canvas as TControlCanvas).Control;
    //(tmpcanvas as TControlCanvas).Control := canvas.;
    tmpcanvas.Handle := canvas.Handle;
    tmpcanvas.Font.Assign(canvas.Font);
    tmpcanvas.Brush.Assign(canvas.Brush);
    tmpcanvas.Pen.Assign(canvas.Pen);

    if (Rotate mod 360 = 0) then
      Rotate := 0;
    Rotate := Rotate - 360*(Rotate div 360);
    if (Rotate < 0) then
      Rotate := Rotate + 360;
    {$IFDEF FPC}
    //TCanvas.GetTextMetrics
    {if (not IsFontTrueType(fnt)) then
      Rotate := 0;
      }
    {$ELSE}
    if (not IsFontTrueType(fnt)) then
      Rotate := 0;
    {$ENDIF}
    //новый шрифт
    if (Rotate <> 0) then
    begin
      Angle := Rotate * Pi / 180;
      sinA := sin(Angle);
      cosA := cos(Angle);
      {$IFDEF FPC}
      canvas.Font.Orientation := Rotate * 10;
      {$ELSE}
      FillChar(Lf, SizeOf(Lf), 0);
      GetObject(fnt.Handle, SizeOf(LF), @LF);

      LF.lfEscapement := Rotate * 10;
      LF.lfQuality := 4;
      LF.lfOrientation := LF.lfEscapement; //w9x
      canvas.Font.Handle := CreateFontIndirect(LF);
      {$ENDIF}
    end;

    canvas.Font.Color := fntColor;

    text := text + #13#10;
    s := '';
    l := length(text);
    for i := 1 to l do
    begin
      case text[i] of
        #10:
           begin
             if (i <> l-1) then
             begin
               CheckArraySize();
               if (length(s) = 0) then
                 s := ' ';
               strl[kol].value := s;
               strl[kol].Size := tmpcanvas.TextExtent(s);
               strIndex[kol] := kol;
               inc(kol);
               s := '';
             end;
           end;
        #13:{Хм...};
        else
          s := s + text[i];
      end; //case
    end; //for

    _width := abs(Rct.Right - Rct.Left);
    _height := abs(Rct.Bottom - Rct.Top);
    calc_height := 0;
    calc_width := 0;
    calc_heightN := 0;
    calc_widthN := 0;

    _spaceSize := tmpcanvas.TextExtent(' ');
    _dotSize := tmpcanvas.TextExtent('.');

    isSizingH := (Params and ZCF_SIZINGH) = ZCF_SIZINGH;
    isSizingW := (Params and ZCF_SIZINGW) = ZCF_SIZINGW;
    isCalcOnly := (Params and ZCF_CALCONLY) = ZCF_CALCONLY;
    isDotted := (Params and ZCF_DOTTED) = ZCF_DOTTED;
    isDottedLine := (Params and ZCF_DOTTED_LINE) = ZCF_DOTTED_LINE;
    isSymbolWrap := ((Params and ZCF_SYMBOL_WRAP) = ZCF_SYMBOL_WRAP) and WordWrap;
    isTransparent := not ((Params and ZCF_NO_TRANSPARENT) = ZCF_NO_TRANSPARENT);
    isClip := not ((Params and ZCF_NO_CLIP) = ZCF_NO_CLIP);
    if (isSizingH) then
      isDotted := false;
    if (isSizingW) then
      isDottedLine := false;

    if (not HA in [0..3]) then
      HA := 0;
    if (not VA in [0..2]) then
      VA := 0;

    if ((HA = 3) and (isSymbolWrap)) then
      HA := 0;

    if (WordWrap) then
      DoWordWrap();
    CalcWords();
    CalcRectSize();
    if (not isCalcOnly) then
      DoDrawTxt();
    result := true;
  finally
    FreeAndNil(tmpcanvas);
    if (ArraySizeRes <> 0) then
    begin
      SetLength(strl, 0);
      SetLength(strIndex, 0);
      SetLength(WordsSize, 0);
      for i := 0 to ArraySizeRes - 1 do
      begin
        SetLength(WordsArray[i], 0);
        WordsArray[i] := nil;
      end;
      SetLength(WordsArray, 0);
    end;
    strl := nil;
    strIndex := nil;
    WordsSize := nil;
    WordsArray := nil;
  end;
end;

//Выводит в canvas текст
//INPUT
//      canvas: TCanvas       - холст для рисования
//      text: string          - текст
//      fnt: TFont            - шрифт
//      HA: integer           - горизонтальное выравнивание (0 - по левому краю, 1 - по центру, 2 - по правому краю, 3 - распределённый)
//      VA: integer           - вертикальное выравнивание (0 - сверху, 1 - по центру, 2 - по низу)
//      WordWrap: boolean     - переносить-ли текст по словам, если не помещается в область
//  var Rct: TRect            - область для рисования
//      IH: byte              - отступ по горизонтали (учитывает HA)
//      IV: byte              - отступ по вертикали (учитывает VA)
//      Params: integer       - дополнительные параметры
//      LineSpacing: integer  - межстрочный интервал в пикселях
//      Rotate: integer       - угол поворота текста (игнорируется, если шрифт не TrueType)
//RETURN
//      true  - ok
//      false - что-то не так
function ZCWriteTextFormatted(canvas: TCanvas; text: string; fnt: TFont; HA: integer; VA: integer; WordWrap: boolean; var Rct: TRect; IH, IV: byte; Params: integer; LineSpacing: integer; Rotate: integer = 0): boolean;
begin
  result := false;
  if (Assigned(fnt)) then
    result := ZCWriteTextFormatted(canvas, text, fnt, fnt.Color, HA, VA, WordWrap, Rct, IH, IV, Params, LineSpacing, Rotate);
end;

end.
