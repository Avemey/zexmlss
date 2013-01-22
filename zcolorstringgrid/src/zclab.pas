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

unit zclab;

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

{$I 'compver.inc'}

interface

uses
  Classes, SysUtils, controls, graphics, zcftext
  {$IFDEF FPC}
  ,lresources
  {$ELSE}
  ,Windows, Messages
  {$ENDIF} 
  ;

type
  TZCLabel = class(TGraphicControl)
  private
    FAlignmentVertical: byte;
    FAlignmentHorizontal: byte;
    FAutoSizeHeight: boolean;
    FAutoSizeWidth: boolean;
    FAutoSizeGrowOnly: boolean;
    FIndentVert: byte;
    FIndentHor: byte;
    FLineSpacing: integer;
    FRotate: integer;
    FSymbolWrap: boolean;
    FTransparent: boolean;
    FWordWrap: boolean;
    {$IFNDEF FPC}
      {$IFNDEF DELPHI_UNICODE}
    FOnMouseLeave: TNotifyEvent;
    FOnMouseEnter: TNotifyEvent;
      {$ENDIF}
    {$ENDIF}

    procedure SetAlignmentHorizontal(Value: byte);
    procedure SetAlignmentVertical(Value: byte);
    procedure SetAutoSizeHeight(Value: boolean);
    procedure SetAutoSizeWidth(Value: boolean);
    procedure SetAutoSizeGrowOnly(Value: boolean);
    procedure SetIndentVert(Value: byte);
    procedure SetIndentHor(Value: byte);
    procedure SetLineSpacing(Value: integer);
    procedure SetRotate(Value: integer);
    procedure SetSymbolWrap(Value: boolean);
    procedure SetTransparent(Value: boolean);
    procedure SetWordWrap(Value: boolean);
    {$IFNDEF FPC}
    procedure CMTextChanged(var Message: TMessage); message CM_TEXTCHANGED;
      {$IFNDEF DELPHI_UNICODE}
    procedure CMMouseEnter(var Message: TMessage); message CM_MOUSEENTER;
    procedure CMMouseLeave(var Message: TMessage); message CM_MOUSELEAVE;
      {$ENDIF}
    {$ENDIF}
  protected
    procedure Paint(); override;
    {$IFDEF FPC}
    procedure TextChanged(); override;
    {$ENDIF}
  public
    constructor Create(AOwner: TComponent); override;
    procedure DrawTextOn(ACanvas: TCanvas; AText: string; var ARect: TRect; CalcOnly: boolean; ClipArea: boolean = false); overload;
    procedure DrawTextOn(ACanvas: TCanvas; AText: string; var ARect: TRect; AColor: TColor; CalcOnly: boolean; ClipArea: boolean = false); overload;
    property Canvas;
  published
    property AlignmentVertical: byte read FAlignmentVertical write SetAlignmentVertical default 0;
    property AlignmentHorizontal: byte read FAlignmentHorizontal write SetAlignmentHorizontal default 0;
    property AutoSizeHeight: boolean read FAutoSizeHeight write SetAutoSizeHeight default false;
    property AutoSizeWidth: boolean read FAutoSizeWidth write SetAutoSizeWidth default false;
    property AutoSizeGrowOnly: boolean read FAutoSizeGrowOnly write SetAutoSizeGrowOnly default false;
    property Caption;
    property Color;
    property Font;
    property IndentVert: byte read FIndentVert write SetIndentVert default 0;
    property IndentHor: byte read FIndentHor write SetIndentHor default 3;
    property LineSpacing: integer read FLineSpacing write SetLineSpacing default 0;
    property Rotate: integer read FRotate write SetRotate default 0;
    property SymbolWrap: boolean read FSymbolWrap write SetSymbolWrap default false;
    property Transparent: boolean read FTransparent write SetTransparent default true;
    property WordWrap: boolean read FWordWrap write SetWordWrap default true;

    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property ShowHint;

    property ParentColor;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;

    property Visible;
    property OnClick;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEndDrag;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    {$IFDEF FPC}
    property OnMouseEnter;
    property OnMouseLeave;
    {$ELSE}
      {$IFDEF DELPHI_UNICODE}
     property OnMouseEnter;
     property OnMouseLeave;
      {$ELSE}
     property OnMouseEnter: TNotifyEvent read FOnMouseEnter write FOnMouseEnter;
     property OnMouseLeave: TNotifyEvent read FOnMouseLeave write FOnMouseLeave;
      {$ENDIF}
    {$ENDIF}

    property OnContextPopup;
    property OnResize;
    property OnStartDrag;
  end;

procedure Register();

implementation

//////////////////////// TZCLabel ////////////

constructor TZCLabel.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Width := 65;
  Height := 17;
  FTransparent := true;
  FWordWrap := true;
  FLineSpacing := 0;
  FAlignmentHorizontal := 0;
  FAlignmentVertical := 0;
  FAutoSizeHeight := false;
  FAutoSizeWidth := false;
  FAutoSizeGrowOnly := false;
  {$IFDEF FPC}
//    Canvas.Brush.Style := bsClear; //почему-то без этой строчки TextOut глючит с прозрачностью
  {$ENDIF}
end;

procedure TZCLabel.SetAlignmentHorizontal(Value: byte);
begin
  if (FAlignmentHorizontal <> Value) then
    if ((Value >=0) and (Value <= 3)) then
    begin
      FAlignmentHorizontal := Value;
      Invalidate();
    end;
end;

procedure TZCLabel.SetAlignmentVertical(Value: byte);
begin
  if (FAlignmentVertical <> Value) then
    if ((Value >=0) and (Value <= 2)) then
    begin
      FAlignmentVertical := Value;
      Invalidate();
    end;
end;

procedure TZCLabel.SetAutoSizeHeight(Value: boolean);
begin
  if (FAutoSizeHeight <> Value) then
  begin
    FAutoSizeHeight := Value;
    Invalidate();
  end;
end;

procedure TZCLabel.SetAutoSizeWidth(Value: boolean);
begin
  if (FAutoSizeWidth <> Value) then
  begin
    FAutoSizeWidth := Value;
    Invalidate();
  end;
end;

procedure TZCLabel.SetAutoSizeGrowOnly(Value: boolean);
begin
  if (FAutoSizeGrowOnly <> Value) then
  begin
    FAutoSizeGrowOnly := Value;
    Invalidate();
  end;
end;

procedure TZCLabel.SetIndentVert(Value: byte);
begin
  if (FIndentVert <> Value) then
  begin
    FIndentVert := Value;
    Invalidate();
  end;
end;

procedure TZCLabel.SetIndentHor(Value: byte);
begin
  if (FIndentHor <> Value) then
  begin
    FIndentHor := Value;
    Invalidate();
  end;
end;

procedure TZCLabel.SetLineSpacing(Value: integer);
begin
  if (FLineSpacing <> Value) then
  begin
    FLineSpacing := Value;
    Invalidate();
  end;
end;

procedure TZCLabel.SetRotate(Value: integer);
begin
  if (FRotate <> Value) then
  begin
    FRotate := Value;
    Invalidate();
  end;
end;

procedure TZCLabel.SetSymbolWrap(Value: boolean);
begin
  if (FSymbolWrap <> Value) then
  begin
    FSymbolWrap := Value;
    Invalidate();
  end;
end;

procedure TZCLabel.SetTransparent(Value: boolean);
begin
  if (FTransparent <> Value) then
  begin
    FTransparent := Value;
    Invalidate();
  end;
end;

procedure TZCLabel.SetWordWrap(Value: boolean);
begin
  if (FWordWrap <> Value) then
  begin
    FWordWrap := Value;
    Invalidate();
  end;
end;

//Рисует текст AText на канвасе ACanvas
//INPUT
//        ACanvas: TCanvas  - холст для рисования
//        AText: string     - текст
//    var ARect: TRect      - прямоугольник для текста
//        AColor: TColor    - цвет надписи
//        CalcOnly: boolean - если true, то только рассчитывает ARect
//        ClipArea: boolean - если true, то обрезает текст, который не поместился в ARect
procedure TZCLabel.DrawTextOn(ACanvas: TCanvas; AText: string; var ARect: TRect; AColor: TColor; CalcOnly: boolean; ClipArea: boolean = false);
var
  params: integer;

begin
  params := 0;
  if (AutoSizeHeight) then
    params := params or ZCF_SIZINGH;
  if (AutoSizeWidth) then
    params := params or ZCF_SIZINGW;
  if (CalcOnly) then
    params := params or ZCF_CALCONLY;
  if (SymbolWrap) then
    params := params or ZCF_SYMBOL_WRAP;
  if (not ClipArea) then
    params := params or ZCF_NO_CLIP;

  ZCWriteTextFormatted(ACanvas,
                       AText,
                       Font,
                       AColor,
                       AlignmentHorizontal,
                       AlignmentVertical,
                       WordWrap,
                       ARect,
                       IndentHor,
                       IndentVert,
                       params,
                       LineSpacing,
                       Rotate);
end;

//Рисует текст AText на канвасе ACanvas. Цвет берёт из шрифта.
//INPUT
//        ACanvas: TCanvas  - холст для рисования
//        AText: string     - текст
//    var ARect: TRect      - прямоугольник для текста
//        CalcOnly: boolean - если true, то только рассчитывает ARect
//        ClipArea: boolean - если true, то обрезает текст, который не поместился в ARect
procedure TZCLabel.DrawTextOn(ACanvas: TCanvas; AText: string; var ARect: TRect; CalcOnly: boolean; ClipArea: boolean = false);
begin
  DrawTextOn(ACanvas, AText, ARect, Font.Color, CalcOnly, ClipArea);
end;

procedure TZCLabel.Paint();
var
  Rct: TRect;

  {$IFDEF FPC} //можно lclproc подключить ...
  procedure OffsetRect(var Rct: TRect; dx, dy: integer);
  begin
    inc(Rct.Top, dy);
    inc(Rct.Bottom, dy);
    inc(Rct.Left, dx);
    inc(Rct.Right, dx);
  end;
  {$ENDIF}

begin
  Rct := ClientRect;
  if (not FAutoSizeGrowOnly) then
  begin
    if (FAutoSizeWidth) then
      Rct.Right := Rct.Left;
    if (FAutoSizeHeight) then
      Rct.Bottom := Rct.Top;
  end;

  DrawTextOn(Canvas, Caption, Rct, true);

  Width := Rct.Right - Rct.Left;
  Height := Rct.Bottom - Rct.Top;

  if (not Transparent) then
  with Canvas do
  begin
    Brush.Color := Self.Color;
    Brush.Style := bsSolid;
    FillRect(Rct);
    Brush.Style := bsClear;
  end;

  if (Enabled) then
    DrawTextOn(Canvas, Caption, Rct, Font.Color, false)
  else
  begin
    OffsetRect(Rct, 1, 1);
    DrawTextOn(Canvas, Caption, Rct, clBtnHighlight, false);
    OffsetRect(Rct, -1, -1);
    DrawTextOn(Canvas, Caption, Rct, clBtnShadow, false);
  end;

end;

{$IFDEF FPC}
procedure TZCLabel.TextChanged();
begin
  Invalidate;
end;

{$ELSE}
procedure TZCLabel.CMTextChanged(var Message: TMessage);
begin
  Invalidate;
end;

{$IFNDEF DELPHI_UNICODE}
procedure TZCLabel.CMMouseEnter(var Message: TMessage);
begin
  inherited;
  if (Assigned(FOnMouseEnter)) then
    FOnMouseEnter(Self);
end;

procedure TZCLabel.CMMouseLeave(var Message: TMessage);
begin
  inherited;
  if (Assigned(FOnMouseLeave)) then
    FOnMouseLeave(Self);
end;
{$ENDIF}

{$ENDIF}

//Регистрация
procedure Register();
begin
  RegisterComponents('zcolor', [TZCLabel]);
end;

initialization
{$IFDEF FPC}
  {$I 'zclabel.lrs'}
{$ENDIF}

end.

