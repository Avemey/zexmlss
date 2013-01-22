//Пример совместного ипользования ZColorStringGrid и ZEXMLSS
unit unit_main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ZColorStringGrid, ExtCtrls, StdCtrls,
  Buttons, math, zclab, ComCtrls;

type
  TfrmMain = class(TForm)
    ZSG: TZColorStringGrid;
    PanelTop: TPanel;
    btnVTop: TBitBtn;
    btnVCenter: TBitBtn;
    btnVBottom: TBitBtn;
    btnHLeft: TBitBtn;
    btnHCenter: TBitBtn;
    btnHRight: TBitBtn;
    LabelV: TLabel;
    LabelH: TLabel;
    btnFont: TBitBtn;
    btnBGColor: TBitBtn;
    btnMerge: TButton;
    FntDialog: TFontDialog;
    BGColorDialog: TColorDialog;
    Bevel1: TBevel;
    ZCLabel1: TZCLabel;
    ZCLabel2: TZCLabel;
    ZCLabel3: TZCLabel;
    EditRotate: TEdit;
    Label1: TLabel;
    RotateUP: TUpDown;
    btnSetRotate: TButton;
    procedure btnVTopClick(Sender: TObject);
    procedure btnVCenterClick(Sender: TObject);
    procedure btnVBottomClick(Sender: TObject);
    procedure btnHLeftClick(Sender: TObject);
    procedure btnHCenterClick(Sender: TObject);
    procedure btnHRightClick(Sender: TObject);
    procedure btnFontClick(Sender: TObject);
    procedure btnBGColorClick(Sender: TObject);
    procedure btnMergeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnSetRotateClick(Sender: TObject);
  private
    procedure GridSelection();
    procedure SetVA(value: TVerticalAlignment);
    procedure SetHA(value: TAlignment);
    procedure AfterButton();
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

//Действие после нажатия кнопок
procedure TfrmMain.AfterButton();
begin
  if zsg.CanFocus then
    zsg.SetFocus();
end;

//Если выбрана объединённая ячейка - установить selection размер объединённой ячейки
procedure TfrmMain.GridSelection();
var
  t: integer;

begin
  if not (abs(ZSG.Selection.Top - ZSG.Selection.Bottom) > 0) or
     (abs(ZSG.Selection.Left - ZSG.Selection.right) > 0) then
  begin
    t := ZSG.MergeCells.InMergeRange(ZSG.Col, ZSG.Row);
    if t >= 0 then
      ZSG.Selection := TGridRect(ZSG.MergeCells.Items[t]);
  end;
end;

//Установить выравнивание по вертикали
procedure TfrmMain.SetVA(value: TVerticalAlignment);
var
  i, j: integer;

begin
  GridSelection();
  for i:=min(zsg.Selection.left, zsg.Selection.Right) to
    max(zsg.Selection.left, zsg.Selection.Right)  do
  for j:=min(zsg.Selection.top, zsg.Selection.Bottom) to
    max(zsg.Selection.top, zsg.Selection.Bottom) do
  zsg.CellStyle[i, j].VerticalAlignment := value;
  AfterButton();
end;

//Установить выравнивание по горизонтали
procedure TfrmMain.SetHA(value: TAlignment);
var
  i, j: integer;

begin
  GridSelection();
  for i:=min(zsg.Selection.left, zsg.Selection.Right) to
    max(zsg.Selection.left, zsg.Selection.Right)  do
  for j:=min(zsg.Selection.top, zsg.Selection.Bottom) to
    max(zsg.Selection.top, zsg.Selection.Bottom) do
  zsg.CellStyle[i, j].HorizontalAlignment := value;
  AfterButton();
end;

procedure TfrmMain.btnVTopClick(Sender: TObject);
begin
  SetVA(vaTop);
end;

procedure TfrmMain.btnVCenterClick(Sender: TObject);
begin
  SetVA(vaCenter);
end;

procedure TfrmMain.btnVBottomClick(Sender: TObject);
begin
  SetVA(vaBottom);
end;

procedure TfrmMain.btnHLeftClick(Sender: TObject);
begin
  SetHa(taLeftJustify);
end;

procedure TfrmMain.btnHCenterClick(Sender: TObject);
begin
  SetHa(taCenter);
end;

procedure TfrmMain.btnHRightClick(Sender: TObject);
begin
  SetHa(taRightJustify)
end;

//Установить шрифт
procedure TfrmMain.btnFontClick(Sender: TObject);
var
  i, j: integer;

begin
  GridSelection();
  FntDialog.Font.Assign(ZSG.CellStyle[ZSG.Col, ZSG.Row].Font);
  if FntDialog.Execute then
  begin
    for i:=min(zsg.Selection.left, zsg.Selection.Right) to
      max(zsg.Selection.left, zsg.Selection.Right)  do
    for j:=min(zsg.Selection.top, zsg.Selection.Bottom) to
        max(zsg.Selection.top, zsg.Selection.Bottom) do
      ZSG.CellStyle[i, j].Font := FntDialog.Font;
  end;
  AfterButton();
end;

//Установить цвет фона ячейки
procedure TfrmMain.btnBGColorClick(Sender: TObject);
var
  i, j: integer;
  
begin
  GridSelection();
  BGColorDialog.Color := ZSG.CellStyle[ZSG.Col, ZSG.Row].BGColor;
  if BGColorDialog.Execute then
  begin
    for i:=min(zsg.Selection.left, zsg.Selection.Right) to
      max(zsg.Selection.left, zsg.Selection.Right)  do
    for j:=min(zsg.Selection.top, zsg.Selection.Bottom) to
      max(zsg.Selection.top, zsg.Selection.Bottom) do
    ZSG.CellStyle[i, j].BGColor := BGColorDialog.Color;
  end;
  AfterButton();
end;

//Объединить/разъединить ячейки
procedure TfrmMain.btnMergeClick(Sender: TObject);
var
  i, j, k: integer;
  haveMerge: boolean;
  
begin
  GridSelection();
  haveMerge := false;
  //пробегаем по выделенной области, если в ней есть объединённые ячейки -
  //удаляем их, если таких ячеек небыло - добавляем объединённую ячейку.
  for i:=min(zsg.Selection.left, zsg.Selection.Right) to
      max(zsg.Selection.left, zsg.Selection.Right)  do
  for j:=min(zsg.Selection.top, zsg.Selection.Bottom) to
      max(zsg.Selection.top, zsg.Selection.Bottom) do
  begin
    //номер объединённой области (-1 - не входит в объединённые ячейки)
    k := zsg.MergeCells.InMergeRange(i, j);
    if k >= 0 then
    begin
      //удаляем область (разъединяем ячейки)
      zsg.MergeCells.DeleteItem(k);
      haveMerge := true;
    end;
  end;
  if not haveMerge then
    zsg.MergeCells.AddRectXY(zsg.Selection.Left,
                            zsg.Selection.top,
                            zsg.Selection.Right,
                            zsg.Selection.Bottom);
  AfterButton();
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  i: integer;
  
begin
  zclabel2.Caption := 'Поворот'#13#10+'многострочного'#13#10+'текста на'#13#10+'заданный'#13#10+'угол';
  ZSG.CellStyle[0, 2].HorizontalAlignment := taRightJustify;
  ZSG.CellStyleCol[0, true] := ZSG.CellStyle[0, 2];
  for i := 2 to ZSG.RowCount - 1 do
    ZSG.Cells[0, i] := IntToStr(i-1)+'.';
  ZSG.ColWidths[0] := 30;

  ZSG.MergeCells.AddRectXY(0, 0, 0, 1);
  ZSG.Cells[0, 0] := '№';

  ZSG.MergeCells.AddRectXY(2, 0, 4, 0);
  ZSG.MergeCells.AddRectXY(5, 0, 7, 0);
  ZSG.Cells[2, 0] := 'Заголовок';
  ZSG.Cells[5, 0] := 'Заголовок (180)';
  ZSG.CellStyle[5, 0].Rotate := 180;

  ZSG.MergeCells.AddRectXY(1, 0, 1, 1);
  ZSG.Cells[1, 0] := '300';
  ZSG.CellStyle[1, 0].Font.Size := 12;
  ZSG.CellStyle[1, 0].Rotate := -60;

  ZSG.MergeCells.AddRectXY(1, 2, 1, 6);
  ZSG.CellStyle[1, 2].Rotate := 90;
  ZSG.Cells[1, 2] := 'Какой-то'#10' текст';
  ZSG.MergeCells.AddRectXY(1, 7, 1, 12);
  ZSG.CellStyle[1, 7].Rotate := 270;
  ZSG.Cells[1, 7] := ZSG.Cells[1, 2] + ' 270';

  for i := 0 to 10 do
  begin
    ZSG.CellStyle[2, 2 + i].IndentH := i;
    ZSG.Cells[2, 2 + i] := 'Отступ ' + IntToStr(i);
  end;
end;

procedure TfrmMain.btnSetRotateClick(Sender: TObject);
var
  i, j: integer;

begin
  GridSelection();
  for i:=min(zsg.Selection.left, zsg.Selection.Right) to
    max(zsg.Selection.left, zsg.Selection.Right)  do
  for j:=min(zsg.Selection.top, zsg.Selection.Bottom) to
    max(zsg.Selection.top, zsg.Selection.Bottom) do
    zsg.CellStyle[i, j].Rotate := RotateUP.Position;
  AfterButton();
end;

end.
