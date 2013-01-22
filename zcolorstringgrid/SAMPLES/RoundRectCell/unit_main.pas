//Пример использования ZColorStringGrid
//Ячейки с закруглёнными рамками 
unit unit_main;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, ZColorStringGrid;

type
  TfrmMain = class(TForm)
    ZSG: TZColorStringGrid;
    procedure FormCanResize(Sender: TObject; var NewWidth,
      NewHeight: Integer; var Resize: Boolean);
    procedure ZSGBeforeTextDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.DFM}

procedure TfrmMain.FormCanResize(Sender: TObject; var NewWidth,
  NewHeight: Integer; var Resize: Boolean);
begin
  if (NewHeight < 599) or (NewWidth < 799) then Resize := false;
  ZSG.Width := Width - 15;
  ZSG.Height := Height - 40;
end;

procedure TfrmMain.ZSGBeforeTextDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
  if (acol>0) and (arow>0) then
  with (Sender as TZColorStringGrid).canvas do
  begin
    pen.Color := clGray;
    //рисуем закруглённый прямоугольник
    RoundRect(rect.left, rect.top, rect.right, rect.bottom, 9, 9);
    SetBkMode(zsg.Canvas.Handle, TRANSPARENT);
  end;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  i,j: integer;

begin
  //заполняем грид
  for i := 1 to zsg.RowCount - 1 do
  for j := 1 to zsg.ColCount - 1 do zsg.Cells[j, i] :=inttostr(i*j);
  for j := 1 to zsg.ColCount - 1 do zsg.Cells[j, 0] :=inttostr(j);
  for i := 1 to zsg.RowCount - 1 do zsg.Cells[0, i] :=inttostr(i);
end;

end.
