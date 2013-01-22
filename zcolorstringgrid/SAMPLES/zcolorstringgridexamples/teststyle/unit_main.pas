unit unit_main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ZColorStringGrid;

type
  TfrmMain = class(TForm)
    ZColorStringGrid1: TZColorStringGrid;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.FormCreate(Sender: TObject);
var
  i, j: integer;

begin
  //заполняем грид
  for i := 1 to ZColorStringGrid1.ColCount - 1 do
  for j := 1 to ZColorStringGrid1.RowCount - 1 do
    ZColorStringGrid1.Cells[i, j] := IntToStr(i * j);

  //Цвет фона
  ZColorStringGrid1.CellStyle[1, 1].BGColor := clYellow;
  ZColorStringGrid1.CellStyle[2, 1].BGColor := clGreen;
  ZColorStringGrid1.CellStyle[3, 1].BGColor := clLime;

  //Меняем стиль у колонки
  ZColorStringGrid1.CellStyleCol[1, false] := ZColorStringGrid1.CellStyle[1, 1];
  ZColorStringGrid1.CellStyleCol[2, true] := ZColorStringGrid1.CellStyle[2, 1];

  //Меняем стиль у ряда
  ZColorStringGrid1.CellStyleRow[3, false] := ZColorStringGrid1.CellStyle[3, 1];

  //Шрифт
  ZColorStringGrid1.CellStyle[2, 3].Font.Size := 12;
  ZColorStringGrid1.CellStyle[2, 3].Font.Name := 'Tahoma';
  ZColorStringGrid1.CellStyle[2, 3].Font.Style := [fsBold, fsItalic];
  ZColorStringGrid1.CellStyle[2, 2].Font.Color := clWhite;

  //Рамка ячейки
  ZColorStringGrid1.CellStyle[3, 3].BorderCellStyle := sgLowered;
  ZColorStringGrid1.CellStyle[4, 3].BorderCellStyle := sgRaised;
end;

end.
