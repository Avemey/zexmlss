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
  //Set all the cells TrueType font
  for i := 0 to ZColorStringGrid1.ColCount - 1 do
  for j := 0 to ZColorStringGrid1.RowCount - 1 do
  begin
    ZColorStringGrid1.CellStyle[i, j].Font.Name := 'Tahoma';
    ZColorStringGrid1.CellStyle[i, j].Font.Size := 12;
  end;
  //Set rotate for merged cells
  ZColorStringGrid1.MergeCells.AddRectXY(0, 1, 0, 4);
  ZColorStringGrid1.CellStyle[0, 1].Rotate := 90;
  ZColorStringGrid1.Cells[0, 1] := 'Rotate' + sLineBreak +
    'text' + sLineBreak + '90 degrees';

  ZColorStringGrid1.MergeCells.AddRectXY(1, 0, 3, 0);
  ZColorStringGrid1.CellStyle[1, 0].Rotate := 180;
  ZColorStringGrid1.Cells[1, 0] := 'Rotate 180 degrees';

  ZColorStringGrid1.MergeCells.AddRectXY(2, 1, 4, 6);
  ZColorStringGrid1.CellStyle[2, 1].Rotate := 60;
  ZColorStringGrid1.CellStyle[2, 1].Font.Size := 16;
  ZColorStringGrid1.CellStyle[2, 1].HorizontalAlignment := taCenter;
  ZColorStringGrid1.Cells[2, 1] := 'Rotating' + sLineBreak +
    'multi-line' + sLineBreak + 'text at' +
    sLineBreak + '60 degrees';

  //Rotate in cells
  ZColorStringGrid1.CellStyle[1, 1].Rotate := 180;
  ZColorStringGrid1.Cells[1, 1] := 'Text';


{
  //Усталёўваны ўсім вочкам TrueType шрыфт
  for i := 0 to ZColorStringGrid1.ColCount - 1 do
  for j := 0 to ZColorStringGrid1.RowCount - 1 do
  begin
    ZColorStringGrid1.CellStyle[i, j].Font.Name := 'Tahoma';
    ZColorStringGrid1.CellStyle[i, j].Font.Size := 12;
  end;
  //Заваротак у аб'яднаных вочках
  ZColorStringGrid1.MergeCells.AddRectXY(0, 1, 0, 4);
  ZColorStringGrid1.CellStyle[0, 1].Rotate := 90;
  ZColorStringGrid1.Cells[0, 1] := 'Заваротак' + sLineBreak +
    'тэксту на' + sLineBreak + '90 градусаў';

  ZColorStringGrid1.MergeCells.AddRectXY(1, 0, 3, 0);
  ZColorStringGrid1.CellStyle[1, 0].Rotate := 180;
  ZColorStringGrid1.Cells[1, 0] := 'Заваротак на 180 градусаў';

  ZColorStringGrid1.MergeCells.AddRectXY(2, 1, 4, 6);
  ZColorStringGrid1.CellStyle[2, 1].Rotate := 60;
  ZColorStringGrid1.CellStyle[2, 1].Font.Size := 16;
  ZColorStringGrid1.CellStyle[2, 1].HorizontalAlignment := taCenter;
  ZColorStringGrid1.Cells[2, 1] := 'Заваротак' + sLineBreak +
    'шматрадковага' + sLineBreak + 'тэксту на' +
    sLineBreak + '60 градусаў';

  //Заваротак у звычайных вочках
  ZColorStringGrid1.CellStyle[1, 1].Rotate := 180;
  ZColorStringGrid1.Cells[1, 1] := 'Тэкст';
  }
end;

end.
