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
  i, j, r: byte;

begin
  //Выравнивание: горизонталь - слева, вертикаль - центр
  ZColorStringGrid1.Cells[1, 0] := 'H:L + V:C';
  //Выравнивание: горизонталь - справа, вертикаль - центр
  ZColorStringGrid1.Cells[2, 0] := 'H:R + V:C';
  //Выравнивание: горизонталь - слева, вертикаль - сверху
  ZColorStringGrid1.Cells[3, 0] := 'H:L + V:T';
  //Выравнивание: горизонталь - по центру, вертикаль - снизу
  ZColorStringGrid1.Cells[4, 0] := 'H:С + V:D';
  for i := 0 to 5 do
  begin
    r := i + 1;
    for j := 1 to 4 do
    begin
      ZColorStringGrid1.Cells[j, r] := IntToStr(i);
      //Отступ по горизонтали
      ZColorStringGrid1.CellStyle[j, r].IndentH := i;
    end;
    //Справа
    ZColorStringGrid1.CellStyle[2, r].HorizontalAlignment := taRightJustify;
    //Сверху
    ZColorStringGrid1.CellStyle[3, r].VerticalAlignment := vaTop;
    //по центру
    ZColorStringGrid1.CellStyle[4, r].HorizontalAlignment := taCenter;
    //Снизу
    ZColorStringGrid1.CellStyle[4, r].VerticalAlignment := vaBottom;

    //Отступ по вертикали
    for j := 3 to 4 do
      ZColorStringGrid1.CellStyle[j, r].IndentV := i;
  end;
end;

end.
