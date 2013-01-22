//Пример использования ZColorStringGrid-а
unit UnitMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ZColorStringGrid, StdCtrls, Buttons;

type
  TfrmMain = class(TForm)
    ZSG: TZColorStringGrid;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure ZSGKeyPress(Sender: TObject; var Key: Char);
    procedure ZSGSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
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
  i: integer;

begin
  ZSG.Cells[0, 0] := '№';
  ZSG.MergeCells.AddRectXY(0,0,0,1);
  ZSG.Cells[1, 0] := 'Название';
  i := ZSG.MergeCells.AddRectXY(1,0,1,1);
  // При необходимости можно делать проверку
  if i <> 0 then
    messagebox(0,'Не могу добавить объединённую ячейку/область','=_=',MB_OK);
  ZSG.Cells[2, 0] := 'Весна';
  ZSG.Cells[2, 1] := 'Март';
  ZSG.Cells[3, 1] := 'Апрель';
  ZSG.Cells[4, 1] := 'Май';

  ZSG.MergeCells.AddRectXY(2,0,4,0);
  ZSG.Cells[5, 0] := 'Лето';
  ZSG.MergeCells.AddRectXY(5,0,7,0);
  ZSG.Cells[5, 1] := 'Июнь';
  ZSG.Cells[6, 1] := 'Июль';
  ZSG.Cells[7, 1] := 'Август';
  ZSG.MergeCells.AddRectXY(8,0,8,1);
  ZSG.Cells[8, 0] := 'Всего';


  ZSG.CellStyle[1, 2].HorizontalAlignment := taLeftJustify;
  ZSG.CellStyle[1, 2].Font.Name := 'Times New Roman';
  ZSG.CellStyleCol[1, false] := ZSG.CellStyle[1, 2];
  for i := 2 to ZSG.RowCount - 1 do
    ZSG.Cells[0, i] := inttostr(i-1)+'.';
  ZSG.ColWidths[1] := 200;

end;

procedure TfrmMain.ZSGKeyPress(Sender: TObject; var Key: Char);
begin
  with (Sender as TZColorStringGrid) do
  begin
    if Col = ColCount -1 then  Key := #0;
  end;
end;

//изменение текста в ячейке 
procedure TfrmMain.ZSGSetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
var
  i: integer;
  summ: real;
  s : string;

begin
  if (ACol > 1)and(ARow>1) then
  with (Sender As TZColorStringGrid) do
  begin
    summ := 0;
    for  i:= 2 to 7 do
    //про try не забываем, но если запуск из IDE...
    Try
      s := Cells[i, ARow];
      s := trim(s);
      if s = '' then s := '0';
      summ := summ + strtofloat(s);
    except
      on EConvertError do
      begin
        Cells[8, ARow] := 'O_o';
        exit;
      end
    end;
    if abs(summ) - 0.0000001 <= 0 then Cells[8, ARow] := '' else
    Cells[8, ARow] := floattostr(summ);
  end;
end;

end.
