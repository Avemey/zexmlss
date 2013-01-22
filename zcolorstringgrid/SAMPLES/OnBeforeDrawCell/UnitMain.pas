//Пример использования TZColorStringGrid-а.
//Показывает разницу между событиями OnBeforeTextDrawCell и OnDrawCell
//
//
unit UnitMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ZColorStringGrid, StdCtrls, ImgList;

type
  TfrmMain = class(TForm)
    ZSGBefore: TZColorStringGrid;
    ZSGAfter: TZColorStringGrid;
    Label1: TLabel;
    Label2: TLabel;
    ImgList: TImageList;
    procedure FormCreate(Sender: TObject);
    procedure ZSGBeforeBeforeTextDrawCell(Sender: TObject; ACol,
      ARow: Integer; Rect: TRect; State: TGridDrawState);
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
  //заполним гриды
  for i := 1 to ZSGBefore.ColCount do
  begin
    ZSGBefore.Cells[i, 0] := 'SText'+inttostr(i);
    ZSGAfter.Cells[i, 0] := ZSGBefore.Cells[i, 0];
  end;
  for i := 1 to ZSGBefore.RowCount do
  begin
   ZSGBefore.Cells[0, i] := '='+inttostr(i)+'=';
   ZSGAfter.Cells[0, i] := ZSGBefore.Cells[0, i];
   for j := 1 to ZSGBefore.ColCount do
   begin
     ZSGAfter.Cells[j, i] := 'TeXt'+inttostr(i*j);
     ZSGBefore.Cells[j, i] := ZSGAfter.Cells[j, i] ;
   end;
  end;
end;

//OnBeforeTextDrawCell и OnDrawCell
procedure TfrmMain.ZSGBeforeBeforeTextDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
  //Берём канвас ...
  with (Sender as TZColorStringGrid).Canvas do
  begin
   pen.Color := clGray;
   pen.Width := 5;
   MoveTo(Rect.Right, Rect.Bottom);
   lineTo(rect.Left, rect.Top);
   // для примера - в последнюю колонку не будем устанавливать TRANSPARENT
   // и сравним результаты
   if ACol <> (Sender as TZColorStringGrid).ColCount - 1 then
     SetBkMode(Handle, TRANSPARENT);
  end;
  //рисунок из ImegeList помещаем в верхнюю строку
  if ARow = 0 then
    ImgList.Draw((Sender as TZColorStringGrid).Canvas, (rect.left + rect.Right) div 2 - 8, rect.top +(rect.Bottom-rect.top) div 2 - 8, ACol, true);
end;

end.
