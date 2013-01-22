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
  //������������: ����������� - �����, ��������� - �����
  ZColorStringGrid1.Cells[1, 0] := 'H:L + V:C';
  //������������: ����������� - ������, ��������� - �����
  ZColorStringGrid1.Cells[2, 0] := 'H:R + V:C';
  //������������: ����������� - �����, ��������� - ������
  ZColorStringGrid1.Cells[3, 0] := 'H:L + V:T';
  //������������: ����������� - �� ������, ��������� - �����
  ZColorStringGrid1.Cells[4, 0] := 'H:� + V:D';
  for i := 0 to 5 do
  begin
    r := i + 1;
    for j := 1 to 4 do
    begin
      ZColorStringGrid1.Cells[j, r] := IntToStr(i);
      //������ �� �����������
      ZColorStringGrid1.CellStyle[j, r].IndentH := i;
    end;
    //������
    ZColorStringGrid1.CellStyle[2, r].HorizontalAlignment := taRightJustify;
    //������
    ZColorStringGrid1.CellStyle[3, r].VerticalAlignment := vaTop;
    //�� ������
    ZColorStringGrid1.CellStyle[4, r].HorizontalAlignment := taCenter;
    //�����
    ZColorStringGrid1.CellStyle[4, r].VerticalAlignment := vaBottom;

    //������ �� ���������
    for j := 3 to 4 do
      ZColorStringGrid1.CellStyle[j, r].IndentV := i;
  end;
end;

end.
