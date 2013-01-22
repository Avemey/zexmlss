//������ ������������� zexmlss
// ������ �������� � �������� ���������
unit unit_main;

{$IFDEF FPC}
  {$MODE Delphi}
{$ENDIF}  

interface

uses
  {$IFDEF FPC}
  LCLIntf, LResources,
  {$ELSE}
  Windows, Messages, Variants,
  {$ENDIF}
  StdCtrls,  SysUtils, Dialogs, Classes, Graphics, Controls, Forms,
  zexmlss, zexmlssutils, zsspxml;

type

  { TfrmMain }

  TfrmMain = class(TForm)
    btnCreate: TButton;
    procedure btnCreateClick(Sender: TObject);
  private
  public
    procedure CreateEXMLSS();
  end;

var
  frmMain: TfrmMain;

implementation

{$IFNDEF FPC}
  {$R *.dfm}
{$ENDIF}  

//������� �������� �� ��������� ��������� � ������ (��� ������� - � utf8)
function my_local_to_some_encoding(value: ansistring): ansistring;
begin
  result := AnsiToUtf8(value);
end;

//������� ����
procedure TfrmMain.CreateEXMLSS();
var
  sd: TSaveDialog;
  tz: TZEXMLSS;
  i, j: integer;
  TextConverter: TAnsiToCPConverter;

begin
  TextConverter := nil;
  {$IFNDEF FPC}
    {$IF CompilerVersion < 20} // < RAD Studio 2009
  TextConverter := @my_local_to_some_encoding;
    {$IFEND}
  {$ENDIF}
  sd := TSaveDialog.Create(nil);
  sd.Filter := 'Excel XML (*.xml)|*.xml';
  sd.DefaultExt := 'xml';
  sd.Options := sd.Options + [ofOverwritePrompt];
  if sd.Execute then
  begin
    tz := TZEXMLSS.Create(nil);
    try
      //� ��������� 2 ��������
      tz.Sheets.Count := 2;
      tz.Sheets[0].Title := '������� ��������';
      //������� �����
      tz.Styles.Count := 6;
      //0 - ��� ��������� (20)
      tz.Styles[0].Font.Size := 20;
      tz.Styles[0].Font.Style := [fsBold];
      tz.Styles[0].Font.Name := 'Tahoma';
      tz.Styles[0].BGColor := $CCFFCC;
      tz.Styles[0].CellPattern := ZPSolid;
      tz.Styles[0].Alignment.Horizontal := ZHCenter;
      tz.Styles[0].Alignment.Vertical := ZVCenter;
      tz.Styles[0].Alignment.WrapText := true;
      //1 - ����� ��� �������
      tz.Styles[1].Border[0].Weight := 1;
      tz.Styles[1].Border[0].LineStyle := ZEContinuous;
      for i := 1 to 3 do tz.Styles[1].Border[i].Assign(tz.Styles[1].Border[0]);
      //2 - ����� ��� ��������� ������� (������ �� ������)
      tz.Styles[2].Assign(tz.Styles[1]);
      tz.Styles[2].Font.Style := [fsBold];
      tz.Styles[2].Alignment.Horizontal := ZHCenter;
      //3-�� �����
      tz.Styles[3].Font.Size := 16;
      tz.Styles[3].Font.Name := 'Arial Black';
      //4-�� �����
      tz.Styles[4].Font.Size := 18;
      tz.Styles[4].Font.Name := 'Arial';
      //5-�� �����
      tz.Styles[5].Font.Size := 14;
      tz.Styles[5].Font.Name := 'Arial Black';

      with tz.Sheets[0] do
      begin
        //��������� ���������� ����� � ��������
        RowCount := 20;
        ColCount := 20;

        Cell[0, 0].CellStyle := 3;
        Cell[0, 0].Data := '������ ������������� zexmlss';

        //������ � �������
        Cell[0, 2].CellStyle := 4;
        Cell[0, 2].Data := '�������� ��������: http://avemey.com';
        Cell[0, 2].Href := 'http://avemey.com';
        Cell[0, 2].HRefScreenTip := '�������� �� ������'#13#10'(����� ����������� ��������� � ������)';
        MergeCells.AddRectXY(0, 2, 10, 2);

        //����������
        Cell[0, 3].CellStyle := 5;
        Cell[0, 3].Data := '���� ����������!';
        Cell[0, 3].Comment := '����� ���������� 1';
        Cell[0, 3].CommentAuthor := '���-�� 1 ������� ����������';
        Cell[0, 3].ShowComment := true;

        Cell[16, 3].CellStyle := 5;
        Cell[16, 3].Data := '���� ����������!';
        Cell[16, 3].Comment := '����� ���������� 2';
        Cell[16, 3].CommentAuthor := '���-�� 2 ������� ����������';
        Cell[16, 3].ShowComment := true;
        Cell[16, 3].AlwaysShowComment := true;
        ColWidths[16] := 160; //������ �������

        Cell[1, 1].CellStyle := 0;
        Cell[1, 1].Data := '������� ��������� (������� ��������)';
        MergeCells.AddRectXY(1, 1, 16, 1);

        Cell[0, 6].Data := '�����������'#13#10'������!';
        Cell[0, 6].CellStyle := 0;
        MergeCells.AddRectXY(0, 6, 3, 14);

        Cell[5, 5].CellStyle := 2;
        for i := 1 to 10 do
        begin
          Cell[5, 5 + i].CellStyle := 2;
          Cell[5, 5 + i].Data := inttostr(i) + '.4';
          Cell[5, 5 + i].CellType := ZENumber;
          Cell[5 + i, 5].Assign(Cell[5, 5 + i]);
          for j := 1 to 10 do
          begin
            Cell[5 + i, 5 + j].CellStyle := 1;
            Cell[5 + i, 5 + j].Data := inttostr(i * j);
            Cell[5 + i, 5 + j].CellType := ZENumber;
          end;
        end;

        SheetOptions.PaperSize := 8; //A3
        SheetOptions.PortraitOrientation := false;
      end;

      //�������� ������ � 0 �� 1-�� ��������
      tz.Sheets[1].Assign(tz.Sheets[0]);
      tz.Sheets[1].Title := '������� �������� (�������)';

      //�� ������ �������� ����� ������������ �������.
      //���������� ������� ���� R1C1
      with tz.Sheets[1] do
      for i := 1 to 10 do
      for j := 1 to 10 do
        Cell[5 + i, 5 + j].Formula := '=R6C*RC6'; // ���������� � �������� �� 1 ������

      //��������� 0-�� � 1-�� �������� � ����
      //��������� - utf8, ��� ���������='utf8' (��� utf8 ����� ''), BOM=''
      SaveXmlssToEXML(tz, sd.FileName, [0, 1], [], @TextConverter, 'utf8');
    finally
      tz.free();
    end;
  end;
  sd.Free();
end;

procedure TfrmMain.btnCreateClick(Sender: TObject);
begin
  CreateEXMLSS();
end;

{$IFDEF FPC}
initialization
  {$i unit_main.lrs}
{$ENDIF}

end.

