//������������ ZEXMLSS
//������ �������� ����� ������� excel XML spreadsheet (SpreadsheetML)
//� StringGrid

unit unit_main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls,
  zexmlss {���������},
  zexmlssutils {��������/���������� � ������� Excel XML};

type
  TfrmMain = class(TForm)
    SGtest: TStringGrid;
    btnOpen: TButton;
    ODxml: TOpenDialog;
    CBList: TComboBox;
    lblList: TLabel;
    ZEXMLSStest: TZEXMLSS;
    procedure btnOpenClick(Sender: TObject);
    procedure CBListSelect(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure OpenEXMLSS();
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

//��������� ���� � ZEXMLSStest
procedure TfrmMain.OpenEXMLSS();
var
  i: integer;

begin
  if ODxml.Execute then
  begin
    //������ ����
    if ReadEXMLSS(ZEXMLSStest, ODxml.FileName) <> 0 then
      messagebox(0, '��� ������ ��������� �������� ������!', '������!', mb_Ok + mb_iconerror);

    CBList.Clear();
    CBList.Enabled := false;
    if ZEXMLSStest.Sheets.Count > 0 then
    begin
      for i := 0 to ZEXMLSStest.Sheets.Count - 1 do
        CBList.Items.Add(ZEXMLSStest.Sheets[i].Title);
      CBList.ItemIndex := 0;
      CBList.OnSelect(self);
      CBList.Enabled := true;  
    end;
  end;
end;

procedure TfrmMain.btnOpenClick(Sender: TObject);
begin
  OpenEXMLSS();
end;

//��� ������ ����� ��������� � SGTest ������
procedure TfrmMain.CBListSelect(Sender: TObject);
var
  PageNum, ColCount, RowCount: integer;

begin
  PageNum := CBList.ItemIndex; // ��������� ����
  ColCount := ZEXMLSStest.Sheets[PageNum].ColCount;
  RowCount := ZEXMLSStest.Sheets[PageNum].RowCount;
  SGTest.ColCount := ColCount; //������������� ���-�� ��������/�����
  SGTest.RowCount := RowCount;
  XmlSSToGrid(SGTest,     //����
              ZEXMLSStest,//���������
              PageNum,    //����� ����� � ���������
              0, 0,       //���������� ������� � �����
              0, 0,       //����� ������� ���� ���������� ������� �� ����� ���������
              //������ ������ ���� ���������� ������� �� ����� ���������
              ColCount - 1, RowCount - 1,
              0           //������ ������� (0 - ��� ������)
              );
end;

end.
