//������ ����������� ������������ ZColorStringGrid � ZEXMLSS
unit unit_main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ZColorStringGrid, ExtCtrls, StdCtrls,
  Buttons, math,
  zexmlss {���������},  
  zexmlssutils {��������/���������� � ������� Excel XML},
  zsspxml;

type
  TfrmMain = class(TForm)
    ZSG: TZColorStringGrid;
    PanelTop: TPanel;
    btnLoad: TButton;
    btnSave: TButton;
    ODxml: TOpenDialog;
    SDxml: TSaveDialog;
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
    ZEXMLSStest: TZEXMLSS;
    BGColorDialog: TColorDialog;
    RGShift: TRadioGroup;
    Bevel1: TBevel;
    CBInsertStart: TCheckBox;
    CBCopyBG: TCheckBox;
    CBCopyFont: TCheckBox;
    CBCopyMerge: TCheckBox;
    procedure btnLoadClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnVTopClick(Sender: TObject);
    procedure btnVCenterClick(Sender: TObject);
    procedure btnVBottomClick(Sender: TObject);
    procedure btnHLeftClick(Sender: TObject);
    procedure btnHCenterClick(Sender: TObject);
    procedure btnHRightClick(Sender: TObject);
    procedure btnFontClick(Sender: TObject);
    procedure btnBGColorClick(Sender: TObject);
    procedure btnMergeClick(Sender: TObject);
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

//�������� ����� ������� ������
procedure TfrmMain.AfterButton();
begin
  if zsg.CanFocus then
    zsg.SetFocus();
end;

//���� ������� ����������� ������ - ���������� selection ������ ����������� ������
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

//���������� ������������ �� ���������
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

//���������� ������������ �� �����������
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

//�������� �� �����
procedure TfrmMain.btnLoadClick(Sender: TObject);
var
  PageNum, ColCount, RowCount: integer;
  StartRow, StartCol: integer;
  InsertMode: integer; //������ �������
  StyleCopy: integer;

begin
  if ODxml.Execute then
  begin
    //������ ����
    if ReadEXMLSS(ZEXMLSStest, ODxml.FileName) <> 0 then
      messagebox(0, '��� ������ ��������� �������� ������!', '������!', mb_Ok + mb_iconerror);

    if ZEXMLSStest.Sheets.Count > 0 then
    begin
      StartRow := 1;
      StartCol := 1;
      if CBInsertStart.Checked then
      begin
        StartRow := ZSG.Row;
        StartCol := ZSG.Col;
      end;

      InsertMode := RGShift.ItemIndex;

      StyleCopy := 950; //(2 or 4 or 16 or 32 or 128 or 256 or 512)
      if CBCopyBG.Checked then StyleCopy := StyleCopy or 1;
      if CBCopyFont.Checked then StyleCopy := StyleCopy or 8;
      if CBCopyMerge.Checked then StyleCopy := StyleCopy or 64;

      PageNum := 0; // ��������� ����                                             
      RowCount := ZEXMLSStest.Sheets[PageNum].RowCount;
      ColCount := ZEXMLSStest.Sheets[PageNum].ColCount;
      XmlSSToGrid(ZSG,        //����
              ZEXMLSStest,    //���������
              PageNum,        //����� ����� � ���������
              StartCol, StartRow,       //���������� ������� � �����
              0, 0,           //����� ������� ���� ���������� ������� �� ����� ���������
              //������ ������ ���� ���������� ������� �� ����� ���������
              ColCount - 1, RowCount - 1,
              InsertMode,     //������ ������� (0 - ��� ������)
              StyleCopy       //��� �� ����� ����������
              );
    end;
  end;
  AfterButton();
end;

//��������� � �����
procedure TfrmMain.btnSaveClick(Sender: TObject);
var
  i: integer;
  lastCol, lastRow: integer;
  TextConverter: TAnsiToCPConverter;

begin
  TextConverter := nil;
  {$IFNDEF FPC}
    {$IF CompilerVersion < 20} // < RAD Studio 2009
  TextConverter := @AnsiToUtf8;
    {$IFEND}
  {$ENDIF}
  try
    if SDxml.Execute then
    begin
      //������� �����
      ZEXMLSStest.Sheets.Count := 0;
      //������� �����
      ZEXMLSStest.Styles.Clear();
      //������ ���-�� ������
      ZEXMLSStest.Sheets.Count := 6;

      //������ ������ ������
      lastCol := ZSG.ColCount - 1;
      lastRow := ZSG.RowCount - 1;

      //�������� �� ����� ��������� �����������
      GridToXmlSS(
                  ZEXMLSStest,  //���������
                  0,            //����� ��������
                  ZSG,          //����������
                  0, 0,         //���� ���������
                  1, 1,         //����� ������� ���� ���������� ������� � �����������
                  //������ ������ ���� ���������� ������� � ������������
                  lastCol, lastRow,
                  false,        //������������ ���� ������
                  1             //��������� ����� ������
                 );
      GridToXmlSS(ZEXMLSStest, 1, ZSG, 0, 0, 1, 1, lastCol, lastRow, false, 0);
      GridToXmlSS(ZEXMLSStest, 2, ZSG, 0, 0, 0, 0, lastCol, lastRow, false, 2);
      GridToXmlSS(ZEXMLSStest, 3, ZSG, 0, 0, 1, 1, lastCol, lastRow, true,  0);
      GridToXmlSS(ZEXMLSStest, 4, ZSG, 0, 0, 1, 1, lastCol, lastRow, true,  1);
      GridToXmlSS(ZEXMLSStest, 5, ZSG, 0, 0, 0, 0, lastCol, lastRow, true,  2);
      for i := 0 to 5 do
        ZEXMLSStest.Sheets[i].Title := '������ ' + inttostr(i+1);
      SaveXmlssToEXML(ZEXMLSStest, SDxml.FileName, [], [], TextConverter, 'utf8');
    end;
  except
  end;
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

//���������� �����
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

//���������� ���� ���� ������
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

//����������/����������� ������
procedure TfrmMain.btnMergeClick(Sender: TObject);
var
  i, j, k: integer;
  haveMerge: boolean;
  
begin
  GridSelection();
  haveMerge := false;
  //��������� �� ���������� �������, ���� � ��� ���� ����������� ������ -
  //������� ��, ���� ����� ����� ������ - ��������� ����������� ������.
  for i:=min(zsg.Selection.left, zsg.Selection.Right) to
      max(zsg.Selection.left, zsg.Selection.Right)  do
  for j:=min(zsg.Selection.top, zsg.Selection.Bottom) to
      max(zsg.Selection.top, zsg.Selection.Bottom) do
  begin
    //����� ����������� ������� (-1 - �� ������ � ����������� ������)
    k := zsg.MergeCells.InMergeRange(i, j);
    if k >= 0 then
    begin
      //������� ������� (����������� ������)
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

end.
