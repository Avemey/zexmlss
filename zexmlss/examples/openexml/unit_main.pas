//Демонстрация ZEXMLSS
//Пример загрузки файла формата excel XML spreadsheet (SpreadsheetML)
//в StringGrid

unit unit_main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls,
  zexmlss {хранилище},
  zexmlssutils {загрузка/сохранение в формате Excel XML};

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

//Загружает файл в ZEXMLSStest
procedure TfrmMain.OpenEXMLSS();
var
  i: integer;

begin
  if ODxml.Execute then
  begin
    //читаем файл
    if ReadEXMLSS(ZEXMLSStest, ODxml.FileName) <> 0 then
      messagebox(0, 'При чтении документа возникла ошибка!', 'Ошибка!', mb_Ok + mb_iconerror);

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

//При выборе листа загружаем в SGTest данные
procedure TfrmMain.CBListSelect(Sender: TObject);
var
  PageNum, ColCount, RowCount: integer;

begin
  PageNum := CBList.ItemIndex; // выбранный лист
  ColCount := ZEXMLSStest.Sheets[PageNum].ColCount;
  RowCount := ZEXMLSStest.Sheets[PageNum].RowCount;
  SGTest.ColCount := ColCount; //устанавливаем кол-во столбцов/строк
  SGTest.RowCount := RowCount;
  XmlSSToGrid(SGTest,     //Грид
              ZEXMLSStest,//Хранилище
              PageNum,    //Номер листа в хранилище
              0, 0,       //Координаты вставки в гриде
              0, 0,       //Левый верхний угол копируемой области из листа хранилища
              //правый нижний угол копируемой области из листа хранилища
              ColCount - 1, RowCount - 1,
              0           //способ вставки (0 - без сдвига)
              );
end;

end.
