unit Unit1; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  EditBtn, zexmlss, zeodfs, zexmlssutils, zeformula, zsspxml;

type

  { TForm1 }

  TForm1 = class(TForm)
    btnCreate: TButton;
    btnFormula: TButton;
    edSavePath: TDirectoryEdit;
    Label1: TLabel;
    procedure btnCreateClick(Sender: TObject);
    procedure btnFormulaClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation
uses zexlsx, zeZippy, zeZippyLazip, zeSave;

{$R *.lfm}

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
begin

end;

procedure TForm1.btnCreateClick(Sender: TObject);
var
  XMLSS: TZEXMLSS;
  i, j: integer;
  TextConverter: TAnsiToCPConverter;
  sEOL: string;
  Path: TFileName;
begin
  Path := edSavePath.Directory;
  ForceDirectories(ExcludeTrailingPathDelimiter(Path));
  Path := IncludeTrailingPathDelimiter(Path);

  TextConverter := nil;
  {$IFNDEF FPC}
    {$IF CompilerVersion < 20} // < RAD Studio 2009
  TextConverter := @AnsiToUtf8;
    {$IFEND}
  {$ENDIF}

  {$IFDEF FPC}
  sEOL := LineEnding;
  {$ELSE}
  sEOL := sLineBreak;
  {$ENDIF}

  XMLSS := TZEXMLSS.Create(nil);
  try
    //В документе 2 страницы
    XMLSS.Sheets.Count := 2;
    XMLSS.Sheets[0].Title := 'Тестовая таблица';
    //добавим стили
    XMLSS.Styles.Count := 13;
    //0 - для заголовка (20)
    XMLSS.Styles[0].Font.Size := 20;
    XMLSS.Styles[0].Font.Style := [fsBold];
    XMLSS.Styles[0].Font.Name := 'Tahoma';
    XMLSS.Styles[0].BGColor := $CCFFCC;
    XMLSS.Styles[0].CellPattern := ZPSolid;
    XMLSS.Styles[0].Alignment.Horizontal := ZHCenter;
    XMLSS.Styles[0].Alignment.Vertical := ZVCenter;
    XMLSS.Styles[0].Alignment.WrapText := true;
    //1 - стиль для таблицы
    XMLSS.Styles[1].Border[0].Weight := 1;
    XMLSS.Styles[1].Border[0].LineStyle := ZEContinuous;
    for i := 1 to 3 do XMLSS.Styles[1].Border[i].Assign(
    	XMLSS.Styles[1].Border[0]);
    //2 - стиль для заголовка таблицы (жирный по центру)
    XMLSS.Styles[2].Assign(XMLSS.Styles[1]);
    XMLSS.Styles[2].Font.Style := [fsBold];
    XMLSS.Styles[2].Alignment.Horizontal := ZHCenter;
    //3-ий стиль
    XMLSS.Styles[3].Font.Size := 16;
    XMLSS.Styles[3].Font.Name := 'Arial Black';
    //4-ый стиль
    XMLSS.Styles[4].Font.Size := 18;
    XMLSS.Styles[4].Font.Name := 'Arial';
    //5-ый стиль
    XMLSS.Styles[5].Font.Size := 14;
    XMLSS.Styles[5].Font.Name := 'Arial Black';
    //6-ой
    XMLSS.Styles[6].Assign(XMLSS.Styles[1]);
    for i := 1 to 4 do
      XMLSS.Styles[6].Border[i].Weight := 3;
    XMLSS.Styles[6].Border[5].Weight := 2;
    XMLSS.Styles[6].Border[0].Color := clRed;
    XMLSS.Styles[6].Border[5].LineStyle := ZEContinuous;

    XMLSS.Styles[6].Border[5].Color := clGreen;
    XMLSS.Styles[6].BGColor := clYellow;
    XMLSS.Styles[6].CellPattern := ZPSolid;
    XMLSS.Styles[6].Font.Color := clRed;

    //Копируется стиль таблицы
    for i := 7 to 12 do
      XMLSS.Styles[i].Assign(XMLSS.Styles[1]);

    //Горизонтальное выравнивание (слева, центр, справа);
    XMLSS.Styles[7].Alignment.Horizontal := ZHLeft;
    XMLSS.Styles[8].Alignment.Horizontal := ZHCenter;
    XMLSS.Styles[9].Alignment.Horizontal := ZHRight;

    //Вертикальное выравнивание (сверху, центр, снизу);
    XMLSS.Styles[10].Alignment.Vertical := ZVTop;
    XMLSS.Styles[11].Alignment.Vertical := ZVCenter;
    XMLSS.Styles[12].Alignment.Vertical := ZVBottom;

    with XMLSS.Sheets[0] do
    begin
      //количество строк и столбцов
      RowCount := 50;
      ColCount := 20;

      Cell[0, 0].CellStyle := 3;
      Cell[0, 0].Data := 'Пример использования zexmlss';

      //ячейка с ссылкой
      Cell[0, 2].CellStyle := 4;
      Cell[0, 2].Data := 'Домашняя страница: http://avemey.com';
      Cell[0, 2].Href := 'http://avemey.com';
      Cell[0, 2].HRefScreenTip := 'Клацните по ссылке' + sEOL +
      	'(текст всплывающей подсказки к ссылке)';
      //Объединённая ячейка
      MergeCells.AddRectXY(0, 2, 10, 2);

      //Примечание
      Cell[0, 3].CellStyle := 5;
      Cell[0, 3].Data := 'Есть примечание!';
      Cell[0, 3].Comment := 'Текст примечания 1';
      Cell[0, 3].CommentAuthor := 'Кто-то 1 оставил примечание';
      Cell[0, 3].ShowComment := true;

      Cell[10, 3].CellStyle := 5;
      Cell[10, 3].Data := 'Тоже примечание!';
      Cell[10, 3].Comment := 'Текст примечания 2';
      Cell[10, 3].CommentAuthor := 'Кто-то 2 оставил примечание';
      Cell[10, 3].ShowComment := true;
      Cell[10, 3].AlwaysShowComment := true;
      ColWidths[10] := 160; //Ширина столбца
      Columns[10].WidthMM := 40; //~40 мм

      //Объединённые ячейки
      Cell[2, 1].CellStyle := 0;
      Cell[2, 1].Data := 'Какой-то заголовок';
      MergeCells.AddRectXY(2, 1, 12, 1);

      Cell[0, 6].Data := 'Объединённая' + sEOL + 'ячейка!';
      Cell[0, 6].CellStyle := 0;
      MergeCells.AddRectXY(0, 6, 3, 14);

      Cell[9, 6].Data := '№ стиля';
      Cell[9, 6].CellStyle := 2;

      Cell[10, 6].Data := 'Пример стиля';
      Cell[10, 6].CellStyle := 2;

      j := 8;

      for i := 0 to 6 do
      begin
        Cell[9,  j].Data := IntToStr(i) + '-ой';
        Cell[10, j].Data := 'текст';
        Cell[10, j].CellStyle := i;
        inc(j, 2);
      end;

      Columns[5].WidthMM := 30;
      Columns[6].WidthMM := 30;

      Cell[5, 5].Data := 'Выравнивание';
      MergeCells.AddRectXY(5, 5, 6, 5);

      Cell[5, 6].Data := 'Горизонталь';
      Cell[5, 6].CellStyle := 2;

      Cell[6, 6].Data := 'Вертикаль';
      Cell[6, 6].CellStyle := 2;

      for i := 5 to 6 do
      for j := 7 to 9 do
        Cell[i, j].Data := 'text';

      for i := 7 to 9 do
      begin
        Cell[5, i].CellStyle := i;
        Cell[6, i].CellStyle := i + 3;
        Rows[i].HeightMM := 14;
      end;

      for i := 1 to 10 do
         Cell[i, 26].AsDouble := 5.0 / i;
    end;

    //копирование с 0 на 1-ую страницу
    XMLSS.Sheets[1].Assign(XMLSS.Sheets[0]);
    XMLSS.Sheets[1].Title := 'Тестовая таблица (копия)';

    //сохранить 0-ую и 1-ую страницу
    //сохранить как excel xml
    //SaveXmlssToEXML(XMLSS, {какой-то путь} Path +'save_test.xml',
    //	[0, 1], [], TextConverter, 'UTF-8');
    //Примеры путей:
    //	/home/user_name/some_path/
    //	d:\some_path\
    //Путь должен существовать!
    //сохранить как незапакованный ods
    //SaveXmlssToODFSPath(XMLSS, Path+'ods_path\',
    //	[0, 1], [], TextConverter, 'UTF-8');
    //{$IFDEF FPC}
    ////Пока только для лазаруса
    ////сохранить как ods
    //SaveXmlssToODFS(XMLSS, Path+'save_test.ods',
    //	[0, 1], [], TextConverter, 'UTF-8');
    //{$ENDIF}
    //
    //SaveXmlssToXLSX(XMLSS, Path+'save_test_laz.xlsx',
    //	[0, 1], [], TextConverter, 'UTF-8');
    //
    //SaveXmlssToXLSXPath(XMLSS, Path+'xlsx_path_laz\',
    //	[0, 1], [], TextConverter, 'UTF-8');
    //
    //ExportXmlssToXLSX( XMLSS, Path+'xlsx_path_generic',
    //	[0, 1], [], TextConverter, 'UTF-8', '', true, QueryDummyZipGen);

    //ExportXmlssToXLSX( XMLSS, Path+'save_test_generic.xlsx',
    //	[0, 1], [], TextConverter, 'UTF-8');

    TZXMLSSave.From(XMLSS).Save(Path+'LazZexSample.xml');
    TZXMLSSave.From(XMLSS).Save(Path+'LazZexSample.xlsx');
    TZXMLSSave.From(XMLSS).Save(Path+'LazZexSample.ods');

    TZXMLSSave.From(XMLSS).Pages([0,1]).NoZip.To_(Path+'LazZexSample.Dir.xml').Save;
    TZXMLSSave.From(XMLSS).Pages([0,1]).NoZip.To_(Path+'LazZexSample.Dir.xlsx').Save;
    TZXMLSSave.From(XMLSS).Pages([0,1]).NoZip.To_(Path+'LazZexSample.Dir.ods').Save;
  finally
    FreeAndNil(XMLSS);
  end;
end;

procedure TForm1.btnFormulaClick(Sender: TObject);
var
  XMLSS: TZEXMLSS;
  i, j: integer;
  TextConverter: TAnsiToCPConverter;
  (*sEOL,*) s: string;

begin
  TextConverter := nil;
  {$IFNDEF FPC}
    {$IF CompilerVersion < 20} // < RAD Studio 2009
  TextConverter := @AnsiToUtf8;
    {$IFEND}
  {$ENDIF}

 (* {$IFDEF FPC}
  sEOL := LineEnding;
  {$ELSE}
  sEOL := sLineBreak;
  {$ENDIF}      *)

  XMLSS := TZEXMLSS.Create(nil);
  try
    //В документе 2 страницы
    XMLSS.Sheets.Count := 2;
    XMLSS.Sheets[0].Title := 'Формулы (ODF)';
    //добавим стили
    XMLSS.Styles.Count := 3;
    //0 - для заголовка (20)
    XMLSS.Styles[0].Font.Size := 20;
    XMLSS.Styles[0].Font.Style := [fsBold];
    XMLSS.Styles[0].Font.Name := 'Tahoma';
    XMLSS.Styles[0].BGColor := $CCFFCC;
    XMLSS.Styles[0].CellPattern := ZPSolid;
    XMLSS.Styles[0].Alignment.Horizontal := ZHCenter;
    XMLSS.Styles[0].Alignment.Vertical := ZVCenter;
    XMLSS.Styles[0].Alignment.WrapText := true;
    //1 - стиль для таблицы
    XMLSS.Styles[1].Border[0].Weight := 1;
    XMLSS.Styles[1].Border[0].LineStyle := ZEContinuous;
    for i := 1 to 3 do
      XMLSS.Styles[1].Border[i].Assign(XMLSS.Styles[1].Border[0]);
    //2 - стиль для заголовка таблицы (жирный по центру)
    XMLSS.Styles[2].Assign(XMLSS.Styles[1]);
    XMLSS.Styles[2].Font.Style := [fsBold];
    XMLSS.Styles[2].Alignment.Horizontal := ZHCenter;

    with XMLSS.Sheets[0] do
    begin
      //количество строк и столбцов
      RowCount := 50;
      ColCount := 20;

      Cell[0, 0].CellStyle := 0;
      Cell[0, 0].Data := 'Пример использования формул в zexmlss';
      MergeCells.AddRectXY(0, 0, 10, 0);

      for  i := 0 to 4 do
      begin
        Cell[i, 1].CellStyle := 2;
        Cell[i, 12].CellStyle := 2;
      end;
      Cell[0, 1].Data := '№';
      for i := 1 to 3 do
         Cell[i, 1].Data := 'Value ' + IntToStr(i);
      Cell[4, 1].Data := 'Formula';
      Cell[0, 12].Data := 'Total';

      for i := 2 to 11 do
      begin
        Cell[0, i].Data := IntToStr(i);
        Cell[0, i].CellStyle := 1;
        for j := 1 to 4 do
        begin
          Cell[j, i].Data := IntToStr(Random(100));
          Cell[j, i].CellStyle := 1;
          Cell[j, i].CellType := ZENumber;
        end;
      end;
    end;

    //копирование с 0 на 1-ую страницу
    XMLSS.Sheets[1].Assign(XMLSS.Sheets[0]);
    XMLSS.Sheets[1].Title := 'Формулы (Excel XML)';

    //формулы для ODF
    with XMLSS.Sheets[0] do
    begin
      for i := 2 to 11 do
      begin
        s := IntToStr(i + 1);
        Cell[4, i].Formula := '=B' + s + '*C' + s + ' - D' + s; //=B*C-D
      end;
      for i := 1 to 4 do
      begin
        s := ZEGetA1byCol(i); //Получить наименование столбца по его номеру
        Cell[i, 12].Formula := '=SUM(' + s +'3:' + s + '12)';
      end;
    end;

    //формулы для Excel XML SpreadSheet
    with XMLSS.Sheets[1] do
    begin
      for i := 2 to 11 do
        Cell[4, i].Formula := '=RC[-3] * RC[-2] - RC[-1]'; //=B*C-D
      for i := 1 to 4 do
        Cell[i, 12].Formula := '=SUM(R2C:R12C)';
    end;

    //сохранить 0-ую и 1-ую страницу
    SaveXmlssToEXML(XMLSS, 'excelxml_formula.xml',
    	[0, 1], [], TextConverter, 'UTF-8');
    SaveXmlssToODFSPath(XMLSS, 'd:\work\ods_path\',
    	[0, 1], [], TextConverter, 'UTF-8');
    {$IFDEF FPC}
    SaveXmlssToODFS(XMLSS, 'ods_formula.ods',
    	[0, 1], [], TextConverter, 'UTF-8');
    {$ENDIF}

  finally
    FreeAndNil(XMLSS);
  end;
end;

end.

