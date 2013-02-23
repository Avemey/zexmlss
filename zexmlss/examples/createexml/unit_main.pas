//пример использования zexmlss
// создаёт документ с таблицей умножения
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
  zexmlss, zexmlssutils, zsspxml, zeformula, StrUtils;

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

uses
 {$IfDef Unicode} AnsiStrings, {$EndIf} TypInfo,
 {$IFNDEF FPC}
 {$IF CompilerVersion > 22}zeZippyXE2, {$ELSE} zeZippyAB,{у кого чего стоит} {$IFEND} //<XE2 have not zip!
 {$Else} zeZippyLaz,{$EndIf}
zeSave, zeSaveODS, zeSaveXLSX, zeSaveEXML;
{$IFNDEF FPC}
  {$R *.dfm}
{$ENDIF}

//Функция перевода из локальной кодировки в нужную (для примера - в utf8)
function my_local_to_some_encoding(value: ansistring): ansistring;
begin
  result := AnsiToUtf8(value);
end;

//Создать файл
procedure TfrmMain.CreateEXMLSS();
var
  sd: TSaveDialog;
  tz: TZEXMLSS;
  i, j: integer;
  TextConverter: TAnsiToCPConverter;
  va: TZVerticalAlignment;
  ha: TZHorizontalAlignment;
  s: string;
  _ConvertParams: integer;

begin
  TextConverter := nil;
  {$IFNDEF FPC}
    {$IF CompilerVersion < 20} // < RAD Studio 2009
  TextConverter := @my_local_to_some_encoding;
    {$IFEND}
  {$ENDIF}
  sd := TSaveDialog.Create(nil);
  sd.Filter := 'Excel XML 2002 (*.xml)|*.xml|OpenDocument|*.ods;*.fods|Office OpenXML|*.xlsx;*.xlsm';
  sd.DefaultExt := 'xml';
  sd.Options := sd.Options + [ofOverwritePrompt];
  if sd.Execute then
  begin
    tz := TZEXMLSS.Create(nil);
    try
      //В документе 2 страницы
      tz.Sheets.Count := 2;
      tz.Sheets[0].Title := 'Таблица Пифагора';
      //добавим стили
      tz.Styles.Count := 7;
      //0 - для заголовка (20)
      tz.Styles[0].Font.Size := 20;
      tz.Styles[0].Font.Style := [fsBold];
      tz.Styles[0].Font.Name := 'Tahoma';
      tz.Styles[0].BGColor := $CCFFCC;
      tz.Styles[0].CellPattern := ZPSolid;
      tz.Styles[0].Alignment.Horizontal := ZHCenter;
      tz.Styles[0].Alignment.Vertical := ZVCenter;
      tz.Styles[0].Alignment.WrapText := true;
      //1 - стиль для таблицы
      tz.Styles[1].Border[0].Weight := 1;
      tz.Styles[1].Border[0].LineStyle := ZEContinuous;
      for i := 1 to 3 do tz.Styles[1].Border[i].Assign(tz.Styles[1].Border[0]);

      //6 - разные границы в одной таблице
      tz.Styles[6].Assign(tz.Styles[1]);
      for i := 0 to 3 do tz.Styles[6].Border[i].Weight := 2;

      //2 - стиль для заголовка таблицы (жирный по центру)
      tz.Styles[2].Assign(tz.Styles[1]);
      tz.Styles[2].Font.Style := [fsBold];
      tz.Styles[2].Alignment.Horizontal := ZHCenter;
      //3-ий стиль
      tz.Styles[3].Font.Size := 16;
      tz.Styles[3].Font.Name := 'Arial Black';
      //4-ый стиль
      tz.Styles[4].Font.Size := 18;
      tz.Styles[4].Font.Name := 'Arial';
      //5-ый стиль
      tz.Styles[5].Font.Size := 14;
      tz.Styles[5].Font.Name := 'Arial Black';

      with tz.Sheets[0] do
      begin
        //установим количество строк и столбцов
        RowCount := 20 + Ord(High(TZVerticalAlignment )) + 2;
        ColCount := 20;

        Cell[0, 0].CellStyle := 3;
        Cell[0, 0].Data := 'Пример использования zexmlss';

        //ячейка с ссылкой
        Cell[0, 2].CellStyle := 4;
        Cell[0, 2].Data := 'Домашняя страница: http://avemey.com';
        Cell[0, 2].Href := 'http://avemey.com';
        Cell[0, 2].HRefScreenTip := 'Клацните по ссылке'#13#10'(текст всплывающей подсказки к ссылке)';
        MergeCells.AddRectXY(0, 2, 10, 2);

        //Примечание
        Cell[0, 3].CellStyle := 5;
        Cell[0, 3].Data := 'Есть примечание!';
        Cell[0, 3].Comment := 'Текст примечания 1';
        Cell[0, 3].CommentAuthor := 'Кто-то 1 оставил примечание';
        Cell[0, 3].ShowComment := true;

        Cell[16, 3].CellStyle := 5;
        Cell[16, 3].Data := 'Тоже примечание!';
        Cell[16, 3].Comment := 'Текст примечания 2';
        Cell[16, 3].CommentAuthor := 'Кто-то 2 оставил примечание';
        Cell[16, 3].ShowComment := true;
        Cell[16, 3].AlwaysShowComment := true;
        ColWidths[16] := 160; //Ширина столбца

        Cell[1, 1].CellStyle := 0;
        Cell[1, 1].Data := 'Таблица умножения (таблица Пифагора)';
        MergeCells.AddRectXY(1, 1, 16, 1);

        Cell[0, 6].Data := 'Объединённая'#13#10'ячейка!';
        Cell[0, 6].CellStyle := 0;
        MergeCells.AddRectXY(0, 6, 3, 14);


        Cell[5, 5].CellStyle := 2;
        for i := 1 to 10 do
        begin
          Cell[5, 5 + i].CellStyle := 2;
          Cell[5, 5 + i].Data := inttostr(i) ;
          Cell[5, 5 + i].CellType := ZENumber;
          Cell[5 + i, 5].Assign(Cell[5, 5 + i]);
          for j := 1 to 10 do
            with Cell[5 + i, 5 + j] do
          begin
            if i = j
               then CellStyle := 6
               else CellStyle := 1;
            AsInteger := (i * j);
          end;
        end;

        SheetOptions.PaperSize := 8; //A3
        SheetOptions.PortraitOrientation := false;

        i := tz.Styles.Count; j := i;
          Inc(i, (Ord(High(TZVerticalAlignment))+1)*(Ord(High(TZHorizontalAlignment))+1) );
          tz.Styles.Count := i;
        i := j;
        for va := Low(TZVerticalAlignment) to High(TZVerticalAlignment) do
          for ha := Low(TZHorizontalAlignment) to High(TZHorizontalAlignment) do
            begin
              with tz.Styles[i].Alignment do begin
                   Vertical := va;
                   Horizontal := ha;
              end;
              Inc(i);
            end;

        i := 19;
        for va := Low(TZVerticalAlignment) to High(TZVerticalAlignment) do begin
          Rows[i+Ord(va)].HeightMM := 20;
          for ha := Low(TZHorizontalAlignment) to High(TZHorizontalAlignment) do
            begin
              with Cell[i+Ord(ha)-14, i+Ord(va)] do begin
                   Data := 'xxx';
                   CellStyle := j;
              end;
              Inc(j);
            end;
        end;

        tz.Styles.Count := j + 2;
        with tz.Styles[j] do begin
           Alignment.Horizontal := ZHRight;
           Alignment.Rotate :=
                     // 360 - 45; // in OpenOffice.org
                      -45;       // in Excel 2010
           with Border.Right do begin
               LineStyle := ZEDot;
               Weight := 1;
           end;
        end;
        for va := Low(TZVerticalAlignment) to High(TZVerticalAlignment) do
          with Cell [i-1-14, i+Ord(va)] do
           begin
               CellStyle := j;
               Data := GetEnumName(TypeInfo(TZVerticalAlignment), Ord(va))
           end;

        Inc(j);
        with tz.Styles[j] do begin
           Alignment.Vertical := ZVTop;
           Alignment.VerticalText := True;
           with Border.Top do begin
               LineStyle := ZEDot;
               Weight := 1;
           end;
        end;
        for ha := Low(TZHorizontalAlignment) to High(TZHorizontalAlignment) do
          with Cell[i+Ord(ha)-14, i+Ord(High(va))+1] do begin
               CellStyle := j;
               Data := GetEnumName(TypeInfo(TZHorizontalAlignment), Ord(ha))
          end;
      end;

     //копируем данные с 0 на 1-ую страницу
     tz.Sheets[1].Assign(tz.Sheets[0]);
     tz.Sheets[1].Title := 'Таблица Пифагора (формулы)';

     //На второй странице будем использовать формулы.
     //Используем формулы вида R1C1 (они проще всего переводятся в другие)
     s := '=R6C*RC6'; // координаты в табличке на 1 больше

     if (AnsiEndsText('.xml', sd.FileName)) then
      with tz.Sheets[1] do
       for i := 1 to 10 do
       for j := 1 to 10 do
        Cell[5 + i, 5 + j].Formula := s
     else
     begin
       _ConvertParams := 0;
       if (AnsiEndsText('.ods', sd.FileName)) then
         _ConvertParams := ZE_RTA_ODF or ZE_RTA_ODF_PREFIX;

        with tz.Sheets[1] do
         for i := 1 to 10 do
         for j := 1 to 10 do
          Cell[5 + i, 5 + j].Formula := ZER1C1ToA1(s, 5 + i, 5 + j, _ConvertParams);
     end;

      //сохраняем 0-ую и 1-ую страницу в файл
      //кодировка - utf8, имя кодировки='utf8' (для utf8 можно ''), BOM=''
//      SaveXmlssToEXML(tz, sd.FileName, [0, 1], [], @TextConverter, 'utf8');

      TZXMLSSave.From(tz).Charset('utf-8', TextConverter).Save(sd.FileName);
      // Page 1 - formulae - only XML SS format

      // formulae would fail in XLSX format and Excel would complain on "corrupt worksheet"
      // may be worked around byt only saving 0th page: .Pages([0])

      //       TZXMLSSave.From(tz).Save(sd.FileName);
    finally
      tz.free();
    end;
  end;
  sd.Free();
end;

procedure TfrmMain.btnCreateClick(Sender: TObject);
begin
  btnCreate.Hide;
  try
    CreateEXMLSS();
  finally
    btnCreate.Show;
  end;
end;

{$IFDEF FPC}
initialization
  {$i unit_main.lrs}
{$ENDIF}

end.

