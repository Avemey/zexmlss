{
implementation: function for XLSX zip (delphi)
using TurboPower Abbrevia http://sourceforge.net/projects/tpabbrevia/
}

//Сохраняет незапакованный документ в формате Office Open XML (OOXML)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//      TextConverter: TAnsiToCPConverter - конвертер
//      CodePageName: string              - имя кодировки
//      BOM: ansistring                   - Byte Order Mark
//RETURN
//      integer
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
var
  tmpdir: string;
  b: boolean;
  _isCreated: boolean;

begin
  _isCreated := false;
  try
    _isCreated := ZECreateUniqueTmpDir(ZEGetTempDir(), tmpdir);
    result := SaveXmlssToXLSXPath(XMLSS, tmpdir, SheetsNumbers, SheetsNames, TextConverter, CodePageName, BOM);
    b := ZEZipPackAb(FileName, tmpdir);
    if (not b) then
      result := -1;
  finally
    if (_isCreated) then
      ZEDelTree(tmpdir);
  end;
end; //SaveXmlssToODFS

//Сохраняет документ в формате Open Office XML (xlsx)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      FileName: string                  - имя файла для сохранения
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//RETURN
//      integer
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string): integer; overload;
begin
  result := SaveXmlssToXLSX(XMLSS, FileName, SheetsNumbers, SheetsNames, ZEGetDefaultUTF8Converter(), 'UTF-8', '');
end; //SaveXmlssToXLSX

//Сохраняет документ в формате Open Office XML (xlsx)
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      FileName: string                  - имя файла для сохранения
//RETURN
//      integer
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
begin
  result := SaveXmlssToXLSX(XMLSS, FileName, [], []);
end; //SaveXmlssToXLSX

//Читает xlsx
//INPUT
//  var XMLSS: TZEXMLSS - хранилище
//  FileName: string    - имя файла
//RETURN
//      integer - номер ошибки (0 - всё OK)
function ReadXLSX(var XMLSS: TZEXMLSS; FileName: string): integer;
var
  tmpdir: string;
  b: boolean;
  _isCreated: boolean;

begin
  result := 0;
  _isCreated := false;
  try
    _isCreated := ZECreateUniqueTmpDir(ZEGetTempDir(), tmpdir);
    b := ZEZipUnpackAb(FileName, tmpdir);
    if (b) then
      result := ReadXLSXPath(XMLSS, tmpdir)
    else
      result := -1;
  finally
    if (_isCreated) then
      ZEDelTree(tmpdir);
  end;
end; //ReadODFS
