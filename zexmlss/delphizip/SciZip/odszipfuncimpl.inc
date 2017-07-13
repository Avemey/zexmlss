{
implementation: function for ODS zip (delphi)
using SciZipFile.pas
}

function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
var
  tmpdir: string;
  _isCreated: boolean;
begin
  _isCreated := false;
  try
    _isCreated := ZECreateUniqueTmpDir(ZEGetTempDir(), tmpdir);
    Result := SaveXmlssToODFSPath(XMLSS, tmpdir, SheetsNumbers, SheetsNames, TextConverter, CodePageName, BOM);

    if not ZEZipPackSciZipFile(FileName, tmpdir) then
     begin
       Result := -1;
     end;
  finally
    if _isCreated then
     begin
       ZEDelTree(tmpdir);
     end;
  end;
end; //SaveXmlssToODFS

function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string): integer; overload;
begin
  result := SaveXmlssToODFS(XMLSS, FileName, SheetsNumbers, SheetsNames, ZEGetDefaultUTF8Converter(), 'UTF-8', '');
end; //SaveXmlssToODFS

//Сохраняет документ в формате Open Document
//INPUT
//  var XMLSS: TZEXMLSS                   - хранилище
//      FileName: string                  - имя файла для сохранения
//RETURN
//      integer
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
begin
  result := SaveXmlssToODFS(XMLSS, FileName, [], []);
end; //SaveXmlssToODFS

function ReadODFS(var XMLSS: TZEXMLSS; FileName: string): integer;
var
  tmpdir: string;
  _isCreated: boolean;
begin
  _isCreated := false;
  try
    _isCreated := ZECreateUniqueTmpDir(ZEGetTempDir(), tmpdir);

    if ZEZipUnpackSciZipFile(FileName, tmpdir) then
     begin
       Result := ReadODFSPath(XMLSS, tmpdir)
     end
    else
     begin
       Result := -1;
     end;
  finally
    if _isCreated then
     begin
       ZEDelTree(tmpdir);
     end;
  end;
end; //ReadODFS
