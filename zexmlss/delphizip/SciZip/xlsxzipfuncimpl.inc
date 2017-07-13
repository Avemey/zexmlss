{
implementation: function for XLSX zip (delphi)
using SciZipFile.pas
}

function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
var
  tmpdir: string;
  _isCreated: boolean;
begin
  _isCreated := false;
  try
    _isCreated := ZECreateUniqueTmpDir(ZEGetTempDir(), tmpdir);
    Result := SaveXmlssToXLSXPath(XMLSS, tmpdir, SheetsNumbers, SheetsNames, TextConverter, CodePageName, BOM);

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
end; //SaveXmlssToXLSX

function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string): integer; overload;
begin
  result := SaveXmlssToXLSX(XMLSS, FileName, SheetsNumbers, SheetsNames, ZEGetDefaultUTF8Converter(), 'UTF-8', '');
end; //SaveXmlssToXLSX

function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
begin
  result := SaveXmlssToXLSX(XMLSS, FileName, [], []);
end; //SaveXmlssToXLSX


function ReadXLSX(var XMLSS: TZEXMLSS; FileName: string): integer;
var
  tmpdir: string;
  _isCreated: boolean;
begin
  _isCreated := false;
  try
    _isCreated := ZECreateUniqueTmpDir(ZEGetTempDir(), tmpdir);

    if ZEZipUnpackSciZipFile(FileName, tmpdir) then
     begin
       Result := ReadXLSXPath(XMLSS, tmpdir)
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
end; //ReadXLSX
