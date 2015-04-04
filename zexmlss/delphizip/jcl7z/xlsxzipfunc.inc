{
interface: functions for XLSX zip (delphi)
JEDI Code library (via www.7-zip.org) : http://sourceforge.net/projects/jcl/
}
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string): integer; overload;
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string): integer; overload;

function ReadXLSX(var XMLSS: TZEXMLSS; FileName: string): integer;