{
interface: functions for XLSX zip (delphi)
using TurboPower Abbrevia http://sourceforge.net/projects/tpabbrevia/
}
function SaveXmlssToXLSX(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
function ReadXLSX(var XMLSS: TZEXMLSS; FileName: string): integer;