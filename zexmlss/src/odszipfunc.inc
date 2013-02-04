{
interface: functions for ODS zip (delphi)
usign none.

}

{
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
function ReadODFS(var XMLSS: TZEXMLSS; FileName: string): integer;
}