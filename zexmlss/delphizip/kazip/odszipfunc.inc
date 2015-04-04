{
interface: functions for ODS zip (delphi)
usign kazip.pas https://raw.githubusercontent.com/JoseJimeniz/KaZip/master/KaZip.pas

}
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string): integer; overload;
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
                         
function ReadODFS(var XMLSS: TZEXMLSS; FileName: string): integer;