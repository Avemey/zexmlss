{
interface: functions for ODS zip (delphi)
using JEDI Code library (via www.7-zip.org) : http://sourceforge.net/projects/jcl/
}
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string; const SheetsNumbers:array of integer;
                         const SheetsNames: array of string): integer; overload;
function SaveXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string): integer; overload;
                         
function ReadODFS(var XMLSS: TZEXMLSS; FileName: string): integer;