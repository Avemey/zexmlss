(* Simplistic interface for uniform workbook saving
   Bridge object template for different avemey.com
      export routines.

   (c) the Arioch, licensed under zLib license *)
unit zeSaveEXML;

interface

implementation
uses zeSave, zexmlss, Types, zexmlssutils;

type TZxXMLSSSaver = class(TZXMLSSave)
     protected
        function DoSave: integer; override;
        class function FormatDescriptions: TStringDynArray; override;
end;

class function TZxXMLSSSaver.FormatDescriptions: TStringDynArray;
begin
  Result := SplitString(
   '.XML*SpreadsheetML*Excel XML*XMLSS*XML SS*Microsoft Office 2003 XML*Office 2003 XML*'+
   'application/vnd.ms-excel*application/xml', '*');
// http://support.microsoft.com/kb/288130
end;

function TZxXMLSSSaver.DoSave: integer;
begin
  Result := SaveXmlssToEXML( fBook, FFile, GetPageNumbers, GetPageTitles, fConv, fCharSet, fBOM);
end;

initialization
    TZxXMLSSSaver.Register;
//    TZXMLSSave.RegisterFormat(TZxXMLSSSaver);
finalization
    TZxXMLSSSaver.UnRegister;
//    TZXMLSSave.UnRegisterFormat(TZxXMLSSSaver);
end.
