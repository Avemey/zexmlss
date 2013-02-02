(* Simplistic interface for uniform workbook saving
   Bridge object template for different avemey.com
      export routines.

   (c) the Arioch, licensed under zLib license *)
unit zeSaveXLSX;

interface

implementation
uses zeSave, zexmlss, zexlsx, Types,
{$IfDef Delphi_Unicode}
  StrUtils;
{$Else}       // SplitString
  zexmlssutils;
{$EndIf}

type TZxXlsxSaver = class(TZXMLSSave)
     protected
        function DoSave: integer; override;
        class function FormatDescriptions: TStringDynArray; override;
end;

{ TZxXlsxSaver }

class function TZxXlsxSaver.FormatDescriptions: TStringDynArray;
begin
  Result := SplitString(
   '.XLSX*.XLSM*application/vnd.openxmlformats-officedocument.spreadsheetml.sheet*'+
   'Excel 2007*Excel 2010*Office 2007*Office 2010*'+
   'Office Open XML*OOXML*OpenXML*ECMA-376*ISO/IEC 29500*ISO 29500', '*');
end;

function TZxXlsxSaver.DoSave: integer;
begin
  Result := ExportXmlssToXLSX(
    fBook, FFile, GetPageNumbers, GetPageTitles, fConv, fCharSet, fBOM,
    false, FZipGen);
end;

initialization
    TZxXlsxSaver.Register;
//    TZXMLSSave.RegisterFormat(TZxXlsxSaver);
finalization
    TZxXlsxSaver.UnRegister;
//    TZXMLSSave.UnRegisterFormat(TZxXlsxSaver);
end.
