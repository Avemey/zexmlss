(* Simplistic interface for uniform workbook saving
   Bridge object template for different avemey.com
      export routines.

   (c) the Arioch, licensed under zLib license *)
unit zeSave;

interface uses zexmlss,  zsspxml, zeZippy, sysutils, classes, Types;

type TZxPageInfo = record name: string; no: integer; end;

type IZXMLSSave = interface
        function ExportFormat(const fmt: string): iZXMLSSave;
        function As_(const fmt: string): iZXMLSSave;

        function ExportTo(const fname: TFileName): iZXMLSSave;
        function To_(const fname: TFileName): iZXMLSSave;

        function Pages(const pages: array of TZxPageInfo): iZXMLSSave; overload;
        function Pages(const numbers: array of integer): iZXMLSSave; overload;
        function Pages(const titles: array of string): iZXMLSSave; overload;

        function CharSet(const cs: AnsiString): iZXMLSSave; overload;
        function CharSet(const converter: TAnsiToCPConverter): iZXMLSSave; overload;
        function CharSet(const cs: AnsiString; const converter: TAnsiToCPConverter): iZXMLSSave; overload;
        function CharSet(const codepage: word): iZXMLSSave; overload;

        function BOM(const Unicode_BOM: AnsiString): iZXMLSSave; // better rawbytestring ?

        function ZipWith(const ZipGenerator: CZxZipGens): iZXMLSSave;
        function NoZip: iZXMLSSave;   // save to folder

        /// returns zero on success, according to original
        ///     description for SaveXmlssToEXML
        function Save: integer;
        function InternalSave: integer; // ugly, but implementation-specific typcast would be ugly too
end;

type TZXMLSSave = class; CZXMLSSaveClass = class of TZXMLSSave;

     TZXMLSSave = class (tInterfacedObject, IzXMLSSave)
     public
        constructor Create (const zxbook: TZEXMLSS); overload;
        constructor Create (const zxsaver: TZXMLSSave); overload; virtual;

        function ExportFormat(const fmt: string): iZXMLSSave;
        function As_(const fmt: string): iZXMLSSave; inline;

        function ExportTo(const fname: TFileName): iZXMLSSave;
        function To_(const fname: TFileName): iZXMLSSave; inline;

        function Pages(const APages: array of TZxPageInfo): iZXMLSSave; overload;
        function Pages(const numbers: array of integer): iZXMLSSave; overload;
        function Pages(const titles: array of string): iZXMLSSave; overload;

        function CharSet(const cs: AnsiString): iZXMLSSave; overload;
        function CharSet(const converter: TAnsiToCPConverter): iZXMLSSave; overload;
        function CharSet(const cs: AnsiString; const converter: TAnsiToCPConverter): iZXMLSSave; overload;
        function CharSet(const codepage: word): iZXMLSSave; overload;

        function BOM(const Unicode_BOM: AnsiString): iZXMLSSave; // better rawbytestring ?

        function ZipWith(const ZipGenerator: CZxZipGens): iZXMLSSave;
        function NoZip: iZXMLSSave;  // save to folder

        /// returns zero on success, according to original
        ///     description for SaveXmlssToEXML
        function Save: integer;

     protected
        fBook: TZEXMLSS;
        fPages: array of TZxPageInfo;

        fBOM: ansistring;
        fCharSet: AnsiString;
        fConv: TAnsiToCPConverter;

        FFile, FPath: TFileName;

        FZipGen: CZxZipGens;

        function GetPageNumbers: TIntegerDynArray;
        function GetPageTitles:  TStringDynArray;
        function CreateSaverForDescription(const desc: string): IZXMLSSave;

        /// returns zero on success, according to original
        ///     description for SaveXmlssToEXML
        ///
        /// tries to guess format by filename in the base class
        function InternalSave: integer; virtual;
        class function FormatDescriptions: TStringDynArray; virtual;
        class procedure RegisterFormat(const sv: CZXMLSSaveClass);
        class procedure UnRegisterFormat(const sv: CZXMLSSaveClass);
     public
        class procedure Register;
        class procedure UnRegister;
     end;

     EZXSaveException = class (Exception);

implementation
uses
{$IfDef MSWINDOWS}Registry, Windows, {$EndIf}
  Contnrs {$IfNDef FPC}, AnsiStrings{$EndIf};

var SaveClasses: TClassList;


/// todo - add implementation for non-Windows platforms
/// if anyone would need it :-)
/// probably via http://sourceforge.net/projects/natspec/
function CharSetByCodePage(const cp: Word): AnsiString;
{$IfDef MSWINDOWS }
var reg: TRegistry;
begin
// HKEY_CLASSES_ROOT\MIME\DataBase\Codepage
  reg := TRegistry.Create(KEY_READ);
  try
    reg.RootKey := HKEY_CLASSES_ROOT;
    if reg.OpenKeyReadOnly('MIME\DataBase\Codepage\' + IntToStr(cp)) then begin
       Result := Trim(AnsiString(reg.ReadString('WebCharset'))); // This key prevails, see #1251 for example
       if Result = '' then
          Result := Trim(AnsiString(reg.ReadString('BodyCharset')));
       if Result > '' then exit;
    end;
  finally
    reg.Free;
  end;
  Raise EZXSaveException.Create('No charset (MIME id) found for codepage '+IntToStr(cp));
end;
{$Else}
begin
  Raise EZXSaveException.Create('Cannot get charset by numeric codepage on this platform.');
  //not implemented, perhaps http://sourceforge.net/projects/natspec/ ?
end;
{$EndIf}

{ TZXMLSSave }

function TZXMLSSave.BOM(const Unicode_BOM: AnsiString): iZXMLSSave;
begin
   fBOM := Unicode_BOM;
   Result := self;
end;

function TZXMLSSave.CharSet(const cs: AnsiString;
  const converter: TAnsiToCPConverter): iZXMLSSave;
begin
  Result := CharSet(cs);
  fConv := converter;
end;

function TZXMLSSave.CharSet(const codepage: word): iZXMLSSave;
begin
  fCharSet := CharSetByCodePage(codepage);
  Result := Self;
end;

function TZXMLSSave.CharSet(const cs: AnsiString): iZXMLSSave;
begin
  fCharSet := cs; // check that encoding is real ???
  Result := self;
end;

function TZXMLSSave.CharSet(const converter: TAnsiToCPConverter): iZXMLSSave;
begin
  fConv := converter;
  Result := Self;
end;

constructor TZXMLSSave.Create(const zxbook: TZEXMLSS);
begin
  if zxbook = nil then raise EZXSaveException.Create ('Cannot export nil book');
  fBook := zxbook;
  fCharSet := 'UTF-8';
end;

constructor TZXMLSSave.Create(const zxsaver: TZXMLSSave);
begin
  Self.fBook    := zxsaver.fBook;
  Self.fBOM     := zxsaver.fBOM;
  Self.fCharSet := zxsaver.fCharSet;
  Self.fConv    := zxsaver.fConv;
  Self.FFile    := zxsaver.FFile;
  Self.FPath    := zxsaver.FPath;
  Self.fPages   := zxsaver.fPages;
  Self.FZipGen  := zxsaver.FZipGen;
end;

function TZXMLSSave.CreateSaverForDescription(const desc: String): IZXMLSSave;
var tgt: String; ss: TStringDynArray;  //s: string;
    cs: CZXMLSSaveClass; cc: TClass;
    i, j: integer;
begin
   tgt := UpperCase(Trim(desc)); // target

   for i := 0 to SaveClasses.Count - 1 do begin
       cc := SaveClasses[i];
       if not cc.InheritsFrom( TZXMLSSave )
          then continue;
       cs := CZXMLSSaveClass(cc);
//       for s in cs.FormatDescriptions do begin
//           if UpperCase(Trim(AnsiString(s))) = tgt then begin
       ss := cs.FormatDescriptions;
       for j := Low(ss) to High(ss) do begin
           if UpperCase(Trim(ss[j])) = tgt then begin
              Result := cs.Create(self);
              exit;
           end;
       end;
   end;

   raise EZXSaveException.Create('Doesn''t know how to save the workbook as '+desc);
end;

function TZXMLSSave.ExportFormat(const fmt: string): iZXMLSSave;
begin
  Result := CreateSaverForDescription(fmt);
end;

function TZXMLSSave.As_(const fmt: string): iZXMLSSave;
begin
  Result := ExportFormat(fmt);
end;

function TZXMLSSave.To_(const fname: TFileName): iZXMLSSave;
begin
  Result := ExportTo(fname);
end;

function TZXMLSSave.ExportTo(const fname: TFileName): iZXMLSSave;
var fp: TFileName;
begin
   fp := ExtractFileDir(fname);

   If not DirectoryExists(fp) then raise EZXSaveException.Create ('No such path: '+ fp);

   FPath := fp;
   FFile := fname;

   Result := Self;
end;

class function TZXMLSSave.FormatDescriptions: TStringDynArray;
begin
  Result := nil;  // make compiler happy
end;

function TZXMLSSave.GetPageNumbers: TIntegerDynArray; var i: integer;
begin
  SetLength(Result, Length(fPages));
  for i := 0 to High(Result) do
      Result[i] := fPages[i].no;
end;

function TZXMLSSave.GetPageTitles: TStringDynArray; var i: integer;
begin
  SetLength(Result, Length(fPages));
  for i := 0 to High(Result) do
      Result[i] := fPages[i].name;
end;

function TZXMLSSave.Pages(const APages: array of TZxPageInfo): iZXMLSSave;
var i, c: integer;
begin
  c := fBook.Sheets.Count - 1;

  for i := Low(APages) to High(APages) do
    if (APages[i].no < 0) or (APages[i].no > c) then
       raise EZXSaveException.Create ('There is no sheet #'+IntToStr(APages[i].no) +' in the book');

  SetLength(fPages, Length(APages));
  for i := 0 to High(APages) do
      fPages[i] := APages[i];

  Result := Self;
end;

function TZXMLSSave.Pages(const numbers: array of integer): iZXMLSSave;
var i, c: integer;
begin
  c := fBook.Sheets.Count - 1;

  for i in numbers do
    if (i < 0) or (i > c) then
       raise EZXSaveException.Create ('There is no sheet #'+IntToStr(i) +' in the book');

  SetLength(fPages, Length(numbers));
  for i := 0 to High(numbers) do with fPages[i] do begin
      no := numbers[i];
      name := '';
  end;

  Result := Self;
end;

function TZXMLSSave.Pages(const titles: array of string): iZXMLSSave;
var i: integer;
begin
  if Length(fPages) <> Length(titles) then
     raise EZXSaveException.Create('Sheets titles do not match sheet numbers');

  for i := 0 to high(titles) do
      fPages[i].name := titles[i];

  Result := self;
end;

function TZXMLSSave.Save:integer;
var i: integer;
begin
  if FZipGen = nil then FZipGen := TZxZipGen.QueryZipGen;
  if fCharSet = '' then fCharSet := 'UTF-8';
//todo - default Ansi-converter (fConv)?

  if Length(fPages) = 0 then begin
     SetLength(fPages, fBook.Sheets.Count);
     for i := 0 to  fBook.Sheets.Count - 1 do begin
        fPages[i].no := i;
        fPages[i].name := fBook.Sheets[i].Title;
     end;
  end;

  Result := InternalSave;
end;

function TZXMLSSave.InternalSave:integer;
begin
  Result := CreateSaverForDescription(ExtractFileExt(FFile)).InternalSave;
end;


class procedure TZXMLSSave.UnRegister;
begin
  UnRegisterFormat(Self);
end;

class procedure TZXMLSSave.UnRegisterFormat(const sv: CZXMLSSaveClass);
begin
  SaveClasses.Remove(sv);
end;

class procedure TZXMLSSave.Register;
begin
  RegisterFormat(Self);
end;

class procedure TZXMLSSave.RegisterFormat(const sv: CZXMLSSaveClass);
begin
  SaveClasses.Add(sv);
end;

function TZXMLSSave.NoZip: iZXMLSSave;
begin
  FZipGen := TZxZipGen.QueryDummyZipGen;
  Result := Self;
end;

function TZXMLSSave.ZipWith(const ZipGenerator: CZxZipGens): iZXMLSSave;
begin
  FZipGen := ZipGenerator;
  Result := Self;
end;

initialization
  SaveClasses := TClassList.Create;

// Ersatz-testing below
//  if '' = CharSetByCodePage(866) then ;
//  if '' = CharSetByCodePage(1251) then ;
//  if '' = CharSetByCodePage(20866) then ;
finalization
  SaveClasses.Free;
end.
