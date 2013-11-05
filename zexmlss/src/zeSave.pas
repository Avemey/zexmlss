(* Simplistic interface for uniform workbook saving
   Bridge object template for different avemey.com
      export routines.

   (c) the Arioch, licensed under zLib license *)
unit zeSave;

interface uses zexmlss,  zsspxml, zeZippy, sysutils, classes, Types;

type TZxPageInfo = record name: string; no: integer; end;

type IZXMLSSave = interface
        function ExportFormat(const fmt: string): IZXMLSSave;
        function As_(const fmt: string): IZXMLSSave;

        function ExportTo(const fname: TFileName): IZXMLSSave;
        function To_(const fname: TFileName): IZXMLSSave;

        function Pages(const pages: array of TZxPageInfo): IZXMLSSave; overload;
        function Pages(const numbers: array of integer): IZXMLSSave; overload;
        function Pages(const titles: array of string): IZXMLSSave; overload;

        function CharSet(const cs: AnsiString): IZXMLSSave; overload;
        function CharSet(const converter: TAnsiToCPConverter): IZXMLSSave; overload;
        function CharSet(const cs: AnsiString; const converter: TAnsiToCPConverter): IZXMLSSave; overload;
        function CharSet(const codepage: word): IZXMLSSave; overload;

        function BOM(const Unicode_BOM: AnsiString): IZXMLSSave; // better rawbytestring ?

        function ZipWith(const ZipGenerator: CZxZipGens): IZXMLSSave;
        function NoZip: IZXMLSSave;   // save to folder

        /// returns zero on success, according to original
        ///     description for SaveXmlssToEXML
        function Save: integer;  overload;
        function Save(Const FileName: TFileName): integer;  overload;

        /// Marks instance as intentionally discarded and cleared to be freed.
        ///  Otherwise you should only free it via .Save call/
        procedure Discard;
     end;

     // Pity it is public, but at least in user-targeted interface API
     IZXMLSSaveImpl = interface(IZXMLSSave) ['{DAB9A318-ADD4-466C-AFE3-CE8623EC8977}']
        function InternalSave: integer;
     end;



type TZXMLSSave = class; CZXMLSSaveClass = class of TZXMLSSave;

     { TZXMLSSave }

     TZXMLSSave = class (TInterfacedObject, IZXMLSSave, IZXMLSSaveImpl)
     protected
        (* two functions below are implementors interface and should be
           mandatory overrode and re-implemented *)

        /// returns zero on success, according to original
        ///     description for SaveXmlssToEXML
        ///
        /// tries to guess format by filename in the base class
        function DoSave: integer; virtual;
        class function FormatDescriptions: TStringDynArray; virtual;

     public
        procedure AfterConstruction; override;
        procedure BeforeDestruction; override;

        /// Factory function. I wish it could be default class property.
        ///   Ancestors may chose to override it to return their instances.
        class function From(const zxbook: TZEXMLSS): IZXMLSSave; virtual;
     protected
        constructor Create (const zxbook: TZEXMLSS); overload;
        constructor Create (const zxsaver: TZXMLSSave); overload; virtual;

        function ExportFormat(const fmt: string): IZXMLSSave;
        function As_(const fmt: string): IZXMLSSave; //inline;

        function ExportTo(const fname: TFileName): IZXMLSSave;
        function To_(const fname: TFileName): IZXMLSSave; //inline;

        function Pages(const APages: array of TZxPageInfo): IZXMLSSave; overload;
        function Pages(const numbers: array of integer): IZXMLSSave; overload;
        function Pages(const titles: array of string): IZXMLSSave; overload;

        function CharSet(const cs: AnsiString): IZXMLSSave; overload;
        function CharSet(const converter: TAnsiToCPConverter): IZXMLSSave; overload;
        function CharSet(const cs: AnsiString; const converter: TAnsiToCPConverter): IZXMLSSave; overload;
        function CharSet(const codepage: word): IZXMLSSave; overload;

        function BOM(const Unicode_BOM: AnsiString): IZXMLSSave; // better rawbytestring ?

        function ZipWith(const ZipGenerator: CZxZipGens): IZXMLSSave;
        function NoZip: IZXMLSSave;  // save to folder

        function OnErrorRaise(): IZXMLSSave;
        function OnErrorRetCode(): IZXMLSSave;

        /// returns zero on success, according to original
        ///     description for SaveXmlssToEXML
        function Save: integer;  overload;
        function Save(Const FileName: TFileName): integer;  overload;
        procedure Discard; virtual;
        function InternalSave: integer;
     protected
        fBook: TZEXMLSS;
        fPages: array of TZxPageInfo;

        fBOM: ansistring;
        fCharSet: AnsiString;
        fConv: TAnsiToCPConverter;

        FFile, FPath: TFileName;

        FZipGen: CZxZipGens;

        FDoNotDestroyMe: Boolean; // guard in case the user loses reference unexpectedly
        FRaiseOnError: Boolean;

        function GetPageNumbers: TIntegerDynArray;
        function GetPageTitles:  TStringDynArray;
        function CreateSaverForDescription(const desc: string): IZXMLSSave;
        procedure CheckSaveRetCode (Result: integer);

        {$IfOpt D+}
        function _AddRef: Integer; stdcall;
        function _Release: Integer; stdcall;
        {$EndIf}

     protected
        class procedure RegisterFormat(const sv: CZXMLSSaveClass);
        class procedure UnRegisterFormat(const sv: CZXMLSSaveClass);
     public
        class procedure Register;
        class procedure UnRegister;
     end;

     EZXSaveException = class (Exception);

{$IfNDef MSWINDOWS}
var ZxCharSetByCodePage: function(const cp: Word): AnsiString;
{$EndIf}

implementation
uses
{$IfDef MSWINDOWS}Registry, Windows, {$EndIf}
  Contnrs {$IfNDef FPC}{$IfDef Delphi_Unicode}, AnsiStrings{$EndIf}{$EndIf};

var SaveClasses: TClassList;


/// todo - add implementation for non-Windows platforms
/// if anyone would need it :-)
/// probably via http://sourceforge.net/projects/natspec/
{$IfDef MSWINDOWS }
function ZxCharSetByCodePage(const cp: Word): AnsiString;
var reg: TRegistry;
begin
// HKEY_CLASSES_ROOT\MIME\DataBase\Codepage
  reg := TRegistry.Create(KEY_READ);
  try
    reg.RootKey := HKEY_CLASSES_ROOT;
    if reg.OpenKeyReadOnly('MIME\DataBase\Codepage\' + IntToStr(cp)) then begin
       Result := AnsiString(Trim(reg.ReadString('WebCharset'))); // This key prevails, see #1251 for example
       if Result = '' then
          Result := AnsiString(Trim(reg.ReadString('BodyCharset')));
       if Result > '' then exit;
    end;
  finally
    reg.Free;
  end;
  Raise EZXSaveException.Create('No charset (MIME id) found for codepage '+IntToStr(cp));
end;
{$Else}
function ZxCharSetByOops(const cp: Word): AnsiString;
begin
  Raise EZXSaveException.Create('Cannot get charset by numeric codepage on this platform.'#13#10 + 'Make wrapper for something like http://sf.net/projects/natspec/');
end;

//Recomendation: DO NOT USE zeSave on FPC at all!!!
function ZxCharSetByCodePageSimple(const cp: Word): AnsiString;
begin
  case cp of
    852: result := 'ibm852';
    866: result := 'cp866';
    1250: result := 'windows-1250';
    1251: result := 'windows-1251';
    1252: result := 'iso-8859-1';
    20866: result := 'koi8-r';
    21866: result := 'koi8-ru';
    65000: result := 'utf-7';
    65001: result := 'UTF-8';
    else
      result := 'UTF-8';
  end;
end;
{$EndIf}

{ TZXMLSSave }

class function TZXMLSSave.From(const zxbook: TZEXMLSS): IZXMLSSave;
begin
  Result := TZXMLSSave.Create(zxbook);
end;

function TZXMLSSave.BOM(const Unicode_BOM: AnsiString): IZXMLSSave;
begin
   fBOM := Unicode_BOM;
   Result := self;
end;

function TZXMLSSave.CharSet(const cs: AnsiString;
  const converter: TAnsiToCPConverter): IZXMLSSave;
begin
  Result := CharSet(cs);
  fConv := converter;
end;

function TZXMLSSave.CharSet(const codepage: word): IZXMLSSave;
begin
  fCharSet := ZxCharSetByCodePage(codepage);
  Result := Self;
end;

function TZXMLSSave.CharSet(const cs: AnsiString): IZXMLSSave;
begin
  fCharSet := cs; // check that encoding is real ???
  Result := self;
end;

function TZXMLSSave.CharSet(const converter: TAnsiToCPConverter): IZXMLSSave;
begin
  fConv := converter;
  Result := Self;
end;

procedure TZXMLSSave.AfterConstruction;
begin
  if fBook = nil then raise EZXSaveException.Create ('Cannot export nil book. Do not use inherited TObject.Create');
  inherited;
  FDoNotDestroyMe := True;
end;

procedure TZXMLSSave.Discard;
begin
  FDoNotDestroyMe := false;
end;

procedure TZXMLSSave.BeforeDestruction;
begin
  inherited;
  if FDoNotDestroyMe then begin
     FDoNotDestroyMe := false; // breaking infinite loop
     raise EzXSaveException.Create('Premature exporter destroying: worksheet export process should end either with .Save or with .Discard.');
  end;
end;

constructor TZXMLSSave.Create(const zxbook: TZEXMLSS);
begin
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
  Self.FRaiseOnError := zxsaver.FRaiseOnError;
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

function TZXMLSSave.ExportFormat(const fmt: string): IZXMLSSave;
begin
  Result := CreateSaverForDescription(fmt);
end;


function TZXMLSSave.As_(const fmt: string): IZXMLSSave;
begin
  Result := ExportFormat(fmt);
end;

function TZXMLSSave.To_(const fname: TFileName): IZXMLSSave;
begin
  Result := ExportTo(fname);
end;

function TZXMLSSave.ExportTo(const fname: TFileName): IZXMLSSave;
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

function TZXMLSSave.Pages(const APages: array of TZxPageInfo): IZXMLSSave;
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

function TZXMLSSave.Pages(const numbers: array of integer): IZXMLSSave;
var i, j, c: integer;
begin
  c := fBook.Sheets.Count - 1;

//  for i in numbers do  //  Screw Delphi 7 !
  for j := 0 to High(numbers) do begin
    i := numbers[j];
    if (i < 0) or (i > c) then
       raise EZXSaveException.Create ('There is no sheet #'+IntToStr(i) +' in the workbook');
  end;

  SetLength(fPages, Length(numbers));
  for i := 0 to High(numbers) do with fPages[i] do begin
      no := numbers[i];
      name := '';
  end;

  Result := Self;
end;

function TZXMLSSave.Pages(const titles: array of string): IZXMLSSave;
var i: integer;
begin
  if Length(fPages) <> Length(titles) then
     raise EZXSaveException.Create('Sheets titles do not match sheet numbers');

  for i := 0 to high(titles) do
      fPages[i].name := titles[i];

  Result := self;
end;

function TZXMLSSave.OnErrorRaise: IZXMLSSave;
begin
  Result := Self;
  Self.FRaiseOnError := True;
end;

function TZXMLSSave.OnErrorRetCode: IZXMLSSave;
begin
  Result := Self;
  Self.FRaiseOnError := False;
end;

procedure TZXMLSSave.CheckSaveRetCode(Result: integer);
begin
  if Result <> 0 then
     if FRaiseOnError then
        raise EZXSaveException.Create('Error #'+IntToStr(Result)+': cannot save ' + Self.FFile);
end;

function TZXMLSSave.Save(const FileName: TFileName): integer;
begin
  Result := Self.ExportTo( FileName ).Save();
  CheckSaveRetCode(Result);
end;

function TZXMLSSave.Save: integer;
var i: integer;
begin
  FDoNotDestroyMe := false;
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
  CheckSaveRetCode(Result);
end;

function TZXMLSSave.InternalSave: integer;
begin
  FDoNotDestroyMe := false;
  Result := DoSave;
end;

function TZXMLSSave.DoSave: integer;
begin
  Result := (
    CreateSaverForDescription(ExtractFileExt(FFile))
    as IZXMLSSaveImpl ).InternalSave;
 // should reset DoNotDestroy flag in secondary class as well
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

function TZXMLSSave.NoZip: IZXMLSSave;
begin
  FZipGen := TZxZipGen.QueryDummyZipGen;
  Result := Self;
end;

function TZXMLSSave.ZipWith(const ZipGenerator: CZxZipGens): IZXMLSSave;
begin
  FZipGen := ZipGenerator;
  Result := Self;
end;

{$IfOpt D+}
// debug-only hooks for setting breakpoints
// and superwising lifetime of saving API-wrapping objects
function TZXMLSSave._AddRef: Integer; StdCall;
begin
  Result := inherited _AddRef;
end;

function TZXMLSSave._Release: Integer; StdCall;
begin
  Result := inherited _Release;
end;
{$EndIf}

initialization
  SaveClasses := TClassList.Create;

// Ersatz-testing below
//  if '' = ZxCharSetByCodePage(866) then ;
//  if '' = ZxCharSetByCodePage(1251) then ;
//  if '' = ZxCharSetByCodePage(20866) then ;

{$IfNDef MSWindows}
  ZxCharSetByCodePage := @ZxCharSetByCodePageSimple;//ZxCharSetByOops;
{$EndIf}
finalization
  SaveClasses.Free;
end.
