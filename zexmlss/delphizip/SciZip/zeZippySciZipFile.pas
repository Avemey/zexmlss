//****************************************************************
// Simplistic interface for creating simplistic Zip files.
//   Bridge object for zexmlss components.
//   This one bridges to SciZipFile.pas
//****************************************************************

unit zeZippySciZipFile;

interface

uses
  zeZippy, SysUtils, Classes, SciZipFile;

type
  TZxZipSciZipFile = class(TZxZipGen)
  public
    // Implementations should try to create the file for writing
    // and throw exception if they can not.
    constructor Create(const ZipFile: TFileName); override;
    procedure BeforeDestruction(); override;

  protected
    FZipFile: TZipFile;
    FZipFileName: string;

    procedure DoAbortAndDelete;  override;
    procedure DoSaveAndSeal;  override;

    // Returns True if the stream was flushed and clearance is given to free it.
    // Otherwise is transmitted to the sealed list
    function  DoSealStream(const Data: TStream; const RelName: TFileName): boolean;  override;
  end;


implementation

{ TZxZipSciZipFile }

procedure TZxZipSciZipFile.BeforeDestruction;
begin
  FZipFile.Free;

  inherited;
end;

constructor TZxZipSciZipFile.Create(const ZipFile: TFileName);
begin
  inherited;

  FZipFile := TZipFile.Create;
  FZipFileName := ZipFile;
end;

procedure TZxZipSciZipFile.DoAbortAndDelete;
begin
  inherited;

  //??
end;

procedure TZxZipSciZipFile.DoSaveAndSeal;
begin
  FZipFile.SaveToFile(FZipFileName);
end;

function TZxZipSciZipFile.DoSealStream(const Data: TStream;
  const RelName: TFileName): boolean;
var
  buffer: TBytes;
begin
  FZipFile.AddFile(AnsiString(RelName));
  SetLength(buffer, Data.Size);
  Data.Position := 0;
  Data.ReadBuffer(buffer[0], Data.Size);
  FZipFile.Bytes[FZipFile.Count - 1] := buffer;

  Result := true;
end;

initialization
  TZxZipSciZipFile.Register;

finalization
  TZxZipSciZipFile.UnRegister;

end.
