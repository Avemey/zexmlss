unit zeZippyXE2;
(* Simplistic interface for creating simplistic Zip files.
   Bridge object for avemey.com components.
   This one bridges to Embarcadero Delphi XE2 and above

   (c) the Arioch, licensed under zLib license *)

interface uses zeZippy, SysUtils, Classes, System.Zip;

type TZxZipXE2 = class (TZxZipGen)
     public

       /// Implementations should try to create the file for writing
       ///    and throw exception if they can not.
       constructor Create(const ZipFile: TFileName); override;
       procedure BeforeDestruction; override;

     protected
       FZ: TZipFile;

       procedure DoAbortAndDelete;  override;
       procedure DoSaveAndSeal;  override;

       /// Returns True if the stream was flushed and clearance is given to free it.
       /// Otherwise is transmitted to the sealed list
       function  DoSealStream(const Data: TStream; const RelName: TFileName): boolean;  override;
   end;

implementation

{ TZxZipXE2 }

procedure TZxZipXE2.BeforeDestruction;
begin
  FZ.Free;
  inherited;
end;

constructor TZxZipXE2.Create(const ZipFile: TFileName);
begin
  inherited;
  FZ := TZipFile.Create;
  FZ.Open(ZipFileName, zmWrite);
end;

procedure TZxZipXE2.DoAbortAndDelete;
begin
  inherited;

  FZ.Close;
  DeleteFile(ZipFileName);
end;

procedure TZxZipXE2.DoSaveAndSeal;
begin
  FZ.Close;
end;

function TZxZipXE2.DoSealStream(const Data: TStream;
  const RelName: TFileName): boolean;
begin
  Data.Position := 0;
  // like TZipper of Lazarus, XE2's zip goes nuts if not resetting a position
  // it even generates broken zip file with fake CRC
  FZ.Add(Data, RelName);
  Result := true;
end;

initialization
  TZxZipXE2.Register;
finalization
  TZxZipXE2.UnRegister;
end.
