unit zeZippyXE2;
(* Simplistic interface to creating simplistic Zip files.
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
       function  DoNewStream(const RelativeName: TFileName): TStream; override;

       /// Returns True if the stream was flushed and clearance is given to free it.
       /// Otherwise is transmitted to the sealed list
       function  DoSealStream(const Data: TStream; const RelName: TFileName): boolean;  override;
   end;

implementation

{ TZxZipXE2 }

procedure TZxZipXE2.BeforeDestruction;
begin
  inherited;  // before freeing - may throw exception on unsealed data
  FZ.Free;    // perhaps there can be second attempt ?
end;

constructor TZxZipXE2.Create(const ZipFile: TFileName);
begin
  inherited;
  FZ := TZipFile.Create;
  FZ.Open(ZipFileName, zmWrite);
end;

procedure TZxZipXE2.DoAbortAndDelete;
var i: integer;
begin
  for i := FActiveStreams.Count - 1 downto 0 do begin
      FActiveStreams.Objects[i].Free; // can be nil - that is okay for .Free
      FActiveStreams.Delete(i);
  end;
  FZ.Close;
  DeleteFile(ZipFileName);
end;

function TZxZipXE2.DoNewStream(const RelativeName: TFileName): TStream;
begin
  Result := TMemoryStream.Create;
end;

procedure TZxZipXE2.DoSaveAndSeal;
begin
  FZ.Close;
end;

function TZxZipXE2.DoSealStream(const Data: TStream;
  const RelName: TFileName): boolean;
begin
  FZ.Add(Data, RelName);
  Result := true;
end;

initialization
  RegisterZipGen(TZxZipXE2);
end.
