unit zeZippyJCL7z;

(* Simplistic interface for creating simplistic Zip files.
   Bridge object for avemey.com components.

   This one bridges www.7-zip.org DLL via wrapper by jcl.sf.net
      The wrapper by Henri Gourvest (progdigy.com) is not supported.

   (c) the Arioch, licensed under zLib license *)

interface uses zeZippy, SysUtils, Classes, JclCompression;

type TZxZipJcl7z = class (TZxZipGen)
     public

       /// Implementations should try to create the file for writing
       ///    and throw exception if they can not.
       constructor Create(const ZipFile: TFileName); override;
       procedure BeforeDestruction; override;

     protected
       FZ: TJclCompressArchive;

       procedure DoSaveAndSeal;  override;

       /// Returns True if the stream was flushed and clearance is given to free it.
       /// Otherwise the stream object is transmitted to the sealed list
       function  DoSealStream(const Data: TStream; const RelName: TFileName): boolean;  override;
   end;

implementation

{ TZxZipJcl7z }

procedure TZxZipJcl7z.BeforeDestruction;
begin
  FZ.Free;
  inherited;
end;

constructor TZxZipJcl7z.Create(const ZipFile: TFileName);
begin
  inherited;
  FZ := TJclZipCompressArchive.Create(self.ZipFileName);
end;

procedure TZxZipJcl7z.DoSaveAndSeal;
begin
  FZ.Compress;
end;

function TZxZipJcl7z.DoSealStream(const Data: TStream;
  const RelName: TFileName): boolean;
begin
  Data.Position := 0;
  FZ.AddFile(RelName, Data, False {do not own});
  Result := false; // do not Free
end;

initialization
  TZxZipJcl7z.Register;
finalization
  TZxZipJcl7z.UnRegister;
end.
