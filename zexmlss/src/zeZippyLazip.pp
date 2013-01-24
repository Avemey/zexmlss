unit zeZippyLazip;
(* Simplistic interface to creating simplistic Zip files.
   Bridge object for avemey.com components.
   This one bridges to Lazaurs TZippes

   (c) the Arioch, licensed under zLib license

   Not tested - i know nothing of TZipper!  *)


{$mode objfpc}{$H+}

interface uses zeZippy, SysUtils, Classes, zipper;

type

{ TZxZipLazip }

 TZxZipLazip = class (TZxZipGen)
     public

       /// Implementations should try to create the file for writing
       ///    and throw exception if they can not.
       constructor Create(const ZipFile: TFileName); override;
       procedure BeforeDestruction; override;

     protected
       FZ: TZipper;

       procedure DoAbortAndDelete;  override;
       procedure DoSaveAndSeal;  override;
       function  DoNewStream(const RelativeName: TFileName): TStream; override;

       /// Returns True if the stream was flushed and clearance is given to free it.
       /// Otherwise is transmitted to the sealed list
       function  DoSealStream(const Data: TStream; const RelName: TFileName): boolean;  override;
   end;

implementation

{ TZxZipLazip }

constructor TZxZipLazip.Create(const ZipFile: TFileName);
begin
  inherited Create(ZipFile);

  FZ := TZipper.Create();
  FZ.FileName := self.ZipFileName;
end;

procedure TZxZipLazip.BeforeDestruction;
begin
  inherited;  // before freeing - may throw exception on unsealed data
  FZ.Free;    // perhaps there can be second attempt ?
end;

procedure TZxZipLazip.DoAbortAndDelete;
begin
  // nothing to do ? Seems TZipper creates nothing until SaveToFile called
end;

procedure TZxZipLazip.DoSaveAndSeal;
begin
  FZ.ZipAllFiles;
  FZ.Clear;
// We also can iterate FSealedStreams and free them here
//   yet after ZipGen is donem there is nothing one can do with it but .Free
//   and then the base class would free the streams.
end;

function TZxZipLazip.DoNewStream(const RelativeName: TFileName): TStream;
begin
  Result := TMemoryStream.Create();
end;

function TZxZipLazip.DoSealStream(const Data: TStream; const RelName: TFileName
  ): boolean;
begin
  Data.Position := 0;
  FZ.Entries.AddFileEntry(Data, RelName);
  Result := false; // do not delete stream yet, until ZipAllFiles
end;

initialization
  RegisterZipGen(TZxZipLazip);
end.

