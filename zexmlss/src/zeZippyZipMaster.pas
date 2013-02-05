unit zeZippyZipMaster;
(* Simplistic interface for creating simplistic Zip files.
   Bridge object for avemey.com components.
   This one bridges to TZipMaster by www.delphizip.org

   (c) the Arioch, licensed under zLib license *)

interface
uses zeZippy, SysUtils, Classes, ZipMstr;

type TZxZipMastered = class (TZxZipGen)
     public
       /// Implementations should try to create the file for writing
       ///    and throw exception if they can not.
       constructor Create(const ZipFile: TFileName); override;
       procedure BeforeDestruction; override;

     protected
       FZM: TZipMaster;

       procedure DoAbortAndDelete;  override;
       procedure DoSaveAndSeal;  override;

       /// Returns True if the stream was flushed and clearance is given to free it.
       /// Otherwise is transmitted to the sealed list
       function  DoSealStream(const Data: TStream; const RelName: TFileName): boolean;  override;
   end;


implementation

{ TZxZipMastered }

procedure TZxZipMastered.BeforeDestruction;
begin
  FZM.Free;
  inherited;
end;

constructor TZxZipMastered.Create(const ZipFile: TFileName);
begin
  inherited;

  FZM := TZipMaster.Create(nil);
  FZM.ZipFileName := Self.ZipFileName;
  FZM.AddOptions := [AddDirNames{, AddSafe}];
  FZM.WriteOptions := [zwoForceDest];
end;

procedure TZxZipMastered.DoAbortAndDelete;
begin
  inherited;

  FZM.Clear;
  FZM.ZipFileName := ''; // hopefully it would close it
end;

function TZxZipMastered.DoSealStream(const Data: TStream; const RelName: TFileName): boolean;
begin
//  FZM.AddStreamToStream(Data as TMemoryStream);
  FZM.ZipStream.LoadFromStream(Data);
  FZM.AddStreamToFile(RelName,0,0);
  Result := True;
end;

procedure TZxZipMastered.DoSaveAndSeal;
begin
//  FZM.Add;
   // maybe nothign at all needed ?
   // Add*** methods loads and unloads DLL
end;

initialization
  TZxZipMastered.Register;
finalization
  TZxZipMastered.UnRegister;
end.
