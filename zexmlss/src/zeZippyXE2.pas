unit zeZippyXE2;
(* Simplistic interface to creating simplistic Zip files.
   Bridge object for avemey.com components.
   This one bridges to Embarcadero Delphi XE2 and above

   (c) the Arioch, licensed under zLib license *)

interface uses zeZippy, SysUtils, Classes;

type TZxZipXE2 = class (TZxZipGen)
     public

       /// Implementations should try to create the file for writing
       ///    and throw exception if they can not.
       constructor Create(const ZipFile: TFileName); override;

     protected
       procedure DoAbortAndDelete;  override;
       procedure DoSaveAndSeal;  override;
       function  DoNewStream(const RelativeName: TFileName): TStream; override;

       /// Returns True if the stream was flushed and clearance is given to free it.
       /// Otherwise is transmitted to the sealed list
       function  DoSealStream(const Data: TStream; const RelName: TFileName): boolean;  override;
   end;

implementation

{ TZxZipXE2 }

constructor TZxZipXE2.Create(const ZipFile: TFileName);
begin
  inherited;

end;

procedure TZxZipXE2.DoAbortAndDelete;
begin
  inherited;

end;

function TZxZipXE2.DoNewStream(const RelativeName: TFileName): TStream;
begin

end;

procedure TZxZipXE2.DoSaveAndSeal;
begin
  inherited;

end;

function TZxZipXE2.DoSealStream(const Data: TStream;
  const RelName: TFileName): boolean;
begin

end;

end.
