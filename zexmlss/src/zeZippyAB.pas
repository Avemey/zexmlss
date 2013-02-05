//****************************************************************
// Simplistic interface for creating simplistic Zip files.
//   Bridge object for zexmlss components.
//   This one bridges to TurboPower Abbrevia
//        http://sourceforge.net/projects/tpabbrevia/
//****************************************************************

unit zeZippyAB;

interface

uses
  zeZippy, SysUtils, Classes, AbZipper, AbArcTyp, AbUtils;

type
  TZxZipAbbrevia = class(TZxZipGen)
  public
    /// Implementations should try to create the file for writing
    ///    and throw exception if they can not.
    constructor Create(const ZipFile: TFileName); override;
    procedure BeforeDestruction(); override;

  protected
    FAb: TAbZipper;

    procedure DoAbortAndDelete();  override;
    procedure DoSaveAndSeal();  override;

    /// Returns True if the stream was flushed and clearance is given to free it.
    /// Otherwise is transmitted to the sealed list
    function  DoSealStream(const Data: TStream; const RelName: TFileName): boolean;  override;
  end;


implementation

{ TZxZipAbbrevia }

procedure TZxZipAbbrevia.BeforeDestruction();
begin
  FreeAndNil(FAb);
  inherited;
end;

constructor TZxZipAbbrevia.Create(const ZipFile: TFileName);
begin
  inherited;
  FAb := TAbZipper.Create(nil);
  FAb.ArchiveType := atZip;
  FAb.ForceType := true;
  FAb.FileName := ZipFile;
end;

procedure TZxZipAbbrevia.DoAbortAndDelete();
begin
  inherited;

  //??

end;

procedure TZxZipAbbrevia.DoSaveAndSeal();
begin
  //???

end;

function TZxZipAbbrevia.DoSealStream(const Data: TStream;
  const RelName: TFileName): boolean;
begin
  FAb.AddFromStream(RelName, Data);
  result := true;
end;

initialization
  TZxZipAbbrevia.Register();

finalization
  TZxZipAbbrevia.UnRegister;

end.
