program zcolorsg;

uses
  Forms,
  unit_main in 'unit_main.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
