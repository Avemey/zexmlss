program roundrectcell;

uses
  Forms,
  unit_main in 'unit_main.pas' {frmMain};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
