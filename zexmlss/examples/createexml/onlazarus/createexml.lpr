program createexml;

{$MODE Delphi}

uses
  Forms, LResources, Interfaces,
  unit_main in 'unit_main.pas' {frmMain}, zexmlsslib;

{$R *.res}

begin
  {$I createexml.lrs}
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
