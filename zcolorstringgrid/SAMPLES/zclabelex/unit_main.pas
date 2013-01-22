unit unit_main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, zclab;

type

  { TfrmMain }

  TfrmMain = class(TForm)
    ZCLabel1: TZCLabel;
    ZCLabel10: TZCLabel;
    ZCLabel2: TZCLabel;
    ZCLabel3: TZCLabel;
    ZCLabel4: TZCLabel;
    ZCLabel5: TZCLabel;
    ZCLabel6: TZCLabel;
    ZCLabel8: TZCLabel;
    ZCLabel9: TZCLabel;
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  frmMain: TfrmMain;

implementation

{$R *.lfm}

{ TfrmMain }

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  ZCLabel1.Caption := 'ZClabel пример выравнивания' + LineEnding +'текст' + LineEnding + '123';
  ZCLabel2.Caption := ZCLabel1.Caption;
  ZCLabel3.Caption := ZCLabel1.Caption;
  ZCLabel4.Caption := ZCLabel1.Caption;
  ZCLabel8.Caption := 'ZCLabel example ' + LineEnding + 'text';
end;

end.

