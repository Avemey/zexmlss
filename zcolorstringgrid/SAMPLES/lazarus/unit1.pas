unit Unit1; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, zclab;

type

  { TForm1 }

  TForm1 = class(TForm)
    ZCLabel1: TZCLabel;
    procedure ZCLabel1Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.ZCLabel1Click(Sender: TObject);
begin

end;

end.

