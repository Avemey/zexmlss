unit unit_main;

{$IFDEF FPC}
  { $MODE Delphi}
{$ENDIF}

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, zclab;

type
  TfrmMain = class(TForm)
    ZCLabel1: TZCLabel;
    ZCLabel2: TZCLabel;
    ZCLabel3: TZCLabel;
    ZCLabel4: TZCLabel;
    ZCLabel5: TZCLabel;
    ZCLabel6: TZCLabel;
    ZCLabel7: TZCLabel;
    ZCLabel8: TZCLabel;
    ZCLabel9: TZCLabel;
    ZCLabel10: TZCLabel;
    ZCLabel11: TZCLabel;
    ZCLabel12: TZCLabel;
    ZCLabel13: TZCLabel;
    ZCLabel14: TZCLabel;
    ZCLabel15: TZCLabel;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$IFNDEF FPC}
  {$R *.dfm}
{$ELSE}
  {$R *.lfm}
{$ENDIF}

procedure TfrmMain.FormCreate(Sender: TObject);
var
  i: integer;
  num: integer;
  ZCL: TZCLabel;
  s: string;
  ang: integer;
  sEOL: string;

begin
  num := 0;
  ang := 0;
  {$IFDEF FPC}
  sEOL := LineEnding;
  {$ELSE}
  sEOL := sLineBreak;
  {$ENDIF}

  s := 'Text example' + sEOL + 'smttext.';
  //Iterate through the components
  for i := 0 to ComponentCount - 1 do
    if (Components[i] is TZCLabel) then
    begin
      ZCL := Components[i] as TZCLabel;
      ZCL.Font.Size := 10;
      ZCL.Caption := s;
      inc(num);
      if (num mod 5 in [1, 2, 3]) then
      begin
        ZCL.Transparent := false;
        ZCL.Color := clYellow;
      end else
      begin
        ZCL.Font.Size := 14;
        inc(ang, 50);
        //Set rotate angle
        ZCL.Rotate := ang;
      end;
      //Horizontal alignment
      ZCL.AlignmentHorizontal := num mod 3;
      //Vertical alignment
      ZCL.AlignmentVertical := num mod 5;
    end; //if



end;

end.
