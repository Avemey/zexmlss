object frmMain: TfrmMain
  Left = 325
  Top = 191
  Width = 368
  Height = 237
  Caption = 'zcolorstringgrid style example'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object ZColorStringGrid1: TZColorStringGrid
    Left = 0
    Top = 0
    Width = 352
    Height = 199
    Align = alClient
    DefaultDrawing = False
    FixedColor = clBtnFace
    RowCount = 8
    TabOrder = 0
    DefaultCellStyle.Font.Charset = DEFAULT_CHARSET
    DefaultCellStyle.Font.Color = clWindowText
    DefaultCellStyle.Font.Height = -11
    DefaultCellStyle.Font.Name = 'MS Sans Serif'
    DefaultCellStyle.Font.Style = []
    DefaultCellStyle.BGColor = clWindow
    DefaultFixedCellStyle.Font.Charset = DEFAULT_CHARSET
    DefaultFixedCellStyle.Font.Color = clWindowText
    DefaultFixedCellStyle.Font.Height = -11
    DefaultFixedCellStyle.Font.Name = 'MS Sans Serif'
    DefaultFixedCellStyle.Font.Style = []
    DefaultFixedCellStyle.BGColor = clBtnFace
    DefaultFixedCellStyle.BorderCellStyle = sgRaised
    LineDesign.LineUpColor = clWhite
  end
end
