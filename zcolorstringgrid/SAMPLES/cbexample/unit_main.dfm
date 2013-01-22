object frmMain: TfrmMain
  Left = 192
  Top = 107
  Width = 696
  Height = 568
  Caption = 'zcolorstringgrid and zclabel example (http://avemey.com)'
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
  object ZCLabel1: TZCLabel
    Left = 440
    Top = 384
    Width = 129
    Height = 137
    AlignmentVertical = 1
    AlignmentHorizontal = 1
    Caption = 'ZCLabel example'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -27
    Font.Name = 'Times New Roman'
    Font.Style = []
    IndentHor = 0
    Rotate = 45
    ParentFont = False
  end
  object ZSG: TZColorStringGrid
    Left = 0
    Top = 0
    Width = 688
    Height = 361
    Align = alTop
    ColCount = 20
    DefaultDrawing = False
    FixedColor = clBtnFace
    RowCount = 20
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
    TabOrder = 0
    DefaultCellStyle.Font.Charset = DEFAULT_CHARSET
    DefaultCellStyle.Font.Color = clWindowText
    DefaultCellStyle.Font.Height = -16
    DefaultCellStyle.Font.Name = 'Arial'
    DefaultCellStyle.Font.Style = []
    DefaultCellStyle.BGColor = clWindow
    DefaultFixedCellStyle.Font.Charset = DEFAULT_CHARSET
    DefaultFixedCellStyle.Font.Color = clWindowText
    DefaultFixedCellStyle.Font.Height = -13
    DefaultFixedCellStyle.Font.Name = 'Tahoma'
    DefaultFixedCellStyle.Font.Style = []
    DefaultFixedCellStyle.BGColor = clBtnFace
    DefaultFixedCellStyle.BorderCellStyle = sgRaised
    LineDesign.LineColor = clSilver
    LineDesign.LineUpColor = clWhite
    SizingHeight = True
    SizingWidth = True
    WordWrap = True
  end
end
