object frmMain: TfrmMain
  Left = 192
  Top = 107
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'ZColorStringGrid Sample'
  ClientHeight = 616
  ClientWidth = 871
  Color = clBtnFace
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 14
  object Label1: TLabel
    Left = 247
    Top = 3
    Width = 374
    Height = 22
    Caption = #1055#1088#1080#1084#1077#1088' '#1075#1088#1080#1076#1072' '#1089' '#1086#1073#1098#1077#1076#1080#1085#1105#1085#1085#1099#1084#1080' '#1103#1095#1077#1081#1082#1072#1084#1080
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object ZSG: TZColorStringGrid
    Left = 8
    Top = 32
    Width = 857
    Height = 577
    ColCount = 9
    DefaultDrawing = False
    FixedColor = clBtnFace
    RowCount = 22
    FixedRows = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
    TabOrder = 0
    OnKeyPress = ZSGKeyPress
    OnSetEditText = ZSGSetEditText
    DefaultCellStyle.Font.Charset = RUSSIAN_CHARSET
    DefaultCellStyle.Font.Color = clWindowText
    DefaultCellStyle.Font.Height = -13
    DefaultCellStyle.Font.Name = 'Comic Sans MS'
    DefaultCellStyle.Font.Style = []
    DefaultCellStyle.BGColor = clInfoBk
    DefaultCellStyle.HorizontalAlignment = taRightJustify
    DefaultFixedCellStyle.Font.Charset = RUSSIAN_CHARSET
    DefaultFixedCellStyle.Font.Color = clWindowText
    DefaultFixedCellStyle.Font.Height = -16
    DefaultFixedCellStyle.Font.Name = 'Times New Roman'
    DefaultFixedCellStyle.Font.Style = [fsBold]
    DefaultFixedCellStyle.BGColor = clMoneyGreen
    DefaultFixedCellStyle.BorderCellStyle = sgRaised
    DefaultFixedCellStyle.HorizontalAlignment = taCenter
    LineDesign.LineUpColor = clWhite
    SelectedColors.BGColor = clYellow
    SelectedColors.FontColor = clWindowFrame
    SizingHeight = True
    SizingWidth = True
    WordWrap = True
  end
end
