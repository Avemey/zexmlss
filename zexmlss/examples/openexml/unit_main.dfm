object frmMain: TfrmMain
  Left = 192
  Top = 107
  BorderStyle = bsSingle
  Caption = #1047#1072#1075#1088#1091#1079#1082#1072' '#1076#1086#1082#1091#1084#1077#1085#1090#1086#1074'  excel XML '#1074' stringgrid'
  ClientHeight = 551
  ClientWidth = 762
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object lblList: TLabel
    Left = 176
    Top = 8
    Width = 31
    Height = 13
    Caption = #1051#1080#1089#1090' :'
  end
  object SGtest: TStringGrid
    Left = 16
    Top = 32
    Width = 729
    Height = 457
    FixedCols = 0
    FixedRows = 0
    TabOrder = 0
  end
  object btnOpen: TButton
    Left = 232
    Top = 496
    Width = 281
    Height = 49
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' Excel XML'
    TabOrder = 1
    OnClick = btnOpenClick
  end
  object CBList: TComboBox
    Left = 208
    Top = 5
    Width = 217
    Height = 21
    Style = csDropDownList
    Enabled = False
    ItemHeight = 13
    TabOrder = 2
    OnSelect = CBListSelect
  end
  object ODxml: TOpenDialog
    DefaultExt = 'xml'
    Filter = 'excel XML|*.xml'
    Left = 528
  end
  object ZEXMLSStest: TZEXMLSS
    Styles.DefaultStyle.Font.Charset = DEFAULT_CHARSET
    Styles.DefaultStyle.Font.Color = clBlack
    Styles.DefaultStyle.Font.Height = -13
    Styles.DefaultStyle.Font.Name = 'MS Sans Serif'
    Styles.DefaultStyle.Font.Style = []
    Styles.DefaultStyle.Alignment.Rotate = 0
    Styles.DefaultStyle.NumberFormat = 'General'
    DocumentProperties.Author = 'none'
    DocumentProperties.LastAuthor = 'none'
    DocumentProperties.Created = 40365.648251469910000000
    DocumentProperties.Company = 'none'
    DocumentProperties.Version = '11.9999'
    HorPixelSize = 0.265000000000000000
    VertPixelSize = 0.265000000000000000
    Left = 568
  end
end
