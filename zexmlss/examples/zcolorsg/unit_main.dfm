object frmMain: TfrmMain
  Left = 192
  Top = 107
  Width = 800
  Height = 781
  Caption = 'ZColorStringGrid + ZEXMLSS'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object ZSG: TZColorStringGrid
    Left = 0
    Top = 113
    Width = 792
    Height = 641
    Align = alClient
    DefaultRowHeight = 17
    DefaultDrawing = False
    FixedColor = clBtnFace
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing, goEditing]
    TabOrder = 0
    DefaultCellStyle.Font.Charset = DEFAULT_CHARSET
    DefaultCellStyle.Font.Color = clWindowText
    DefaultCellStyle.Font.Height = -11
    DefaultCellStyle.Font.Name = 'MS Sans Serif'
    DefaultCellStyle.Font.Style = []
    DefaultCellStyle.BGColor = clWindow
    DefaultCellStyle.SizingHeight = True
    DefaultCellStyle.SizingWidth = True
    DefaultFixedCellStyle.Font.Charset = DEFAULT_CHARSET
    DefaultFixedCellStyle.Font.Color = clWindowText
    DefaultFixedCellStyle.Font.Height = -11
    DefaultFixedCellStyle.Font.Name = 'MS Sans Serif'
    DefaultFixedCellStyle.Font.Style = []
    DefaultFixedCellStyle.BGColor = clOlive
    DefaultFixedCellStyle.BorderCellStyle = sgRaised
    LineDesign.LineUpColor = clWhite
    SizingHeight = True
    SizingWidth = True
    UseCellSizingHeight = True
    UseCellSizingWidth = True
  end
  object PanelTop: TPanel
    Left = 0
    Top = 0
    Width = 792
    Height = 113
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object LabelV: TLabel
      Left = 112
      Top = 0
      Width = 45
      Height = 16
      Caption = 'Vertical'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object LabelH: TLabel
      Left = 183
      Top = 0
      Width = 60
      Height = 16
      Caption = 'Horizontal'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Bevel1: TBevel
      Left = 92
      Top = 18
      Width = 162
      Height = 93
    end
    object btnLoad: TButton
      Left = 8
      Top = 32
      Width = 75
      Height = 25
      Hint = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1080#1079' '#1092#1072#1081#1083#1072' Excel XML'
      Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = btnLoadClick
    end
    object btnSave: TButton
      Left = 8
      Top = 64
      Width = 75
      Height = 25
      Hint = #1057#1086#1093#1088#1072#1080#1090#1100' '#1074' '#1092#1072#1081#1083#1077' '#1092#1086#1088#1084#1072#1090#1072' Excel XML'
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = btnSaveClick
    end
    object btnVTop: TBitBtn
      Left = 94
      Top = 20
      Width = 75
      Height = 25
      Caption = 'Top'
      TabOrder = 2
      OnClick = btnVTopClick
    end
    object btnVCenter: TBitBtn
      Left = 94
      Top = 52
      Width = 75
      Height = 25
      Caption = 'Center'
      TabOrder = 3
      OnClick = btnVCenterClick
    end
    object btnVBottom: TBitBtn
      Left = 94
      Top = 84
      Width = 75
      Height = 25
      Caption = 'Bottom'
      TabOrder = 4
      OnClick = btnVBottomClick
    end
    object btnHLeft: TBitBtn
      Left = 177
      Top = 20
      Width = 75
      Height = 25
      Caption = 'Left'
      TabOrder = 5
      OnClick = btnHLeftClick
    end
    object btnHCenter: TBitBtn
      Left = 177
      Top = 52
      Width = 75
      Height = 25
      Caption = 'Center'
      TabOrder = 6
      OnClick = btnHCenterClick
    end
    object btnHRight: TBitBtn
      Left = 177
      Top = 84
      Width = 75
      Height = 25
      Caption = 'Right'
      TabOrder = 7
      OnClick = btnHRightClick
    end
    object btnFont: TBitBtn
      Left = 259
      Top = 20
      Width = 75
      Height = 25
      Caption = 'Font'
      TabOrder = 8
      OnClick = btnFontClick
    end
    object btnBGColor: TBitBtn
      Left = 259
      Top = 52
      Width = 75
      Height = 25
      Caption = 'BGColor'
      TabOrder = 9
      OnClick = btnBGColorClick
    end
    object btnMerge: TButton
      Left = 259
      Top = 84
      Width = 75
      Height = 25
      Caption = #1054#1073#1098#1077#1076#1080#1085#1080#1090#1100
      TabOrder = 10
      OnClick = btnMergeClick
    end
    object RGShift: TRadioGroup
      Left = 342
      Top = 5
      Width = 139
      Height = 104
      Caption = #1044#1086#1082#1091#1084#1077#1085#1090' '#1074#1089#1090#1072#1074#1083#1103#1090#1100':'
      ItemIndex = 0
      Items.Strings = (
        #1055#1086#1074#1077#1088#1093
        #1057#1076#1074#1080#1075' '#1087#1088#1072#1074#1086
        #1057#1076#1074#1080#1075' '#1074#1085#1080#1079
        #1057#1076#1074#1080#1075' '#1074#1085#1080#1079' '#1080' '#1074#1087#1088#1072#1074#1086)
      TabOrder = 11
    end
    object CBInsertStart: TCheckBox
      Left = 496
      Top = 8
      Width = 233
      Height = 17
      Caption = #1042#1089#1090#1072#1074#1083#1103#1090#1100' '#1085#1072#1095#1080#1085#1072#1103' '#1089' '#1072#1082#1090#1080#1074#1085#1086#1081' '#1103#1095#1077#1081#1082#1080
      TabOrder = 12
    end
    object CBCopyBG: TCheckBox
      Left = 496
      Top = 32
      Width = 177
      Height = 17
      Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100' BGColor'
      Checked = True
      State = cbChecked
      TabOrder = 13
    end
    object CBCopyFont: TCheckBox
      Left = 496
      Top = 56
      Width = 145
      Height = 17
      Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100' Font'
      Checked = True
      State = cbChecked
      TabOrder = 14
    end
    object CBCopyMerge: TCheckBox
      Left = 496
      Top = 80
      Width = 177
      Height = 17
      Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100' Merge Area'
      Checked = True
      State = cbChecked
      TabOrder = 15
    end
  end
  object ODxml: TOpenDialog
    DefaultExt = 'xml'
    Filter = 'Excel XML|*.xml'
    Left = 760
    Top = 32
  end
  object SDxml: TSaveDialog
    DefaultExt = 'xml'
    Filter = 'Excel XML|*.xml'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 760
    Top = 64
  end
  object FntDialog: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Left = 720
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
    DocumentProperties.Created = 40385.667887152770000000
    DocumentProperties.Company = 'none'
    DocumentProperties.Version = '11.9999'
    HorPixelSize = 0.265000000000000000
    VertPixelSize = 0.265000000000000000
    Left = 760
  end
  object BGColorDialog: TColorDialog
    Left = 720
    Top = 32
  end
end
