object frmMain: TfrmMain
  Left = 192
  Top = 107
  Width = 800
  Height = 781
  Caption = 'ZColorStringGrid + ZCLabel example (http://avemey.com)'
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
  object ZSG: TZColorStringGrid
    Left = 0
    Top = 145
    Width = 792
    Height = 609
    Align = alClient
    ColCount = 30
    DefaultRowHeight = 17
    DefaultDrawing = False
    FixedColor = clBtnFace
    FixedCols = 2
    RowCount = 60
    FixedRows = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing, goEditing]
    TabOrder = 0
    DefaultCellStyle.Font.Charset = DEFAULT_CHARSET
    DefaultCellStyle.Font.Color = clWindowText
    DefaultCellStyle.Font.Height = -13
    DefaultCellStyle.Font.Name = 'Times New Roman'
    DefaultCellStyle.Font.Style = []
    DefaultCellStyle.BGColor = clWindow
    DefaultCellStyle.SizingHeight = True
    DefaultCellStyle.SizingWidth = True
    DefaultCellStyle.WordWrap = True
    DefaultFixedCellStyle.Font.Charset = DEFAULT_CHARSET
    DefaultFixedCellStyle.Font.Color = clWindowText
    DefaultFixedCellStyle.Font.Height = -16
    DefaultFixedCellStyle.Font.Name = 'Tahoma'
    DefaultFixedCellStyle.Font.Style = []
    DefaultFixedCellStyle.BGColor = clMoneyGreen
    DefaultFixedCellStyle.BorderCellStyle = sgRaised
    DefaultFixedCellStyle.HorizontalAlignment = taCenter
    DefaultFixedCellStyle.SizingHeight = True
    DefaultFixedCellStyle.SizingWidth = True
    DefaultFixedCellStyle.WordWrap = True
    LineDesign.LineUpColor = clWhite
    SelectedColors.BGColor = clYellow
    SelectedColors.FontColor = clBlack
    SizingHeight = True
    SizingWidth = True
    UseCellSizingHeight = True
    UseCellSizingWidth = True
    UseCellWordWrap = True
  end
  object PanelTop: TPanel
    Left = 0
    Top = 0
    Width = 792
    Height = 145
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
    object ZCLabel1: TZCLabel
      Left = 0
      Top = 8
      Width = 89
      Height = 97
      AlignmentVertical = 1
      AlignmentHorizontal = 1
      Caption = 'ZCLabel example '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -20
      Font.Name = 'Times New Roman'
      Font.Style = [fsBold]
      IndentHor = 0
      Rotate = 90
      ParentFont = False
    end
    object ZCLabel2: TZCLabel
      Left = 408
      Top = 0
      Width = 265
      Height = 145
      AlignmentVertical = 1
      AlignmentHorizontal = 1
      Caption = 'zclabel'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -21
      Font.Name = 'Tahoma'
      Font.Style = []
      IndentHor = 0
      Rotate = 45
      ParentFont = False
    end
    object ZCLabel3: TZCLabel
      Left = 88
      Top = 112
      Width = 289
      Height = 30
      Caption = #1055#1086#1074#1086#1088#1072#1095#1080#1074#1072#1090#1100' '#1090#1077#1082#1089#1090' '#1084#1086#1078#1085#1086' '#1090#1086#1083#1100#1082#1086' '#1076#1083#1103' TrueType '#1096#1088#1080#1092#1090#1086#1074
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      IndentHor = 0
      Rotate = 180
      ParentFont = False
    end
    object Label1: TLabel
      Left = 344
      Top = 8
      Width = 46
      Height = 13
      Caption = #1055#1086#1074#1086#1088#1086#1090':'
    end
    object btnVTop: TBitBtn
      Left = 94
      Top = 20
      Width = 75
      Height = 25
      Caption = 'Top'
      TabOrder = 0
      OnClick = btnVTopClick
    end
    object btnVCenter: TBitBtn
      Left = 94
      Top = 52
      Width = 75
      Height = 25
      Caption = 'Center'
      TabOrder = 1
      OnClick = btnVCenterClick
    end
    object btnVBottom: TBitBtn
      Left = 94
      Top = 84
      Width = 75
      Height = 25
      Caption = 'Bottom'
      TabOrder = 2
      OnClick = btnVBottomClick
    end
    object btnHLeft: TBitBtn
      Left = 177
      Top = 20
      Width = 75
      Height = 25
      Caption = 'Left'
      TabOrder = 3
      OnClick = btnHLeftClick
    end
    object btnHCenter: TBitBtn
      Left = 177
      Top = 52
      Width = 75
      Height = 25
      Caption = 'Center'
      TabOrder = 4
      OnClick = btnHCenterClick
    end
    object btnHRight: TBitBtn
      Left = 177
      Top = 84
      Width = 75
      Height = 25
      Caption = 'Right'
      TabOrder = 5
      OnClick = btnHRightClick
    end
    object btnFont: TBitBtn
      Left = 259
      Top = 20
      Width = 75
      Height = 25
      Caption = 'Font'
      TabOrder = 6
      OnClick = btnFontClick
    end
    object btnBGColor: TBitBtn
      Left = 259
      Top = 52
      Width = 75
      Height = 25
      Caption = 'BGColor'
      TabOrder = 7
      OnClick = btnBGColorClick
    end
    object btnMerge: TButton
      Left = 259
      Top = 84
      Width = 75
      Height = 25
      Caption = #1054#1073#1098#1077#1076#1080#1085#1080#1090#1100
      TabOrder = 8
      OnClick = btnMergeClick
    end
    object EditRotate: TEdit
      Left = 344
      Top = 24
      Width = 49
      Height = 21
      TabOrder = 9
      Text = '0'
    end
    object RotateUP: TUpDown
      Left = 393
      Top = 24
      Width = 15
      Height = 21
      Associate = EditRotate
      Max = 359
      TabOrder = 10
    end
    object btnSetRotate: TButton
      Left = 344
      Top = 48
      Width = 75
      Height = 25
      Caption = #1055#1086#1074#1077#1088#1085#1091#1090#1100
      TabOrder = 11
      OnClick = btnSetRotateClick
    end
  end
  object FntDialog: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Left = 600
    Top = 8
  end
  object BGColorDialog: TColorDialog
    Left = 600
    Top = 40
  end
end
