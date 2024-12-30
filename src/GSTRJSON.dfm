object frmImpGSTR: TfrmImpGSTR
  Left = 252
  Top = 30
  Width = 804
  Height = 219
  Caption = 'Import GSTRJson Data'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -15
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 16
  object lbl1: TLabel
    Left = 40
    Top = 30
    Width = 43
    Height = 20
    Caption = 'Open '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object lbl2: TLabel
    Left = 40
    Top = 74
    Width = 57
    Height = 20
    Caption = 'Save as'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Info: TLabel
    Left = 40
    Top = 106
    Width = 37
    Height = 16
    Caption = 'Status'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Info1: TLabel
    Left = 480
    Top = 74
    Width = 81
    Height = 20
    Caption = 'Format'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object edtReqFileName: TEdit
    Left = 128
    Top = 30
    Width = 612
    Height = 32
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    Text = 'Sales.json'
  end
  object cmbdbName: TComboBox
    Left = 590
    Top = 74
    Width = 150
    Height = 28
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 20
    ItemIndex = 0
    ParentFont = False
    TabOrder = 1
    Text = 'Tally2AB.xls'
    Items.Strings = (
      'Tally2AB.xls'
      'Sample.xls'
      '')
  end
  object btn1: TButton
    Left = 309
    Top = 139
    Width = 171
    Height = 25
    Caption = 'Import Json Data'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    OnClick = btn1Click
  end
  object edt1: TEdit
    Left = 344
    Top = 104
    Width = 121
    Height = 24
    TabOrder = 3
    Text = 'edt1'
  end
  object edtdbName: TEdit
    Left = 128
    Top = 74
    Width = 121
    Height = 28
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    Text = 'Tally2AB.xls'
  end
  object OpenDialog1: TOpenDialog
    Left = 376
    Top = 64
  end
  object frmEzReSizer: TFormResizer
    MinFontSize = 8
    MaxFontSize = 48
    Left = 424
    Top = 264
  end
end
