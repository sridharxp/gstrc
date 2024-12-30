object frmFC2: TfrmFC2
  Left = 642
  Top = 141
  Width = 561
  Height = 202
  Caption = 'Reconcile and  Compare  '
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = mm1
  OldCreateOrder = False
  Position = poDefaultSizeOnly
  PixelsPerInch = 96
  TextHeight = 13
  object lbl1: TLabel
    Left = 12
    Top = 21
    Width = 45
    Height = 16
    Caption = 'Tally IP'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object lbl2: TLabel
    Left = 260
    Top = 21
    Width = 24
    Height = 16
    Caption = 'Port'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object edtHost: TEdit
    Left = 129
    Top = 21
    Width = 118
    Height = 28
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    Text = '127.0.0.1'
  end
  object edtPort: TEdit
    Left = 325
    Top = 21
    Width = 81
    Height = 28
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    Text = '9000'
  end
  object mm1: TMainMenu
    Left = 424
    Top = 16
    object mniool1: TMenuItem
      Caption = 'Tool'
      object mniReconcile2A2B1: TMenuItem
        Caption = 'Compare Purchase'
        OnClick = mniReconcile2A2B1Click
      end
      object mniImportTallyData1: TMenuItem
        Caption = 'Import Tally Data'
        OnClick = mniImportTallyData1Click
      end
      object mniImportGSTRjson1: TMenuItem
        Caption = 'Import GSTR JSON'
        OnClick = mniImportGSTRjson1Click
      end
    end
    object mniHelp1: TMenuItem
      Caption = 'Help'
      object mniTestTally: TMenuItem
        Caption = 'TestTally'
        OnClick = mniTestTallyClick
      end
      object mniFindIP1: TMenuItem
        Caption = 'Find IP'
        OnClick = mniFindIP1Click
      end
      object mniFC2: TMenuItem
        Caption = 'Developer'
        OnClick = mniFC2Click
      end
    end
  end
end
