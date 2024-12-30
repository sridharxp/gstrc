object frmOfTlyP: TfrmOfTlyP
  Left = 236
  Top = 118
  Width = 557
  Height = 447
  Caption = 'Import from Tally'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -16
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 20
  object Label1: TLabel
    Left = 37
    Top = 95
    Width = 76
    Height = 20
    Caption = 'From Date'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 327
    Top = 95
    Width = 18
    Height = 20
    Caption = 'To'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 37
    Top = 215
    Width = 163
    Height = 40
    Caption = 'Party Ledger (Optional)'#13#10
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Info: TLabel
    Left = 21
    Top = 366
    Width = 30
    Height = 13
    Caption = 'Status'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label4: TLabel
    Left = 37
    Top = 28
    Width = 98
    Height = 20
    Caption = 'Voucher Type'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label6: TLabel
    Left = 37
    Top = 101
    Width = 3
    Height = 13
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label8: TLabel
    Left = 37
    Top = 157
    Width = 140
    Height = 20
    Caption = 'Company (Optional)'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object btnImport: TButton
    Left = 272
    Top = 275
    Width = 237
    Height = 20
    Caption = 'Import Tally Data to Books'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    OnClick = btnImportClick
  end
  object DateTimePicker2: TDateTimePicker
    Left = 379
    Top = 95
    Width = 130
    Height = 32
    Date = 45747.474290173610000000
    Format = 'dd/MM/yyyy'
    Time = 45747.474290173610000000
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
  end
  object cmbLedGroup: TComboBox
    Left = 177
    Top = 28
    Width = 118
    Height = 32
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 24
    ParentFont = False
    TabOrder = 0
    Text = 'Purchase'
    Items.Strings = (
      ''
      'Purchase'
      'Sales')
  end
  object cmbVendor: TComboBox
    Left = 241
    Top = 215
    Width = 268
    Height = 32
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 24
    ParentFont = False
    TabOrder = 4
  end
  object btnRefresh: TButton
    Left = 49
    Top = 275
    Width = 184
    Height = 20
    Caption = 'Import Ledger List'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
    OnClick = btnRefreshClick
  end
  object cmbFirm: TComboBox
    Left = 241
    Top = 154
    Width = 268
    Height = 32
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 24
    ParentFont = False
    TabOrder = 3
  end
  object DateTimePicker1: TDateTimePicker
    Left = 177
    Top = 95
    Width = 130
    Height = 32
    Date = 45383.474290173610000000
    Format = 'dd/MM/yyyy'
    Time = 45383.474290173610000000
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
  end
  object cbFaxInv: TCheckBox
    Left = 49
    Top = 320
    Width = 208
    Height = 17
    Caption = 'b2b Tax Invoices Only'
    Enabled = False
    TabOrder = 7
  end
  object frmEzReSizer: TFormResizer
    MinFontSize = 8
    MaxFontSize = 48
    Left = 448
    Top = 344
  end
end
