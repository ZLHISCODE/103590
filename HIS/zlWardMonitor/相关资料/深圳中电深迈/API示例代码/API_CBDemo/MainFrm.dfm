object Form1: TForm1
  Left = 0
  Top = 0
  Width = 1024
  Height = 738
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 65
    Width = 1016
    Height = 639
    Align = alClient
    Caption = 'Panel1'
    TabOrder = 0
    OnResize = Panel1Resize
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 1016
    Height = 65
    Align = alTop
    TabOrder = 1
    object Label1: TLabel
      Left = 560
      Top = 9
      Width = 27
      Height = 13
      Caption = #22995#21517':'
    end
    object Label2: TLabel
      Left = 673
      Top = 11
      Width = 27
      Height = 13
      Caption = #24180#40836':'
    end
    object Label3: TLabel
      Left = 8
      Top = 8
      Width = 57
      Height = 13
      AutoSize = False
      Caption = #26381#21153#22120'IP'
    end
    object Label4: TLabel
      Left = 166
      Top = 8
      Width = 52
      Height = 13
      AutoSize = False
      Caption = #31471#21475#21495
    end
    object Label5: TLabel
      Left = 8
      Top = 40
      Width = 73
      Height = 13
      AutoSize = False
      Caption = #30149#21382#32534#21495':'
    end
    object Label6: TLabel
      Left = 200
      Top = 40
      Width = 57
      Height = 13
      AutoSize = False
      Caption = 'HIS'#24202#21495':'
    end
    object Button2: TButton
      Left = 334
      Top = 2
      Width = 67
      Height = 25
      Caption = #26174#31034#31383#21475
      TabOrder = 0
      OnClick = Button2Click
    end
    object Edit1: TEdit
      Left = 405
      Top = 4
      Width = 41
      Height = 24
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
    object Edit2: TEdit
      Left = 447
      Top = 4
      Width = 90
      Height = 24
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
    end
    object Edit3: TEdit
      Left = 594
      Top = 5
      Width = 63
      Height = 21
      TabOrder = 3
      Text = #36755#20837#22995#21517
    end
    object Edit4: TEdit
      Left = 706
      Top = 5
      Width = 41
      Height = 21
      MaxLength = 2
      TabOrder = 4
      Text = '25'
    end
    object bntConnect: TButton
      Left = 256
      Top = 2
      Width = 73
      Height = 25
      Caption = #36830#25509#26381#21153
      TabOrder = 5
      OnClick = bntConnectClick
    end
    object edtIp: TEdit
      Left = 66
      Top = 5
      Width = 91
      Height = 21
      TabOrder = 6
      Text = '192.168.1.206'
    end
    object edtPort: TEdit
      Left = 211
      Top = 5
      Width = 42
      Height = 21
      TabOrder = 7
      Text = '5000'
    end
    object BtnList: TButton
      Left = 912
      Top = 5
      Width = 93
      Height = 25
      Caption = #21462#30417#25252#21015#34920
      TabOrder = 8
      OnClick = BtnListClick
    end
    object edtCaseNo: TEdit
      Left = 73
      Top = 35
      Width = 109
      Height = 21
      TabOrder = 9
    end
    object edtHisNo: TEdit
      Left = 253
      Top = 35
      Width = 109
      Height = 21
      TabOrder = 10
    end
    object btnSelectNo: TButton
      Left = 375
      Top = 34
      Width = 83
      Height = 25
      Caption = 'HIS2DEV'
      TabOrder = 11
      OnClick = btnSelectNoClick
    end
    object edtSelectCaseNo: TEdit
      Left = 558
      Top = 36
      Width = 98
      Height = 21
      TabOrder = 12
    end
    object radgType: TRadioGroup
      Left = 671
      Top = 30
      Width = 313
      Height = 33
      Columns = 3
      ItemIndex = 2
      Items.Strings = (
        #25353#30417#25252#20202#24202#21495
        #25353'HIS'#24202#21495
        #25353#30149#21382#21495)
      TabOrder = 13
    end
    object radgSex: TRadioGroup
      Left = 763
      Top = 0
      Width = 94
      Height = 30
      Columns = 2
      ItemIndex = 0
      Items.Strings = (
        #30007
        #22899)
      TabOrder = 14
    end
    object btnDev2His: TButton
      Left = 465
      Top = 34
      Width = 82
      Height = 25
      Caption = 'DEV2HIS'
      TabOrder = 15
      OnClick = btnDev2HisClick
    end
  end
end
