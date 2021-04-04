object Form1: TForm1
  Left = 534
  Top = 144
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  BorderWidth = 8
  Caption = 'pacs'#25509#21475#35843#29992
  ClientHeight = 597
  ClientWidth = 657
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clBlack
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 12
  object PageControl1: TPageControl
    Left = 0
    Top = 105
    Width = 657
    Height = 492
    ActivePage = TabSheet3
    Align = alClient
    TabOrder = 0
    object TabSheet3: TTabSheet
      Caption = #33719#21462#26816#26597#37096#20301
      ImageIndex = 2
      object lvPacsStudyBodyPart: TListView
        Left = 8
        Top = 8
        Width = 633
        Height = 313
        Columns = <>
        ReadOnly = True
        RowSelect = True
        TabOrder = 0
        ViewStyle = vsReport
        OnClick = lvRequestInfClick
      end
      object Button3: TButton
        Left = 8
        Top = 328
        Width = 105
        Height = 25
        Caption = #33719#21462#26816#26597#37096#20301'(&S)'
        TabOrder = 1
        OnClick = Button3Click
      end
    end
    object TabSheet6: TTabSheet
      Caption = #33719#21462#31185#23460#20449#24687
      ImageIndex = 5
      object lvDeptItemInf: TListView
        Left = 8
        Top = 8
        Width = 633
        Height = 313
        Columns = <>
        ReadOnly = True
        RowSelect = True
        TabOrder = 0
        ViewStyle = vsReport
        OnClick = lvRequestInfClick
      end
      object Button7: TButton
        Left = 8
        Top = 328
        Width = 129
        Height = 25
        Caption = #33719#21462#26816#26597#31185#23460#20449#24687'(&D)'
        TabOrder = 1
        OnClick = Button7Click
      end
    end
    object TabSheet1: TTabSheet
      Caption = #33719#21462#30003#35831#21333
      object Label1: TLabel
        Left = 8
        Top = 176
        Width = 84
        Height = 12
        Caption = #35786#30103#39033#30446#26126#32454#65306
      end
      object Label2: TLabel
        Left = 192
        Top = 401
        Width = 48
        Height = 12
        Caption = #26597#35810#20540#65306
      end
      object Label9: TLabel
        Left = 8
        Top = 400
        Width = 60
        Height = 12
        Caption = #26597#35810#31867#22411#65306
      end
      object Label24: TLabel
        Left = 8
        Top = 288
        Width = 60
        Height = 12
        Caption = #36153#29992#26126#32454#65306
      end
      object Shape1: TShape
        Left = 8
        Top = 424
        Width = 633
        Height = 3
        Brush.Color = clBlack
      end
      object Label31: TLabel
        Left = 8
        Top = 436
        Width = 60
        Height = 12
        Caption = #24320#22987#26085#26399#65306
      end
      object Label32: TLabel
        Left = 181
        Top = 436
        Width = 60
        Height = 12
        Caption = #24320#22987#26085#26399#65306
      end
      object lvRequestInf: TListView
        Left = 8
        Top = 8
        Width = 633
        Height = 161
        Columns = <>
        ReadOnly = True
        RowSelect = True
        TabOrder = 0
        ViewStyle = vsReport
        OnClick = lvRequestInfClick
      end
      object Button1: TButton
        Left = 400
        Top = 393
        Width = 89
        Height = 25
        Caption = #26597#35810#30003#35831#20449#24687
        TabOrder = 1
        OnClick = Button1Click
      end
      object Edit1: TEdit
        Left = 240
        Top = 397
        Width = 145
        Height = 20
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        TabOrder = 2
      end
      object cbxQueryType: TComboBox
        Left = 64
        Top = 396
        Width = 105
        Height = 20
        Style = csDropDownList
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        ItemHeight = 12
        TabOrder = 3
        Items.Strings = (
          #30149#20154'id'
          #20303#38498#21495
          #38376#35786#21495
          #23601#35786#21345#21495
          #36523#20221#35777#21495
          #20581#24247#21495
          #22995#21517
          #21307#22065'ID')
      end
      object lvAdviceFees: TListView
        Left = 8
        Top = 304
        Width = 633
        Height = 81
        Columns = <>
        ReadOnly = True
        RowSelect = True
        TabOrder = 4
        ViewStyle = vsReport
      end
      object lvAdviceItems: TListView
        Left = 8
        Top = 192
        Width = 633
        Height = 89
        Columns = <>
        ReadOnly = True
        RowSelect = True
        TabOrder = 5
        ViewStyle = vsReport
      end
      object DateTimePicker1: TDateTimePicker
        Left = 64
        Top = 432
        Width = 105
        Height = 20
        Date = 40757.549892638890000000
        Time = 40757.549892638890000000
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        TabOrder = 6
      end
      object DateTimePicker2: TDateTimePicker
        Left = 240
        Top = 432
        Width = 145
        Height = 20
        Date = 40757.549892638890000000
        Time = 40757.549892638890000000
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        TabOrder = 7
      end
      object Button13: TButton
        Left = 400
        Top = 432
        Width = 89
        Height = 25
        Caption = #26597#35810#30003#35831#20449#24687
        TabOrder = 8
        OnClick = Button13Click
      end
    end
    object TabSheet2: TTabSheet
      Caption = #33719#21462#30149#20154#20449#24687
      ImageIndex = 1
      object Label8: TLabel
        Left = 8
        Top = 336
        Width = 60
        Height = 12
        Caption = #26597#35810#31867#22411#65306
      end
      object Label10: TLabel
        Left = 192
        Top = 337
        Width = 48
        Height = 12
        Caption = #26597#35810#20540#65306
      end
      object lvPatientInf: TListView
        Left = 8
        Top = 8
        Width = 633
        Height = 313
        Columns = <>
        ReadOnly = True
        RowSelect = True
        TabOrder = 0
        ViewStyle = vsReport
        OnClick = lvRequestInfClick
      end
      object butQueryPatientInf: TButton
        Left = 424
        Top = 329
        Width = 105
        Height = 25
        Caption = #26597#35810#30149#20154#20449#24687'(&P)'
        TabOrder = 1
        OnClick = butQueryPatientInfClick
      end
      object cbxPatientQueryType: TComboBox
        Left = 64
        Top = 332
        Width = 105
        Height = 20
        Style = csDropDownList
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        ItemHeight = 12
        ItemIndex = 4
        TabOrder = 2
        Text = #36523#20221#35777#21495
        Items.Strings = (
          #30149#20154'id'
          #20303#38498#21495
          #38376#35786#21495
          #23601#35786#21345#21495
          #36523#20221#35777#21495
          #20581#24247#21495
          #22995#21517)
      end
      object edtPatientValue: TEdit
        Left = 240
        Top = 333
        Width = 145
        Height = 20
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        TabOrder = 3
        Text = '500383198803212226'
      end
    end
    object TabSheet5: TTabSheet
      Caption = #30003#35831#25509#25910
      ImageIndex = 4
      object GroupBox5: TGroupBox
        Left = 8
        Top = 8
        Width = 633
        Height = 113
        TabOrder = 0
        object Label15: TLabel
          Left = 16
          Top = 20
          Width = 60
          Height = 12
          Caption = #25191#34892#31185#23460#65306
        end
        object Label16: TLabel
          Left = 224
          Top = 20
          Width = 60
          Height = 12
          Caption = #26816' '#26597' '#21495#65306
        end
        object Label17: TLabel
          Left = 440
          Top = 20
          Width = 60
          Height = 12
          Caption = #26816#26597#35774#22791#65306
        end
        object Label18: TLabel
          Left = 16
          Top = 52
          Width = 60
          Height = 12
          Caption = #36523'    '#39640#65306
        end
        object Label19: TLabel
          Left = 224
          Top = 52
          Width = 60
          Height = 12
          Caption = #20307'    '#37325#65306
        end
        object Label20: TLabel
          Left = 440
          Top = 52
          Width = 60
          Height = 12
          Caption = #26816#26597#21307#29983#65306
        end
        object Label21: TLabel
          Left = 16
          Top = 83
          Width = 60
          Height = 12
          Caption = #25191#34892#26085#26399#65306
        end
        object Label22: TLabel
          Left = 224
          Top = 84
          Width = 60
          Height = 12
          Caption = #25191#34892#35828#26126#65306
        end
        object edtExeRoom: TEdit
          Left = 72
          Top = 16
          Width = 121
          Height = 20
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 0
          Text = 'CT'#23460
        end
        object edtStudyNo: TEdit
          Left = 280
          Top = 16
          Width = 121
          Height = 20
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 1
          Text = '0001'
        end
        object edtDevice: TEdit
          Left = 496
          Top = 16
          Width = 121
          Height = 20
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 2
          Text = 'Philips CT'
        end
        object edtHeight: TEdit
          Left = 72
          Top = 48
          Width = 121
          Height = 20
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 3
          Text = '173'
        end
        object edtWeight: TEdit
          Left = 280
          Top = 48
          Width = 121
          Height = 20
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 4
          Text = '72'
        end
        object edtStudyDoctor: TEdit
          Left = 496
          Top = 48
          Width = 121
          Height = 20
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 5
          Text = #26446'XXX'
        end
        object dtpRequestDate: TDateTimePicker
          Left = 72
          Top = 80
          Width = 121
          Height = 20
          Date = 40535.477761226850000000
          Time = 40535.477761226850000000
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 6
        end
        object edtExeDes: TEdit
          Left = 280
          Top = 80
          Width = 337
          Height = 20
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 7
        end
      end
      object GroupBox6: TGroupBox
        Left = 8
        Top = 128
        Width = 633
        Height = 65
        Caption = #25104#22871#25191#34892#21307#22065
        TabOrder = 1
        object Label23: TLabel
          Left = 32
          Top = 30
          Width = 48
          Height = 12
          Caption = #21307#22065'ID'#65306
        end
        object edtRecevieAdviceId: TEdit
          Left = 80
          Top = 26
          Width = 121
          Height = 20
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 0
        end
        object butRecevieAdvice: TButton
          Left = 312
          Top = 24
          Width = 75
          Height = 25
          Caption = #25509#25910#30003#35831'(&R)'
          TabOrder = 1
          OnClick = butRecevieAdviceClick
        end
        object Button8: TButton
          Left = 392
          Top = 24
          Width = 75
          Height = 25
          Caption = #20462#25913#30003#35831'(&M)'
          TabOrder = 2
          OnClick = Button8Click
        end
        object Button9: TButton
          Left = 472
          Top = 24
          Width = 75
          Height = 25
          Caption = #25764#38144#30003#35831'(&C)'
          TabOrder = 3
          OnClick = Button9Click
        end
      end
      object GroupBox7: TGroupBox
        Left = 8
        Top = 208
        Width = 633
        Height = 65
        Caption = #20998#37096#20301#25191#34892#21307#22065
        TabOrder = 2
        object Label25: TLabel
          Left = 32
          Top = 30
          Width = 48
          Height = 12
          Caption = #21307#22065'ID'#65306
        end
        object edtReceiveAdviceIDOne: TEdit
          Left = 80
          Top = 26
          Width = 121
          Height = 20
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 0
        end
        object btnReceiveRequestOne: TButton
          Left = 312
          Top = 24
          Width = 75
          Height = 25
          Caption = #25509#25910#30003#35831'(&R)'
          TabOrder = 1
          OnClick = butRecevieAdviceClick
        end
        object btnModifyReqestOne: TButton
          Left = 392
          Top = 24
          Width = 75
          Height = 25
          Caption = #20462#25913#30003#35831'(&M)'
          TabOrder = 2
          OnClick = Button8Click
        end
        object btnCancelRequestOne: TButton
          Left = 472
          Top = 24
          Width = 75
          Height = 25
          Caption = #25764#38144#30003#35831'(&C)'
          TabOrder = 3
          OnClick = Button9Click
        end
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'PACS'#25253#21578#20445#23384
      ImageIndex = 3
      object Label13: TLabel
        Left = 8
        Top = 400
        Width = 48
        Height = 12
        Caption = #21307#22065'ID'#65306
      end
      object Label14: TLabel
        Left = 184
        Top = 400
        Width = 60
        Height = 12
        Caption = #25253#21578#21307#29983#65306
      end
      object GroupBox2: TGroupBox
        Left = 8
        Top = 8
        Width = 633
        Height = 153
        Caption = #35786#26029#20449#24687
        TabOrder = 0
        object Label11: TLabel
          Left = 16
          Top = 24
          Width = 60
          Height = 12
          Caption = #26816#26597#25152#35265#65306
        end
        object Label12: TLabel
          Left = 16
          Top = 88
          Width = 60
          Height = 12
          Caption = #35786#26029#24847#35265#65306
        end
        object memStudyView: TMemo
          Left = 72
          Top = 24
          Width = 553
          Height = 57
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 0
        end
        object memAdvice: TMemo
          Left = 72
          Top = 88
          Width = 553
          Height = 57
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 1
        end
      end
      object GroupBox3: TGroupBox
        Left = 8
        Top = 168
        Width = 633
        Height = 105
        Caption = #25253#21578#22270#20687
        TabOrder = 1
        object memReportImage: TMemo
          Left = 16
          Top = 24
          Width = 521
          Height = 65
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 0
        end
        object Button5: TButton
          Left = 544
          Top = 24
          Width = 83
          Height = 65
          Caption = #28155#21152#22270#20687'(&I)'
          TabOrder = 1
          OnClick = Button5Click
        end
      end
      object GroupBox4: TGroupBox
        Left = 8
        Top = 280
        Width = 633
        Height = 105
        Caption = #25253#21578#38468#20214
        TabOrder = 2
        object memAffix: TMemo
          Left = 16
          Top = 24
          Width = 521
          Height = 65
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 0
        end
        object Button6: TButton
          Left = 544
          Top = 24
          Width = 83
          Height = 65
          Caption = #28155#21152#38468#20214'(&F)'
          TabOrder = 1
          OnClick = Button6Click
        end
      end
      object Button4: TButton
        Left = 384
        Top = 394
        Width = 75
        Height = 25
        Caption = #20445#23384#25253#21578'(&S)'
        TabOrder = 3
        OnClick = Button4Click
      end
      object edtAdviceId: TEdit
        Left = 54
        Top = 397
        Width = 121
        Height = 20
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        TabOrder = 4
      end
      object edtReportDoctor: TEdit
        Left = 243
        Top = 396
        Width = 121
        Height = 20
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        TabOrder = 5
        Text = #24352'XXX'
      end
      object btnDeleteReport: TButton
        Left = 472
        Top = 394
        Width = 75
        Height = 25
        Caption = #21024#38500#25253#21578'(&D)'
        TabOrder = 6
        OnClick = btnDeleteReportClick
      end
    end
    object TabSheet7: TTabSheet
      Caption = #24515#30005#25253#21578#20445#23384
      ImageIndex = 6
      object Label28: TLabel
        Left = 8
        Top = 400
        Width = 48
        Height = 12
        Caption = #21307#22065'ID'#65306
      end
      object Label29: TLabel
        Left = 184
        Top = 400
        Width = 60
        Height = 12
        Caption = #25253#21578#21307#29983#65306
      end
      object Label30: TLabel
        Left = 16
        Top = 320
        Width = 60
        Height = 12
        Caption = #25253#21578#26631#39064#65306
      end
      object GroupBox8: TGroupBox
        Left = 8
        Top = 8
        Width = 633
        Height = 153
        Caption = #35786#26029#20449#24687
        TabOrder = 0
        object Label26: TLabel
          Left = 16
          Top = 24
          Width = 60
          Height = 12
          Caption = #26816#26597#25152#35265#65306
        end
        object Label27: TLabel
          Left = 16
          Top = 88
          Width = 60
          Height = 12
          Caption = #35786#26029#24847#35265#65306
        end
        object memECGResult: TMemo
          Left = 72
          Top = 24
          Width = 553
          Height = 57
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 0
        end
        object memECGAdvice: TMemo
          Left = 72
          Top = 88
          Width = 553
          Height = 57
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 1
        end
      end
      object GroupBox9: TGroupBox
        Left = 8
        Top = 168
        Width = 633
        Height = 105
        Caption = #25253#21578#22270#20687
        TabOrder = 1
        object memECGImage: TMemo
          Left = 16
          Top = 24
          Width = 521
          Height = 65
          ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
          TabOrder = 0
        end
        object Button10: TButton
          Left = 544
          Top = 24
          Width = 83
          Height = 65
          Caption = #28155#21152#22270#20687'(&I)'
          TabOrder = 1
          OnClick = Button10Click
        end
      end
      object edtECGAdviceId: TEdit
        Left = 54
        Top = 397
        Width = 121
        Height = 20
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        TabOrder = 2
      end
      object edtECGReport: TEdit
        Left = 243
        Top = 396
        Width = 121
        Height = 20
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        TabOrder = 3
        Text = #24352'XXX'
      end
      object Button11: TButton
        Left = 384
        Top = 394
        Width = 75
        Height = 25
        Caption = #20445#23384#25253#21578'(&S)'
        TabOrder = 4
        OnClick = Button11Click
      end
      object Button12: TButton
        Left = 472
        Top = 394
        Width = 75
        Height = 25
        Caption = #21024#38500#25253#21578'(&D)'
        TabOrder = 5
        OnClick = Button12Click
      end
      object edtECGName: TEdit
        Left = 80
        Top = 320
        Width = 289
        Height = 20
        ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
        TabOrder = 6
        Text = #20013#32852#24515#30005#26816#26597#25253#21578
      end
    end
  end
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 657
    Height = 105
    Align = alTop
    Caption = #21021#22987#21270#25968#25454#24211#36830#25509
    Color = clBtnFace
    ParentColor = False
    TabOrder = 1
    object Label3: TLabel
      Left = 8
      Top = 27
      Width = 96
      Height = 12
      Caption = 'oracle'#23454#20363#21517#31216#65306
    end
    object Label4: TLabel
      Left = 8
      Top = 56
      Width = 96
      Height = 12
      Caption = 'oracle'#29992#25143#21517#31216#65306
    end
    object Label5: TLabel
      Left = 248
      Top = 24
      Width = 48
      Height = 12
      Caption = #31995#32479#21495#65306
    end
    object Label6: TLabel
      Left = 440
      Top = 24
      Width = 84
      Height = 12
      Caption = #25968#25454#24211#25152#26377#32773#65306
    end
    object Label7: TLabel
      Left = 248
      Top = 56
      Width = 72
      Height = 12
      Caption = 'oracle'#23494#30721#65306
    end
    object Label33: TLabel
      Left = 8
      Top = 80
      Width = 96
      Height = 12
      Caption = '    '#24403#21069#37096#38376'ID'#65306
    end
    object Label34: TLabel
      Left = 229
      Top = 83
      Width = 225
      Height = 12
      Caption = #35831#22312#37096#38376#26694#20013#24405#20837#24403#21069#31185#23460#23545#24212#30340#37096#38376'ID'
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = #23435#20307
      Font.Style = []
      ParentFont = False
    end
    object edtOracleInstanceName: TEdit
      Left = 104
      Top = 24
      Width = 121
      Height = 20
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      TabOrder = 0
    end
    object edtOracleUserName: TEdit
      Left = 104
      Top = 56
      Width = 121
      Height = 20
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      TabOrder = 1
      Text = 'zlhis'
    end
    object edtSysNum: TEdit
      Left = 296
      Top = 24
      Width = 121
      Height = 20
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      TabOrder = 4
      Text = '100'
    end
    object edtDbOwner: TEdit
      Left = 520
      Top = 24
      Width = 121
      Height = 20
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      TabOrder = 5
      Text = 'zlhis'
    end
    object edtOraclePwd: TEdit
      Left = 328
      Top = 56
      Width = 121
      Height = 20
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      TabOrder = 2
    end
    object Button2: TButton
      Left = 552
      Top = 56
      Width = 89
      Height = 25
      Caption = #36830#25509#25968#25454#24211'(&C)'
      TabOrder = 3
      OnClick = Button2Click
    end
    object edtPartId: TEdit
      Left = 104
      Top = 80
      Width = 121
      Height = 20
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      TabOrder = 6
    end
  end
  object OpenDialog1: TOpenDialog
    Options = [ofHideReadOnly, ofAllowMultiSelect, ofEnableSizing]
    Left = 244
    Top = 392
  end
end
