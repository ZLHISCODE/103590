object Form1: TForm1
  Left = 512
  Top = 243
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = #24405#38899#26426
  ClientHeight = 275
  ClientWidth = 360
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 134
    Width = 60
    Height = 13
    Caption = #23384#20648#36335#24452#65306
  end
  object TMCIAudio1: TTMCIAudio
    Left = 16
    Top = 16
    Width = 321
    Height = 100
    ParentColor = False
    TabOrder = 0
    ControlData = {
      5450463009545265636F7264657200044C656674021003546F70021005576964
      746803410106486569676874026409466F726D617454616702010D4269747350
      657253616D706C6507056270733136084368616E6E656C730708616353746572
      656F0A53616D706C65526174650444AC0000094E6F53616D706C657303000407
      5365704374726C0808446576696365496402000E417564696F4C696E65436F6C
      6F720706636C4C696D65084D6178436F6C6F720705636C5265640B53616D706C
      65436F756E7403C8000D447261774672657175656E637903F4010B4275666665
      72436F756E740704627566380D53706C69744368616E6E656C73080954726967
      4C6576656C03800009547269676765726564080D4973436F6D70726573734D70
      33090F4D7033436F6D707265737352617465024005436F6C6F720707636C426C
      61636B0000}
  end
  object butFormatCfg: TButton
    Left = 16
    Top = 168
    Width = 75
    Height = 25
    Caption = #26684#24335#35774#32622'(&C)'
    TabOrder = 1
    OnClick = butFormatCfgClick
  end
  object Button2: TButton
    Left = 312
    Top = 128
    Width = 25
    Height = 25
    Caption = '...'
    TabOrder = 2
    OnClick = Button2Click
  end
  object edtPath: TEdit
    Left = 80
    Top = 131
    Width = 233
    Height = 21
    TabOrder = 3
  end
  object butStart: TButton
    Left = 96
    Top = 168
    Width = 75
    Height = 25
    Caption = #24320#22987#24405#38899'(&S)'
    TabOrder = 4
    OnClick = butStartClick
  end
  object butPause: TButton
    Left = 184
    Top = 168
    Width = 75
    Height = 25
    Caption = #26242#20572#24405#38899'(&P)'
    TabOrder = 5
    OnClick = butPauseClick
  end
  object butStop: TButton
    Left = 264
    Top = 168
    Width = 75
    Height = 25
    Caption = #20572#27490#24405#38899'(&A)'
    TabOrder = 6
    OnClick = butStopClick
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 256
    Width = 360
    Height = 19
    Panels = <
      item
        Text = #20934#22791#23601#32490'...'
        Width = 80
      end
      item
        Text = #24405#21046#38271#24230#65306'0('#31186')'
        Width = 50
      end>
  end
  object RadioButton1: TRadioButton
    Tag = 32
    Left = 16
    Top = 208
    Width = 89
    Height = 17
    Caption = '32kbps'#21387#32553
    TabOrder = 8
    OnClick = RadioButton1Click
  end
  object RadioButton2: TRadioButton
    Tag = 64
    Left = 136
    Top = 208
    Width = 89
    Height = 17
    Caption = '64kbps'#21387#32553
    Checked = True
    TabOrder = 9
    TabStop = True
    OnClick = RadioButton1Click
  end
  object RadioButton3: TRadioButton
    Tag = 128
    Left = 248
    Top = 208
    Width = 89
    Height = 17
    Caption = '128kbps'#21387#32553
    TabOrder = 10
    OnClick = RadioButton1Click
  end
  object RadioButton4: TRadioButton
    Tag = 256
    Left = 136
    Top = 232
    Width = 89
    Height = 17
    Caption = '256kbps'#21387#32553
    TabOrder = 11
    OnClick = RadioButton1Click
  end
  object RadioButton5: TRadioButton
    Tag = 320
    Left = 248
    Top = 232
    Width = 89
    Height = 17
    Caption = '320kbps'#21387#32553
    TabOrder = 12
    OnClick = RadioButton1Click
  end
  object RadioButton6: TRadioButton
    Tag = 192
    Left = 16
    Top = 232
    Width = 89
    Height = 17
    Caption = '192kbps'#21387#32553
    TabOrder = 13
    OnClick = RadioButton1Click
  end
  object Timer1: TTimer
    OnTimer = Timer1Timer
    Left = 320
    Top = 48
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = '*.mp3'
    Filter = '(*.mp3)|*.mp3|(*.wav)|*.wav|(*.*)|*.*'
    Left = 264
    Top = 48
  end
end
