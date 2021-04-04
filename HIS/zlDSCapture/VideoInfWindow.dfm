object frmVideoInf: TfrmVideoInf
  Left = 365
  Top = 151
  BorderStyle = bsDialog
  BorderWidth = 5
  Caption = #35270#39057#23646#24615
  ClientHeight = 369
  ClientWidth = 452
  Color = clBtnFace
  Font.Charset = GB2312_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object RichEdit1: TRichEdit
    Left = 0
    Top = 0
    Width = 452
    Height = 321
    Align = alTop
    Color = clBtnFace
    ReadOnly = True
    TabOrder = 0
  end
  object Button1: TButton
    Left = 376
    Top = 336
    Width = 75
    Height = 25
    Caption = #30830' '#23450'(&S)'
    TabOrder = 1
    OnClick = Button1Click
  end
end
