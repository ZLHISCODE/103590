VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisReportLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����ֽ�Ų�������
'����
'######################################################################################################################

Public Enum USERCOLOR
    ��ɫ = 16711680
    ��ɫ = 16777215
    ��ɫ = 0
    ����ɫ = 10191921
    ������ɫ = 5880731
    ���񱳾�ɫ = 15790320
    
    ��Ŀ���ⱳ��ɫ = -2147483636
    
    ��ɫ = &HFF&
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E4E7
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
End Enum

Public Type USERRECT
    X0      As Single
    Y0      As Single
    X1      As Single
    Y1      As Single
    R0      As Single                            '�м��
    B0      As Single                            '�߽�
    Page    As Integer                           'ҳ��
End Type

Public Type USERFONT
    Name            As String
    Size            As Single
    Bold            As Boolean
    Italic          As Boolean
    Underline       As Boolean
    Strikethrough   As Boolean
    ForeColor       As Long
    BackColor       As Long
    LineWidth       As Integer
    LineStyle       As Integer
    HAlignment      As Byte                              '�������
    VAlignment      As Byte                              '�������
    WordWarp        As Boolean                           '�Զ�����
    MuliLine        As Boolean                           '�����ı�
    AutoFit         As Boolean                           '�Զ���Ӧ
    Rows            As Integer                           '����
End Type

Public Type USERPAPER
    Width As Single
    Height As Single
    BorderLeft As Single
    BorderRight As Single
    BorderTop As Single
    BorderBottom As Single
    PageHead As Single
    PageFoot As Single
    SpaceTop As Single
    SpaceBottom As Single
    PaperSize As Integer
    Orientation As Byte
End Type

Public Event FormatData()

Public Property Get RECT() As USERRECT
    RECT = gobjRect
End Property

Public Property Let RECT(vData As USERRECT)
    gobjRect = vData
End Property

Public Property Get Font() As USERFONT
    Font = gobjFont
End Property

Public Property Let Font(vData As USERFONT)
    gobjFont = vData
    
    gobjDraw.FontName = gobjFont.Name
    gobjDraw.FontSize = gobjFont.Size
    gobjDraw.FontBold = gobjFont.Bold
    gobjDraw.FontItalic = gobjFont.Italic
    
End Property

Public Property Get Paper() As USERPAPER
    Paper = gobjPaper
End Property

Public Property Let Paper(vData As USERPAPER)
    gobjPaper = vData
End Property

Public Property Get Draw() As Object
    Set Draw = gobjDraw
End Property

Public Function GetTextWidth(ByVal strText As String) As Single
    GetTextWidth = gobjDraw.TextWidth(strText)
End Function

Public Function GetTextHeight(ByVal strText As String) As Single
    GetTextHeight = gobjDraw.TextHeight(strText)
End Function

Public Function GetTwipsX(ByVal sglNumber As Single) As Single
    GetTwipsX = gobjDraw.ScaleX(sglNumber, vbCentimeters, vbTwips)
End Function

Public Function GetTwipsY(ByVal sglNumber As Single) As Single
    GetTwipsY = gobjDraw.ScaleY(sglNumber, vbCentimeters, vbTwips)
End Function

Public Function GetCentimetersX(ByVal sglNumber As Single) As Single
    GetCentimetersX = gobjDraw.ScaleX(sglNumber, vbTwips, vbCentimeters)
End Function

Public Function GetCentimetersY(ByVal sglNumber As Single) As Single
    GetCentimetersY = gobjDraw.ScaleY(sglNumber, vbTwips, vbCentimeters)
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Private Sub Class_Initialize()
    gobjFont.Name = "����"
    gobjFont.Size = 12
    gobjFont.Bold = False
    gobjFont.Underline = False
    gobjFont.Strikethrough = False
    gobjFont.Italic = False
End Sub
