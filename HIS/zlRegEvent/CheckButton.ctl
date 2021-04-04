VERSION 5.00
Begin VB.UserControl CheckButton 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   ScaleHeight     =   630
   ScaleWidth      =   1005
   Begin VB.Label lblTittle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��һ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   360
   End
   Begin VB.Shape shpLine 
      BackColor       =   &H8000000F&
      BorderStyle     =   6  'Inside Solid
      Height          =   555
      Left            =   15
      Top             =   30
      Width           =   945
   End
End
Attribute VB_Name = "CheckButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ȱʡ����ֵ:
Const m_def_BackColor = &H8000000F
Const m_def_BackColorSeling = &H80000003
Const m_def_BackColorSeled = &H8000000B

Const m_def_Enabled = True
Const m_def_BorderStyle = 0
Const m_def_Value = False
'���Ա���:
Dim m_DefaultFont As Font
Dim m_BackColor As OLE_COLOR
Dim m_FontSeled As Font
Dim m_FontSeling As Font
Dim m_BackColorSeling As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_BorderStyle As Integer
Dim m_Value As Boolean
'�¼�����:
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "���û���ӵ�н���Ķ����ϰ��������ʱ������"
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "���û����º��ͷ� ANSI ��ʱ������"
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "���û���ӵ�н���Ķ������ͷż�ʱ������"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "���û���ӵ�н���Ķ����ϰ�����갴ťʱ������"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "���û��ƶ����ʱ������"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "���û���ӵ�н���Ķ������ͷ���귢����"
Event Click()

Private Sub lblTittle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblTittle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblTittle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Or m_Enabled = False Then Exit Sub
    '0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
'    Call PicShowFlat(-1)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled = False Then Exit Sub
    If Tag = "In" Then
        If X < 0 Or Y < 0 Or X > Width Or Y > Height Then
            Tag = "": ReleaseCapture
'            UserControl.BackColor = IIf(m_Value, shpLine.BackColor, m_BackColor)
            Set lblTittle.Font = IIf(m_Value, FontSeled, DefaultFont)
            shpLine.Visible = m_Value
        End If
    Else
        Tag = "In"
        SetCapture Hwnd
        shpLine.Visible = True
'        UserControl.BackColor = m_BackColorSeling
        Set lblTittle.Font = FontSeling
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Or m_Enabled = False Then Exit Sub
    Tag = ""
    m_Value = Not m_Value
    UserControl.BackColor = IIf(m_Value, shpLine.BackColor, m_BackColor)
    Set lblTittle.Font = IIf(m_Value, FontSeled, DefaultFont)
    shpLine.Visible = m_Value
    RaiseEvent Click
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����/����һ��ֵ������һ�������Ƿ���Ӧ�û������¼���"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "����/���ö���ı߿���ʽ��"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,False
Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    UserControl.BackColor = IIf(m_Value, shpLine.BackColor, m_BackColor)
    Set lblTittle.Font = IIf(m_Value, FontSeled, DefaultFont)
    PropertyChanged "Value"
    shpLine.Visible = m_Value
End Property

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_BorderStyle = m_def_BorderStyle
    m_Value = m_def_Value
    m_BackColorSeling = m_def_BackColorSeling
    m_BackColor = m_def_BackColor
    
    BackColorSeled = m_def_BackColorSeled
    Set m_FontSeled = Ambient.Font
    Set m_FontSeling = Ambient.Font
    Set m_DefaultFont = Ambient.Font
    shpLine.Visible = m_Value
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    shpLine.BorderColor = PropBag.ReadProperty("BorderColorSeled", -2147483640)
    shpLine.BackColor = PropBag.ReadProperty("BackColorSeled", &H80000005)
    m_BackColorSeling = PropBag.ReadProperty("BackColorSeling", m_def_BackColorSeling)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    
    Set m_FontSeled = PropBag.ReadProperty("FontSeled", Ambient.Font)
    Set m_FontSeling = PropBag.ReadProperty("FontSeling", Ambient.Font)
    Set m_DefaultFont = PropBag.ReadProperty("DefaultFont", Ambient.Font)
    
    shpLine.Visible = m_Value

    lblTittle.Caption = PropBag.ReadProperty("Caption", "��һ")
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With lblTittle
        .Left = (ScaleWidth - .Width) \ 2
        .Top = (ScaleHeight - .Height) \ 2
    End With
    With shpLine
        .Left = ScaleLeft
        .Top = ScaleTop
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
    
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("BorderColorSeled", shpLine.BorderColor, -2147483640)
    Call PropBag.WriteProperty("BackColorSeled", shpLine.BackColor, &H8000000F)
    Call PropBag.WriteProperty("FillStyleSeled", shpLine.FillStyle, 1)
    Call PropBag.WriteProperty("BorderStyleSeled", shpLine.BackStyle, 1)
    Call PropBag.WriteProperty("BackColorSeling", m_BackColorSeling, m_def_BackColorSeling)
    Call PropBag.WriteProperty("FontSeled", m_FontSeled, Ambient.Font)
    Call PropBag.WriteProperty("FontSeling", m_FontSeling, Ambient.Font)
    Call PropBag.WriteProperty("DefaultFont", m_DefaultFont, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Caption", lblTittle.Caption, "��һ")
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=shpLine,shpLine,-1,BorderColor
Public Property Get BorderColorSeled() As OLE_COLOR
Attribute BorderColorSeled.VB_Description = "����/���ö���ı߿���ɫ��"
    BorderColorSeled = shpLine.BorderColor
End Property

Public Property Let BorderColorSeled(ByVal New_BorderColorSeled As OLE_COLOR)
    shpLine.BorderColor = New_BorderColorSeled
    PropertyChanged "BorderColorSeled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=shpLine,shpLine,-1,BackColor
Public Property Get BackColorSeled() As OLE_COLOR
Attribute BackColorSeled.VB_Description = "����/���ö������ı���ͼ�εı���ɫ��"
    BackColorSeled = shpLine.BackColor
End Property

Public Property Let BackColorSeled(ByVal New_BackColorSeled As OLE_COLOR)
    shpLine.BackColor = New_BackColorSeled
    PropertyChanged "BackColorSeled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=shpLine,shpLine,-1,FillStyle
Public Property Get FillStyleSeled() As Integer
Attribute FillStyleSeled.VB_Description = "����/����һ�� shape �ؼ��������ʽ��"
    FillStyleSeled = shpLine.FillStyle
End Property

Public Property Let FillStyleSeled(ByVal New_FillStyleSeled As Integer)
    shpLine.FillStyle = New_FillStyleSeled
    PropertyChanged "FillStyleSeled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=shpLine,shpLine,-1,BackStyle
Public Property Get BorderStyleSeled() As Integer
Attribute BorderStyleSeled.VB_Description = "ָ�� Label �� Shape �ı�����ʽ��͸���Ļ��ǲ�͸���ġ�"
    BorderStyleSeled = shpLine.BackStyle
End Property

Public Property Let BorderStyleSeled(ByVal New_BorderStyleSeled As Integer)
    shpLine.BackStyle = New_BorderStyleSeled
    PropertyChanged "BorderStyleSeled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=10,0,0,0
Public Property Get BackColorSeling() As OLE_COLOR
    BackColorSeling = m_BackColorSeling
End Property

Public Property Let BackColorSeling(ByVal New_BackColorSeling As OLE_COLOR)
    m_BackColorSeling = New_BackColorSeling
    PropertyChanged "BackColorSeling"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=6,0,0,0
Public Property Get FontSeled() As Font
    Set FontSeled = m_FontSeled
End Property

Public Property Set FontSeled(ByVal New_FontSeled As Font)
    Set m_FontSeled = New_FontSeled
    PropertyChanged "FontSeled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=6,0,0,0
Public Property Get FontSeling() As Font
    Set FontSeling = m_FontSeling
End Property

Public Property Set FontSeling(ByVal New_FontSeling As Font)
    Set m_FontSeling = New_FontSeling
    PropertyChanged "FontSeling"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=6,0,0,0
Public Property Get DefaultFont() As Font
    Set DefaultFont = m_DefaultFont
End Property

Public Property Set DefaultFont(ByVal New_DefaultFont As Font)
    Set m_DefaultFont = New_DefaultFont
    PropertyChanged "DefaultFont"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=10,0,0,&H8000000F&
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=lblTittle,lblTittle,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "����/���ö���ı������л�ͼ��������ı���"
    Caption = lblTittle.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblTittle.Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Resize
End Property

