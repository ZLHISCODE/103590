VERSION 5.00
Begin VB.UserControl ctlButton 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   KeyPreview      =   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   3795
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   90
      MouseIcon       =   "ctlButton.ctx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   855
      ScaleWidth      =   2640
      TabIndex        =   0
      Top             =   60
      Width           =   2640
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "按钮"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   360
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image img 
         Height          =   300
         Left            =   0
         Picture         =   "ctlButton.ctx":030A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   300
      End
   End
End
Attribute VB_Name = "ctlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_ShowPicture As Boolean
Private m_State As Long
Private m_Check As Boolean
Private m_Key As String
Private m_Key2 As String
Private m_DisableColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private mvarAutoSize As Boolean
Private mbytTextAligment As Byte            '0-左对齐;1-居中对齐;2-右对齐
Private mblnDrawColor As Boolean


Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event CommandClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picBack_MouseDown(1, 0, 10, 10)
End Sub

Private Sub img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picBack_MouseMove(0, 0, 10, 10)
End Sub

Private Sub img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picBack_MouseUp(1, 0, 10, 10)
End Sub

Private Sub picBack_GotFocus()
    '
    
End Sub

Private Sub picBack_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RaiseEvent CommandClick
    End If
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    'If Button = 1 And m_State <> -1 Then
    If Button = 1 Then
        RaisEffect picBack, -1, "", IIf(m_ShowPicture, 180, 30)
        
        PrintText picBack, lbl.Caption, mbytTextAligment, IIf(m_ShowPicture, 390, 30)
'        For i = 1 To 3
'            Beep
'        Next
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If m_State = -1 Then Exit Sub
    
'    If x < 0 Or y < 0 Or x > picBack.Width Or y > picBack.Height Then
'        ReleaseCapture
'        If picBack.Tag = "1" Then
'            RaisEffect picBack, 0, lbl.Caption, IIf(m_ShowPicture, 180, 30)
'            picBack.Tag = ""
'
'        End If
'    Else
'        SetCapture picBack.hWnd
'        If picBack.Tag <> "1" Then
'            RaisEffect picBack, 1, lbl.Caption, IIf(m_ShowPicture, 180, 30)
'            picBack.Tag = "1"
'
'        End If
'    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 1 And m_State <> -1 Then
    If Button = 1 Then
        RaisEffect picBack, 1, "", IIf(m_ShowPicture, 180, 30)
        PrintText picBack, lbl.Caption, mbytTextAligment, IIf(m_ShowPicture, 390, 30)
        If picBack.Enabled Then RaiseEvent CommandClick
    End If
End Sub

Private Sub picBack_Paint()

    If mblnDrawColor Then Call DrawColorToColor(picBack, picBack.BackColor, &HFFC0C0)
        
    RaisEffect picBack, 1, "", IIf(m_ShowPicture, 180, 30)
    PrintText picBack, lbl.Caption, mbytTextAligment, IIf(m_ShowPicture, 390, 30)
    
    picBack.Tag = ""
End Sub


Private Sub UserControl_Initialize()
    m_ShowPicture = True
    img.Width = 240
    img.Height = 240
    
    lbl.FontName = "宋体"
    lbl.FontBold = False
    lbl.FontItalic = False
    lbl.FontSize = 12
    lbl.FontStrikethru = False
    lbl.FontUnderline = False

    picBack.FontName = "宋体"
    picBack.FontBold = False
    picBack.FontItalic = False
    picBack.FontSize = 12
    picBack.FontStrikethru = False
    picBack.FontUnderline = False
    
    m_DisableColor = &H808080
    m_ForeColor = &HFF0000
    
    mlngBackColor = RGB(255, 255, 255)
    
    mblnDrawColor = True
    
    mbytTextAligment = 1
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Call picBack_KeyPress(KeyAscii)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lbl.Caption = PropBag.ReadProperty("Caption", "平面按钮")
    picBack.BackColor = PropBag.ReadProperty("BackColor", &HE0E0E0)
    picBack.ForeColor = PropBag.ReadProperty("ForeColor", &HFF0000)
    
    picBack.FontName = PropBag.ReadProperty("FontName", "宋体")
    picBack.FontSize = PropBag.ReadProperty("FontSize", 12)
    picBack.FontBold = PropBag.ReadProperty("FontBold", False)
    picBack.FontItalic = PropBag.ReadProperty("FontItalic", False)
    picBack.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
    picBack.FontUnderline = PropBag.ReadProperty("FontUnderline", False)

    m_ShowPicture = PropBag.ReadProperty("ShowPicture", True)
    mvarAutoSize = PropBag.ReadProperty("AutoSize", True)
    lbl.Height = PropBag.ReadProperty("ButtonHeight", 300)
    
    mbytTextAligment = PropBag.ReadProperty("TextAligment", 1)
    
    mblnDrawColor = PropBag.ReadProperty("DrawColor", True)
    
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
            
    lbl.Top = 60
    
    img.Visible = m_ShowPicture
    img.Left = 60
    img.Top = lbl.Top + (lbl.Height - img.Height) / 2
    
    lbl.Left = IIf(m_ShowPicture, img.Left + img.Width + 60, 60)
                    
    UserControl.Height = lbl.Top + lbl.Height + 60
        
    If mvarAutoSize = True Then
        UserControl.Width = lbl.Left + picBack.TextWidth(lbl.Caption) + 60
    End If
    
    picBack.Left = 0
    picBack.Top = 0
    picBack.Height = UserControl.Height
    picBack.Width = UserControl.Width
End Sub

'属性:Cpation
Public Property Let Caption(vData As String)
    lbl.Caption = vData
    UserControl.Width = lbl.Left + lbl.Width + 120
    PropertyChanged "Caption"
    Call UserControl_Resize
End Property

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Picture(vData As ListImage)
    img.Picture = vData.Picture
    img.Width = 300
    img.Height = 300
End Property

Public Property Let ShowPicture(vData As Boolean)
    m_ShowPicture = vData
    PropertyChanged "ShowPicture"
    Call UserControl_Resize
End Property

Public Property Get ShowPicture() As Boolean
    ShowPicture = m_ShowPicture
End Property

Public Property Let BackColor(vData As OLE_COLOR)
    UserControl.BackColor = vData
    picBack.BackColor = vData
    PropertyChanged "BackColor"
    lbl.BackColor = vData
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = picBack.BackColor
End Property

Public Property Let Key(vData As String)
    m_Key = vData
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key2(vData As String)
    m_Key2 = vData
End Property

Public Property Get Key2() As String
    Key2 = m_Key2
End Property

Public Property Let ForeColor(vData As OLE_COLOR)
    picBack.ForeColor = vData
    m_ForeColor = vData
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let State(ByVal vData As Long)
    m_State = vData
'    Call picBack_Paint
End Property

Public Property Get State() As Long
    State = m_State
End Property

Public Property Let Check(vData As Boolean)
    m_Check = vData
End Property

Public Property Get Check() As Boolean
    Check = m_Check
End Property

Public Property Let Enabled(vData As Boolean)
    picBack.ForeColor = IIf(vData, m_ForeColor, m_DisableColor)
    picBack.Enabled = vData
    Call picBack_Paint
End Property

Public Property Let AutoSize(vData As Boolean)
    mvarAutoSize = vData
    PropertyChanged "AutoSize"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = mvarAutoSize
End Property

Public Property Get Enabled() As Boolean
    Enabled = picBack.Enabled
End Property

Public Property Set Font(vData As StdFont)
    Set lbl.Font = vData
    Set picBack.Font = vData
    PropertyChanged "Font"
    Call UserControl_Resize
End Property

Public Property Get Font() As StdFont
    Set Font = picBack.Font
        
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lbl.Caption, "平面按钮")
    Call PropBag.WriteProperty("Caption", m_ShowPicture, True)
    Call PropBag.WriteProperty("BackColor", picBack.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("ForeColor", picBack.ForeColor, &HFF0000)
    
    Call PropBag.WriteProperty("FontName", picBack.FontName, "宋体")
    Call PropBag.WriteProperty("FontSize", picBack.FontSize, 12)
    Call PropBag.WriteProperty("FontBold", picBack.FontBold, False)
    Call PropBag.WriteProperty("FontItalic", picBack.FontItalic, False)
    Call PropBag.WriteProperty("FontStrikethru", picBack.FontStrikethru, False)
    Call PropBag.WriteProperty("FontUnderline", picBack.FontUnderline, False)
    
    Call PropBag.WriteProperty("AutoSize", mvarAutoSize, True)
    Call PropBag.WriteProperty("ButtonHeight", lbl.Height, 300)
    
    Call PropBag.WriteProperty("TextAligment", mbytTextAligment, 1)
    Call PropBag.WriteProperty("DrawColor", mblnDrawColor, True)
        
End Sub

Public Property Let ButtonHeight(vData As Single)
    lbl.Height = vData
    PropertyChanged "ButtonHeight"
    Call UserControl_Resize
End Property

Public Property Get ButtonHeight() As Single
    ButtonHeight = lbl.Height
End Property

Public Property Let BorderStyle(ByVal vData As Byte)
    m_BorderStyle = vData
    PropertyChanged "m_BorderStyle"
End Property

Public Property Get BorderStyle() As Byte
    BorderStyle = m_BorderStyle
End Property

Public Property Let TextAligment(ByVal vData As Byte)
    mbytTextAligment = vData
    PropertyChanged "m_BorderStyle"
End Property

Public Property Get TextAligment() As Byte
    TextAligment = mbytTextAligment
End Property


Public Property Let DrawColor(ByVal vData As Boolean)
    mblnDrawColor = vData
    PropertyChanged "DrawColor"
End Property

Public Property Get DrawColor() As Boolean
    DrawColor = mblnDrawColor
End Property
