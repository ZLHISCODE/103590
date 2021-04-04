VERSION 5.00
Begin VB.UserControl PictureButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   ScaleHeight     =   1770
   ScaleWidth      =   4410
   ToolboxBitmap   =   "PictureButton.ctx":0000
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   3570
      TabIndex        =   0
      Top             =   0
      Width           =   3570
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Left            =   1545
         MouseIcon       =   "PictureButton.ctx":0312
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   525
         Width           =   420
      End
      Begin VB.Image imgFlag 
         Height          =   420
         Left            =   420
         MouseIcon       =   "PictureButton.ctx":061C
         MousePointer    =   99  'Custom
         Top             =   570
         Width           =   450
      End
      Begin VB.Image imgBack 
         Height          =   435
         Left            =   405
         MouseIcon       =   "PictureButton.ctx":0926
         MousePointer    =   99  'Custom
         Top             =   345
         Width           =   1020
      End
   End
End
Attribute VB_Name = "PictureButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const GWL_HWNDPARENT = (-8)

Private Type POINT
    x As Long
    y As Long
End Type

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINT) As Long

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

Private mstdNormal As StdPicture
Private mstdPressed As StdPicture
Private mstdSelect As StdPicture
Private mstdDisable As StdPicture

Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event CommandClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long


    If Button = 1 Then
    
        If mstdPressed Is Nothing Then
            imgBack.Left = 15
            imgBack.Top = 15
            lbl.Left = lbl.Left + 15
            lbl.Top = lbl.Top + 15
        Else
            Set imgBack.Picture = mstdPressed
        End If

    End If
End Sub

Private Sub imgBack_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
        If mstdPressed Is Nothing Then
            imgBack.Left = 0
            imgBack.Top = 0
            lbl.Left = lbl.Left - 15
            lbl.Top = lbl.Top - 15
        Else
            
            Set imgBack.Picture = mstdNormal
            
        End If

        If lbl.Enabled Then RaiseEvent CommandClick
    End If
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgBack_MouseDown Button, Shift, 10, 10
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgBack_MouseUp Button, Shift, 10, 10
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    With picBack
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height
    End With
    
    imgBack.Left = 0
    imgBack.Top = 0
    imgBack.Width = UserControl.Width
    imgBack.Height = UserControl.Height
    
    imgFlag.Left = 60
    imgFlag.Width = 240
    imgFlag.Height = 240
    
    imgFlag.Top = (UserControl.Height - imgFlag.Height) / 2
    
    Select Case mbytTextAligment
    Case 0
        lbl.Left = 240
    Case 1
        lbl.Left = (UserControl.Width - lbl.Width) / 2
    Case 2
        lbl.Left = UserControl.Width - lbl.Width - 240
    End Select
    
    
    lbl.Top = (UserControl.Height - lbl.Height) / 2
    
End Sub

'Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '
'End Sub
'
'Private Sub img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '
'End Sub
'
'Private Sub img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    Call picBack_MouseUp(1, 0, 10, 10)
'End Sub

'Private Sub picBack_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        RaiseEvent CommandClick
'    End If
'End Sub

'Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim i As Long
'
'    'If Button = 1 And m_State <> -1 Then
'    If Button = 1 Then
'        RaisEffect picBack, -1, "", IIf(m_ShowPicture, 180, 30)
'
'        PrintText picBack, lbl.Caption, mbytTextAligment, IIf(m_ShowPicture, 390, 30)
''        For i = 1 To 3
''            Beep
''        Next
'    End If
'End Sub
'
'Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    'If m_State = -1 Then Exit Sub
'
''    If x < 0 Or y < 0 Or x > picBack.Width Or y > picBack.Height Then
''        ReleaseCapture
''        If picBack.Tag = "1" Then
''            RaisEffect picBack, 0, lbl.Caption, IIf(m_ShowPicture, 180, 30)
''            picBack.Tag = ""
''
''        End If
''    Else
''        SetCapture picBack.hWnd
''        If picBack.Tag <> "1" Then
''            RaisEffect picBack, 1, lbl.Caption, IIf(m_ShowPicture, 180, 30)
''            picBack.Tag = "1"
''
''        End If
''    End If
'    RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub
'
'Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    'If Button = 1 And m_State <> -1 Then
'    If Button = 1 Then
'        RaisEffect picBack, 1, "", IIf(m_ShowPicture, 180, 30)
'        PrintText picBack, lbl.Caption, mbytTextAligment, IIf(m_ShowPicture, 390, 30)
'        If picBack.Enabled Then RaiseEvent CommandClick
'    End If
'End Sub
'
'Private Sub picBack_Paint()
'
'    If mblnDrawColor Then Call DrawColorToColor(picBack, picBack.BackColor, &HFFC0C0)
'
'    RaisEffect picBack, 1, "", IIf(m_ShowPicture, 180, 30)
'    PrintText picBack, lbl.Caption, mbytTextAligment, IIf(m_ShowPicture, 390, 30)
'
'    picBack.Tag = ""
'End Sub


Private Sub UserControl_Initialize()
    m_ShowPicture = True
    
    m_DisableColor = &H808080
    m_ForeColor = &HFF0000
    
    mlngBackColor = RGB(255, 255, 255)
        
    mblnDrawColor = True
    mbytTextAligment = 1
            
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'属性:Cpation
Public Property Let Caption(vData As String)
    lbl.Caption = vData
'    UserControl.Width = lbl.Left + lbl.Width + 120
    PropertyChanged "Caption"
    Call UserControl_Resize
End Property

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Picture(vData As ListImage)
    With imgFlag
        .Picture = vData.Picture
        
        .Width = 240
        .Height = 240
    End With
End Property


Public Property Set PictureNormal(vData As StdPicture)
    '正常图片
    Set mstdNormal = vData
    
    Set imgBack.Picture = mstdNormal
    
End Property

Public Property Get PictureNormal() As StdPicture
    '正常图片
    
    Set PictureNormal = mstdNormal
    
End Property

Public Property Set PictureSelect(vData As StdPicture)
    '选中图片
    Set mstdSelect = vData
    
End Property

Public Property Get PictureSelect() As StdPicture
    '选中图片
    
    Set PictureSelect = mstdSelect
    
End Property

Public Property Set PicturePressed(vData As StdPicture)
    '按下图片
    Set mstdPressed = vData
    
End Property

Public Property Get PicturePressed() As StdPicture
    '选中图片
    
    Set PicturePressed = mstdPressed
    
End Property

Public Property Set PictureDisable(vData As StdPicture)
    '禁用图片
    Set mstdDisable = vData
    
End Property

Public Property Get PictureDisable() As StdPicture
    '禁用图片
    
    Set PictureDisable = mstdDisable
    
End Property

Public Property Let ShowPicture(vData As Boolean)
    m_ShowPicture = vData
    
    imgFlag.Visible = vData
    
    PropertyChanged "ShowPicture"
    Call UserControl_Resize
End Property

Public Property Get ShowPicture() As Boolean
    ShowPicture = m_ShowPicture
End Property

'Public Property Let BackColor(vData As OLE_COLOR)
'    PropertyChanged "BackColor"
'End Property
'
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property

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
    
    lbl.ForeColor = vData
    m_ForeColor = vData
    
    PropertyChanged "ForeColor"
    
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let State(ByVal vData As Long)
    m_State = vData
    
    If m_State = 1 Then
        
        If mstdSelect Is Nothing Then Exit Property
        Set imgBack.Picture = mstdSelect
    Else
        Set imgBack.Picture = mstdNormal
    End If
    
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
    
    imgBack.Enabled = vData
    
    lbl.ForeColor = IIf(vData, m_ForeColor, m_DisableColor)
    lbl.Enabled = vData
    
    If vData = False Then
        Set imgBack.Picture = mstdDisable
    Else
        Set imgBack.Picture = mstdNormal
    End If
    
End Property

Public Property Let AutoSize(vData As Boolean)
    mvarAutoSize = vData
    PropertyChanged "AutoSize"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = mvarAutoSize
End Property

Public Property Get Enabled() As Boolean
    Enabled = lbl.Enabled
End Property

Public Property Set Font(vData As StdFont)
    Set lbl.Font = vData
    'Set picBack.Font = vData
    PropertyChanged "Font"
    Call UserControl_Resize
End Property

Public Property Get Font() As StdFont
    Set Font = lbl.Font
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lbl.Caption = PropBag.ReadProperty("Caption", "平面按钮")
    'lbl.BackColor = PropBag.ReadProperty("BackColor", &HE0E0E0)
    lbl.ForeColor = PropBag.ReadProperty("ForeColor", &HFF0000)
    
    lbl.FontName = PropBag.ReadProperty("FontName", "宋体")
    lbl.FontSize = PropBag.ReadProperty("FontSize", 12)
    lbl.FontBold = PropBag.ReadProperty("FontBold", False)
    lbl.FontItalic = PropBag.ReadProperty("FontItalic", False)
    lbl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
    lbl.FontUnderline = PropBag.ReadProperty("FontUnderline", False)

    m_ShowPicture = PropBag.ReadProperty("ShowPicture", True)
    mvarAutoSize = PropBag.ReadProperty("AutoSize", True)
    'lbl.Height = PropBag.ReadProperty("ButtonHeight", 300)
    
    mbytTextAligment = PropBag.ReadProperty("TextAligment", 1)
    
    mblnDrawColor = PropBag.ReadProperty("DrawColor", True)
    
    Set mstdNormal = PropBag.ReadProperty("PictureNormal", Nothing)
    Set mstdSelect = PropBag.ReadProperty("PictureSelect", Nothing)
    Set mstdPressed = PropBag.ReadProperty("PicturePressed", Nothing)
    Set mstdDisable = PropBag.ReadProperty("PictureDisable", Nothing)
    
    Set imgBack.Picture = mstdNormal
    imgFlag.Visible = m_ShowPicture
        
    Call UserControl_Resize
End Sub

Private Sub UserControl_Show()
    Dim lngParentContainer As Long
    Dim lngLoop As Long
    Dim objPoint As POINT
    Dim objPointDraw As POINT
    Dim objDraw As Object
        
    On Error Resume Next
    
    lngParentContainer = GetWindowLong(UserControl.hwnd, GWL_HWNDPARENT)
        
    For lngLoop = 1 To UserControl.ParentControls.Count - 1
        Set objDraw = UserControl.ParentControls(lngLoop)
        If TypeName(objDraw) = "PictureBox" Then
            If objDraw.hwnd = lngParentContainer Then
            
                ClientToScreen objDraw.hwnd, objPointDraw
                ClientToScreen UserControl.hwnd, objPoint
                
                picBack.PaintPicture objDraw.Image, 0, 0, picBack.Width, picBack.Height, (objPoint.x - objPointDraw.x) * 15, (objPoint.y - objPointDraw.y) * 15, UserControl.Width, UserControl.Height
                
                GoTo Over
            End If
        End If
    Next
    
Over:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lbl.Caption, "平面按钮")
    Call PropBag.WriteProperty("ShowPicture", m_ShowPicture, True)
    'Call PropBag.WriteProperty("BackColor", lbl.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("ForeColor", lbl.ForeColor, &HFF0000)
    
    Call PropBag.WriteProperty("FontName", lbl.FontName, "宋体")
    Call PropBag.WriteProperty("FontSize", lbl.FontSize, 12)
    Call PropBag.WriteProperty("FontBold", lbl.FontBold, False)
    Call PropBag.WriteProperty("FontItalic", lbl.FontItalic, False)
    Call PropBag.WriteProperty("FontStrikethru", lbl.FontStrikethru, False)
    Call PropBag.WriteProperty("FontUnderline", lbl.FontUnderline, False)
    
    Call PropBag.WriteProperty("AutoSize", mvarAutoSize, True)
    'Call PropBag.WriteProperty("ButtonHeight", lbl.Height, 300)
    
    Call PropBag.WriteProperty("TextAligment", mbytTextAligment, 1)
    Call PropBag.WriteProperty("DrawColor", mblnDrawColor, True)
    
    Call PropBag.WriteProperty("PictureNormal", mstdNormal, Nothing)
    Call PropBag.WriteProperty("PictureSelect", mstdSelect, Nothing)
    Call PropBag.WriteProperty("PicturePressed", mstdPressed, Nothing)
    Call PropBag.WriteProperty("PictureDisable", mstdDisable, Nothing)
        
End Sub

'Public Property Let ButtonHeight(vData As Single)
'    lbl.Height = vData
'    PropertyChanged "ButtonHeight"
'    Call UserControl_Resize
'End Property
'
'Public Property Get ButtonHeight() As Single
'    ButtonHeight = lbl.Height
'End Property

Public Property Let BorderStyle(ByVal vData As Byte)
    m_BorderStyle = vData
    PropertyChanged "m_BorderStyle"
End Property

Public Property Get BorderStyle() As Byte
    BorderStyle = m_BorderStyle
End Property

Public Property Let TextAligment(ByVal vData As Byte)
    mbytTextAligment = vData
    
    Call UserControl_Resize
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
