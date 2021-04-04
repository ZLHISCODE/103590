VERSION 5.00
Begin VB.UserControl CommandEx 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ScaleHeight     =   1275
   ScaleWidth      =   3990
   Begin VB.PictureBox picDown 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   630
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   0
      Top             =   540
      Width           =   750
      Begin VB.Shape shapBack 
         BorderColor     =   &H80000005&
         Height          =   495
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.Image imgDown 
      Height          =   120
      Left            =   1815
      Picture         =   "CommandEx.ctx":0000
      Top             =   750
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "CommandEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum EM_DrawStyle
    DW_Flat = 0  '= 平面
    Dw_SubKen = -1 '= 凹下
    Dw_Heave = 1  '= 凸起
    Dw_Deepen_Subken = -2 '= 深凹下,
    Dw_Deepen_Heave = 2 ' = 深凸起
End Enum
Public Enum gAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum
Public Enum Em_Appearance_Button
    Show_3D = 1     '3D显示
    Show_Flat = 0   '平面
    Show_Flat_Line = 2 ''平面方式，以线型反应
End Enum
'缺省属性值:
Const m_def_Appearance = 1
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'属性变量:
Dim m_Appearance As Em_Appearance_Button
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_Picture As Picture
'事件声明:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

 
Private Sub picDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserControl.Enabled = False Then Exit Sub
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button <> 1 Then Exit Sub
    DawCommandEx Dw_SubKen
End Sub
Private Sub picDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserControl.Enabled = False Then Exit Sub
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Appearance = Show_3D Then Exit Sub
    
    If picDown.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > picDown.Width Or Y > picDown.Height Then
            picDown.Tag = ""
            ReleaseCapture
            Select Case Appearance
            Case Show_Flat_Line '平面方式，以线型反应
                 shapBack.BorderColor = vbWhite
            Case Else
                DawCommandEx DW_Flat
            End Select
        End If
    Else
        picDown.Tag = "In"
        SetCapture picDown.hWnd
        Select Case Appearance
        Case Show_Flat_Line '平面方式，以线型反应
             shapBack.BorderColor = &H80000003
        Case Else
            DawCommandEx Dw_Heave
        End Select
    End If
End Sub
Private Sub picDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserControl.Enabled = False Then Exit Sub
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button <> 1 Then Exit Sub
    If Appearance = Show_3D Then
        DawCommandEx Dw_Heave   '突出按钮
    Else
        Select Case Appearance
        Case Show_Flat_Line '平面方式，以线型反应
             shapBack.BorderColor = vbWhite
        Case Else
            DawCommandEx DW_Flat
        End Select
    End If
    Call ClearTag
End Sub
Private Sub ClearTag()
     'shapBack.BorderColor = &H8000000A
      picDown.Tag = ""
      
End Sub
Private Sub DrawIco()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绘制图标
    '编制:刘兴洪
    '日期:2014-12-19 14:52:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngX As Single, sngY As Single, sngPicW As Single, sngPicH As Single
    Dim sdPic As StdPicture
    
    picDown.BorderStyle = 0
    If m_Picture Is Nothing Then Exit Sub
    If m_Picture = 0 Then Exit Sub
    
    Set sdPic = m_Picture
    With sdPic
        sngPicW = sdPic.Width / 1.766667
        sngPicH = sdPic.Height / 1.766667
        sngX = (picDown.ScaleWidth - sngPicW) / 2
        sngY = (picDown.ScaleHeight - sngPicH) / 2
    End With
    picDown.PaintPicture sdPic, sngX, sngY
End Sub
 

'注意！不要删除或修改下列被注释的行！
'MemberInfo=11,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As OLE_COLOR
      ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Public Sub Refresh()
     
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Call DrawIco
    PropertyChanged "Picture"
End Property

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        shapBack.Top = .ScaleTop
        shapBack.Height = .ScaleHeight
        shapBack.Left = .ScaleLeft
        shapBack.Width = .ScaleWidth
    End With
    Call ReSetFace
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_BackColor = &H8000000F
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    Set m_Picture = imgDown.Picture  'LoadPicture("")
    m_Appearance = m_def_Appearance
    Call DrawIco
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    
    UserControl.Enabled = m_Enabled
    
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    Set m_Picture = PropBag.ReadProperty("Picture", imgDown.Picture)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    ReSetFace
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With UserControl
        picDown.Top = .ScaleTop
        picDown.Left = .ScaleLeft
        picDown.Width = .ScaleWidth
        picDown.Height = .ScaleHeight
    End With
End Sub
Private Sub ReSetFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置面版
    '编制:刘兴洪
    '日期:2018-01-02 17:06:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    shapBack.Visible = m_Appearance = Show_Flat_Line
    picDown.Cls
     If m_Appearance = Show_3D Then
        DawCommandEx Dw_Heave   '突出按钮
     ElseIf m_Appearance = Show_Flat Then
        DawCommandEx DW_Flat
     End If
     Call DrawIco
End Sub




'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Picture", m_Picture, imgDown.Picture)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
End Sub
Public Sub DawCommandEx(Optional intStyle As EM_DrawStyle, _
    Optional strName As String = "", Optional TxtAlignment As gAlignment = 1)
    '功能：将PictureBox模拟成3D平面按钮
    'intStyle=0=平面,-1=凹下,1=凸起,-2=深凹下,2=深凸起
    Dim PicRect As RECT
    Dim lngTmp As Long, intStyleTmp As Integer
    ' intStyle=0=平面,-1=凹下,1=凸起,-2=深凹下,2=深凸起
    intStyleTmp = intStyle
    Call zlRaisEffect(picDown, intStyleTmp, strName, TxtAlignment)
    Call DrawIco
End Sub

Private Sub picDown_Click()
    RaiseEvent Click
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property



'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,1
Public Property Get Appearance() As Em_Appearance_Button
Attribute Appearance.VB_Description = "返回/设置一个对象在运行时是否以 3D 效果显示。"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Em_Appearance_Button)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    Call ReSetFace
End Property

