VERSION 5.00
Begin VB.UserControl TextEx 
   BackColor       =   &H80000005&
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   ScaleHeight     =   855
   ScaleWidth      =   4095
   Begin VB.TextBox txtExpend 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   300
      TabIndex        =   0
      Top             =   195
      Width           =   2070
   End
   Begin VB.Shape shapBack 
      BorderColor     =   &H80000003&
      Height          =   825
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Line lneUnder 
      Visible         =   0   'False
      X1              =   2940
      X2              =   3360
      Y1              =   735
      Y2              =   735
   End
End
Attribute VB_Name = "TextEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum Em_InputMode
    InPut_Chars = 0 '字符
    InPut_Numbers = 1   '负数
    Input_Moneys = 2    '金额(不包含负数)
    Input_NegativeMoneys = 4 '含负数金额
End Enum
Public Enum Em_Appearance
    Show_3D = 1     '3D显示
    Show_Flat = 0   '平面
End Enum

'事件声明:
Event Change() 'MappingInfo=txtExpend,txtExpend,-1,Change
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtExpend,txtExpend,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtExpend,txtExpend,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtExpend,txtExpend,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtExpend,txtExpend,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtExpend,txtExpend,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtExpend,txtExpend,-1,MouseUp
'缺省属性值:
Const m_def_GotFocusSelAll = True
Const m_def_Appearance = Em_Appearance.Show_3D
Const m_def_BorderStyle = 0
Const m_def_InputMode = Em_InputMode.InPut_Chars
Const m_def_MouseRightMenus = True
'属性变量:
Dim m_GotFocusSelAll As Boolean
Dim m_Appearance As Em_Appearance
Dim m_Showline As Integer
Dim m_BorderStyle As Integer
Dim m_InputMode As Em_InputMode
Dim m_MouseRightMenus As Boolean


Private Sub txtExpend_GotFocus()
    TxtSelAll txtExpend
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    Set txtExpend.Font = UserControl.Font
    m_MouseRightMenus = m_def_MouseRightMenus
    m_InputMode = m_def_InputMode
    m_BorderStyle = m_def_BorderStyle
    m_Appearance = m_def_Appearance
    m_Showline = 0
    m_GotFocusSelAll = m_def_GotFocusSelAll
    Call ReSetFace
    
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtExpend.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtExpend.BackColor() = UserControl.BackColor
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    txtExpend.Enabled = UserControl.Enabled
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    txtExpend.Alignment = PropBag.ReadProperty("Alignment", 0)
    shapBack.BorderColor = PropBag.ReadProperty("BorderColor", -2147483638)
    lneUnder.BorderColor = shapBack.BorderColor
    txtExpend.IMEMode = PropBag.ReadProperty("IMEMode", 0)
    txtExpend.Locked = PropBag.ReadProperty("Locked", False)
    txtExpend.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtExpend.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtExpend.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtExpend.SelText = PropBag.ReadProperty("SelText", "")
    m_MouseRightMenus = PropBag.ReadProperty("MouseRightMenus", m_def_MouseRightMenus)
    m_InputMode = PropBag.ReadProperty("InputMode", m_def_InputMode)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_Showline = PropBag.ReadProperty("Showline", 0)
    txtExpend.Text = PropBag.ReadProperty("Text", "Text1")
    Set txtExpend.Font = UserControl.Font
    
    m_GotFocusSelAll = PropBag.ReadProperty("GotFocusSelAll", m_def_GotFocusSelAll)
    txtExpend.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    
    Call ReSetFace
End Sub

Private Sub UserControl_Resize()
    Dim sngTop As Single, lngStep As Long
    
    Err = 0: On Error Resume Next
    With shapBack
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    lneUnder.Y1 = ScaleHeight - 15
    lneUnder.Y2 = ScaleHeight - 15
    lneUnder.X1 = ScaleLeft
    lneUnder.X2 = ScaleWidth
    Call SetTxtHeight
    lngStep = IIf(m_Appearance = Show_3D, 0, 10)
    sngTop = (ScaleHeight - txtExpend.Height) \ 2 - lngStep
    If sngTop < 0 Then sngTop = ScaleTop
    sngTop = sngTop + lngStep
    
    With UserControl
        txtExpend.Left = ScaleLeft + lngStep
        txtExpend.Top = sngTop
        txtExpend.Width = ScaleWidth - lngStep * 4
    End With
    Err = 0: On Error GoTo 0
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtExpend.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Alignment", txtExpend.Alignment, 0)
    Call PropBag.WriteProperty("BorderColor", shapBack.BorderColor, -2147483638)
    Call PropBag.WriteProperty("IMEMode", txtExpend.IMEMode, 0)
    Call PropBag.WriteProperty("Locked", txtExpend.Locked, False)
    Call PropBag.WriteProperty("MaxLength", txtExpend.MaxLength, 0)
    Call PropBag.WriteProperty("SelLength", txtExpend.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtExpend.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtExpend.SelText, "")
    Call PropBag.WriteProperty("MouseRightMenus", m_MouseRightMenus, m_def_MouseRightMenus)
    Call PropBag.WriteProperty("InputMode", m_InputMode, m_def_InputMode)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("Showline", m_Showline, 0)
    Call PropBag.WriteProperty("Text", txtExpend.Text, "Text1")
    Call PropBag.WriteProperty("GotFocusSelAll", m_GotFocusSelAll, m_def_GotFocusSelAll)
    Call PropBag.WriteProperty("PasswordChar", txtExpend.PasswordChar, "")
End Sub
Private Sub SetTxtHeight()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置文本的高度
    '编制:刘兴洪
    '日期:2014-12-26 14:13:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
     txtExpend.Height = TextHeight("刘") * 1.2 \ Screen.TwipsPerPixelY
     txtExpend.Refresh
End Sub
 

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    txtExpend.BackColor() = UserControl.BackColor
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtExpend.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtExpend.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    txtExpend.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Set txtExpend.Font = UserControl.Font
    Call UserControl_Resize
End Property
 
'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,Alignment
Public Property Get Alignment() As Integer
    Alignment = txtExpend.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    txtExpend.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=shapBack,shapBack,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = shapBack.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shapBack.BorderColor() = New_BorderColor
    lneUnder.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

Private Sub txtExpend_Change()
    RaiseEvent Change
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,IMEMode
Public Property Get IMEMode() As Integer
    IMEMode = txtExpend.IMEMode
End Property

Public Property Let IMEMode(ByVal New_IMEMode As Integer)
    txtExpend.IMEMode() = New_IMEMode
    PropertyChanged "IMEMode"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
    Set Image = UserControl.Image
End Property

Private Sub txtExpend_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtExpend_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    Call InputGovern(txtExpend, KeyAscii, m_InputMode)
End Sub

Private Sub txtExpend_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,Locked
Public Property Get Locked() As Boolean
    Locked = txtExpend.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtExpend.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = txtExpend.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtExpend.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Private Sub txtExpend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button <> 2 Then Exit Sub
    If m_MouseRightMenus And txtExpend.Locked = False Then Exit Sub
    
    glngTXTProc = GetWindowLong(UserControl.hWnd, GWL_WNDPROC)
    Call SetWindowLong(txtExpend.hWnd, GWL_WNDPROC, AddressOf NotRightMenuMessage)
End Sub

Private Sub txtExpend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtExpend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button <> 2 Then Exit Sub
    If m_MouseRightMenus And txtExpend.Locked = False Then Exit Sub
    Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = txtExpend.SelLength
End Property
Public Property Let SelLength(ByVal New_SelLength As Long)
    txtExpend.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = txtExpend.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtExpend.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,SelText
Public Property Get SelText() As String
    SelText = txtExpend.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtExpend.SelText() = New_SelText
    PropertyChanged "SelText"
End Property
 

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get MouseRightMenus() As Boolean
    MouseRightMenus = m_MouseRightMenus
End Property

Public Property Let MouseRightMenus(ByVal New_MouseRightMenus As Boolean)
    m_MouseRightMenus = New_MouseRightMenus
    PropertyChanged "MouseRightMenus"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get InputMode() As Em_InputMode
    InputMode = m_InputMode
End Property

Public Property Let InputMode(ByVal New_InputMode As Em_InputMode)
    m_InputMode = New_InputMode
    PropertyChanged "InputMode"
End Property


Public Sub InputGovern(ByVal objCtl As Object, KeyAscii As Integer, ByVal intInputMode As Em_InputMode)
    '------------------------------------------------------------------------------------------------------------------
    '功能:输入控制
    '参数:
    '   objctl:限制控件
    '   Keyascii:
    '         Keyascii:8 (退格)
    '   intInputMode:(0-文本式;1-数字式;2-金额式,3-金额式(含负金额))
    '返回:一个KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo errHand:
    
    If intInputMode = InPut_Chars Then
        If KeyAscii = Asc("'") Then KeyAscii = 0
        Exit Sub
    End If

    If Not (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then Exit Sub
    If KeyAscii = vbKeyReturn Then Exit Sub ' 回车
    If KeyAscii = 8 Then Exit Sub    '退格
    If KeyAscii = Asc(".") Then '小数点控制
        If intInputMode = Input_Moneys Or intInputMode = Input_NegativeMoneys Then
            If InStr(objCtl, ".") <> 0 Then     '只能存在一个小数点
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
        Exit Sub
    End If
    If KeyAscii = Asc("-") Then '负数控制
        If intInputMode <> Input_NegativeMoneys Then KeyAscii = 0: Exit Sub
        If Trim(objCtl.Text) = "" Then Exit Sub
        If objCtl.SelStart <> 0 Then KeyAscii = 0: Exit Sub      '光标不存第一位,不能输入负数
        If InStr(1, objCtl.Text, "-") <> 0 Then KeyAscii = 0: Exit Sub     '只能存在一个负数
        Exit Sub
    End If
    '非数字
    KeyAscii = 0
    Exit Sub
errHand:
    KeyAscii = 0
End Sub
 
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
'MemberInfo=7,0,0,0
Public Property Get Appearance() As Em_Appearance
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Em_Appearance)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    Call ReSetFace
    
End Property

Private Sub ReSetFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置控件面版
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-02 15:17:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If m_Appearance = Show_3D Then
        UserControl.BorderStyle = 1
        shapBack.Visible = False
        lneUnder.Visible = False
    Else
        UserControl.BorderStyle = 0
        If m_Showline = 0 Then
            shapBack.Visible = True
            lneUnder.Visible = False
        Else
            shapBack.Visible = False
            lneUnder.Visible = True
        End If
    End If
    Call UserControl_Resize
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get Showline() As Integer
    Showline = m_Showline
End Property

Public Property Let Showline(ByVal New_ShowLine As Integer)
    m_Showline = New_ShowLine
    PropertyChanged "Showline"
    Call ReSetFace
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,Text
Public Property Get Text() As String
    Text = txtExpend.Text
End Property
Public Property Let Text(ByVal New_Text As String)
    txtExpend.Text() = New_Text
    PropertyChanged "Text"
End Property
Private Sub UserControl_GotFocus()
    Err = 0: On Error Resume Next
    If txtExpend.Enabled And txtExpend.Visible Then txtExpend.SetFocus
    Err = 0: On Error GoTo 0
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,true
Public Property Get GotFocusSelAll() As Boolean
    GotFocusSelAll = m_GotFocusSelAll
End Property

Public Property Let GotFocusSelAll(ByVal New_GotFocusSelAll As Boolean)
    m_GotFocusSelAll = New_GotFocusSelAll
    PropertyChanged "GotFocusSelAll"
End Property

Public Sub TxtSelAll(objTxt As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将编辑框的的文本全部选中
    '入参:objTxt=需要全选的编辑控件,该控件具有SelStart,SelLength属性
    '编制:刘兴洪
    '日期:2015-01-12 14:03:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    If TypeName(objTxt) = "TextBox" Then
        If objTxt.MultiLine Then
            SendMessage objTxt.hWnd, WM_VSCROLL, SB_TOP, 0
        End If
    End If
End Sub
'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtExpend,txtExpend,-1,PasswordChar
Public Property Get PasswordChar() As String
    PasswordChar = txtExpend.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtExpend.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property



