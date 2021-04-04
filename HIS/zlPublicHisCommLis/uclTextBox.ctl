VERSION 5.00
Begin VB.UserControl uclTextBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   60
      ScaleHeight     =   2805
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   30
      Width           =   4665
      Begin VB.TextBox txtShow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   420
         TabIndex        =   2
         Top             =   0
         Width           =   1695
      End
      Begin VB.PictureBox picFoundButton 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2220
         MouseIcon       =   "uclTextBox.ctx":0000
         MousePointer    =   99  'Custom
         Picture         =   "uclTextBox.ctx":030A
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   60
         Width           =   255
      End
      Begin VB.Label lblCation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标题"
         Height          =   180
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   360
      End
      Begin VB.Line lineBottom 
         X1              =   420
         X2              =   1950
         Y1              =   450
         Y2              =   450
      End
   End
End
Attribute VB_Name = "uclTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum m_e_DataType
    Number_Type = 0
    Char_Type
    Date_Type
End Enum


'缺省属性值:
Const m_def_MaxValue = ""
Const m_def_MinValue = ""
Const m_def_Text1 = ""
Const m_def_ShowFoundButton = True
Const m_def_DataType = m_e_DataType.Char_Type
'Const m_def_Text1 = ""
'Const m_def_DataType = m_e_DataType.Char_Type
'Const m_def_SQLDataFrom = m_e_SQLDataFrom.LISDB
'Const m_def_SQL = ""
'Const m_def_DataType = m_e_DataType.Char_Type
'Const m_def_ShowFoundButton = True
'属性变量:
Dim m_MaxValue As String
Dim m_MinValue As String
Dim m_Text1 As String
Dim m_ShowFoundButton As Boolean
Dim m_DataType As m_e_DataType
'Dim m_Text1 As Variant
'Dim m_DataType As m_e_DataType
'Dim m_SQLDataFrom As Byte
'Dim m_SQL As String
'Dim m_DataType As m_e_DataType
'Dim m_ShowFoundButton As Boolean
'事件声明:
Event Click()    'MappingInfo=picFoundButton,picFoundButton,-1,Click
Event DblClick()    'MappingInfo=txtShow,txtShow,-1,DblClick
Attribute DblClick.VB_Description = "当用户在一个对象上按下并释放鼠标按钮后再次按下并释放鼠标按钮时发生。"
Event KeyDown(KeyCode As Integer, Shift As Integer)    'MappingInfo=txtShow,txtShow,-1,KeyDown
Attribute KeyDown.VB_Description = "当用户在拥有焦点的对象上按下任意键时发生。"
Event KeyPress(KeyAscii As Integer)    'MappingInfo=txtShow,txtShow,-1,KeyPress
Attribute KeyPress.VB_Description = "当用户按下和释放 ANSI 键时发生。"
Event KeyUp(KeyCode As Integer, Shift As Integer)    'MappingInfo=txtShow,txtShow,-1,KeyUp
Attribute KeyUp.VB_Description = "当用户在拥有焦点的对象上释放键时发生。"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)    'MappingInfo=picFoundButton,picFoundButton,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)    'MappingInfo=picFoundButton,picFoundButton,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)    'MappingInfo=picFoundButton,picFoundButton,-1,MouseUp



'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtShow,txtShow,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = txtShow.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtShow.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtShow,txtShow,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = txtShow.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtShow.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtShow,txtShow,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = txtShow.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtShow.Font = New_Font
    Set lblCation.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub picFoundButton_Click()
    RaiseEvent Click
End Sub

Private Sub picMain_Resize()
    On Error Resume Next

    picFoundButton.Left = UserControl.Width - picFoundButton.Width
    lblCation.Left = 0

    With txtShow
        .Width = picFoundButton.Left - lblCation.Width
        .Left = picFoundButton.Left - .Width
        .Top = 0
        .Height = picMain.Height - 100
        picFoundButton.Top = (txtShow.Height - picFoundButton.Height) / 2
        lblCation.Top = (txtShow.Height - lblCation.Height) / 2
    End With
    With lineBottom
        .X1 = txtShow.Left
        .Y1 = txtShow.Top + txtShow.Height + 10
        .X2 = .X1 + txtShow.Width
        .Y2 = .Y1
    End With
End Sub

Private Sub txtShow_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtShow_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtShow_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtShow_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picFoundButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picFoundButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picFoundButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=lblCation,lblCation,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "返回/设置对象的标题栏中或图标下面的文本。"
    Caption = lblCation.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCation.Caption() = New_Caption
    Call picMain_Resize
    PropertyChanged "Caption"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtShow,txtShow,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "决定控件是否可编辑。"
    Locked = txtShow.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtShow.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtShow,txtShow,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "返回/设置一个控件中可以输入的字符的最大数。"
    MaxLength = txtShow.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtShow.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtShow,txtShow,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "返回/设置选定的字符数。"
    SelLength = txtShow.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtShow.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtShow,txtShow,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "返回/设置选定文本的起始点。"
    SelStart = txtShow.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtShow.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtShow,txtShow,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "返回/设置控件中包含的文本。"
    Text = txtShow.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtShow.Text() = New_Text
    PropertyChanged "Text"
End Property
'
''注意！不要删除或修改下列被注释的行！
''MemberInfo=0,0,0,True
'Public Property Get ShowFoundButton() As Boolean
'    ShowFoundButton = m_ShowFoundButton
'End Property
'
'Public Property Let ShowFoundButton(ByVal New_ShowFoundButton As Boolean)
'    m_ShowFoundButton = New_ShowFoundButton
'    picFoundButton.Visible = New_ShowFoundButton
'    PropertyChanged "ShowFoundButton"
'End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
'    m_ShowFoundButton = m_def_ShowFoundButton
'    m_SQL = m_def_SQL
'    m_DataType = m_def_DataType
'    m_SQLDataFrom = m_def_SQLDataFrom
'    m_DataType = m_def_DataType
    m_ShowFoundButton = m_def_ShowFoundButton
    m_DataType = m_def_DataType
'    m_Text1 = m_def_Text1
    m_Text1 = m_def_Text1
    m_MaxValue = m_def_MaxValue
    m_MinValue = m_def_MinValue
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtShow.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtShow.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtShow.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set lblCation.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCation.Caption = PropBag.ReadProperty("Caption", "")
    txtShow.Locked = PropBag.ReadProperty("Locked", False)
    txtShow.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtShow.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtShow.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtShow.Text = PropBag.ReadProperty("Text", "")
'    m_ShowFoundButton = PropBag.ReadProperty("ShowFoundButton", m_def_ShowFoundButton)
    picFoundButton.Visible = PropBag.ReadProperty("ShowFoundButton", m_def_ShowFoundButton)
    picMain.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtShow.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    picFoundButton.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
'    m_SQL = PropBag.ReadProperty("SQL", m_def_SQL)
'    m_DataType = PropBag.ReadProperty("DataType", m_def_DataType)
'    m_SQLDataFrom = PropBag.ReadProperty("SQLDataFrom", m_def_SQLDataFrom)
'    m_DataType = PropBag.ReadProperty("DataType", m_def_DataType)
    m_ShowFoundButton = PropBag.ReadProperty("ShowFoundButton", m_def_ShowFoundButton)
    picFoundButton.Visible = m_ShowFoundButton
    m_DataType = PropBag.ReadProperty("DataType", m_def_DataType)
'    m_Text1 = PropBag.ReadProperty("Text1", m_def_Text1)
    m_Text1 = PropBag.ReadProperty("Text1", m_def_Text1)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
End Sub

Private Sub UserControl_Resize()
    With picMain
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height
    End With
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", txtShow.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtShow.Enabled, True)
    Call PropBag.WriteProperty("Font", txtShow.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", lblCation.Caption, "")
    Call PropBag.WriteProperty("Locked", txtShow.Locked, False)
    Call PropBag.WriteProperty("MaxLength", txtShow.MaxLength, 0)
    Call PropBag.WriteProperty("SelLength", txtShow.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtShow.SelStart, 0)
    Call PropBag.WriteProperty("Text", txtShow.Text, "")
'    Call PropBag.WriteProperty("ShowFoundButton", m_ShowFoundButton, m_def_ShowFoundButton)
    Call PropBag.WriteProperty("BackColor", picMain.BackColor, &H80000005)
'    Call PropBag.WriteProperty("SQL", m_SQL, m_def_SQL)
'    Call PropBag.WriteProperty("DataType", m_DataType, m_def_DataType)
'    Call PropBag.WriteProperty("SQLDataFrom", m_SQLDataFrom, m_def_SQLDataFrom)
'    Call PropBag.WriteProperty("DataType", m_DataType, m_def_DataType)
    Call PropBag.WriteProperty("ShowFoundButton", m_ShowFoundButton, m_def_ShowFoundButton)
    Call PropBag.WriteProperty("DataType", m_DataType, m_def_DataType)
'    Call PropBag.WriteProperty("Text1", m_Text1, m_def_Text1)
    Call PropBag.WriteProperty("Text1", m_Text1, m_def_Text1)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
End Sub


'注意！不要删除或修改下列被注释的行！
'MappingInfo=picMain,picMain,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = picMain.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picMain.BackColor() = New_BackColor
    txtShow.BackColor() = New_BackColor
    picFoundButton.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
''
'''注意！不要删除或修改下列被注释的行！
'''MemberInfo=13,0,0,
''Public Property Get SQL() As String
''    SQL = m_SQL
''End Property
''
''Public Property Let SQL(ByVal New_SQL As String)
''    m_SQL = New_SQL
''    PropertyChanged "SQL"
''End Property
''
'''注意！不要删除或修改下列被注释的行！
'''MemberInfo=14,0,0,0
''Public Property Get DataType() As m_e_DataType
''    DataType = m_DataType
''End Property
''
''Public Property Let DataType(ByVal New_DataType As m_e_DataType)
''    m_DataType = New_DataType
''    PropertyChanged "DataType"
''End Property
''
'''注意！不要删除或修改下列被注释的行！
'''MemberInfo=1,0,0,0
''Public Property Get SQLDataFrom() As m_e_SQLDataFrom
''    SQLDataFrom = m_SQLDataFrom
''End Property
''
''Public Property Let SQLDataFrom(ByVal New_SQLDataFrom As m_e_SQLDataFrom)
''    m_SQLDataFrom = New_SQLDataFrom
''    PropertyChanged "SQLDataFrom"
''End Property
''
''注意！不要删除或修改下列被注释的行！
''MemberInfo=23,0,0,0
'Public Property Get DataType() As m_e_DataType
'    DataType = m_DataType
'End Property
'
'Public Property Let DataType(ByVal New_DataType As m_e_DataType)
'    m_DataType = New_DataType
'    PropertyChanged "DataType"
'End Property
'
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get ShowFoundButton() As Boolean
Attribute ShowFoundButton.VB_Description = "返回/设置是否显示查找按钮"
    ShowFoundButton = m_ShowFoundButton
End Property

Public Property Let ShowFoundButton(ByVal New_ShowFoundButton As Boolean)
    m_ShowFoundButton = New_ShowFoundButton
    picFoundButton.Visible = New_ShowFoundButton
    PropertyChanged "ShowFoundButton"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=23,0,0,0
Public Property Get DataType() As m_e_DataType
Attribute DataType.VB_Description = "返回/设置文本框的数据类型"
    DataType = m_DataType
End Property

Public Property Let DataType(ByVal New_DataType As m_e_DataType)
    m_DataType = New_DataType
    PropertyChanged "DataType"
End Property
'
''注意！不要删除或修改下列被注释的行！
''MemberInfo=14,0,0,0
'Public Property Get Text1() As Variant
'    Text1 = m_Text1
'End Property
'
'Public Property Let Text1(ByVal New_Text1 As Variant)
'    m_Text1 = New_Text1
'    PropertyChanged "Text1"
'End Property
'
'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get Text1() As String
Attribute Text1.VB_Description = "返回/设置控件的另一个Text属性"
    Text1 = m_Text1
End Property

Public Property Let Text1(ByVal New_Text1 As String)
    m_Text1 = New_Text1
    PropertyChanged "Text1"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get MaxValue() As String
    MaxValue = m_MaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As String)
    m_MaxValue = New_MaxValue
    PropertyChanged "MaxValue"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get MinValue() As String
    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As String)
    m_MinValue = New_MinValue
    PropertyChanged "MinValue"
End Property

