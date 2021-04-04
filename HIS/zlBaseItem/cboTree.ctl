VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl cboTree 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2925
   ScaleHeight     =   300
   ScaleWidth      =   2925
   Begin VB.PictureBox picCmd 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2655
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   240
      Begin VB.Image img 
         Height          =   150
         Left            =   30
         Picture         =   "cboTree.ctx":0000
         Top             =   45
         Width           =   135
      End
   End
   Begin VB.TextBox txtThis 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2925
   End
   Begin MSComctlLib.ImageList ilsDown 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cboTree.ctx":015A
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cboTree.ctx":0476
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cboTree.ctx":0690
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cboTree.ctx":09AC
            Key             =   "Dept_No"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3000
      Left            =   945
      TabIndex        =   2
      Top             =   1590
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5292
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.TreeView tvwDown 
      Height          =   4575
      Left            =   2445
      TabIndex        =   3
      Top             =   1395
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   8070
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ilsDown"
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "cboTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'下列API及类型定义，用于将选择器显示于父窗体中
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const SM_CYMENU = 15
Private Const GWL_HWNDPARENT = (-8)
Private Const HWND_TOP = 0
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private rect_Client As RECT
Private pt_Pop As POINTAPI
Private rect_Parent As RECT


Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'事件声明:
Event Click() 'MappingInfo=txtThis,txtThis,-1,Click
Attribute Click.VB_Description = "当用户在一个对象上按下并释放鼠标按钮时发生。"
Event DblClick() 'MappingInfo=txtThis,txtThis,-1,DblClick
Attribute DblClick.VB_Description = "当用户在一个对象上按下并释放鼠标按钮后再次按下并释放鼠标按钮时发生。"
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtThis,txtThis,-1,KeyDown
Attribute KeyDown.VB_Description = "当用户在拥有焦点的对象上按下任意键时发生。"
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtThis,txtThis,-1,KeyPress
Attribute KeyPress.VB_Description = "当用户按下和释放 ANSI 键时发生。"
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtThis,txtThis,-1,KeyUp
Attribute KeyUp.VB_Description = "当用户在拥有焦点的对象上释放键时发生。"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtThis,txtThis,-1,MouseDown
Attribute MouseDown.VB_Description = "当用户在拥有焦点的对象上按下鼠标按钮时发生。"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtThis,txtThis,-1,MouseMove
Attribute MouseMove.VB_Description = "当用户移动鼠标时发生。"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtThis,txtThis,-1,MouseUp
Attribute MouseUp.VB_Description = "当用户在拥有焦点的对象上释放鼠标发生。"
Event DownClick()
Event Change() 'MappingInfo=txtThis,txtThis,-1,Change
Attribute Change.VB_Description = "当控件内容改变时发生。"
Event ItemClick(strID As String)        '选中项目时,激发该事件

'本地变量设置
Private mrsDataSource As ADODB.Recordset        '源记录集
Private mstrFileter As String
Private mstrShowFields As String        '多选时的显示字段
Private mstrSaveTvwKey As String
Private mblnSelect As Boolean
Private mintPreCol As Integer               '前一次单据头的排序列
Private mintsort As Integer                 '前一次单据头的排序
Private msngSelDownWidth As Single             '下拉选择的宽度
Private msngSelDownHeight As Single             '下拉选择的高度
Private mSplitString As String              '编码与名称的分离符号
Private mTopShowDown As Boolean             '在上显示下拉列表
Private lngParent As Long       '保存父窗体指针

Private blnPop As Boolean       '是否已经弹出


Private Sub MouseDown()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按下时，Pic凹下
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    If Enabled = False Then Exit Sub
    zlControl.PicShowFlat picCmd, -1
End Sub
Private Sub MouseUp()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按下时，Pic凸起
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    If Enabled = False Then Exit Sub
    zlControl.PicShowFlat picCmd, 1
End Sub

Private Sub img_Click()
    picCmd_Click
End Sub

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        MouseDown
End Sub

Private Sub img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
       MouseUp
End Sub

Private Sub picCmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
       MouseDown
End Sub

Private Sub picCmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        MouseUp
End Sub

Private Sub picCmd_Resize()
    With img
        .Left = (picCmd.ScaleWidth - .Width) / 2
        .Top = (picCmd.ScaleHeight - .Height) / 2
    End With
End Sub

Private Sub picCmd_Click()
   If txtThis.Enabled = False Then Exit Sub
     
    PicClick
    If txtThis.Enabled Then
        RaiseEvent DownClick
    End If
End Sub

Private Sub tvwDown_DblClick()
    If tvwDown.SelectedItem Is Nothing Then
        txtThis.Text = ""
        '3.关闭
        tvwDown.Visible = False
        RaiseEvent ItemClick("")
        Exit Sub
    End If
    If mstrSaveTvwKey <> tvwDown.SelectedItem.Key Then
        mstrSaveTvwKey = tvwDown.SelectedItem.Key
        '1.
        txtThis.Text = tvwDown.SelectedItem.Text
        tvwDown.Visible = False
        RaiseEvent ItemClick(Mid(mstrSaveTvwKey, 2))
    End If
    '3.关闭
    tvwDown.Visible = False
End Sub

Private Sub tvwDown_GotFocus()
    If tvwDown.SelectedItem Is Nothing Then
        mstrSaveTvwKey = ""
    Else
        mstrSaveTvwKey = tvwDown.SelectedItem.Key
    End If
End Sub

Private Sub tvwDown_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
'        If mstrSaveTvwKey <> tvwDown.SelectedItem.Key Then
'            Err = 0
'            On Error Resume Next
'            tvwDown.Nodes(mstrSaveTvwKey).Selected = True
'            txtThis.Text = tvwDown.SelectedItem.Text
'        Else
'            txtThis.Text = tvwDown.SelectedItem.Text
'        End If
         tvwDown.Visible = False
         TxtSelAll txtThis
    End If
End Sub

Private Sub tvwDown_LostFocus()
    tvwDown.Visible = False
End Sub

Private Sub tvwDown_NodeClick(ByVal Node As MSComctlLib.Node)
        Call tvwDown_DblClick
End Sub

Private Sub txtThis_GotFocus()
    If blnPop Then
        blnPop = False
        Exit Sub
    End If
    
    tvwDown.Visible = False
    TxtSelAll txtThis
End Sub

Private Sub UserControl_ExitFocus()
    tvwDown.Visible = False
    mshSelect.Visible = False
    If tvwDown.SelectedItem Is Nothing Then Exit Sub
    If tvwDown.SelectedItem.Text <> txtThis.Text And Trim(txtThis.Text) <> "" Then
        txtThis.Text = tvwDown.SelectedItem.Text
    End If
End Sub

Private Sub UserControl_Paint()
    Err = 0
    On Error Resume Next
     With txtThis
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
        Width = .Width
        picCmd.Left = IIF(.Width - picCmd.Width - 30 <= 0, 0, .Width - picCmd.Width - 30)
        picCmd.Height = .Height - 50
        zlControl.PicShowFlat picCmd, 1
    End With

End Sub

Private Sub UserControl_Resize()
    Err = 0
    On Error Resume Next
    With txtThis
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
        
        picCmd.Left = IIF(.Width - picCmd.Width - 30 <= 0, 0, .Width - picCmd.Width - 30)
        picCmd.Height = .Height - 50
        zlControl.PicShowFlat picCmd, 1
    End With
    mshSelect.SelectionMode = flexSelectionByRow
'    If mshSelect.Visible Then
'        LocaleCtl mshSelect
'    End If
'    If tvwDown.Visible Then
'        LocaleCtl tvwDown
'    End If

End Sub
Public Sub redraw()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:重画
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Refresh
   If mshSelect.Visible Then
        mshSelect.Visible = False
    End If
    If tvwDown.Visible Then
        mshSelect.Visible = False
    End If
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtThis.Locked = PropBag.ReadProperty("Locked", False)
    txtThis.Alignment = PropBag.ReadProperty("Alignment", 0)
    txtThis.Appearance = PropBag.ReadProperty("Appearance", 1)
    txtThis.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtThis.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtThis.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtThis.FontBold = PropBag.ReadProperty("FontBold", 0)
    txtThis.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    txtThis.FontName = PropBag.ReadProperty("FontName", "")
    txtThis.FontSize = PropBag.ReadProperty("FontSize", 0)
    txtThis.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    txtThis.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtThis.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    txtThis.Locked = PropBag.ReadProperty("Locked", False)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtThis.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtThis.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtThis.SelText = PropBag.ReadProperty("SelText", "")
    txtThis.Text = PropBag.ReadProperty("Text", "")
    txtThis.SelLength = PropBag.ReadProperty("SelLength", 0)
    mSplitString = PropBag.ReadProperty("SplitString", "】")
    
    msngSelDownWidth = PropBag.ReadProperty("sngSelDownWidth", 1980)
    msngSelDownHeight = PropBag.ReadProperty("sngSelDownHeight", 4575)
    mTopShowDown = PropBag.ReadProperty("TopShowDown", False)
    
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
      
    Call PropBag.WriteProperty("Locked", txtThis.Locked, False)
    Call PropBag.WriteProperty("Alignment", txtThis.Alignment, 0)
    Call PropBag.WriteProperty("Appearance", txtThis.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", txtThis.BackColor, &H80000005)
    Call PropBag.WriteProperty("Enabled", txtThis.Enabled, True)
    Call PropBag.WriteProperty("Font", txtThis.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", txtThis.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", txtThis.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", txtThis.FontName, "")
    Call PropBag.WriteProperty("FontSize", txtThis.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", txtThis.FontStrikethru, 0)
    Call PropBag.WriteProperty("ForeColor", txtThis.ForeColor, &H80000008)
    Call PropBag.WriteProperty("FontUnderline", txtThis.FontUnderline, 0)
    Call PropBag.WriteProperty("Locked", txtThis.Locked, False)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MaxLength", txtThis.MaxLength, 0)
    Call PropBag.WriteProperty("SelStart", txtThis.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtThis.SelText, "")
    Call PropBag.WriteProperty("SelLength", txtThis.SelLength, 0)
    Call PropBag.WriteProperty("Text", txtThis.Text, "")
    Call PropBag.WriteProperty("SplitString", mSplitString, "】")
    Call PropBag.WriteProperty("sngSelDownWidth", msngSelDownWidth, 1980)
    Call PropBag.WriteProperty("sngSelDownHeight", msngSelDownHeight, 4575)
    Call PropBag.WriteProperty("TopShowDown", mTopShowDown, False)
End Sub

Private Sub txtThis_Change()
    DataSourceFilter
    RaiseEvent Change
    '
End Sub
'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,Alignment
Public Property Get SplitString() As String
    SplitString = mSplitString
End Property

Public Property Let SplitString(ByVal New_SplitString As String)
    mSplitString = New_SplitString
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "返回/设置复选框或选项按钮、或一个控件的文本的对齐。"
    Alignment = txtThis.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    txtThis.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "返回/设置一个对象在运行时是否以 3D 效果显示。"
    Appearance = txtThis.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    txtThis.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = txtThis.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtThis.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub txtThis_Click()
    RaiseEvent Click
End Sub

Private Sub txtThis_DblClick()
    RaiseEvent DblClick
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = txtThis.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtThis.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = txtThis.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtThis.Font = New_Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "返回/设置粗体字样式。"
    FontBold = txtThis.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtThis.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "返回/设置斜体字样式。"
    FontItalic = txtThis.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtThis.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "指定给定层的每一行出现的字体名。"
    FontName = txtThis.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    txtThis.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "指定给定层的每一行出现的字体大小(以磅为单位)。"
    FontSize = txtThis.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    txtThis.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "返回/设置删除线字体样式。"
    FontStrikethru = txtThis.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    txtThis.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = txtThis.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtThis.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "返回/设置下划线字体样式。"
    FontUnderline = txtThis.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtThis.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property


Private Sub txtThis_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = Asc("*") Then
        KeyAscii = 0
        '弹出选择
        DataSourceFilter True
        If mshSelect.Visible Then
            mshSelect.SetFocus
        End If
    End If
End Sub

Private Sub txtThis_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "决定控件是否可编辑。"
    Locked = txtThis.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtThis.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

Private Sub txtThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "设置一个自定义鼠标图标。"
    Set MouseIcon = txtThis.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txtThis.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "返回/设置一个控件中可以输入的字符的最大数。"
    MaxLength = txtThis.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtThis.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Private Sub txtThis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "返回/设置选定文本的起始点。"
    SelStart = txtThis.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtThis.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "返回/设置包含当前选定文本的字符串。"
    SelText = txtThis.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtThis.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "返回/设置控件中包含的文本。"
    Text = txtThis.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtThis.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Sub TxtReSize(ByVal sngTxtWidth As Single)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:调整文本大小
    '--入参数:sngTxtWidth-文本框的宽度
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    txtThis.Width = IIF(sngTxtWidth < 0, 0, sngTxtWidth)
    UserControl.Width = IIF(sngTxtWidth < 0, 0, sngTxtWidth) + 10
    
    Call UserControl_Resize
End Sub
'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = txtThis.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtThis.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,SelLength
Public Property Get TopShowDown() As Boolean
    TopShowDown = mTopShowDown
End Property

Public Property Let TopShowDown(ByVal New_TopShowDown As Boolean)
    mTopShowDown = New_TopShowDown
End Property


'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtThis,txtThis,-1,SelLength
Public Property Get SelectItemID() As String
    Dim strKey As String
    If tvwDown.SelectedItem Is Nothing Then Exit Sub
    strKey = Mid(tvwDown.SelectedItem.Key, 2)
    If UCase(strKey) = "OOT" Then
        SelectItemID = ""
    Else
        SelectItemID = Mid(tvwDown.SelectedItem.Key, 2)
    End If
End Property

Public Property Let SelectItemID(ByVal New_SelLength As String)
    Err = 0
    On Error Resume Next
    
    With tvwDown
        If New_SelLength = "" Then
            .Nodes("Root").Selected = True
            .Nodes("Root").EnsureVisible
            .Nodes("Root").Expanded = True
        Else
            .Nodes("K" & New_SelLength).Selected = True
            .Nodes("K" & New_SelLength).EnsureVisible
            .Nodes("K" & New_SelLength).Expanded = True
              
        End If
    End With
    If Err <> 0 Then
        tvwDown.SelectedItem.Selected = False
        txtThis.Text = ""
    End If
    'tvwDown_MouseDown 1, 0, 0, 0
    Call tvwDown_DblClick
End Property
Public Property Get selDownWidth() As Single
    selDownWidth = msngSelDownWidth
End Property
Public Property Let selDownWidth(ByVal sngNewValue As Single)
        msngSelDownWidth = sngNewValue
        tvwDown.Width = sngNewValue
        mshSelect.Width = sngNewValue
End Property

Public Property Get selDownHeight() As Single
    selDownHeight = msngSelDownHeight
End Property
Public Property Let selDownHeight(ByVal sngNewValue As Single)
        msngSelDownHeight = sngNewValue
        tvwDown.Height = sngNewValue
        mshSelect.Height = sngNewValue
End Property

Private Sub PicClick()
    '设置在当前窗体范围内的位置
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
    End If
    With tvwDown
        .Visible = True
        .Left = txtThis.Left
        .Width = IIF(selDownWidth <= 0, UserControl.Width, selDownWidth)
        .Height = IIF(selDownHeight <= 0, .Height, selDownHeight)
        If TopShowDown Then
            .Top = picCmd.Top
        Else
            .Top = picCmd.Top + picCmd.Height
        End If
        .SetFocus
       
        If .SelectedItem Is Nothing Then
        Else
             .Nodes(.SelectedItem.Index).Expanded = True
            .Nodes(.SelectedItem.Index).EnsureVisible
        End If
        
    End With
    
    '转换成屏幕坐标
    pt_Pop.X = tvwDown.Left / Screen.TwipsPerPixelX
    pt_Pop.Y = tvwDown.Top / Screen.TwipsPerPixelY
    Call ClientToScreen(UserControl.hwnd, pt_Pop)
    '设置其父窗体为控件的父窗体
    If lngParent = 0 Then lngParent = GetParentWindow(UserControl.hwnd)
    Call SetParent(tvwDown.hwnd, lngParent)
    '设置为屏幕坐标
    Call GetWindowRect(lngParent, rect_Parent)
    Call GetClientRect(lngParent, rect_Client)
    If TopShowDown Then
        Call SetWindowPos(tvwDown.hwnd, HWND_TOP, pt_Pop.X - rect_Parent.Left - 2.5, pt_Pop.Y - (tvwDown.Height + txtThis.Height) / Screen.TwipsPerPixelY - ((rect_Parent.Top + rect_Client.Top)) - 7.5, 0, 0, SWP_NOSIZE + SWP_SHOWWINDOW)
    Else
        Call SetWindowPos(tvwDown.hwnd, HWND_TOP, pt_Pop.X - rect_Parent.Left - 2.5, pt_Pop.Y - (rect_Parent.Bottom - rect_Client.Bottom) + 5, 0, 0, SWP_NOSIZE + SWP_SHOWWINDOW)
    End If
End Sub
Private Sub LocaleCtl(ByVal objCtl As Object)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:将指定控件移动到相应位置
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------

    
    '设置在当前窗体范围内的位置
    With objCtl
        .Visible = False
        .Left = txtThis.Left
        .Width = IIF(msngSelDownWidth <= 0, Width, msngSelDownWidth)
        If .Height > IIF(msngSelDownHeight <= 0, .Height, msngSelDownHeight) Then
            .Height = IIF(msngSelDownHeight <= 0, .Height, msngSelDownHeight)
        End If
        
        If TopShowDown Then
            .Top = picCmd.Top
        Else
            .Top = picCmd.Top + picCmd.Height
        End If
    End With
    '转换成屏幕坐标
    pt_Pop.X = objCtl.Left / Screen.TwipsPerPixelX
    pt_Pop.Y = objCtl.Top / Screen.TwipsPerPixelY
    Call ClientToScreen(UserControl.hwnd, pt_Pop)
    '设置其父窗体为控件的父窗体
    
    If lngParent = 0 Then lngParent = GetParentWindow(UserControl.hwnd)
    Call SetParent(objCtl.hwnd, lngParent)
    
    '设置为屏幕坐标
    Call GetWindowRect(lngParent, rect_Parent)
    Call GetClientRect(lngParent, rect_Client)
    If TopShowDown Then
        Call SetWindowPos(objCtl.hwnd, HWND_TOP, pt_Pop.X - rect_Parent.Left - 2.5, pt_Pop.Y - (objCtl.Height + txtThis.Height) / Screen.TwipsPerPixelY - ((rect_Parent.Top + rect_Client.Top)) - 7.5, 0, 0, SWP_NOSIZE + SWP_SHOWWINDOW)
    Else
        Call SetWindowPos(objCtl.hwnd, HWND_TOP, pt_Pop.X - rect_Parent.Left - 5, pt_Pop.Y - (rect_Parent.Bottom - rect_Client.Bottom) + 5, 0, 0, SWP_NOSIZE + SWP_SHOWWINDOW)
    End If
    objCtl.Visible = True
End Sub

Public Function FullCboData(ByVal rsDataSource As ADODB.Recordset, ByRef strRootCaption As String, _
                       ByVal strFilterFields As String, ByVal strShowFields As String, _
                       Optional strSelID As String = "", Optional str人员性质 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充Combox数据
    '--入参数:rsDataSource-数据源,以id,上级id,...,编码,名称...顺序进行
    '         strRootCaption-根节点名字
    '         strFilterFields-过滤字段,如:编码,名称,简码
    '         strShowFilds-显示字段,如:编码|1000,名称|2000
    '         strSelID-初始被选择的ID
    '--出参数:
    '--返  回:填序成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strRootKey As String
    Dim strIco As String
    Dim strArr As Variant
    Dim intTmp As Integer
    Dim strLeftSplit As String, strRitghtLeft As String
    Dim objNd As Node
    
    
    Dim objNode As Node
    
     Set mrsDataSource = rsDataSource
    
    Select Case mSplitString
    Case "】"
        strLeftSplit = "【"
        strRitghtLeft = mSplitString
    Case "]"
        strLeftSplit = "["
        strRitghtLeft = mSplitString
    Case ")"
        strLeftSplit = "("
        strRitghtLeft = mSplitString
    Case "〗"
        strLeftSplit = "〖"
        strRitghtLeft = mSplitString
    Case "}"
        strLeftSplit = "{"
        strRitghtLeft = mSplitString
    Case Else
        If mSplitString = "" Then
            strLeftSplit = "【"
            strRitghtLeft = "】"
        Else
            strLeftSplit = ""
            strRitghtLeft = mSplitString
        End If
    End Select
    
    '3.加载树型数据
    tvwDown.Nodes.Clear
   
    '0.无数据时,退出
    If rsDataSource Is Nothing Then Exit Function
    If rsDataSource.RecordCount = 0 Then Exit Function
    '1.分解过滤字段
    If strFilterFields = "" Then
        strFilterFields = "编码,名称"
    End If
    strArr = Split(strFilterFields, ",")
    mstrFileter = ""
    For intTmp = 0 To UBound(strArr)
        mstrFileter = mstrFileter & " or ( " & strArr(intTmp) & " like '[*]%')"
    Next
    If mstrFileter <> "" Then
        mstrFileter = Mid(mstrFileter, 4)
    End If
    '2.初始变量
    mstrShowFields = strShowFields
    If strShowFields = "" Then
        mstrShowFields = "编码|1000,名称|2000"
    End If
    
    '3.加载树型数据
    If strRootCaption <> "" Then
        Set objNode = tvwDown.Nodes.Add(, , "Root", strRootCaption, "Root", "Root")
        
        strRootKey = "Root"
        objNode.Sorted = True
        objNode.Expanded = True
        If strSelID = "" Then
            objNode.Selected = True
            txtThis.Text = objNode.Text
            RaiseEvent ItemClick("0")
        End If
    Else
        strRootKey = ""
    End If
    strIco = "Dept"
    With rsDataSource
        rsDataSource.Filter = 0
        Do Until .EOF
            If IsNull(!上级id) Then
                If strRootKey = "" Then
                    Set objNode = tvwDown.Nodes.Add(, tvwChild, "K" & !ID, strLeftSplit & Nvl(!编码) & strRitghtLeft & Nvl(!名称), strIco, strIco)
                Else
                    Set objNode = tvwDown.Nodes.Add(strRootKey, tvwChild, "K" & !ID, strLeftSplit & Nvl(!编码) & strRitghtLeft & Nvl(!名称), strIco, strIco)
                End If
            Else
                Set objNode = tvwDown.Nodes.Add("K" & Nvl(!上级id), tvwChild, "K" & !ID, strLeftSplit & Nvl(!编码) & strRitghtLeft & Nvl(!名称), strIco, strIco)
            End If
            If strSelID = !ID Then
                mblnSelect = True
                objNode.Selected = True
                objNode.EnsureVisible
                txtThis.Text = objNode.Text
                RaiseEvent ItemClick(strSelID)
                mblnSelect = False
            End If
            objNode.Sorted = True
            objNode.Expanded = True
            .MoveNext
        Loop
    End With
    If str人员性质 = "医生" Then
        For Each objNd In tvwDown.Nodes
            If objNd.Text Like "*卫生技术人员（医疗）*" Then
                objNd.Selected = True
                objNd.EnsureVisible
                Exit For
            End If
        Next
    ElseIf str人员性质 = "护士" Then
        For Each objNd In tvwDown.Nodes
            If objNd.Text Like "*卫生技术（护理）*" Then
                objNd.Selected = True
                objNd.EnsureVisible
                Exit For
            End If
        Next
    End If
    
    If tvwDown.SelectedItem Is Nothing Then
        Err = 0
        On Error GoTo Err:
'        mblnSelect = True
'        tvwDown.Nodes(1).Selected = True
'        strSelID = Mid(tvwDown.SelectedItem.Key, 2)
'
'        txtThis.Text = tvwDown.SelectedItem.Text
'        RaiseEvent ItemClick(strSelID)
'        mblnSelect = False
    End If
Err:
    FullCboData = True
End Function

Private Sub DataSourceFilter(Optional blnAll As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:筛选数据，并以Grd方式显示出来
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim intCol As Integer
    Dim strTmp As String
    Dim sngWidth As Single
    Dim strArr As Variant
    Dim intTmp As Integer
    Dim lngRow As Long
    
    
    If tvwDown.Visible Then Exit Sub
    If mblnSelect = True Then Exit Sub
    If Trim(txtThis.Text) = "" And blnAll = False Then Exit Sub
    
    Err = 0
    On Error GoTo ErrHand:
    
        
    strTmp = Replace(UCase(txtThis.Text), "'", "")
    strTmp = Replace(mstrFileter, "[*]", "%" & strTmp)
  
    mrsDataSource.Filter = 0
    If blnAll = False Then
        mrsDataSource.Filter = strTmp
    End If

    If mrsDataSource.RecordCount > 1 Then
            With mshSelect
                .Cols = 2
                Set .Recordset = mrsDataSource
                ' id,上级id,编码 ,名称,简码
                strArr = Split(mstrShowFields, ",")
                
                For intCol = 0 To .Cols - 1
                    .ColWidth(intCol) = 0
                    For intTmp = 0 To UBound(strArr)
                        If InStr(1, strArr(intTmp), .TextMatrix(0, intCol)) <> 0 Then
                            .ColWidth(intCol) = Split(strArr(intTmp), "|")(1)
                            Exit For
                        End If
                    Next
                Next
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                If blnAll Then
                    Err = 0
                    On Error Resume Next
                    For lngRow = 1 To .Rows - 1
                         If .TextMatrix(lngRow, 0) = Mid(tvwDown.SelectedItem.Key, 2) Then
                            .Row = lngRow
                            .Col = 0
                            .ColSel = .Cols - 1
                            If .RowIsVisible(.Row) = False Then
                                .TopRow = .Row
                            End If
                            Exit For
                         End If
                    Next
                End If
                .Height = (.RowHeight(0) + 30) * (mrsDataSource.RecordCount + 1)
                LocaleCtl mshSelect
                .Visible = True
            End With
    ElseIf mrsDataSource.RecordCount = 1 Then
            tvwDown.Nodes("K" & mrsDataSource!ID).Selected = True
            tvwDown.Nodes("K" & mrsDataSource!ID).Expanded = True
            tvwDown.Nodes("K" & mrsDataSource!ID).EnsureVisible
            mstrSaveTvwKey = ""
            mblnSelect = True
            tvwDown_DblClick 'tvwDown_MouseDown 1, 0, 0, 0
            mshSelect.Visible = False
            TxtSelAll txtThis
            mblnSelect = False
    Else
           mshSelect.Visible = False
    End If
ErrHand:
End Sub

Private Sub TxtSelAll(objTxt As Object)
'功能：将编辑框的的文本全部选中
'参数：objTxt=需要全选的编辑控件,该控件具有SelStart,SelLength属性

    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
End Sub



Private Sub mshSelect_Click()
    With mshSelect
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            SetColumnSort mshSelect, mintPreCol, mintsort
            Exit Sub
         End If
        'Refresh
    End With
End Sub

Private Sub mshSelect_DblClick()
    With mshSelect
        If .Row > 0 And .TextMatrix(.Row, 0) <> "" Then
            mshSelect_KeyPress 13
        End If
    End With
End Sub

Private Sub mshselect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    With mshSelect
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sngWidth = sngWidth + .ColWidth(i)
                    If sngWidth > .Width Then
                        .LeftCol = i + 1
                        .Col = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
            Case vbKeyEscape    '取消选择
                txtThis.Text = tvwDown.SelectedItem.Text
                 mshSelect.Visible = False
                 TxtSelAll txtThis
            Case vbKeyBack      '回滚
                txtThis.SetFocus
                blnPop = True
                SendKeys "{BACKSPACE}"
                'SendKeys vbKeyBack
        End Select
        .redraw = True
        .Refresh
    End With
End Sub



Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    Dim strKey As String
    
    With mshSelect
            If KeyAscii = 13 Then
            strKey = "K" & Trim(.TextMatrix(.Row, 0))
            tvwDown.Nodes(strKey).Selected = True
            tvwDown.Nodes(strKey).Expanded = True '
            tvwDown.Nodes(strKey).EnsureVisible
            
            mstrSaveTvwKey = ""
            mblnSelect = True
            tvwDown_DblClick 'tvwDown_MouseDown 1, 0, 0, 0
            mblnSelect = False
            txtThis.SetFocus
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    If blnPop Then
        Exit Sub
    End If
    mshSelect.Visible = False
    If tvwDown.SelectedItem Is Nothing Then
        Exit Sub
    End If
    If txtThis.Text <> tvwDown.SelectedItem.Text Then
        txtThis.Text = tvwDown.SelectedItem.Text
    End If
End Sub
Private Sub tvwDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call tvwDown_DblClick
    End If
End Sub

Private Sub tvwDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With tvwDown
        Set .DropHighlight = .HitTest(X, Y)
        If .DropHighlight Is Nothing Then Exit Sub
        
        'tvwDown.Nodes(.DropHighlight.Index).Selected = True
    End With
End Sub
Private Sub txtThis_KeyDown(KeyCode As Integer, Shift As Integer)
     Dim strKey As String
     Select Case KeyCode
     Case vbKeyDown
        If Shift = 2 Then
            PicClick
        Else
            If mshSelect.Visible Then
                mblnSelect = False
                With mshSelect
                    
                    .SetFocus
                    If .Rows > 2 Then
                        Refresh
                        If .Row < .Rows - 1 Then
                            .Row = .Row + 1
                        End If
                        .Col = 0
                        .ColSel = .Cols - 1
                        
                    Else
                        .Row = 1
                        .Col = 0
                        .ColSel = .Cols - 1
                    End If
                End With
            End If
         End If
     Case vbKeyEscape
            If tvwDown.Visible = True Then
                txtThis.Text = tvwDown.SelectedItem.Text
                mshSelect.Visible = False
                TxtSelAll txtThis
            End If
     Case vbKeyReturn
        With mshSelect
            If mshSelect.Visible Then
                Err = 0
                On Error Resume Next
                strKey = "K" & Trim(.TextMatrix(1, 0))
                tvwDown.Nodes(strKey).Selected = True
                tvwDown.Nodes(strKey).Expanded = True
                tvwDown.Nodes(strKey).EnsureVisible
                mstrSaveTvwKey = ""
                mblnSelect = True
                tvwDown_DblClick 'tvwDown_MouseDown 1, 0, 0, 0
                mblnSelect = False
                mshSelect.Visible = False
                TxtSelAll txtThis
            End If
        End With
    
     End Select
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub


'对列头进行排序
Private Sub SetColumnSort(ByVal mshFilter As MSHFlexGrid, ByRef intPreCol As Integer, ByRef intPreSort As Integer)
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshFilter
        If .Rows > 1 Then
            .redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
            If InStr(1, .TextMatrix(0, intCol), "数量") <> 0 Or InStr(1, .TextMatrix(0, intCol), "金额") <> 0 Then
                    If intCol = intPreCol And intPreSort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       intPreSort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       intPreSort = flexSortNumericDescending
                    End If
            Else
                    If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       intPreSort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       intPreSort = flexSortStringNoCaseDescending
                    End If
            End If
            intPreCol = intCol
            .Row = FindRow(mshFilter, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

Private Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .Rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function

Private Function GetParentWindow(ByVal hwndFrm As Long, Optional ByVal blnParent As Boolean = False) As Long
    Dim strClass As String * 256
    Dim blnCBR As Boolean   '是Coolbar
    
    On Error Resume Next
    
    '获取指定窗体的父窗体的父窗体
    '如果其父窗体不是Form，继续向上找
    'blnParent表示仅取其父窗体

    Do While True
        hwndFrm = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
        Call GetClassName(hwndFrm, strClass, 255)

        If Not OS.IsDesinMode Then
            If (strClass Like "ThunderRT6FormDC*") Then Exit Do
        Else
            If (strClass Like "ThunderFormDC*") Then Exit Do
        End If
    Loop
    GetParentWindow = hwndFrm
End Function

