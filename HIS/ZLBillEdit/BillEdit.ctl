VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BillEdit 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  '透明
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   6435
   ToolboxBitmap   =   "BillEdit.ctx":0000
   Begin VB.CommandButton CmdSelect 
      Caption         =   "…"
      Height          =   285
      Left            =   1485
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "↓…"
      ToolTipText     =   "请用""*""键激活"
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.MonthView MonView 
      Height          =   2220
      Left            =   2295
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1695
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483629
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   93257729
      TitleBackColor  =   10841698
      TitleForeColor  =   16777215
      CurrentDate     =   36395
   End
   Begin VB.TextBox TxtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   540
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox CboSelect 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "CboSelect"
      Top             =   3375
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSF 
      Height          =   1395
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2461
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorSel    =   10249818
      BackColorBkg    =   -2147483643
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "BillEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'下列API及类型定义，用于将选择器显示于父窗体中
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const GWL_HWNDPARENT = (-8)
Private Const HWND_TOP = 0
Private Type POINTAPI
        x As Long
        y As Long
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'其他API定义
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

'自定义变量
Private lngParent As Long       '保存父窗体指针
Private Lop As Integer
Private UserCol As Integer
Private BlnFlash As Boolean '是否已经全部刷新
Private LngColorPri As Long '上一次的颜色
Private BlnCancel As Boolean
Private PreCellColor As Long
Private strFind As String
Private blnShow As Boolean
'定义公共变量，集合
Public lngColor As Long '设置不修改列的颜色
Public CommColl As Collection '提供给用户的集合
Public AutoRefresh As Boolean '是否启动自动刷新
Public CurCellBackColor As Long '选中单元格的颜色
Public MsfObj As Object
Public cboObj As Object
'缺省属性值:
Const m_def_AllowAddRow = True
Const m_def_CheckChar = "√"
Const M_DEF_TextMask = ""
Const M_DEF_TextLen = 0
Const M_DEF_CellBackColor = 0
Const M_DEF_Active = False
Const M_DEF_CmdEnable = True
Const M_DEF_CboEnable = True
Const M_DEF_TxtCheck = False
Const M_DEF_Enabled = 0
Const M_DEF_BackStyle = 0
Const M_DEF_BorderStyle = 0
Const M_DEF_Col = 0
Const M_DEF_Row = 0
Const M_DEF_PrimaryCol = 0
Const M_DEF_LocateCol = 0
Const M_DEF_LastRow = 0
Const M_DEF_LastCol = 0
Const M_DEF_CellAlignment = 0
Const M_DEF_CboText = ""
Const M_DEF_CmdVisible = False
Const M_DEF_CboVisible = False
Const M_DEF_MonVisible = False
Const M_DEF_MonEnable = True
Const M_DEF_TxtVisible = False
Const M_DEF_TxtEnable = True
'属性变量:
Dim m_AllowAddRow As Boolean
Dim m_CheckChar As String
Dim M_TextMask As String
Dim M_TextLen As Long
Dim M_CellBackColor As Variant
Dim M_Active As Boolean
Dim M_CmdEnable As Boolean
Dim M_CboEnable As Boolean
Dim M_TxtCheck As Boolean
Dim M_Enabled As Boolean
Dim M_BackStyle As Integer
Dim M_BorderStyle As Integer
Dim M_Col As Long
Dim M_Row As Long
Dim M_PrimaryCol As Long
Dim M_LocateCol As Long
Dim M_LastRow As Long
Dim M_LastCol As Long
Dim M_CellAlignment As Long
Dim M_CboText As String
Dim M_CmdVisible As Boolean
Dim M_CboVisible As Boolean
Dim M_MonVisible As Boolean
Dim M_MonEnable As Boolean
Dim M_TxtVisible As Boolean
Dim M_TxtEnable As Boolean
'事件声明:
Event DecideInput(StrInput As String, Cancel As Boolean)
Event AfterDeleteRow()
Event CellCheck(Row As Long, Col As Long)
Event BeforeLostFocus(Cancel As Boolean)
Event BeforeDeleteRow(Row As Long, Cancel As Boolean)
Event BeforeAddRow(Row As Long)
Event AfterAddRow(Row As Long)
Event EnterCell(Row As Long, Col As Long) 'MappingInfo=Msf,Msf,-1,EnterCell
Event LeaveCell(Row As Long, Col As Long)
Event Click()
Attribute Click.VB_Description = "当用户在一个对象上按下并释放鼠标按钮时发生。"
Event DblClick(Cancel As Boolean)
Event KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "当用户按下和释放 ANSI 键时发生。"
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "当用户在拥有焦点的对象上释放键时发生。"
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "当用户在拥有焦点的对象上按下鼠标按钮时发生。"
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "当用户移动鼠标时发生。"
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "当用户在拥有焦点的对象上释放鼠标发生。"
Event CommandClick() 'MappingInfo=CmdSelect,CmdSelect,-1,Click
Event cboKeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=CboSelect,CboSelect,-1,KeyDown
Event cboClick(ListIndex As Long)
Event EditChange(curText As String)  'MappingInfo=TxtEdit,TxtEdit,-1,Change
Event EditKeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TxtEdit,TxtEdit,-1,KeyDown
Event EditKeyPress(KeyAscii As Integer)
Event EditVisible(Text As String)

'------------------------------------------------------------------------------
'刘兴洪 问题:修改了Combox的处理 日期:2010-01-28 10:26:14
'     1.将ComSelect的Style改为了0
'     2.增加属性:Sytle的更改
Private mblnNotAutoSearch As Boolean
Public Enum cboStyleEnum
    DropDownAndEdit = 0
    DropOlnyDown = 1
End Enum
Dim mcboStyle As cboStyleEnum
Const m_def_cboStyle = 1

'-----------------------------------------------------------------------------------------
'刘兴洪加如如下属性
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get cboStyle() As cboStyleEnum
    cboStyle = mcboStyle
End Property

Public Property Let cboStyle(ByVal New_cboStyle As cboStyleEnum)
    mcboStyle = New_cboStyle
    PropertyChanged "cboStyle"
End Property

'-----------------------------------------------------------------------------------------

Private Sub CboSelect_Click()
    RaiseEvent cboClick(CboSelect.ListIndex)
End Sub

Private Sub CboSelect_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    On Error Resume Next
    If mblnNotAutoSearch = True Then Exit Sub
    '问题28718 by lesfeng 2010-06-28 由于vbKeyF1至vbKeyF12的值与输入的p至z的值一致，以致于输入p至z无法进行后续代码处理，因此屏蔽下列代码
'    If KeyAscii >= vbKeyF1 And KeyAscii <= vbKeyF12 Then
'        If KeyAscii <> vbKeyF4 Then Exit Sub
'    End If

    If mcboStyle = DropOlnyDown Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
        End If
        If KeyAscii <> 0 Then
            lngIdx = zlControl.CboMatchIndex(CboSelect.hwnd, KeyAscii)
            If lngIdx = -1 And CboSelect.ListCount > 0 Then lngIdx = 0
            If lngIdx <> -1 Then CboSelect.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub CboSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngR As Long
    RaiseEvent cboKeyDown(KeyCode, Shift)
    If Shift = vbCtrlMask Or Shift = vbAltMask Then mblnNotAutoSearch = True

    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        If CboSelect.ListIndex <> -1 Then
            msf_KeyDown vbKeyReturn, 0
        Else
            Beep
        End If
    ElseIf KeyCode = vbKeyEscape Then
        'CboVisible = False
        'If MSF.Col + 1 <= MSF.Cols - 1 Then MSF.Col = MSF.Col + 1
    ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        Select Case KeyCode
            Case 37
                If MSF.Col - 1 >= 0 Then MSF.Col = MSF.Col - 1
            Case 39
                If MSF.Col + 1 <= MSF.Cols - 1 Then MSF.Col = MSF.Col + 1
        End Select
        MSF.SetFocus
        CboVisible = False
     End If
End Sub

Private Sub CboSelect_KeyUp(KeyCode As Integer, Shift As Integer)
    mblnNotAutoSearch = False
End Sub

Private Sub CmdSelect_GotFocus()
    If MonVisible Then MonVisible = False: MSF.SetFocus
End Sub

Private Sub MonView_LostFocus()
    MonVisible = False: MSF.SetFocus
End Sub

Private Sub MonView_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then MonView_DblClick
    If KeyCode = vbKeyEscape Then MonVisible = False: MSF.SetFocus
End Sub

Private Sub MonView_DateClick(ByVal DateClicked As Date)
    TxtEdit.Text = Format(MonView.Value, "yyyy-MM-dd")
    MSF.TextMatrix(MSF.Row, MSF.Col) = Format(MonView.Value, "yyyy-MM-dd")
End Sub

Private Sub MonView_DblClick()
    TxtEdit.Text = Format(MonView.Value, "yyyy-MM-dd")
    MSF.TextMatrix(MSF.Row, MSF.Col) = Format(MonView.Value, "yyyy-MM-dd")
    msf_KeyDown vbKeyReturn, 0
End Sub

Private Sub MSF_GotFocus()
    Call msf_EnterCell
End Sub

Private Sub Msf_LeaveCell()
    Dim StrInput As String, BlnCancel As Boolean
    
    RaiseEvent LeaveCell(MSF.Row, MSF.Col)
    If PreCellColor <> Empty Then MSF.CellBackColor = PreCellColor
    MSF.FocusRect = flexFocusLight
        
    CmdVisible = False
    CboVisible = False
    TxtVisible = False: TxtEdit.Text = ""
    MonVisible = False
    
    LastCol = MSF.Col
    LastRow = MSF.Row
End Sub

Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub MSF_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub MSF_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub Msf_Scroll()
    CmdVisible = False
    CboVisible = False
    MonVisible = False
    TxtVisible = False: TxtEdit.Text = ""
End Sub

Private Sub TxtEdit_Validate(Cancel As Boolean)
    Dim BlnCancel As Boolean
    
    BlnCancel = False
    '先执行用户的代码
    RaiseEvent KeyDown(vbKeyReturn, 0, BlnCancel)
    '判断当前列值是否为4,2,3；是则赋值
    If BlnCancel Then
        If TxtVisible Then
            TxtEdit.SelStart = 0: TxtEdit.SelLength = Len(TxtEdit.Text)
        End If
    End If
    Cancel = BlnCancel
        
    '为各控件赋值
    If Cancel = False Then
        If MonView.Visible Then
            TxtEdit.Text = Format(MonView.Value, "yyyy-MM-dd")
            MSF.TextMatrix(MSF.Row, MSF.Col) = Format(MonView.Value, "yyyy-MM-dd")
        End If
        If CboSelect.Visible And CboSelect.ListIndex <> -1 Then MSF.TextMatrix(MSF.Row, MSF.Col) = CboSelect.Text
        If TxtEdit.Visible Then MSF.TextMatrix(MSF.Row, MSF.Col) = TxtEdit.Text
    End If
End Sub

Private Sub UserControl_Initialize()
    lngColor = &H80000014
    'CurCellBackColor = &H80000014
    SetColor True, lngColor
    Set MsfObj = MSF
    Set cboObj = CboSelect
    Call SetFixCenter
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If CmdVisible And KeyAscii = Asc("*") Then KeyAscii = 0: Call CmdSelect_Click
End Sub

Private Sub UserControl_LostFocus()
    CmdVisible = False
    CboVisible = False
    TxtVisible = False
    MonVisible = False
    TxtEdit.Text = ""
End Sub

Private Sub UserControl_Resize()
    MSF.Top = 0
    MSF.Left = 0
    MSF.Height = UserControl.Height
    MSF.Width = UserControl.Width
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = M_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    M_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = M_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    M_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
    BorderStyle = M_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    M_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "强制完全重画一个对象。"
     
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=CboSelect,CboSelect,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "添加一项到 Listbox 或 ComboBox 控件，或添加一行到 Grid 控件。"
    CboSelect.AddItem Item, Index
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=CboSelect,CboSelect,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "清除控件或系统剪贴板的内容。"
    CboSelect.Clear
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=CboSelect,CboSelect,-1,RemoveItem
Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "从一个 ListBox 或 ComboBox 控件或一个 Grid 控件中的一行中删除一项。"
    CboSelect.RemoveItem Index
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=CboSelect,CboSelect,-1,ItemData
Public Property Get ItemData(ByVal Index As Integer) As Long
Attribute ItemData.VB_Description = "返回/设置 ComboBox 或 ListBox 控件中每一个项的指定号。"
    ItemData = CboSelect.ItemData(Index)
End Property

Public Property Let ItemData(ByVal Index As Integer, ByVal New_ItemData As Long)
    CboSelect.ItemData(Index) = New_ItemData
    PropertyChanged "ItemData"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=CboSelect,CboSelect,-1,NewIndex
Public Property Get NewIndex() As Integer
Attribute NewIndex.VB_Description = "返回添加到控件中的最近一个项目的索引。"
    NewIndex = CboSelect.NewIndex
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get Col() As Long
Attribute Col.VB_Description = "MSHFlexGrid当前行值"
    Col = MSF.Col
End Property

Public Property Let Col(ByVal New_Col As Long)
    Msf_LeaveCell
    M_Col = New_Col
    MSF.Col = M_Col
    PropertyChanged "Col"
    msf_EnterCell
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get Row() As Long
Attribute Row.VB_Description = "MSHFlexGrid的当前行值"
    Row = MSF.Row
End Property

Public Property Let Row(ByVal New_Row As Long)
    Msf_LeaveCell
    M_Row = New_Row
    MSF.Row = M_Row
    PropertyChanged "Row"
    msf_EnterCell
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get TopRow() As Long
    TopRow = MSF.TopRow
End Property

Public Property Let TopRow(ByVal New_TopRow As Long)
    MSF.TopRow = New_TopRow
    PropertyChanged "TopRow"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get PrimaryCol() As Long
Attribute PrimaryCol.VB_Description = "（主列）如果该列不为空，到最后一列则自动增加行；否则不增加行。"
    PrimaryCol = M_PrimaryCol
End Property

Public Property Let PrimaryCol(ByVal New_PrimaryCol As Long)
    If New_PrimaryCol >= 0 And New_PrimaryCol <= MSF.Cols - 1 Then
        M_PrimaryCol = New_PrimaryCol
        PropertyChanged "PrimaryCol"
    End If
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get LocateCol() As Long
Attribute LocateCol.VB_Description = "当选择了不可选择的列，则定位到该列"
    LocateCol = M_LocateCol
End Property

Public Property Let LocateCol(ByVal New_LocateCol As Long)
    If New_LocateCol >= 0 And New_LocateCol <= MSF.Cols - 1 Then
        M_LocateCol = New_LocateCol
        PropertyChanged "LocateCol"
    End If
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get LastRow() As Long
Attribute LastRow.VB_Description = "跟踪反映上一次的行值"
    LastRow = M_LastRow
End Property

Public Property Let LastRow(ByVal New_LastRow As Long)
    M_LastRow = New_LastRow
    PropertyChanged "LastRow"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get LastCol() As Long
Attribute LastCol.VB_Description = "跟踪反映上一次的列值"
    LastCol = M_LastCol
End Property

Public Property Let LastCol(ByVal New_LastCol As Long)
    M_LastCol = New_LastCol
    PropertyChanged "LastCol"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get CellAlignment() As Long
Attribute CellAlignment.VB_Description = "单元格的对齐方向"
    M_CellAlignment = MSF.CellAlignment
    CellAlignment = M_CellAlignment
End Property

Public Property Let CellAlignment(ByVal New_CellAlignment As Long)
    M_CellAlignment = New_CellAlignment
    MSF.CellAlignment = M_CellAlignment
    PropertyChanged "CellAlignment"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=TxtEdit,TxtEdit,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "返回/设置控件中包含的文本。"
    Text = TxtEdit.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    TxtEdit.Text() = New_Text
    PropertyChanged "Text"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,TextMatrix
Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
Attribute TextMatrix.VB_Description = "Returns or sets the text content of an arbitrary cell (row/column subscripts)."
    TextMatrix = MSF.TextMatrix(Row, Col)
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, ByVal New_TextMatrix As String)
    MSF.TextMatrix(Row, Col) = New_TextMatrix
    PropertyChanged "TextMatrix"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14
Public Function MergeRow(ByVal Row As Long, ByVal VarBool As Boolean) As Boolean
    MSF.MergeRow(Row) = VarBool
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14
Public Function MergeCol(ByVal Col As Long, ByVal VarBool As Boolean) As Boolean
    MSF.MergeCol(Col) = VarBool
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,1,0,
Public Property Get CboText() As String
Attribute CboText.VB_Description = "下拉框的当前值"
    CboText = CboSelect.Text
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14
Public Function MergeCell(ByVal Merge As Long) As Boolean
    On Error Resume Next

    MSF.MergeCells = Merge
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14
Public Function ClearMsf() As Variant
    MSF.Clear
    CmdVisible = False
    CboVisible = False
    TxtVisible = False: TxtEdit.Text = ""
    MonVisible = False
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get CmdVisible() As Boolean
Attribute CmdVisible.VB_Description = "按钮的Visible属性值"
    CmdVisible = M_CmdVisible
End Property

Public Property Let CmdVisible(ByVal New_CmdVisible As Boolean)
    M_CmdVisible = New_CmdVisible
    CmdSelect.Visible = M_CmdVisible
    PropertyChanged "CmdVisible"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get CboVisible() As Boolean
Attribute CboVisible.VB_Description = "下拉框的Visible属性值"
    CboVisible = M_CboVisible
End Property

Public Property Let CboVisible(ByVal New_CboVisible As Boolean)
    M_CboVisible = New_CboVisible
    CboSelect.Visible = M_CboVisible
    PropertyChanged "CboVisible"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get MonVisible() As Boolean
Attribute MonVisible.VB_Description = "日期控件的Visible属性值"
    MonVisible = M_MonVisible
End Property

Public Property Let MonVisible(ByVal New_MonVisible As Boolean)
    M_MonVisible = New_MonVisible
    MonView.Visible = M_MonVisible
    PropertyChanged "MonVisible"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get MonEnable() As Boolean
Attribute MonEnable.VB_Description = "日期控件的Enable属性值"
    MonEnable = M_MonEnable
End Property

Public Property Let MonEnable(ByVal New_MonEnable As Boolean)
    M_MonEnable = New_MonEnable
    MonView.Enabled = M_MonEnable
    PropertyChanged "MonEnable"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get TxtVisible() As Boolean
Attribute TxtVisible.VB_Description = "文本框的Visible属性值"
    TxtVisible = M_TxtVisible
End Property

Public Property Let TxtVisible(ByVal New_TxtVisible As Boolean)
    Dim strTemp As String
    
    M_TxtVisible = New_TxtVisible
    If New_TxtVisible = True Then
        '允许在显示之前决定文本框的值
        strTemp = TxtEdit.Text
        RaiseEvent EditVisible(strTemp)
        TxtEdit.Text = strTemp
        TxtEdit.SelStart = 0: TxtEdit.SelLength = Len(TxtEdit.Text)
    End If
    
    TxtEdit.Visible = M_TxtVisible
    PropertyChanged "TxtVisible"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get TxtEnable() As Boolean
Attribute TxtEnable.VB_Description = "文本框的Enable属性值"
    TxtEnable = M_TxtEnable
End Property

Public Property Let TxtEnable(ByVal New_TxtEnable As Boolean)
    M_TxtEnable = New_TxtEnable
    TxtEdit.Enabled = M_TxtEnable
    PropertyChanged "TxtEnable"
End Property

Private Sub CmdSelect_Click()
    Select Case MSF.ColData(MSF.Col)
        Case 1
            RaiseEvent CommandClick
        Case 2
            On Error Resume Next
            With MonView
                If Not .Visible Then
                    If MSF.TextMatrix(MSF.Row, MSF.Col) <> "" Then
                        If IsDate(MSF.TextMatrix(MSF.Row, MSF.Col)) Then .Value = CDate(MSF.TextMatrix(MSF.Row, MSF.Col))
                    End If
                    .Left = CmdSelect.Left - MonView.Width / 2
                    .Top = CmdSelect.Top + CmdSelect.Height
                    Call AdjustData
                    MonVisible = True
                    .SetFocus
                Else
                    MonVisible = False
                    MSF.SetFocus
                End If
            End With
    End Select
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent EditKeyDown(KeyCode, Shift)
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    M_Enabled = M_DEF_Enabled
    M_BackStyle = M_DEF_BackStyle
    M_BorderStyle = M_DEF_BorderStyle
    M_Col = M_DEF_Col
    M_Row = M_DEF_Row
    M_PrimaryCol = M_DEF_PrimaryCol
    M_LocateCol = M_DEF_LocateCol
    M_LastRow = M_DEF_LastRow
    M_LastCol = M_DEF_LastCol
    M_CellAlignment = M_DEF_CellAlignment
    M_CboText = M_DEF_CboText
    M_CmdVisible = M_DEF_CmdVisible
    M_CboVisible = M_DEF_CboVisible
    M_MonVisible = M_DEF_MonVisible
    M_MonEnable = M_DEF_MonEnable
    M_TxtVisible = M_DEF_TxtVisible
    M_TxtEnable = M_DEF_TxtEnable
    M_TxtCheck = M_DEF_TxtCheck
    M_Active = M_DEF_Active
    M_CmdEnable = M_DEF_CmdEnable
    M_CboEnable = M_DEF_CboEnable
    M_CellBackColor = M_DEF_CellBackColor
    M_TextMask = M_DEF_TextMask
    M_TextLen = M_DEF_TextLen
    m_CheckChar = m_def_CheckChar
    m_AllowAddRow = m_def_AllowAddRow
    mcboStyle = m_def_cboStyle
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer
    On Error GoTo ErrHand
    
    M_Enabled = PropBag.ReadProperty("Enabled", M_DEF_Enabled)
    M_BackStyle = PropBag.ReadProperty("BackStyle", M_DEF_BackStyle)
    M_BorderStyle = PropBag.ReadProperty("BorderStyle", M_DEF_BorderStyle)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    CboSelect.ItemData(Index) = PropBag.ReadProperty("ItemData" & Index, 0)
    MSF.Appearance = PropBag.ReadProperty("Appearance", 1)
    M_Col = PropBag.ReadProperty("Col", M_DEF_Col)
    M_Row = PropBag.ReadProperty("Row", M_DEF_Row)
    M_PrimaryCol = PropBag.ReadProperty("PrimaryCol", M_DEF_PrimaryCol)
    M_LocateCol = PropBag.ReadProperty("LocateCol", M_DEF_LocateCol)
    M_LastRow = PropBag.ReadProperty("LastRow", M_DEF_LastRow)
    M_LastCol = PropBag.ReadProperty("LastCol", M_DEF_LastCol)
    M_CellAlignment = PropBag.ReadProperty("CellAlignment", M_DEF_CellAlignment)
    TxtEdit.Text = PropBag.ReadProperty("Text", "Text1")
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    MSF.TextMatrix(Row, Col) = PropBag.ReadProperty("TextMatrix" & Index, "0")
    M_CboText = PropBag.ReadProperty("CboText", M_DEF_CboText)
    M_CmdVisible = PropBag.ReadProperty("CmdVisible", M_DEF_CmdVisible)
    M_CboVisible = PropBag.ReadProperty("CboVisible", M_DEF_CboVisible)
    M_MonVisible = PropBag.ReadProperty("MonVisible", M_DEF_MonVisible)
    M_MonEnable = PropBag.ReadProperty("MonEnable", M_DEF_MonEnable)
    M_TxtVisible = PropBag.ReadProperty("TxtVisible", M_DEF_TxtVisible)
    M_TxtEnable = PropBag.ReadProperty("TxtEnable", M_DEF_TxtEnable)
    M_TxtCheck = PropBag.ReadProperty("TxtCheck", M_DEF_TxtCheck)
    MonView.MaxDate = PropBag.ReadProperty("MaxDate", 9999 - 12 - 31)
    MonView.MinDate = PropBag.ReadProperty("MinDate", 1753 - 1 - 1)
    MonView.Value = PropBag.ReadProperty("Value", 1999 - 8 - 23)
    M_Active = PropBag.ReadProperty("Active", M_DEF_Active)
    M_CmdEnable = PropBag.ReadProperty("CmdEnable", M_DEF_CmdEnable)
    M_CboEnable = PropBag.ReadProperty("CboEnable", M_DEF_CboEnable)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    MSF.Cols = PropBag.ReadProperty("Cols", 4)
    MSF.Rows = PropBag.ReadProperty("Rows", 2)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    MSF.RowData(Index) = PropBag.ReadProperty("RowData" & Index, 0)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    MSF.ColData(Index) = PropBag.ReadProperty("ColData" & Index, 0)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    MSF.RowHeight(Index) = PropBag.ReadProperty("RowHeight" & Index, 0)
    MSF.RowHeightMin = PropBag.ReadProperty("RowHeightMin", 0)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    MSF.ColWidth(Index) = PropBag.ReadProperty("ColWidth" & Index, 0)
    MSF.BackColor = PropBag.ReadProperty("BackColor", &H80000014)
    MSF.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    MSF.BackColorBkg = PropBag.ReadProperty("BackColorBkg", 2147483668#)
    MSF.BackColorSel = PropBag.ReadProperty("BackColorSel", 2147483661#)
    MSF.BackColorFixed = PropBag.ReadProperty("BackColorFixed", 2147483653#)
    MSF.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", 2147483666#)
    MSF.ForeColorSel = PropBag.ReadProperty("ForeColorSel", 2147483662#)
    MSF.GridColor = PropBag.ReadProperty("GridColor", 2147483666#)
    MSF.GridColorFixed = PropBag.ReadProperty("GridColorFixed", 12632256)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    MSF.ColAlignment(Index) = PropBag.ReadProperty("ColAlignment" & Index, 0)
    CboSelect.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    TxtEdit.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    TxtEdit.SelLength = PropBag.ReadProperty("SelLength", 0)
    TxtEdit.SelStart = PropBag.ReadProperty("SelStart", 0)
    TxtEdit.BackColor = PropBag.ReadProperty("TxtBackColor", &H80000009)
    M_CellBackColor = PropBag.ReadProperty("CellBackColor", M_DEF_CellBackColor)
    m_CheckChar = PropBag.ReadProperty("CheckChar", m_def_CheckChar)
    m_AllowAddRow = PropBag.ReadProperty("AllowAddRow", m_def_AllowAddRow)
    
    '刘兴洪增加
    mcboStyle = PropBag.ReadProperty("cboStyle", m_def_cboStyle)
    
    
    SetColor True, lngColor
    Exit Sub

ErrHand:
    If Err = 381 Then Resume Next
    Set MSF.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set CboSelect.Font = PropBag.ReadProperty("CboFont", Ambient.Font)
    Set TxtEdit.Font = PropBag.ReadProperty("TxtEditFont", Ambient.Font)
    Set DataSource = PropBag.ReadProperty("MsfDataSource", Nothing)
    MSF.DataMember = PropBag.ReadProperty("DataMember", "")
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    M_TextMask = PropBag.ReadProperty("TextMask", M_DEF_TextMask)
    M_TextLen = PropBag.ReadProperty("TextLen", M_DEF_TextLen)
    m_CheckChar = PropBag.ReadProperty("CheckChar", m_def_CheckChar)
    m_AllowAddRow = PropBag.ReadProperty("AllowAddRow", m_def_AllowAddRow)
    
    '刘兴洪增加
    mcboStyle = PropBag.ReadProperty("cboStyle", m_def_cboStyle)
    
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get TxtCheck() As Boolean
    TxtCheck = M_TxtCheck
End Property

Public Property Let TxtCheck(ByVal New_TxtCheck As Boolean)
    M_TxtCheck = New_TxtCheck
    PropertyChanged "TxtCheck"
End Property
'注意！不要删除或修改下列被注释的行！
'MappingInfo=MonView,MonView,-1,MaxDate
Public Property Get MaxDate() As Date
    MaxDate = MonView.MaxDate
End Property

Public Property Let MaxDate(ByVal New_MaxDate As Date)
    MonView.MaxDate() = New_MaxDate
    PropertyChanged "MaxDate"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=MonView,MonView,-1,MinDate
Public Property Get MinDate() As Date
    MinDate = MonView.MinDate
End Property

Public Property Let MinDate(ByVal New_MinDate As Date)
    MonView.MinDate() = New_MinDate
    PropertyChanged "MinDate"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=MonView,MonView,-1,Value
Public Property Get Value() As Date
    Value = MonView.Value
End Property

Public Property Let Value(ByVal New_Value As Date)
    MonView.Value() = New_Value
    PropertyChanged "Value"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get Active() As Boolean
    Active = M_Active
End Property

Public Property Let Active(ByVal New_Active As Boolean)
    Dim intIndex As Integer
    M_Active = New_Active
    PropertyChanged "Active"
    If M_Active = False Then
        CmdVisible = False
        TxtVisible = False: TxtEdit.Text = ""
        intIndex = CboSelect.ListIndex
        CboVisible = False
        '问题:30998
        If intIndex <> CboSelect.ListIndex Then
            DoEvents: CboSelect.ListIndex = intIndex
        End If
        MonVisible = False
    End If
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get CmdEnable() As Boolean
    CmdEnable = M_CmdEnable
End Property

Public Property Let CmdEnable(ByVal New_CmdEnable As Boolean)
    M_CmdEnable = New_CmdEnable
    CmdSelect.Enabled = M_CmdEnable
    PropertyChanged "CmdEnable"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get CboEnable() As Boolean
    CboEnable = M_CboEnable
End Property

Public Property Let CboEnable(ByVal New_CboEnable As Boolean)
    M_CboEnable = New_CboEnable
    CboSelect.Enabled = M_CboEnable
    PropertyChanged "CboEnable"
End Property

Private Sub UserControl_Terminate()
    Set MsfObj = Nothing
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer
    On Error GoTo ErrHand
    
    Call PropBag.WriteProperty("Enabled", M_Enabled, M_DEF_Enabled)
    Call PropBag.WriteProperty("BackStyle", M_BackStyle, M_DEF_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", M_BorderStyle, M_DEF_BorderStyle)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    Call PropBag.WriteProperty("ItemData" & Index, CboSelect.ItemData(Index), 0)
    Call PropBag.WriteProperty("Appearance", MSF.Appearance, 1)
    Call PropBag.WriteProperty("Col", M_Col, M_DEF_Col)
    Call PropBag.WriteProperty("Row", M_Row, M_DEF_Row)
    Call PropBag.WriteProperty("PrimaryCol", M_PrimaryCol, M_DEF_PrimaryCol)
    Call PropBag.WriteProperty("LocateCol", M_LocateCol, M_DEF_LocateCol)
    Call PropBag.WriteProperty("LastRow", M_LastRow, M_DEF_LastRow)
    Call PropBag.WriteProperty("LastCol", M_LastCol, M_DEF_LastCol)
    Call PropBag.WriteProperty("CellAlignment", M_CellAlignment, M_DEF_CellAlignment)
    Call PropBag.WriteProperty("Text", TxtEdit.Text, "Text1")
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    Call PropBag.WriteProperty("TextMatrix" & Index, MSF.TextMatrix(Row, Col), "0")
    Call PropBag.WriteProperty("CboText", M_CboText, M_DEF_CboText)
    Call PropBag.WriteProperty("CmdVisible", M_CmdVisible, M_DEF_CmdVisible)
    Call PropBag.WriteProperty("CboVisible", M_CboVisible, M_DEF_CboVisible)
    Call PropBag.WriteProperty("MonVisible", M_MonVisible, M_DEF_MonVisible)
    Call PropBag.WriteProperty("MonEnable", M_MonEnable, M_DEF_MonEnable)
    Call PropBag.WriteProperty("TxtVisible", M_TxtVisible, M_DEF_TxtVisible)
    Call PropBag.WriteProperty("TxtEnable", M_TxtEnable, M_DEF_TxtEnable)
    Call PropBag.WriteProperty("TxtCheck", M_TxtCheck, M_DEF_TxtCheck)
    Call PropBag.WriteProperty("TxtCheck", M_TxtCheck, M_DEF_TxtCheck)
    Call PropBag.WriteProperty("MaxDate", MonView.MaxDate, 9999 - 12 - 31)
    Call PropBag.WriteProperty("MinDate", MonView.MinDate, 1753 - 1 - 1)
    Call PropBag.WriteProperty("Value", MonView.Value, 1999 - 8 - 23)
    Call PropBag.WriteProperty("Active", M_Active, M_DEF_Active)
    Call PropBag.WriteProperty("CmdEnable", M_CmdEnable, M_DEF_CmdEnable)
    Call PropBag.WriteProperty("CboEnable", M_CboEnable, M_DEF_CboEnable)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    Call PropBag.WriteProperty("Cols", MSF.Cols, 4)
    Call PropBag.WriteProperty("Rows", MSF.Rows, 2)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    Call PropBag.WriteProperty("RowData" & Index, MSF.RowData(Index), 0)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    Call PropBag.WriteProperty("ColData" & Index, MSF.ColData(Index), 0)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    Call PropBag.WriteProperty("RowHeight" & Index, MSF.RowHeight(Index), 0)
    Call PropBag.WriteProperty("RowHeightMin", MSF.RowHeightMin, 0)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    Call PropBag.WriteProperty("ColWidth" & Index, MSF.ColWidth(Index), 0)
    Call PropBag.WriteProperty("BackColor", MSF.BackColor, &H80000014)
    Call PropBag.WriteProperty("ForeColor", MSF.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BackColorBkg", MSF.BackColorBkg, 2147483668#)
    Call PropBag.WriteProperty("BackColorSel", MSF.BackColorSel, 2147483661#)
    Call PropBag.WriteProperty("BackColorFixed", MSF.BackColorFixed, 2147483653#)
    Call PropBag.WriteProperty("ForeColorFixed", MSF.ForeColorFixed, 2147483666#)
    Call PropBag.WriteProperty("ForeColorSel", MSF.ForeColorSel, 2147483662#)
    Call PropBag.WriteProperty("GridColor", MSF.GridColor, 2147483666#)
    Call PropBag.WriteProperty("GridColorFixed", MSF.GridColorFixed, 12632256)
'TO DO: 你映射到的成员包含数据数组。
'   您必须提供代码来保持数组。
'   以下显示原型行:
    Call PropBag.WriteProperty("ColAlignment" & Index, MSF.ColAlignment(Index), 0)
    Call PropBag.WriteProperty("ListIndex", CboSelect.ListIndex, 0)
    Call PropBag.WriteProperty("MaxLength", TxtEdit.MaxLength, 0)
    Call PropBag.WriteProperty("SelLength", TxtEdit.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", TxtEdit.SelStart, 0)
    Call PropBag.WriteProperty("TxtBackColor", TxtEdit.BackColor, &H80000009)
    Call PropBag.WriteProperty("CellBackColor", M_CellBackColor, M_DEF_CellBackColor)
        
    '刘兴洪加入:
    Call PropBag.WriteProperty("cboStyle", mcboStyle, m_def_cboStyle)
    
    Exit Sub

ErrHand:
    If Err = 381 Then Resume Next
    Call PropBag.WriteProperty("Font", MSF.Font, Ambient.Font)
    Call PropBag.WriteProperty("CboFont", CboSelect.Font, Ambient.Font)
    Call PropBag.WriteProperty("TxtEditFont", TxtEdit.Font, Ambient.Font)
    Call PropBag.WriteProperty("MsfDataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataMember", MSF.DataMember, "")
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("TextMask", M_TextMask, M_DEF_TextMask)
    Call PropBag.WriteProperty("TextLen", M_TextLen, M_DEF_TextLen)
    Call PropBag.WriteProperty("CheckChar", m_CheckChar, m_def_CheckChar)
    Call PropBag.WriteProperty("AllowAddRow", m_AllowAddRow, m_def_AllowAddRow)
    '刘兴洪加入:
    Call PropBag.WriteProperty("cboStyle", mcboStyle, m_def_cboStyle)

End Sub
'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,Cols
Public Property Get Cols() As Long
Attribute Cols.VB_Description = "Determines the total number of columns or rows in the Hierarchical FlexGrid."
    Cols = MSF.Cols
End Property

Public Property Let Cols(ByVal New_Cols As Long)
    On Error Resume Next
    MSF.Cols = New_Cols
    PropertyChanged "Cols"
    Call SetFixCenter
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,Rows
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Determines the total number of columns or rows in the Hierarchical FlexGrid."
    Rows = MSF.Rows
End Property

Public Property Let Rows(ByVal New_Rows As Long)
    On Error Resume Next
    
    MSF.Rows = New_Rows
    PropertyChanged "Rows"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,RowData
Public Property Get RowData(ByVal Index As Long) As Long
Attribute RowData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the Hierarchical FlexGrid. Not available at design time."
    RowData = MSF.RowData(Index)
End Property

Public Property Let RowData(ByVal Index As Long, ByVal New_RowData As Long)
    MSF.RowData(Index) = New_RowData
    PropertyChanged "RowData"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,ColData
Public Property Get ColData(ByVal Index As Long) As Long
Attribute ColData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the Hierarchical FlexGrid. Not available at design time."
    ColData = MSF.ColData(Index)
End Property

Public Property Let ColData(ByVal Index As Long, ByVal New_ColData As Long)
    MSF.ColData(Index) = New_ColData
    PropertyChanged "ColData"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,RowHeight
Public Property Get RowHeight(ByVal Index As Long) As Long
Attribute RowHeight.VB_Description = "Returns or sets the height of the specified row, in Twips. Not available at design time."
    RowHeight = MSF.RowHeight(Index)
End Property

Public Property Let RowHeight(ByVal Index As Long, ByVal New_RowHeight As Long)
    MSF.RowHeight(Index) = New_RowHeight
    PropertyChanged "RowHeight"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,RowHeightMin
Public Property Get RowHeightMin() As Long
Attribute RowHeightMin.VB_Description = "Returns or sets a minimum row height for the entire control, in Twips."
    RowHeightMin = MSF.RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
    MSF.RowHeightMin() = New_RowHeightMin
    PropertyChanged "RowHeightMin"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,ColWidth
Public Property Get ColWidth(ByVal Index As Long) As Long
Attribute ColWidth.VB_Description = "Determines the width of the specified column, in Twips. Not available at design time."
    If Index <= MSF.Cols - 1 Then ColWidth = MSF.ColWidth(Index)
End Property

Public Property Let ColWidth(ByVal Index As Long, ByVal New_ColWidth As Long)
    If Index <= MSF.Cols - 1 Then
        MSF.ColWidth(Index) = New_ColWidth
        PropertyChanged "ColWidth"
    End If
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
    BackColor = MSF.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    MSF.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Determines the color used to draw text on each part of the Hierarchical FlexGrid."
    ForeColor = MSF.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    MSF.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,BackColorBkg
Public Property Get BackColorBkg() As Long
Attribute BackColorBkg.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
    BackColorBkg = MSF.BackColorBkg
End Property

Public Property Let BackColorBkg(ByVal New_BackColorBkg As Long)
    MSF.BackColorBkg() = New_BackColorBkg
    PropertyChanged "BackColorBkg"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,BackColorSel
Public Property Get BackColorSel() As Long
Attribute BackColorSel.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
    BackColorSel = MSF.BackColorSel
End Property

Public Property Let BackColorSel(ByVal New_BackColorSel As Long)
    MSF.BackColorSel() = New_BackColorSel
    PropertyChanged "BackColorSel"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,BackColorFixed
Public Property Get BackColorFixed() As Long
Attribute BackColorFixed.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
    BackColorFixed = MSF.BackColorFixed
End Property

Public Property Let BackColorFixed(ByVal New_BackColorFixed As Long)
    MSF.BackColorFixed() = New_BackColorFixed
    PropertyChanged "BackColorFixed"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,ForeColorFixed
Public Property Get ForeColorFixed() As Long
Attribute ForeColorFixed.VB_Description = "Determines the color used to draw text on each part of the Hierarchical FlexGrid."
    ForeColorFixed = MSF.ForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal New_ForeColorFixed As Long)
    MSF.ForeColorFixed() = New_ForeColorFixed
    PropertyChanged "ForeColorFixed"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,ForeColorSel
Public Property Get ForeColorSel() As Long
Attribute ForeColorSel.VB_Description = "Determines the color used to draw text on each part of the Hierarchical FlexGrid."
    ForeColorSel = MSF.ForeColorSel
End Property

Public Property Let ForeColorSel(ByVal New_ForeColorSel As Long)
    MSF.ForeColorSel() = New_ForeColorSel
    PropertyChanged "ForeColorSel"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,GridColor
Public Property Get GridColor() As Long
Attribute GridColor.VB_Description = "Returns or sets the color used to draw the lines between Hierarchical FlexGrid cells."
    GridColor = MSF.GridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As Long)
    MSF.GridColor() = New_GridColor
    PropertyChanged "GridColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,GridColorFixed
Public Property Get GridColorFixed() As Long
Attribute GridColorFixed.VB_Description = "Returns or sets the color used to draw the lines between Hierarchical FlexGrid cells."
    GridColorFixed = MSF.GridColorFixed
End Property

Public Property Let GridColorFixed(ByVal New_GridColorFixed As Long)
    MSF.GridColorFixed() = New_GridColorFixed
    PropertyChanged "GridColorFixed"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,ColAlignment
Public Property Get ColAlignment(ByVal Index As Long) As Integer
Attribute ColAlignment.VB_Description = "Returns or sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString property)."
    If Index <= MSF.Cols - 1 Then ColAlignment = MSF.ColAlignment(Index)
End Property

Public Property Let ColAlignment(ByVal Index As Long, ByVal New_ColAlignment As Integer)
    If Index <= MSF.Cols - 1 Then
        MSF.ColAlignment(Index) = New_ColAlignment
        PropertyChanged "ColAlignment"
    End If
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=CboSelect,CboSelect,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "返回/设置该控件中当前选定项目的索引。"
    ListIndex = CboSelect.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    CboSelect.ListIndex = New_ListIndex
    PropertyChanged "ListIndex"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=CboSelect,CboSelect,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "返回控件的列表部分中的项目数。"
    ListCount = CboSelect.ListCount
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=TxtEdit,TxtEdit,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "返回/设置一个控件中可以输入的字符的最大数。"
    MaxLength = TxtEdit.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    TxtEdit.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=TxtEdit,TxtEdit,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "返回/设置选定的字符数。"
    SelLength = TxtEdit.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    TxtEdit.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=TxtEdit,TxtEdit,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "返回/设置选定文本的起始点。"
    SelStart = TxtEdit.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    TxtEdit.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

Private Sub msf_DblClick()
    If Not Active Then Exit Sub
    
    BlnCancel = False
    RaiseEvent DblClick(BlnCancel)
    If BlnCancel Then
        TxtVisible = False
        Exit Sub
    End If
    
    If MSF.ColData(MSF.Col) = 1 Or MSF.ColData(MSF.Col) = 2 Or MSF.ColData(MSF.Col) = 4 Then
        Select Case MSF.ColAlignment(MSF.Col)
            Case 0, 1, 2
                TxtEdit.Alignment = vbLeftJustify
            Case 3, 4, 5
                TxtEdit.Alignment = vbCenter
            Case 6, 7, 8
                TxtEdit.Alignment = vbRightJustify
            Case Else
                TxtEdit.Alignment = vbLeftJustify
        End Select
        
        CmdVisible = False
        MonVisible = False
        CboVisible = False

        With TxtEdit
            .Left = MSF.Left + MSF.CellLeft + 15
            .Top = MSF.Top + MSF.CellTop + (MSF.CellHeight - .Height) / 2 - 15
            .Width = MSF.CellWidth - 15 - 45
            .BackColor = MSF.CellBackColor
            .Text = MSF.TextMatrix(MSF.Row, MSF.Col)
            .SelStart = 0
            .SelLength = Len(TxtEdit.Text)
        End With
        If TxtEdit.Enabled Then
            TxtVisible = True
            TxtEdit.SetFocus
        End If
    ElseIf MSF.ColData(MSF.Col) = -1 Then
        MSF.CellAlignment = 4
        If Trim(MSF.TextMatrix(MSF.Row, MSF.Col)) = "" Then
            MSF.TextMatrix(MSF.Row, MSF.Col) = CheckChar
        Else
            MSF.TextMatrix(MSF.Row, MSF.Col) = ""
        End If
        RaiseEvent CellCheck(MSF.Row, MSF.Col)
    End If
End Sub

Public Sub msf_EnterCell()
    Static lngLastRow As Long, lngLastCol As Long
    On Error Resume Next
    
    '先执行用户的代码
    RaiseEvent EnterCell(MSF.Row, MSF.Col)
    blnShow = False
    
    '统计可显示列数
    Dim IntCols As Integer, IntShowCols As Integer, BlnLoop As Boolean
    IntShowCols = 0
    For IntCols = 0 To MSF.Cols - 1
        If MSF.ColIsVisible(IntCols) Then IntShowCols = IntShowCols + 1
    Next
    IntShowCols = IIf(MSF.Col + 1 > IntShowCols, MSF.Col - IntShowCols + 2, 0)
    If IntShowCols > MSF.Cols - 1 Then IntShowCols = MSF.Cols - 1
    BlnLoop = True
    Do While BlnLoop
        If (MSF.ColData(IntShowCols) <> 0 And MSF.ColData(IntShowCols) <> -1 And MSF.ColData(IntShowCols) <> 1 And MSF.ColData(IntShowCols) <> 2 And MSF.ColData(IntShowCols) <> 3 And MSF.ColData(IntShowCols) <> 4) Or MSF.ColWidth(IntShowCols) = 0 Then
            IntShowCols = IntShowCols + 1
            If IntShowCols = MSF.Cols - 1 Then Exit Do
        Else
            BlnLoop = False
        End If
    Loop
    If IntShowCols > MSF.Cols - 1 Then IntShowCols = MSF.Cols - 1
    If IntShowCols <> 0 And (MSF.LeftCol < IntShowCols Or MSF.LeftCol > IntShowCols) And Not MSF.ColIsVisible(MSF.Col) Then MSF.LeftCol = IntShowCols
    PreCellColor = MSF.CellBackColor
    
    If Not Active Then Exit Sub
    '如果列值为-1，则为复选框
    '如果列值为1，则为按钮
    '如果列值为2，则为按钮，但显示日期控件
    '如果列值为3，则为下拉框
    '如果列值为4，则为文本框
    '如果列值为0，则用户可以选择
    '如果列值为其它值，则用户不能选择
    
    Select Case MSF.ColData(MSF.Col)
        Case 0
        Case -1
            MSF.FocusRect = flexFocusHeavy
        Case 1, 2
            With CmdSelect
                .Height = MSF.CellHeight
                .Width = .Height
                .Left = MSF.Left + MSF.CellLeft + MSF.CellWidth - CmdSelect.Width
                .Top = MSF.Top + MSF.CellTop - 15
                If MSF.ColData(MSF.Col) = 1 Then
                    .Caption = Right(.Tag, 1)
                Else
                    .Caption = Left(.Tag, 1)
                End If
            End With
            MSF.FocusRect = flexFocusHeavy
            CmdVisible = True
        Case 3
            With CboSelect
                .Left = MSF.Left + MSF.CellLeft - 15
                .Top = MSF.Top + MSF.CellTop - 15
                .Width = MSF.CellWidth - 15
                .BackColor = MSF.CellBackColor
            End With
            CboVisible = True
            CboSelect.SetFocus
            
            '清除上次录入的值
            If MSF.Col <> lngLastCol Or MSF.Row <> lngLastRow Then
                Call zlControl.CboMatchIndex(0, 0)
                lngLastRow = MSF.Row
                lngLastCol = MSF.Col
            End If
            
            '定位下拉框的值
            If MSF.TextMatrix(MSF.Row, MSF.Col) <> "" Then
                For Lop = 0 To CboSelect.ListCount - 1
                    If CboSelect.List(Lop) = MSF.TextMatrix(MSF.Row, MSF.Col) Then CboSelect.ListIndex = Lop: Exit Sub
                Next
            End If
        Case 4
            MSF.FocusRect = flexFocusHeavy
        Case Else
            Call Msf_LeaveCell
            If ColData(LocateCol) = 5 Then ColData(LocateCol) = 0
            MSF.Col = LocateCol
            Call msf_EnterCell
    End Select
End Sub

Private Sub msf_KeyDown(KeyCode As Integer, Shift As Integer)
    BlnCancel = False
    '先执行用户的代码
    RaiseEvent KeyDown(KeyCode, Shift, BlnCancel)
    '判断当前列值是否为4,2,3；是则赋值
    If BlnCancel Then
        If TxtVisible Then
            TxtEdit.SelStart = 0: TxtEdit.SelLength = Len(TxtEdit.Text)
        End If
        Exit Sub
    End If
    
    Dim BlnLoop As Boolean
    If KeyCode = vbKeyReturn Then
        
        If MonView.Visible Then
            TxtEdit.Text = Format(MonView.Value, "yyyy-MM-dd")
            MSF.TextMatrix(MSF.Row, MSF.Col) = Format(MonView.Value, "yyyy-MM-dd")
        End If
        If CboSelect.Visible And CboSelect.ListIndex <> -1 Then MSF.TextMatrix(MSF.Row, MSF.Col) = CboSelect.Text
        If TxtEdit.Visible Then MSF.TextMatrix(MSF.Row, MSF.Col) = TxtEdit.Text
        
        If (MSF.ColData(MSF.Col) = 4 Or MSF.ColData(MSF.Col) = 1) And MSF.TextMatrix(MSF.Row, MSF.Col) = "" And PrimaryCol <> MSF.Col Then
            Call Beep: Exit Sub
        ElseIf (MSF.ColData(MSF.Col) = 4 Or MSF.ColData(MSF.Col) = 1) And MSF.TextMatrix(MSF.Row, MSF.Col) = "" And PrimaryCol = MSF.Col Then
            If MSF.Row = MSF.Rows - 1 And MSF.Rows > 2 Then
                TxtVisible = False: TxtEdit.Text = ""
                CmdVisible = False
                CboVisible = False
                MonVisible = False
                PressKey (vbKeyTab): Exit Sub
            Else
                Call Beep: Exit Sub
            End If
        End If
        
        If MSF.Col = PrimaryCol And MSF.TextMatrix(MSF.Row, PrimaryCol) = "" Then
            If MSF.Rows >= 3 Then
                BlnCancel = False
                RaiseEvent BeforeLostFocus(BlnCancel)
                If Not BlnCancel Then
                    TxtVisible = False: TxtEdit.Text = ""
                    CmdVisible = False
                    CboVisible = False
                    MonVisible = False
                    PressKey (vbKeyTab)
                End If
                Exit Sub
            End If
        End If
        
        Msf_LeaveCell
Here:   If MSF.Col = MSF.Cols - 1 Then
            If MSF.Row = MSF.Rows - 1 Then
                '不允许新增行
                If AllowAddRow = False Then
                    TxtVisible = False: TxtEdit.Text = ""
                    CmdVisible = False
                    CboVisible = False
                    MonVisible = False
                    PressKey (vbKeyTab)
                    Exit Sub
                End If
                
                If Active And MSF.TextMatrix(MSF.Row, PrimaryCol) <> "" Then
                    RaiseEvent BeforeAddRow(MSF.Rows)
                    MSF.Rows = MSF.Rows + 1
                    RaiseEvent AfterAddRow(MSF.Rows - 1)
                    MSF.Row = MSF.Row + 1
                    MSF.Col = 0
                    BlnLoop = True
                    Do While BlnLoop
                        If (MSF.ColData(MSF.Col) <> 0 And MSF.ColData(MSF.Col) <> -1 And MSF.ColData(MSF.Col) <> 1 And MSF.ColData(MSF.Col) <> 2 And MSF.ColData(MSF.Col) <> 3 And MSF.ColData(MSF.Col) <> 4) Or MSF.ColWidth(MSF.Col) = 0 Then
                            MSF.Col = MSF.Col + 1
                        Else
                            BlnLoop = False
                        End If
                    Loop
                    SetColor False, lngColor, MSF.Rows - 1
                Else
                    MSF.Col = 0
                End If
            Else
                MSF.Row = MSF.Row + 1
                MSF.Col = 0
            End If
        Else
            MSF.Col = MSF.Col + 1
            If (MSF.ColData(MSF.Col) <> 0 And MSF.ColData(MSF.Col) <> -1 And MSF.ColData(MSF.Col) <> 1 And MSF.ColData(MSF.Col) <> 2 And MSF.ColData(MSF.Col) <> 3 And MSF.ColData(MSF.Col) <> 4) Or MSF.ColWidth(MSF.Col) = 0 Then GoTo Here
        End If
        
        MSF.SetFocus
        msf_EnterCell
    Else
        If KeyCode = vbKeyRight Then Exit Sub
        If KeyCode = vbKeyDelete And Active Then
            BlnCancel = False
            RaiseEvent BeforeDeleteRow(MSF.Row, BlnCancel)
            If Not BlnCancel Then
                MonVisible = False
                CmdVisible = False
                TxtVisible = False: TxtEdit.Text = ""
                CboVisible = False
                If MSF.Rows > 2 Then
                    MSF.RemoveItem MSF.Row
                    RowHeightMin = MSF.RowHeight(1)
                    RaiseEvent AfterDeleteRow
                Else
                    For Lop = 0 To MSF.Cols - 1
                        MSF.TextMatrix(MSF.Row, Lop) = ""
                    Next
                    MSF.RowData(MSF.Row) = 0
                    RaiseEvent AfterDeleteRow
                End If
            End If
            Exit Sub
        End If
        
        If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyNumlock Then Exit Sub
        If (KeyCode >= vbKeyF1 And KeyCode < vbKeyF12) Or KeyCode = vbKeyEscape Or KeyCode = vbKeyMultiply Or Shift = vbCtrlMask Or Shift = vbAltMask Then Exit Sub
        If Active And (MSF.ColData(MSF.Col) = 1 Or MSF.ColData(MSF.Col) = 2 Or MSF.ColData(MSF.Col) = 4) Then
            CmdVisible = False
            MonVisible = False
            CboVisible = False
            
            Select Case MSF.ColAlignment(MSF.Col)
                Case 0, 1, 2
                    TxtEdit.Alignment = vbLeftJustify
                Case 3, 4, 5
                    TxtEdit.Alignment = vbCenter
                Case 6, 7, 8
                    TxtEdit.Alignment = vbRightJustify
                Case Else
            End Select
            
            On Error Resume Next
            With TxtEdit
                .Left = MSF.Left + MSF.CellLeft + 15
                .Top = MSF.Top + MSF.CellTop + (MSF.CellHeight - TxtEdit.Height) / 2 - 15
                .Width = MSF.CellWidth - 15 - 45
                .BackColor = MSF.CellBackColor
                .Text = MSF.TextMatrix(MSF.Row, MSF.Col)
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            TxtVisible = True
            TxtEdit.SetFocus
        End If
    End If
End Sub

Private Sub msf_KeyPress(KeyAscii As Integer)
    '先执行用户的代码
    RaiseEvent KeyPress(KeyAscii)
    
    If (Not TxtVisible Or Not TxtEnable) And MSF.ColData(MSF.Col) <> -1 Then KeyAscii = 0: Exit Sub
    
    If Not Active Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    If MSF.ColData(MSF.Col) = 1 Or MSF.ColData(MSF.Col) = 2 Or MSF.ColData(MSF.Col) = 4 Then
        CmdVisible = False
        MonVisible = False
        CboVisible = False
        
        Select Case MSF.ColAlignment(MSF.Col)
            Case 0, 1, 2
                TxtEdit.Alignment = vbLeftJustify
            Case 3, 4, 5
                TxtEdit.Alignment = vbCenter
            Case 6, 7, 8
                TxtEdit.Alignment = vbRightJustify
            Case Else
        End Select
        
        On Error Resume Next
        With TxtEdit
            .Left = MSF.Left + MSF.CellLeft + 15
            .Top = MSF.Top + MSF.CellTop + (MSF.CellHeight - TxtEdit.Height) / 2 - 15
            .Width = MSF.CellWidth - 15 - 45
            .BackColor = MSF.CellBackColor
        End With
        TxtVisible = True
        
        Select Case KeyAscii
            Case 0 To 32
                TxtEdit.Text = MSF.TextMatrix(MSF.Row, MSF.Col)
                TxtEdit.SelStart = 0
                TxtEdit.SelLength = Len(TxtEdit.Text)
            Case Else
                If TxtCheck And TextMask <> "" Then
                    If InStr(1, TextMask, Chr(KeyAscii)) > 0 Then
                        TxtEdit.Text = Chr(KeyAscii)
                        TxtEdit.SelStart = 1
                        TxtEdit.SetFocus
                        Exit Sub
                    Else
                        KeyAscii = 0: Beep
                    End If
                Else
                    TxtEdit.Text = Chr(KeyAscii)
                    TxtEdit.SelStart = 1
                    RaiseEvent EditKeyPress(KeyAscii)
                    TxtEdit.SetFocus
                End If
                
        End Select
    ElseIf MSF.ColData(MSF.Col) = -1 And KeyAscii = 32 Then
        MSF.CellAlignment = 4
        If MSF.TextMatrix(MSF.Row, MSF.Col) = "" Then
            MSF.TextMatrix(MSF.Row, MSF.Col) = CheckChar
        Else
            MSF.TextMatrix(MSF.Row, MSF.Col) = ""
        End If
        RaiseEvent CellCheck(MSF.Row, MSF.Col)
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    RaiseEvent EditKeyPress(KeyAscii)
    'If KeyAscii = 39 Then KeyAscii = 0: Beep: Exit Sub '按下单引号
    If KeyAscii = 27 Then TxtVisible = False: TxtEdit.Text = "": MSF.SetFocus: Exit Sub '按下Esc
    If KeyAscii = 13 Then '按下Enter
        KeyAscii = 0
        msf_KeyDown vbKeyReturn, 0
        Exit Sub
    End If
    If KeyAscii = 8 Then Exit Sub '按下退格键
    If TxtCheck Then
        If TextMask <> "" Then
            If InStr(1, TextMask, Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
        End If
        If TextLen <> 0 Then
            If LenB(StrConv(TxtEdit.Text, vbFromUnicode)) = TextLen Then KeyAscii = 0: Beep
        End If
    End If
End Sub
'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "返回/设置一个对象在运行时是否以 3D 效果显示。"
Attribute Appearance.VB_MemberFlags = "100c"
    Appearance = MSF.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    MSF.Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,RemoveItem
Public Sub RemoveMSFItem(ByVal Index As Long)
Attribute RemoveMSFItem.VB_Description = "Removes a row from a Hierarchical FlexGrid control at run time"
    MSF.RemoveItem Index
    RowHeightMin = MSF.RowHeight(1)
End Sub
'注意！不要删除或修改下列被注释的行！
'MappingInfo=TxtEdit,TxtEdit,-1,BackColor
Public Property Get TxtBackColor() As OLE_COLOR
Attribute TxtBackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    TxtBackColor = TxtEdit.BackColor
End Property

Public Property Let TxtBackColor(ByVal New_TxtBackColor As OLE_COLOR)
    TxtEdit.BackColor() = New_TxtBackColor
    PropertyChanged "TxtBackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14,0,1,0
Public Property Get CellBackColor() As Variant
    M_CellBackColor = MSF.CellBackColor
    CellBackColor = M_CellBackColor
End Property

Public Property Let CellBackColor(ByVal New_CellBackColor As Variant)
    If Ambient.UserMode = False Then Err.Raise 387
    
    On Error Resume Next
    
    M_CellBackColor = New_CellBackColor
    MSF.CellBackColor = M_CellBackColor
    PropertyChanged "CellBackColor"
End Property

Public Function SetColor(ByVal BlnAll As Boolean, Optional ByVal lngColor As Long = &H80000014, Optional ByVal LngRow As Long = 1)
'    Dim IntRow As Integer, IntCol As Integer
'
'    If Not AutoRefresh Then Exit Function
'    If BlnAll Then
'        If BlnFlash Then Exit Function
'        LngColorPri = lngColor
'        For IntRow = 1 To MSF.Rows - 1
'            MSF.Row = IntRow
'            For IntCol = 0 To MSF.Cols - 1
'                MSF.Col = IntCol
'                If MSF.ColData(IntCol) = 0 Then
'                    MSF.CellBackColor = lngColor
'                Else
'                    MSF.CellBackColor = &H80000014
'                End If
'            Next
'        Next
'        BlnFlash = True
'    Else
'        BlnFlash = False
'        MSF.Row = LngRow
'        For IntCol = 0 To MSF.Cols - 1
'            MSF.Col = IntCol
'            If MSF.ColData(IntCol) = 0 Then
'                MSF.CellBackColor = lngColor
'            Else
'                MSF.CellBackColor = &H80000014
'            End If
'        Next
'    End If
'
'    MSF.Col = 0 '恢复初始列值
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns or sets the default font or the font for individual cells."
Attribute Font.VB_UserMemId = -512
    Set Font = MSF.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set MSF.Font = New_Font
    PropertyChanged "Font"
End Property
'注意！不要删除或修改下列被注释的行！
'MappingInfo=CboSelect,CboSelect,-1,Font
Public Property Get CboFont() As Font
Attribute CboFont.VB_Description = "返回一个 Font 对象。"
    Set CboFont = CboSelect.Font
End Property

Public Property Set CboFont(ByVal New_CboFont As Font)
    Set CboSelect.Font = New_CboFont
    PropertyChanged "CboFont"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=TxtEdit,TxtEdit,-1,Font
Public Property Get TxtEditFont() As Font
Attribute TxtEditFont.VB_Description = "返回一个 Font 对象。"
    Set TxtEditFont = TxtEdit.Font
End Property

Public Property Set TxtEditFont(ByVal New_TxtEditFont As Font)
    Set TxtEdit.Font = New_TxtEditFont
    PropertyChanged "TxtEditFont"
End Property

Private Sub TxtEdit_Change()
    RaiseEvent EditChange(TxtEdit.Text)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,DataBindings
Public Property Get DataBindings() As DataBindings
Attribute DataBindings.VB_Description = "返回/设置一 DataBindings 集合对象，它收集开发人员可利用的可绑定属性。"
    Set DataBindings = MSF.DataBindings
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,DataMember
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "Returns or sets the data member for the control."
    DataMember = MSF.DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    MSF.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Msf,Msf,-1,DataSource
Public Property Get DataSource() As Recordset
Attribute DataSource.VB_Description = "Returns or sets the data source for the control."
    Set DataSource = MSF.DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As Recordset)
    Set MSF.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get TextMask() As String
    TextMask = M_TextMask
End Property

Public Property Let TextMask(ByVal New_TextMask As String)
    M_TextMask = New_TextMask
    PropertyChanged "TextMask"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get TextLen() As Long
    TextLen = M_TextLen
End Property

Public Property Let TextLen(ByVal New_TextLen As Long)
    M_TextLen = New_TextLen
    PropertyChanged "TextLen"
End Property

Public Sub SetRowColor(Row As Long, Color As Long, Optional Clear As Boolean)
    Dim i As Long, j As Long
    Dim preRow As Long, preCol As Long
    
    If Row >= MSF.Rows Then Exit Sub
    With MSF
        preRow = .Row: preCol = .Col
        .Redraw = False
        If Clear Then
            For i = .FixedRows To .Rows - 1
                .Row = i
                For j = 0 To .Cols - 1
                    .Col = j
                    .CellBackColor = .BackColor
                Next
            Next
        End If
        .Row = Row
        For j = 0 To .Cols - 1
            .Col = j
            .CellBackColor = Color
        Next
        .Row = preRow: .Col = preCol
        .Redraw = True
    End With
End Sub

Public Sub SetColColor(Col As Long, Color As Long, Optional Clear As Boolean)
    Dim i As Long, j As Long
    Dim preRow As Long, preCol As Long
    
    If Row >= MSF.Rows Then Exit Sub
    With MSF
        preRow = .Row: preCol = .Col
        .Redraw = False
        If Clear Then
            For i = .FixedRows To .Rows - 1
                .Row = i
                For j = 0 To .Cols - 1
                    .Col = j
                    .CellBackColor = .BackColor
                Next
            Next
        End If
        .Col = Col
        For i = .FixedRows To .Rows - 1
            .Row = i
            .CellBackColor = Color
        Next
        .Row = preRow: .Col = preCol
        .Redraw = True
    End With
End Sub

Private Sub SetFixCenter()
    Dim j As Long
    Dim preRow As Long, preCol As Long
    preRow = MSF.Row: preCol = MSF.Col
    MSF.Row = 0
    For j = 0 To MSF.Cols - 1
        MSF.Col = j
        MSF.CellAlignment = 4
        If MSF.ColWidth(j) = -1 Then MSF.ColWidth(j) = 1000
    Next
    MSF.Row = preRow: MSF.Col = preCol
End Sub

Public Property Get MouseCol() As Long
    MouseCol = MSF.MouseCol
End Property

Public Property Get MouseRow() As Long
    MouseRow = MSF.MouseRow
End Property

Public Property Get ToolTipText() As String
    ToolTipText = MSF.ToolTipText
End Property

Public Property Let ToolTipText(ByVal vNewValue As String)
    MSF.ToolTipText = vNewValue
End Property

Public Property Get CboHwnd() As Variant
    CboHwnd = CboSelect.hwnd
End Property

Public Property Let Redraw(ByVal vNewValue As Boolean)
    MSF.Redraw = vNewValue
End Property

Public Property Get List(ByVal Index As Integer) As String
    List = CboSelect.List(Index)
End Property

Public Sub ClearBill()
    Dim i As Long, j As Long
    MSF.Redraw = False
    For i = 1 To MSF.Rows - 1
        MSF.RowData(i) = 0
        For j = 0 To MSF.Cols - 1
            MSF.TextMatrix(i, j) = ""
        Next
    Next
    CmdVisible = False
    TxtVisible = False: TxtEdit.Text = ""
    CboVisible = False
    MonVisible = False
    MSF.Rows = 2
    MSF.Row = 1: MSF.Col = MSF.FixedCols
    MSF.Redraw = True
End Sub

Public Property Get CellTop() As Long
    CellTop = MSF.CellTop
End Property

Public Sub TxtSetFocus()
    If TxtVisible Then
        TxtEdit.SelStart = 0: TxtEdit.SelLength = Len(TxtEdit.Text)
        TxtEdit.SetFocus
    End If
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,（
Public Property Get CheckChar() As String
    CheckChar = m_CheckChar
End Property

Public Property Let CheckChar(ByVal New_CheckChar As String)
    m_CheckChar = New_CheckChar
    PropertyChanged "CheckChar"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get AllowAddRow() As Boolean
    AllowAddRow = m_AllowAddRow
End Property

Public Property Let AllowAddRow(ByVal New_AllowAddRow As Boolean)
    m_AllowAddRow = New_AllowAddRow
    PropertyChanged "AllowAddRow"
End Property




Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Private Sub AdjustData()
    '如果当前坐标超出当前窗体，则调整其坐标
    '将在当前窗体中的坐标转换成屏幕坐标
    pt_Pop.x = MonView.Left / Screen.TwipsPerPixelX
    pt_Pop.y = MonView.Top / Screen.TwipsPerPixelY
    Call ClientToScreen(UserControl.hwnd, pt_Pop)
    '设置其父窗体为控件的父窗体
    If lngParent = 0 Then lngParent = GetParentWindow(UserControl.hwnd)
    Call SetParent(MonView.hwnd, lngParent)
    '设置为屏幕坐标
    Call GetWindowRect(lngParent, rect_Parent)
    Call GetClientRect(lngParent, rect_Client)
    pt_Pop.x = pt_Pop.x - rect_Parent.Left
    pt_Pop.y = pt_Pop.y - (rect_Parent.Bottom - rect_Client.Bottom) + 5
    
    '调整坐标
    If pt_Pop.x < 0 Then pt_Pop.x = 0
    If pt_Pop.x + MonView.Width / Screen.TwipsPerPixelX > rect_Parent.Right - rect_Parent.Left Then pt_Pop.x = rect_Parent.Right - MonView.Width / Screen.TwipsPerPixelX - rect_Parent.Left
    If pt_Pop.y + MonView.Height / Screen.TwipsPerPixelY > rect_Parent.Bottom - rect_Parent.Top Then pt_Pop.y = rect_Parent.Bottom - MonView.Height / Screen.TwipsPerPixelY - rect_Parent.Top - 25
    Call SetWindowPos(MonView.hwnd, HWND_TOP, pt_Pop.x, pt_Pop.y, 0, 0, SWP_NOSIZE + SWP_SHOWWINDOW)
End Sub

Private Function GetParentWindow(ByVal hwndFrm As Long) As Long
    Dim strClass As String * 256
    Dim blnCBR As Boolean   '是Coolbar
    
    On Error Resume Next
    '获取指定窗体的父窗体的父窗体
    '如果其父窗体不是Form，继续向上找

    Do While True
        hwndFrm = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
        Call GetClassName(hwndFrm, strClass, 255)
        If (strClass Like "ThunderRT6FormDC*") Or _
        (strClass Like "ThunderFormDC*") Then Exit Do
    Loop
    GetParentWindow = hwndFrm
End Function
'注意！不要删除或修改下列被注释的行！
'MappingInfo=TxtEdit,TxtEdit,-1,hWnd
Public Property Get TxtHwnd() As Long
Attribute TxtHwnd.VB_Description = "返回一个句柄到(from Microsoft Windows)一个对象的窗口。"
    TxtHwnd = TxtEdit.hwnd
End Property

