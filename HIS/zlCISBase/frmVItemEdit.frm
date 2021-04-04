VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmVItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "所见项编辑"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmVItemEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraBase 
      Height          =   2790
      Left            =   105
      TabIndex        =   36
      Top             =   360
      Width           =   5460
      Begin VB.TextBox txt临床意义 
         Height          =   600
         Left            =   1215
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         ToolTipText     =   "要素编辑时的提示内容"
         Top             =   2085
         Width           =   4065
      End
      Begin VB.ComboBox cbo性别域 
         Height          =   300
         ItemData        =   "frmVItemEdit.frx":000C
         Left            =   3780
         List            =   "frmVItemEdit.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1698
         Width           =   1500
      End
      Begin VB.TextBox txt单位 
         Height          =   300
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1698
         Width           =   1170
      End
      Begin VB.ComboBox cbo类型 
         Height          =   300
         ItemData        =   "frmVItemEdit.frx":0010
         Left            =   1215
         List            =   "frmVItemEdit.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1326
         Width           =   1185
      End
      Begin VB.TextBox txt编码 
         Height          =   300
         Left            =   1215
         MaxLength       =   8
         TabIndex        =   4
         Top             =   210
         Width           =   960
      End
      Begin VB.TextBox txt中文名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   6
         Top             =   582
         Width           =   4065
      End
      Begin VB.TextBox txt英文名 
         Height          =   300
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   8
         Top             =   954
         Width           =   4065
      End
      Begin VB.TextBox txt长度 
         Height          =   300
         Left            =   3375
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1326
         Width           =   570
      End
      Begin VB.TextBox txt小数 
         Height          =   300
         Left            =   4785
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1326
         Width           =   495
      End
      Begin VB.Label lbl临床意义 
         AutoSize        =   -1  'True
         Caption         =   "临床意义(&M)"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   2130
         Width           =   990
      End
      Begin VB.Label lbl编码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目编码(&R)"
         Height          =   180
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl中文名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中文名称(&N)"
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   642
         Width           =   990
      End
      Begin VB.Label lbl英文名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "英文名称(&E)"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   1014
         Width           =   990
      End
      Begin VB.Label lbl类型 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数据类型(&T)"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   1386
         Width           =   990
      End
      Begin VB.Label lbl长度 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "长度(&L)"
         Height          =   180
         Left            =   2730
         TabIndex        =   11
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lbl小数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "小数(&D)"
         Height          =   180
         Left            =   4110
         TabIndex        =   13
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lbl单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数值单位(&U)"
         Height          =   180
         Left            =   180
         TabIndex        =   15
         Top             =   1758
         Width           =   990
      End
      Begin VB.Label lbl性别域 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别限制(&X)"
         Height          =   180
         Left            =   2730
         TabIndex        =   17
         Top             =   1755
         Width           =   990
      End
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "&P"
      Height          =   285
      Left            =   5235
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   75
      Width           =   315
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   1
      Top             =   60
      Width           =   4080
   End
   Begin VB.Frame fraScope 
      Height          =   2745
      Left            =   105
      TabIndex        =   35
      Top             =   3030
      Width           =   5460
      Begin VB.CheckBox chkDyn 
         Caption         =   "自定义"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2940
         TabIndex        =   42
         ToolTipText     =   "当“表现方法”为复选/单选时是否允许选项自定义"
         Top             =   2325
         Width           =   915
      End
      Begin VB.CheckBox chkMust 
         Caption         =   "必填"
         Height          =   300
         Left            =   4350
         TabIndex        =   41
         ToolTipText     =   "病历书写检查时为是否必填项"
         Top             =   2325
         Width           =   660
      End
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   1
         Left            =   5055
         Picture         =   "frmVItemEdit.frx":0014
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "向后移动"
         Top             =   1395
         Width           =   345
      End
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   0
         Left            =   5055
         Picture         =   "frmVItemEdit.frx":0161
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "向前移动"
         Top             =   1005
         Width           =   345
      End
      Begin VB.ComboBox cbo表示法 
         Height          =   300
         ItemData        =   "frmVItemEdit.frx":02AE
         Left            =   1215
         List            =   "frmVItemEdit.frx":02B0
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   195
         Width           =   2970
      End
      Begin VB.TextBox txt初始值 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1215
         MaxLength       =   250
         TabIndex        =   27
         Top             =   2325
         Visible         =   0   'False
         Width           =   1230
      End
      Begin ZL9BillEdit.BillEdit msh数值域 
         Height          =   1275
         Left            =   1215
         TabIndex        =   25
         Top             =   975
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   2249
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483633
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox txt数值域 
         Height          =   270
         Index           =   1
         Left            =   2940
         MaxLength       =   250
         TabIndex        =   24
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txt数值域 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   250
         TabIndex        =   22
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label lbl表示法 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "表现方法(&F)"
         Height          =   180
         Left            =   165
         TabIndex        =   19
         Top             =   255
         Width           =   990
      End
      Begin VB.Label lbl初始值 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "初始数值(&I)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   165
         TabIndex        =   26
         Top             =   2385
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   2610
         TabIndex        =   23
         Top             =   690
         Width           =   180
      End
      Begin VB.Label lbl数值域 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "取值范围(&V)"
         Height          =   180
         Left            =   165
         TabIndex        =   21
         Top             =   660
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   105
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5835
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3315
      TabIndex        =   32
      Top             =   5835
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4455
      TabIndex        =   33
      Top             =   5835
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   5730
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   420
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5760
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemEdit.frx":02B2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemEdit.frx":084C
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraWord 
      Height          =   1095
      Left            =   60
      TabIndex        =   37
      Top             =   6135
      Visible         =   0   'False
      Width           =   5460
      Begin VB.ComboBox cbo文字表述 
         Height          =   300
         ItemData        =   "frmVItemEdit.frx":0DE6
         Left            =   2310
         List            =   "frmVItemEdit.frx":0DE8
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   210
         Width           =   2865
      End
      Begin VB.TextBox txt空值文字 
         Height          =   300
         Left            =   2310
         MaxLength       =   100
         TabIndex        =   31
         Top             =   615
         Width           =   2865
      End
      Begin VB.Label lbl文字表述 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转换为文本的表述方法(&Y)"
         Height          =   180
         Left            =   180
         TabIndex        =   28
         Top             =   270
         Width           =   2070
      End
      Begin VB.Label lbl空值文字 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目数值为空时表示为(&W)"
         Height          =   180
         Left            =   180
         TabIndex        =   30
         Top             =   675
         Width           =   2070
      End
   End
   Begin VB.Label lbl分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "项目分类(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   990
   End
End
Attribute VB_Name = "frmVItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、权限、编辑项目的分类ID、ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"增加"、"修改"、"查阅"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private lngClassId As Long       '被编辑的分类ID，上级程序通过ShowMe传递进入
Private lngItemID As Long        '被编辑的项目ID，修改、查阅时由上级程序通过ShowMe传递进入,增加时为0，

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer

Public Sub ShowMe(ByVal frmParent As Object, ByVal byt状态 As Byte, ByVal lng分类id As Long, Optional ByVal lng项目id As Long)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Me.Tag = Switch(byt状态 = 0, "增加", byt状态 = 1, "修改", byt状态 = 2, "查阅")
    lngClassId = lng分类id: lngItemID = lng项目id
    
    '填写需要选择的数据
    aryTemp = Split("0-数值;1-文字;2-日期;3-逻辑", ";")
    Me.cbo类型.Clear
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo类型.AddItem aryTemp(intCount)
    Next
    Me.cbo类型.ListIndex = 0
    
    aryTemp = Split("0-无限制;1-男;2-女", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo性别域.AddItem aryTemp(intCount)
    Next
    Me.cbo性别域.ListIndex = 0
    
    aryTemp = Split("1-项目名+项目值+单位;2-项目值+单位+项目名;3-项目值+单位", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo文字表述.AddItem aryTemp(intCount)
    Next
    Me.cbo文字表述.ListIndex = 0
    
    Err = 0: On Error GoTo errHand
    
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 诊治所见分类" & _
            " Where 性质 =(select 性质 from 诊治所见分类 where ID=[1])" & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngClassId)
    
    With rsTemp
        If .BOF Or .EOF Then MsgBox "请首先建立诊疗分类项目之后增加项目", vbExclamation, gstrSysName: Unload Me: Exit Sub
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Nodes("_" & lng分类id).Selected = True
        Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
        Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End With
    
    '显示窗体
    Me.Show 1, frmParent
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo表示法_Click()
    Call zlSetGround
End Sub

Private Sub cbo表示法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo类型_Click()
    '0-数值；1-文字；2-日期；3-逻辑
    Me.txt长度.Enabled = True: Me.txt小数.Enabled = True
    Select Case Left(Me.cbo类型.Text, 1)
    Case 0
        aryTemp = Split("0-文本;1-上下;2-下拉", ";")
    Case 1
        Me.txt小数.Text = 0: Me.txt小数.Enabled = False
        aryTemp = Split("0-文本;2-下拉;3-复选;4-单选", ";")
    Case 2
        Me.txt长度.Text = 0: Me.txt长度.Enabled = False
        Me.txt小数.Text = 0: Me.txt小数.Enabled = False
        aryTemp = Split("0-文本;2-下拉", ";")
    Case 3
        Me.txt长度.Text = 0: Me.txt长度.Enabled = False
        Me.txt小数.Text = 0: Me.txt小数.Enabled = False
        aryTemp = Split("3-复选", ";")
    End Select
    Me.cbo表示法.Clear
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo表示法.AddItem aryTemp(intCount)
    Next
    Me.cbo表示法.ListIndex = 0
    Call zlSetGround
End Sub

Private Sub cbo类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo文字表述_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo性别域_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        If msh数值域.Row > 1 Then
            
            Call MoveItem(msh数值域.Row, -1)
            msh数值域.Row = msh数值域.Row - 1

            
        End If
    ElseIf msh数值域.Row < msh数值域.Rows - 1 Then
        
        Call MoveItem(msh数值域.Row, 1)
        msh数值域.Row = msh数值域.Row + 1

    End If
'    MSHFlexGrid1.TopRow = msh数值域.Row
    If msh数值域.MsfObj.RowIsVisible(msh数值域.Row) = False Then
        msh数值域.MsfObj.TopRow = msh数值域.Row
    End If
    msh数值域.SetFocus
End Sub

Private Function MoveItem(ByVal intCurRow As Integer, Optional ByVal intMove As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim intCol As Integer
    
    On Error GoTo errHand
    
    strTmp = CStr(msh数值域.RowData(intCurRow))
            
    msh数值域.RowData(intCurRow) = msh数值域.RowData(intCurRow + intMove)
    msh数值域.RowData(intCurRow + intMove) = Val(strTmp)
    
    For intCol = 1 To msh数值域.Cols - 1
        
        strTmp = msh数值域.TextMatrix(intCurRow, intCol)
        
        msh数值域.TextMatrix(msh数值域.Row, intCol) = msh数值域.TextMatrix(intCurRow + intMove, intCol)
        
        msh数值域.TextMatrix(intCurRow + intMove, intCol) = strTmp
        
    Next
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then MsgBox "请输入项目编码！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt编码.Text), vbFromUnicode)) > 8 Then MsgBox "编码超长（最多8个字符）！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If Trim(Me.txt中文名.Text) = "" Then MsgBox "请输入中文名！", vbInformation, gstrSysName: Me.txt中文名.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt中文名.Text), vbFromUnicode)) > 40 Then MsgBox "中文名超长（最多40个字符或20个汉字）！", vbInformation, gstrSysName: Me.txt中文名.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt英文名.Text), vbFromUnicode)) > 40 Then MsgBox "英文名超长（最多40个字符）！", vbInformation, gstrSysName: Me.txt英文名.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt单位.Text), vbFromUnicode)) > 10 Then MsgBox "单位超长（最多10个字符或5个汉字）！", vbInformation, gstrSysName: Me.txt单位.SetFocus: Exit Sub
'    If Me.cbo类型.Text = "0-数值" And IsNumeric(Me.txt初始值) = False Then MsgBox "类型为数值时初始值只能输入数字！", vbInformation, gstrSysName: Me.txt初始值.SetFocus: Exit Sub
'    If Me.cbo类型.Text = "2-日期" And IsDate(Me.txt初始值) = False Then MsgBox "类型为日期时初始值只能输入日期格式！", vbInformation, gstrSysName: Me.txt初始值.SetFocus: Exit Sub
    
    gstrSql = Val(Me.txt分类.Tag) & "," & _
            "'" & Trim(Me.txt编码.Text) & "'," & _
            "'" & Trim(Me.txt中文名.Text) & "'," & _
            "'" & Trim(Me.txt英文名.Text) & "'," & _
            Me.cbo类型.ListIndex & "," & _
            IIf(Me.txt长度.Enabled, Val(Me.txt长度.Text), 0) & "," & _
            IIf(Me.txt小数.Enabled, Val(Me.txt小数.Text), 0) & "," & _
            "'" & Trim(Me.txt单位.Text) & "'," & _
            "'" & Trim(Me.txt临床意义.Text) & "'," & _
            Left(Me.cbo表示法.Text, 1) & "," & _
            Me.cbo性别域.ListIndex & ","
    strTemp = ""
    If Me.txt数值域(0).Enabled Then
        strTemp = Trim(Me.txt数值域(0).Text) & ";" & Me.txt数值域(1).Text
    End If
    If Me.msh数值域.Active Then
        strTemp = ""
        For intCount = 1 To Me.msh数值域.Rows - 1
            If Trim(Me.msh数值域.TextMatrix(intCount, 1)) <> "" Then
                strTemp = strTemp & ";" & Trim(Me.msh数值域.TextMatrix(intCount, 1))
            End If
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
'        If InStr(1, strTemp, Trim(Me.txt初始值.Text)) = 0 Then
'            MsgBox "初始值没有包含在可选数值中！", vbInformation, gstrSysName
'            Me.msh数值域.SetFocus: Exit Sub
'        End If
    End If
    gstrSql = gstrSql & "'" & strTemp & "',"
    gstrSql = gstrSql & "'" & Trim(Me.txt初始值.Text) & "'," & _
        Me.cbo文字表述.ListIndex + 1 & "," & _
        "'" & Trim(Me.txt空值文字.Text) & "'," & chkMust.Value & "," & chkDyn.Value
    '数据保存
    If Me.Tag = "增加" Then
        lngItemID = zlDatabase.GetNextId("诊治所见项目")
        gstrSql = "ZL_所见项目_INSERT(" & lngItemID & "," & gstrSql & ")"
    Else
        gstrSql = "ZL_所见项目_UPDATE(" & lngItemID & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo errHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Unload Me
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd分类_Click()
    With Me.tvwClass
        .Left = Me.txt分类.Left
        .Top = Me.txt分类.Top + Me.txt分类.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    
    '提取执行项目的信息
    Err = 0: On Error GoTo errHand
    
    gstrSql = "select ID,编码,中文名,英文名,nvl(类型,0) as 类型,长度,小数,小数,单位," & _
            "        临床意义,nvl(表示法,0) as 表示法,nvl(性别域,0) as 性别域,数值域,初始值,nvl(文字表述,1) as 文字表述,空值文字,必填,动态域" & _
            " from 诊治所见项目 I" & _
            " where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt编码.Text = !编码
            Me.txt中文名.Text = IIf(IsNull(!中文名), "", !中文名)
            Me.txt英文名.Text = IIf(IsNull(!英文名), "", !英文名)
            For intCount = 0 To Me.cbo类型.ListCount - 1
                If Val(Left(Me.cbo类型.List(intCount), 1)) = !类型 Then
                    Me.cbo类型.ListIndex = intCount: Exit For
                End If
            Next
            Me.txt长度.Text = IIf(IsNull(!长度), 0, !长度)
            Me.txt小数.Text = IIf(IsNull(!小数), 0, !小数)
            Me.txt单位.Text = IIf(IsNull(!单位), "", !单位)
            Me.txt临床意义.Text = IIf(IsNull(!临床意义), "", !临床意义)
            
            For intCount = 0 To Me.cbo表示法.ListCount - 1
                If Val(Left(Me.cbo表示法.List(intCount), 1)) = !表示法 Then
                    Me.cbo表示法.ListIndex = intCount: Exit For
                End If
            Next
            Call zlSetGround
            For intCount = 0 To Me.cbo性别域.ListCount - 1
                If Val(Left(Me.cbo性别域.List(intCount), 1)) = !性别域 Then
                    Me.cbo性别域.ListIndex = intCount: Exit For
                End If
            Next
            aryTemp = Split(IIf(IsNull(!数值域), "", !数值域), ";")
            If Me.txt数值域(0).Enabled And UBound(aryTemp) >= 0 Then
                Me.txt数值域(0).Text = Val(aryTemp(0)): Me.txt数值域(1).Text = 0
                If UBound(aryTemp) > 0 Then Me.txt数值域(1).Text = Val(aryTemp(1))
                If Me.txt数值域(0).Text = 0 Then Me.txt数值域(0).Text = ""
                If Me.txt数值域(1).Text = 0 Then Me.txt数值域(1).Text = ""
            End If
            If Me.msh数值域.Active Then
                With Me.msh数值域
                    .ClearBill
                    .Rows = UBound(aryTemp) + 2
                    For intCount = 0 To UBound(aryTemp)
                        .TextMatrix(intCount + 1, 0) = intCount + 1
                        .TextMatrix(intCount + 1, 1) = aryTemp(intCount)
                    Next
                End With
            End If
            Me.txt初始值.Text = IIf(IsNull(!初始值), "", !初始值)
            For intCount = 0 To Me.cbo文字表述.ListCount - 1
                If Val(Left(Me.cbo文字表述.List(intCount), 1)) = !文字表述 Then
                    Me.cbo文字表述.ListIndex = intCount: Exit For
                End If
            Next
            Me.txt空值文字.Text = IIf(IsNull(!空值文字), "", !空值文字)
            chkMust.Value = !必填
            chkDyn.Value = Nvl(!动态域, 0)
        End If
        
        If Me.Tag = "增加" Then
            lngItemID = 0

            gstrSql = "select nvl(max(I.编码),'00000000') as 编码" & _
                    " From 诊治所见项目 I,诊治所见分类 C" & _
                    " Where I.分类ID=C.ID and C.性质=(select 性质 from 诊治所见分类 where ID=[1])"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngClassId)
            
            If rsTemp.BOF = False Then
                Me.txt编码.Text = Right(String(8, "0") & Val(rsTemp!编码) + 1, Len(rsTemp!编码))
            End If
            '清除命名信息
            Me.txt中文名.Text = "": Me.txt英文名.Text = "": Me.txt临床意义.Text = ""
        End If
        If Me.Tag = "查阅" Then
            Me.fraBase.Enabled = False: Me.fraScope.Enabled = False: Me.fraWord.Enabled = False
            Me.cmd分类.Enabled = False: Me.cmdOK.Visible = False
            Me.cmdCancel.Caption = "关闭(&C)"
        End If
    End With
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvwClass.Visible Then
        Me.tvwClass.Visible = False: Me.txt分类.SetFocus: Exit Sub
    End If
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msh数值域
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 2
        .MsfObj.AllowUserResizing = flexResizeNone
        .MsfObj.ScrollBars = flexScrollBarVertical
        .MsfObj.MergeCells = flexMergeFree
        .TextMatrix(0, 0) = "可选数值" & Space(30)
        .TextMatrix(0, 1) = "可选数值" & Space(30)
        .MsfObj.MergeRow(0) = True
        
        .ColData(0) = 5: .ColAlignment(0) = 1
        .ColData(1) = 4: .ColAlignment(1) = 1
        .ColWidth(0) = 250: .ColWidth(1) = .Width - .ColWidth(0) - 30
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
End Sub

Private Sub msh数值域_AfterAddRow(Row As Long)
    With Me.msh数值域
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msh数值域_AfterDeleteRow()
    With Me.msh数值域
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msh数值域_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If Me.txt初始值.Text = Me.msh数值域.TextMatrix(Row, 1) Then
        Me.txt初始值.Text = ""
    End If
End Sub

Private Sub msh数值域_DblClick(Cancel As Boolean)
    If Me.msh数值域.Active Then
        Me.txt初始值.Text = Me.msh数值域.TextMatrix(Me.msh数值域.Row, 1): Cancel = True
    End If
End Sub

Private Sub msh数值域_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.msh数值域.Active = False Then Exit Sub
    With Me.msh数值域
        If .Col <> 1 Then Exit Sub
        If .TxtVisible = False Then
            If .TextMatrix(.Row, 1) = "" Then
                If .Row = 1 Then Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
        Else
            If Trim(.Text) = "" Then
                If .Row = 1 Then .SetFocus: Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            strTemp = UCase(Trim(.Text))
        End If
        Select Case Left(Me.cbo类型.Text, 1)
        Case 0  '数值
            If strTemp <> "0" And Val(strTemp) = 0 Then
                MsgBox "输入数据不是数值型！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
        Case 1  '文字
            strTemp = Replace(strTemp, "%", "")
            strTemp = Replace(strTemp, "&", "")
            strTemp = Replace(strTemp, ";", "")
            strTemp = Replace(strTemp, "'", "")
        Case 2  '日期
            Err = 0: On Error Resume Next
            strTemp = CDate(strTemp)
            If Err <> 0 Then
                Err = 0
                MsgBox "输入数据不是日期格式！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
        End Select
        .TextMatrix(.Row, 1) = strTemp
    End With
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
    Me.txt分类.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd分类 Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 100
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt长度_GotFocus()
    Me.txt长度.SelStart = 0: Me.txt长度.SelLength = 100
End Sub

Private Sub txt长度_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt初始值_GotFocus()
    Me.txt初始值.SelStart = 0: Me.txt初始值.SelLength = 100
End Sub

Private Sub txt初始值_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt单位_GotFocus()
    Me.txt单位.SelStart = 0: Me.txt单位.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt单位_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt分类_GotFocus()
    Me.txt分类.SelStart = 0: Me.txt分类.SelLength = 100
End Sub

Private Sub txt分类_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt空值文字_GotFocus()
    Me.txt空值文字.SelStart = 0: Me.txt空值文字.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt空值文字_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt空值文字_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt临床意义_GotFocus()
    Me.txt临床意义.SelStart = 0: Me.txt临床意义.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt临床意义_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("%&_|'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt临床意义_LostFocus()
    Me.txt临床意义.Text = Replace(Me.txt临床意义, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt数值域_GotFocus(Index As Integer)
    Me.txt数值域(Index).SelStart = 0: Me.txt数值域(Index).SelLength = 100
End Sub

Private Sub txt数值域_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt小数_GotFocus()
    Me.txt小数.SelStart = 0: Me.txt小数.SelLength = 100
End Sub

Private Sub txt小数_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt英文名_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("&'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt中文名_GotFocus()
    Me.txt中文名.SelStart = 0: Me.txt中文名.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt中文名_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt中文名_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub zlSetGround()
    '----------------------------------------
    '功能：根据表示法和项目类型确定数值范围输入方式
    '----------------------------------------
    '0-文本;1-上下;2-下拉;3-复选;4-单选
    
    cmdMove(0).Enabled = False
    cmdMove(1).Enabled = False
    Select Case Left(Me.cbo表示法.Text, 1)
    Case 0
        '可能为数值、文本、日期
        Select Case Left(Me.cbo类型.Text, 1)
        Case 0  '数值
            Me.txt数值域(0).Enabled = True: Me.txt数值域(1).Enabled = True
        Case 1  '文字
            Me.txt数值域(0).Enabled = False: Me.txt数值域(1).Enabled = False
        Case 2  '日期
            Me.txt数值域(0).Enabled = True: Me.txt数值域(1).Enabled = True
        End Select
        Me.msh数值域.Active = False
        chkDyn.Enabled = False: chkDyn.Value = vbUnchecked
    Case 1
        '只可能是数值类型
        Me.txt数值域(0).Enabled = True: Me.txt数值域(1).Enabled = True
        Me.msh数值域.Active = False
        chkDyn.Enabled = False: chkDyn.Value = vbUnchecked
    Case 2
        '可能为数值、文本、日期，但无显示区别
        Me.txt数值域(0).Enabled = False: Me.txt数值域(1).Enabled = False
        Me.msh数值域.Active = True
        cmdMove(0).Enabled = True
        cmdMove(1).Enabled = True
        chkDyn.Enabled = False: chkDyn.Value = vbUnchecked
    Case 3
        Me.txt数值域(0).Enabled = False: Me.txt数值域(1).Enabled = False
        '可能为文本、逻辑
        Select Case Left(Me.cbo类型.Text, 1)
        Case 1  '文字
            Me.msh数值域.Active = True
        Case 2  '逻辑
            Me.msh数值域.Active = False
        End Select
        cmdMove(0).Enabled = True
        cmdMove(1).Enabled = True
        chkDyn.Enabled = True
    Case 4
        '可能为数值、文本，但无显示区别
        Me.txt数值域(0).Enabled = False: Me.txt数值域(1).Enabled = False
        Me.msh数值域.Active = True
        cmdMove(0).Enabled = True
        cmdMove(1).Enabled = True
        chkDyn.Enabled = True
    End Select
    
    Me.txt初始值.Text = ""
    If Me.txt数值域(0).Enabled = True Then
        Me.txt数值域(0).BackColor = &H80000005
        Me.txt数值域(1).BackColor = &H80000005
    Else
        Me.txt数值域(0).BackColor = &H8000000F
        Me.txt数值域(1).BackColor = &H8000000F
    End If
    
    If Me.msh数值域.Active Then
'        Me.msh数值域.ToolTipText = "双击设置初始数值"
        Call Me.msh数值域.SetColColor(1, &H80000005)
        Me.msh数值域.BackColorBkg = &H80000005
        Me.txt初始值.Enabled = False
        Me.txt初始值.BackColor = &H8000000F
    Else
        Me.msh数值域.ToolTipText = ""
        Call Me.msh数值域.SetColColor(1, &H8000000F)
        Me.msh数值域.BackColorBkg = &H8000000F
        Me.txt初始值.Enabled = False
        Me.txt初始值.BackColor = &H80000005
    End If
    
End Sub

