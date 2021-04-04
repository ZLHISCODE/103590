VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIllItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "疾病代码编辑"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmIllItemEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -45
      TabIndex        =   16
      Top             =   5370
      Width           =   5865
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4485
      TabIndex        =   15
      Top             =   5535
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3255
      TabIndex        =   14
      Top             =   5535
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   2355
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
            Picture         =   "frmIllItemEdit.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllItemEdit.frx":0326
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   4620
      Index           =   0
      Left            =   180
      TabIndex        =   18
      Top             =   555
      Visible         =   0   'False
      Width           =   5280
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   3495
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "统计码"
         Top             =   1815
         Width           =   1680
      End
      Begin VB.CheckBox Chk分娩 
         Caption         =   "是否录入分娩信息"
         Height          =   240
         Left            =   1005
         TabIndex        =   11
         Top             =   2205
         Width           =   1815
      End
      Begin VB.CommandButton cmd分类 
         Caption         =   "…"
         Height          =   270
         Left            =   4890
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4320
         Width           =   255
      End
      Begin VB.ComboBox cbo手术类型 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1815
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1005
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "编码"
         Top             =   135
         Width           =   1395
      End
      Begin VB.ComboBox cmb疗效 
         Height          =   300
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1410
         Width           =   1695
      End
      Begin VB.ComboBox cmb性别 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1410
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4305
         Width           =   3885
      End
      Begin VB.TextBox txtEdit 
         Height          =   1710
         Index           =   4
         Left            =   1005
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   12
         Tag             =   "说明"
         Top             =   2535
         Width           =   4155
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1005
         MaxLength       =   150
         TabIndex        =   2
         Tag             =   "名称"
         Top             =   555
         Width           =   4155
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1005
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "拼音码"
         Top             =   975
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   3210
         MaxLength       =   20
         TabIndex        =   4
         Tag             =   "五笔码"
         Top             =   975
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "附码"
         Top             =   135
         Width           =   1680
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "统计码(&A)"
         Height          =   180
         Index           =   9
         Left            =   2670
         TabIndex        =   9
         Top             =   1875
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "手术类型(&L)"
         Height          =   180
         Index           =   8
         Left            =   0
         TabIndex        =   7
         Top             =   1875
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&J)"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   29
         Top             =   1035
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "说明(&E)"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   28
         Top             =   2535
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&B)"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "附码(&A)"
         Height          =   180
         Index           =   1
         Left            =   2790
         TabIndex        =   26
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "性别限制(&S)"
         Height          =   180
         Index           =   6
         Left            =   0
         TabIndex        =   25
         Top             =   1470
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "提醒疗效(&F)"
         Height          =   180
         Index           =   7
         Left            =   2490
         TabIndex        =   24
         Top             =   1470
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "分类(&T)"
         Height          =   180
         Index           =   5
         Left            =   330
         TabIndex        =   22
         Top             =   4365
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(拼音)"
         Height          =   180
         Left            =   2385
         TabIndex        =   21
         Top             =   1035
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(五笔)"
         Height          =   180
         Left            =   4620
         TabIndex        =   20
         Top             =   1035
         Width           =   540
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   4620
      Index           =   1
      Left            =   180
      TabIndex        =   30
      Top             =   585
      Visible         =   0   'False
      Width           =   5280
      Begin VB.TextBox txtLocate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         TabIndex        =   41
         ToolTipText     =   "查找下一个按F3或回车，定位输入框按F4"
         Top             =   75
         Width           =   1905
      End
      Begin VB.OptionButton opt应用范围 
         Appearance      =   0  'Flat
         Caption         =   "应用于当前分类"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   3690
         TabIndex        =   34
         Top             =   4230
         Width           =   1575
      End
      Begin VB.OptionButton opt应用范围 
         Appearance      =   0  'Flat
         Caption         =   "应用于同级项目"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1860
         TabIndex        =   33
         Top             =   4230
         Width           =   1620
      End
      Begin VB.OptionButton opt应用范围 
         Appearance      =   0  'Flat
         Caption         =   "应用于当前项目"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   30
         TabIndex        =   32
         Top             =   4230
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.CheckBox chkSelectAll 
         Appearance      =   0  'Flat
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   45
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   120
         Width           =   675
      End
      Begin MSComctlLib.ListView Lvw科室 
         Height          =   3645
         Left            =   0
         TabIndex        =   35
         Top             =   495
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   6429
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lblLocate 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Left            =   960
         TabIndex        =   42
         Top             =   135
         Width           =   360
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   4650
      Index           =   2
      Left            =   165
      TabIndex        =   36
      Top             =   585
      Visible         =   0   'False
      Width           =   5280
      Begin VSFlex8Ctl.VSFlexGrid vs病种 
         Height          =   4275
         Left            =   30
         TabIndex        =   40
         Top             =   15
         Width           =   5250
         _cx             =   9260
         _cy             =   7541
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmIllItemEdit.frx":0640
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.OptionButton opt病种应用范围 
         Appearance      =   0  'Flat
         Caption         =   "应用于当前分类"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   3690
         TabIndex        =   39
         Top             =   4395
         Width           =   1575
      End
      Begin VB.OptionButton opt病种应用范围 
         Appearance      =   0  'Flat
         Caption         =   "应用于同级项目"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1860
         TabIndex        =   38
         Top             =   4395
         Width           =   1620
      End
      Begin VB.OptionButton opt病种应用范围 
         Appearance      =   0  'Flat
         Caption         =   "应用于当前项目"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   30
         TabIndex        =   37
         Top             =   4380
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   5145
      Left            =   75
      TabIndex        =   17
      Top             =   180
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   9075
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本项目(&1)"
            Key             =   "K1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "对应科室(&2)"
            Key             =   "K2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "对应病种(&3)"
            Key             =   "K3"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIllItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrID As String             '当前编辑的项目ID
Dim mstr分类ID As String     '当前编辑的分类项目ID
Dim mstr编码类别 As String
Dim mstr原编码   As String       '保存该疾病的原始编码，用于判断是否要重新得到序号
Dim mlng序号     As Long

Dim mblnChange As Boolean  '已修改
Private mlng简码长度 As Long

Private Const mconInt编码 As Integer = 0
Private Const mconInt附码 As Integer = 1
Private Const mconInt名称 As Integer = 2
Private Const mconInt拼音码 As Integer = 3
Private Const mconInt说明 As Integer = 4
Private Const mconInt分类 As Integer = 5
Private Const mconInt五笔码 As Integer = 6
Private Const mconInt统计码 As Integer = 7

Private Sub IniDept()
    Dim rsTemp As ADODB.Recordset
    Dim lngId As Long
    
    lngId = Val(mstrID)
    
    On Error GoTo ErrHandle
    gstrSQL = " Select A.编码 || '-' || A.名称 科室, A.ID, Nvl(B.科室id, 0) 科室id , A.简码" & _
            " From 部门表 A, (Select 科室id From 疾病编码科室 Where 疾病id = [1]) B " & _
            " Where A.ID = B.科室id(+) And " & _
            " A.ID In (Select 部门id From 部门性质说明 Where 工作性质 In ('临床', '检查', '检验', '治疗', '手术', '营养')) And " & _
            " (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By A.编码 || '-' || A.名称 "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取临床、医技类部门", lngId)
    
    With rsTemp
        If .EOF Then
            MsgBox "没有设置临床或医技类部门！（部门管理）", vbInformation, gstrSysName
            Exit Sub
        End If
        Me.Lvw科室.ListItems.Clear
        Do While Not .EOF
            Me.Lvw科室.ListItems.Add , "_" & !ID, !科室, 1, 1
            Me.Lvw科室.ListItems("_" & !ID).Tag = NVL(!简码)
            If !科室ID > 0 Then
                Me.Lvw科室.ListItems("_" & !ID).Checked = True
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Init手术类型(Optional ByVal str手术类型 As String)
    '----------------------------------------------------------------------------
    '功能:初始化手术类型数据
    '参数:str手术类型-指向指定的手术类型
    '返回:
    '编制:刘兴宏
    '日期:2007/08/14
    '----------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select 名称 From 手术类型 order by 编码"
    
    Err = 0: On Error GoTo ErrHand:
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "获取手术类型")
    With rsTemp
        Me.cbo手术类型.AddItem ""
        If str手术类型 = "" Then
            cbo手术类型.ListIndex = cbo手术类型.NewIndex
        End If
        Do While Not .EOF
            cbo手术类型.AddItem !名称
            If str手术类型 = NVL(!名称) Then
                cbo手术类型.ListIndex = cbo手术类型.NewIndex
            End If
            .MoveNext
        Loop
        If str手术类型 <> "" And cbo手术类型.ListIndex < 0 Then
            cbo手术类型.AddItem str手术类型
            cbo手术类型.ListIndex = cbo手术类型.NewIndex
        End If
        If cbo手术类型.ListIndex < 0 Then cbo手术类型.ListIndex = 0
        cbo手术类型.Tag = cbo手术类型.Text
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub TabShow(ByVal i As Integer)
    tabMain.Tabs(i).Selected = True
    tabMain_Click
End Sub

Private Sub cbo手术类型_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkSelectAll_Click()
    Dim n As Integer
    Dim BlnSelect As Boolean
    
    If chkSelectAll.value = 2 Then Exit Sub
    
    BlnSelect = (chkSelectAll.value = 1)
    
    With Lvw科室
        For n = 1 To .ListItems.Count
            .ListItems(n).Checked = BlnSelect
        Next
    End With
End Sub


Private Sub Chk分娩_Click()
    mblnChange = True
End Sub

Private Sub Chk分娩_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub Chk分娩_LostFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub cmb性别_Click()
    mblnChange = True
End Sub

Private Sub cmb疗效_Click()
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save项目() = False Then Exit Sub
    
    '由于可能修改多个节点，所以只有强制刷新
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstrID = ""
    txtEdit(mconInt编码).Text = ""
    txtEdit(mconInt附码).Text = ""
    txtEdit(mconInt名称).Text = ""
    txtEdit(mconInt拼音码).Text = ""
    txtEdit(mconInt说明).Text = ""
    cmb疗效.ListIndex = 0
    cmb性别.ListIndex = 0
    
    Call TabShow(1)
    txtEdit(mconInt编码).SetFocus
    mblnChange = False
End Sub

Private Function IsValid() As Boolean
'功能:分析输入编码类别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    
    If mstr分类ID = "1" Then
        If MsgBox("在该项目下增加疾病会引起系统自带报表的计算错误，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    For i = 0 To 7
        If i <> 5 Then
            If zlCommFun.StrIsValid(txtEdit(i).Text, txtEdit(i).MaxLength, , txtEdit(i).Tag) = False Then
                Call TabShow(1)
                txtEdit(i).SetFocus
                zlControl.TxtSelAll txtEdit(i)
                Exit Function
            End If
        End If
    Next
    
'    If zlCommFun.ActualLen(txtEdit(mconInt统计码).Text) > txtEdit(mconInt统计码).MaxLength Then
'        MsgBox "统计码不能超过" & txtEdit(mconInt统计码).MaxLength & "个字符或" & txtEdit(mconInt统计码).MaxLength / 2 & "个汉字,请检查!"
'        Call TabShow(1)
'        txtEdit(mconInt统计码).SetFocus
'        zlControl.TxtSelAll txtEdit(mconInt统计码)
'        Exit Function
'    End If
    
    txtEdit(mconInt编码).Text = UCase(Trim(txtEdit(mconInt编码).Text))
    txtEdit(mconInt附码).Text = UCase(Trim(txtEdit(mconInt附码).Text))
    txtEdit(mconInt名称).Text = Trim(txtEdit(mconInt名称).Text)
    
    If Len(txtEdit(mconInt编码).Text) = 0 Then
        MsgBox "编码不能为空。", vbExclamation, gstrSysName
        Call TabShow(1)
        txtEdit(mconInt编码).SetFocus
        Exit Function
    End If
    
    If InStr(txtEdit(mconInt附码).Text, "+") > 0 And InStr(txtEdit(mconInt附码).Text, "*") = 0 Then
        MsgBox "请在附码处输入星号编码。", vbExclamation, gstrSysName
        Call TabShow(1)
        zlControl.TxtSelAll txtEdit(mconInt附码)
        txtEdit(mconInt附码).SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtEdit(mconInt名称).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        Call TabShow(1)
        txtEdit(mconInt名称).Text = ""
        txtEdit(mconInt名称).SetFocus
        Exit Function
    End If
    
    '专门针对疾病编码的特点，作硬性的规定
    If mstr编码类别 = "D" Or mstr编码类别 = "Y" Or mstr编码类别 = "M" Then
        '检查所有字母
        If 检查编码("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/*-+. ", txtEdit(mconInt编码).Text) = False Then
            Call TabShow(1)
            txtEdit(mconInt编码).SetFocus
            zlControl.TxtSelAll txtEdit(mconInt编码)
            Exit Function
        End If
        If 检查编码("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/*-+. ", txtEdit(mconInt附码).Text) = False Then
            Call TabShow(1)
            txtEdit(mconInt附码).SetFocus
            zlControl.TxtSelAll txtEdit(mconInt附码)
            Exit Function
        End If
        
        '检查首字母
        Select Case mstr编码类别
            Case "D"
                If InStr("ABCDEFGHIJKLMNOPQRSTUZ", Left(txtEdit(mconInt编码).Text, 1)) = 0 Then
                    MsgBox "编码的首字母错误。", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt编码).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt编码)
                    Exit Function
                End If
            Case "Y"
                If InStr("VWXY", Left(txtEdit(mconInt编码).Text, 1)) = 0 Then
                    MsgBox "外部原因编码的首字母只为VWXY四种之一。", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt编码).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt编码)
                    Exit Function
                End If
            Case "M"
                If "M" <> Left(txtEdit(mconInt编码).Text, 1) Then
                    MsgBox "形态学编码的首字母只是M。", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt编码).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt编码)
                    Exit Function
                End If
        End Select
        '检查后余字母
        Select Case mstr编码类别
            Case "D", "Y"
                If 检查编码("0123456789+-. ", Mid(txtEdit(mconInt编码).Text, 2), True) = False Then
                    Call TabShow(1)
                    txtEdit(mconInt编码).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt编码)
                    Exit Function
                End If
                If 检查编码("0123456789+-*/. ", Mid(txtEdit(mconInt附码).Text, 2), True) = False Then
                    Call TabShow(1)
                    txtEdit(mconInt附码).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt附码)
                    Exit Function
                End If
            Case "M"
                If 检查编码("0123456789/ ", Mid(txtEdit(mconInt编码).Text, 2)) = False Then
                    Call TabShow(1)
                    txtEdit(mconInt编码).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt编码)
                    Exit Function
                End If
                i = InStr(txtEdit(mconInt编码), "/")
                If i = 0 Then
                    MsgBox "请加上肿瘤动态编码。", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt编码).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt编码)
                    Exit Function
                End If
                strTemp = Mid(txtEdit(mconInt编码).Text, i + 1)
                If Len(strTemp) <> 1 Then
                    MsgBox "肿瘤动态编码错误。", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt编码).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt编码)
                    Exit Function
                End If
                If InStr("012369", strTemp) = 0 Then
                    MsgBox "肿瘤动态编码错误。", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt编码).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt编码)
                    Exit Function
                End If
                
        End Select
    End If
    
'    '检查编码的范围
'    Dim arrValues As Variant, blnMatch As Boolean
'    Dim lngCount As Long, lngPos As Long
'    Dim strBegin As String, strEnd As String
'
'    If Trim(txtEdit(mconInt分类).Tag) = "" Then
'        '如果分类没有设置编码范围，那么可以不管理
'        blnMatch = True
'    Else
'        arrValues = Split(txtEdit(mconInt分类).Tag, ",")
'        For lngCount = LBound(arrValues) To UBound(arrValues)
'            lngPos = InStr(arrValues(lngCount), "-")
'            If lngPos = 0 Then
'                '没有短横线，只能以该处开头
'                strBegin = Trim(arrValues(lngCount))
'                If txtEdit(mconInt编码).Text Like (strBegin & "*") Then
'                    blnMatch = True
'                    Exit For
'                End If
'            Else
'                '取值范围
'                strBegin = Trim(Mid(arrValues(lngCount), 1, lngPos - 1))
'                strEnd = Trim(Mid(arrValues(lngCount), lngPos + 1))
'
'                strTemp = Mid(txtEdit(mconInt编码).Text, 1, Len(strBegin))
'                If strTemp >= strBegin And strTemp <= strEnd Then
'                    blnMatch = True
'                    Exit For
'                End If
'            End If
'        Next
'    End If
'
'    If blnMatch = False Then
'        MsgBox "编码错误，当前分类下的编码范围是：" & vbCrLf & vbCrLf & txtEdit(mconInt分类).Tag, vbInformation, gstrSysName
'        Call TabShow(1)
'        txtEdit(mconInt编码).SetFocus
'        zlControl.TxtSelAll txtEdit(mconInt编码)
'        Exit Function
'    End If
    
    '检查
    With vs病种
        If vs病种.Tag <> "DEL" Then
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("病种"))) <> "" And Val(.Cell(flexcpData, i, .ColIndex("病种"))) = 0 Then
                    MsgBox "所输入的病种不正确,请重新输入!", vbInformation + vbDefaultButton1, gstrSysName
                    TabShow (3)
                    .Row = i
                    If .RowIsVisible(i) = False Then
                        .TopRow = i
                    End If
                    .SetFocus
                    Exit Function
                End If
            Next
        End If
    End With
    IsValid = True
End Function

Private Function 检查编码(ByVal strPartten As String, ByVal strCheck As String, Optional blnX As Boolean = False) As Boolean
'参数:blnX 是否支持第4个字母为X，如G01.X55*
    Dim i As Long
    Dim blnIsValid As Boolean
    
    blnIsValid = True
    
    For i = 1 To Len(strCheck)
        If InStr(strPartten, Mid(strCheck, i, 1)) = 0 Then
            If Not (blnX = True And i = 4 And Mid(strCheck, 4, 1) = "X") Then
'                '排除第4个字母为X的情况
'                MsgBox "请在编码中输入合法字母。", vbInformation, gstrSysName
'                Exit Function
                blnIsValid = False
                Exit For
            End If
        End If
    Next
    
    If Not blnIsValid Then
        If MsgBox("在编码或附码中含有非法字母，确认保存吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) <> vbYes Then
            检查编码 = False
            Exit Function
        End If
    End If
    
    
    检查编码 = True
    
    
End Function

Private Function Save项目() As Boolean
'功能:保存编辑的内容到编码类别表中
'参数:
'返回值:成功返回True,否则为False
    Dim rsTemp As New ADODB.Recordset
    Dim lngId As Long
    Dim lng序号 As Long
    Dim lng分类id As Long
    Dim bln属于 As Boolean
    Dim str手术类型 As String
    
    Dim lst As ListItem
    Dim strDeptId As String
    Dim n As Integer
    Dim int应用范围 As Integer
    Dim int病种应用范围 As Integer
    Dim str病种 As String '险类|病种id,险类1|病种id1 ....
    On Error GoTo ErrHandle
    
    
    '首先判断编码的序号
    lng序号 = mlng序号
    Err = 0: On Error GoTo ErrHand:
    
    If mstr原编码 <> txtEdit(mconInt编码).Text Then
    
        gstrSQL = "select max(序号) as 最大号,count(序号) as 数量 from 疾病编码目录 " & _
                   " Where 编码 = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtEdit(mconInt编码).Text)
        
        If rsTemp("数量") > 0 Then
            If MsgBox("编码" & txtEdit(mconInt编码).Text & "已存在，是否要再增加一条？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
            
            lng序号 = rsTemp("最大号") + 1
        End If
        rsTemp.Close
    End If
    
    For n = 1 To Lvw科室.ListItems.Count
        If Lvw科室.ListItems(n).Checked = True Then
            strDeptId = IIF(strDeptId = "", Mid(Lvw科室.ListItems(n).Key, 2), strDeptId & "," & Mid(Lvw科室.ListItems(n).Key, 2))
        End If
    Next
    
    str病种 = ""
    With vs病种
        For n = 1 To .Rows - 1
            If .RowData(n) <> 0 Then
                If Val(.Cell(flexcpData, n, .ColIndex("病种"))) <> 0 Then
                    str病种 = str病种 & "," & .RowData(n) & "|" & Val(.Cell(flexcpData, n, .ColIndex("病种")))
                End If
            End If
        Next
    End With
    If str病种 <> "" Then str病种 = Mid(str病种, 2)
    
    For n = 0 To opt应用范围.UBound
        If opt应用范围(n).value = True Then
            int应用范围 = n
            Exit For
        End If
    Next
    
    For n = 0 To opt病种应用范围.UBound
        If opt病种应用范围(n).value = True Then
            int病种应用范围 = n
            Exit For
        End If
    Next
    
    lng分类id = Val(mstr分类ID)
    If cbo手术类型.Enabled And cbo手术类型.ListIndex > 0 Then
        str手术类型 = cbo手术类型.Text
    Else
        str手术类型 = ""
    End If
    If mstrID = "" Then       '新增一条记录
        
        '对于没有分类的编码还应该得到其分类ID
        If cmd分类.Visible = False And mstr分类ID = "" Then
            gstrSQL = "select ID from 疾病编码分类 where 类别=[1] and rownum<2"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr编码类别)
            
            If rsTemp.EOF Then
                
                lng分类id = zlDatabase.GetNextId("疾病编码分类")
                
                gstrSQL = "ZL_疾病编码分类_INSERT(" & lng分类id & ",NULL,1,'" & Mid(frmIllManage.cmbType.Text, 3) & "','','" & mstr编码类别 & "',1)"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Else
                '得到分类ID
                lng分类id = rsTemp("ID")
            End If
        End If
                
        lngId = zlDatabase.GetNextId("疾病编码目录")
        
        'Zl_疾病编码目录_Insert
        gstrSQL = "Zl_疾病编码目录_Insert("
        '  Id_In       In 疾病编码目录.ID%Type,
        gstrSQL = gstrSQL & "" & lngId & ","
        '  编码_In     In 疾病编码目录.编码%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt编码).Text & "',"
        '  序号_In     In 疾病编码目录.序号%Type,
        gstrSQL = gstrSQL & "" & lng序号 & ","
        '  附码_In     In 疾病编码目录.附码%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt附码).Text & "',"
        '  统计码_In   In 疾病编码目录.统计码%Type,
        gstrSQL = gstrSQL & "" & IIF(Trim(txtEdit(mconInt统计码).Text) = "", "NULL", "'" & Trim(txtEdit(mconInt统计码).Text) & "'") & ","
        '  名称_In     In 疾病编码目录.名称%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt名称).Text & "',"
        '  简码_In     In 疾病编码目录.简码%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt拼音码).Text & "',"
        '  说明_In     In 疾病编码目录.说明%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt说明).Text & "',"
        '  性别限制_In In 疾病编码目录.性别限制%Type,
        gstrSQL = gstrSQL & "'" & cmb性别.Text & "',"
        '  疗效限制_In In 疾病编码目录.疗效限制%Type,
        gstrSQL = gstrSQL & "'" & cmb疗效.Text & "',"
        '  类别_In     In 疾病编码目录.类别%Type,
        gstrSQL = gstrSQL & "'" & mstr编码类别 & "',"
        '  手术类型_In In 疾病编码目录.手术类型%Type,
        gstrSQL = gstrSQL & "" & IIF(str手术类型 = "", "NULL", "'" & str手术类型 & "'") & ","
        '  分类id_In   In 疾病编码目录.分类id%Type,
        gstrSQL = gstrSQL & "" & lng分类id & ","
        '  分娩_In     In 疾病编码目录.分娩%Type := Null,
        gstrSQL = gstrSQL & "'" & Chk分娩.value & "',"
        '  五笔码_In   In 疾病编码目录.五笔码%Type := Null,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt五笔码).Text & "',"
        '  参数_In     In Varchar2, --科室ID串:科室ID1,科室ID2,科室ID3...
        gstrSQL = gstrSQL & "'" & strDeptId & "',"
        '  应用_In     In Number := 0 --应用范围:0-应用于当前项目;1-应用于同级项目;2-应用于当前分类
        gstrSQL = gstrSQL & "" & int应用范围 & ")"
    Else    '修改
        ' Zl_疾病编码目录_Update
        gstrSQL = "Zl_疾病编码目录_Update("
        '  Id_In       In 疾病编码目录.ID%Type,
        gstrSQL = gstrSQL & "" & mstrID & ","
        '  编码_In     In 疾病编码目录.编码%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt编码).Text & "',"
        '  序号_In     In 疾病编码目录.序号%Type,
        gstrSQL = gstrSQL & "" & lng序号 & ","
        '  附码_In     In 疾病编码目录.附码%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt附码).Text & "',"
        '  统计码_In   In 疾病编码目录.统计码%Type,
        gstrSQL = gstrSQL & "" & IIF(Trim(txtEdit(mconInt统计码).Text) = "", "NULL", "'" & Trim(txtEdit(mconInt统计码).Text) & "'") & ","
        '  名称_In     In 疾病编码目录.名称%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt名称).Text & "',"
        '  简码_In     In 疾病编码目录.简码%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt拼音码).Text & "',"
        '  说明_In     In 疾病编码目录.说明%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt说明).Text & "',"
        '  性别限制_In In 疾病编码目录.性别限制%Type,
        gstrSQL = gstrSQL & "'" & cmb性别.Text & "',"
        '  疗效限制_In In 疾病编码目录.疗效限制%Type,
        gstrSQL = gstrSQL & "'" & cmb疗效.Text & "',"
        '  类别_In     In 疾病编码目录.类别%Type,
        gstrSQL = gstrSQL & "'" & mstr编码类别 & "',"
        '  手术类型_In In 疾病编码目录.手术类型%Type,
        gstrSQL = gstrSQL & "" & IIF(str手术类型 = "", "NULL", "'" & str手术类型 & "'") & ","
        '  分类id_In   In 疾病编码目录.分类id%Type,
        gstrSQL = gstrSQL & "" & lng分类id & ","
        '  分娩_In     In 疾病编码目录.分娩%Type,
        gstrSQL = gstrSQL & "'" & Chk分娩.value & "',"
        '  五笔码_In   In 疾病编码目录.五笔码%Type := Null,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt五笔码).Text & "',"
        '  参数_In     In Varchar2, --科室ID串:科室ID1,科室ID2,科室ID3...
        gstrSQL = gstrSQL & "'" & strDeptId & "',"
        '  应用_In     In Number := 0 --应用范围:0-应用于当前项目;1-应用于同级项目;2-应用于当前分类
        gstrSQL = gstrSQL & "" & int应用范围 & ")"
    End If
    
    gcnOracle.BeginTrans
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    ' Zl_疾病病种对应_Update
    gstrSQL = "Zl_疾病病种对应_Update("
    
    If mstrID = "" Then
        '  Id_In     In 疾病编码目录.ID%Type,
        gstrSQL = gstrSQL & "" & lngId & ","
    Else
        '  Id_In     In 疾病编码目录.ID%Type,
        gstrSQL = gstrSQL & "" & mstrID & ","
    End If
    '  类别_In   In 疾病编码目录.类别%Type,
    gstrSQL = gstrSQL & "'" & mstr编码类别 & "',"

    '  分类id_In In 疾病编码目录.分类id%Type,
    gstrSQL = gstrSQL & "" & lng分类id & ","

    '  病种_In   In Varchar2 := Null, --病种id串,险类1|病种id1,险类2|病种id2.....
    gstrSQL = gstrSQL & "'" & str病种 & "',"
    
    '  应用_In   In Number := 0 --病种的应用范围:0-应用于当前项目;1-应用于同级项目;2-应用于当前分类
    
    gstrSQL = gstrSQL & "" & int病种应用范围 & ")"
    
    If vs病种.Tag <> "DEL" Then
        '主要是病案共享时,不应该有这些设置
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    gcnOracle.CommitTrans
    
     Err = 0: On Error GoTo ErrHandle
    '对主界面的内容进行更新
    '首先判断该项是否应该在当前列表中显示
    With frmIllManage
        If .tvwMain_S.SelectedItem Is Nothing Then
            bln属于 = True
        Else
            If .mnuViewAll.Checked = True Then
                '显示该项下的所有
                bln属于 = IsChild(mstr分类ID, .tvwMain_S.SelectedItem)
            Else
                bln属于 = (mstr分类ID = Mid(.tvwMain_S.SelectedItem.Key, 2))
            End If
        End If
    End With
    With frmIllManage.lvwMain
        If mstrID = "" Then
            If bln属于 = True Then
                Set lst = .ListItems.Add(, "K" & lngId, " ", "Item", "Item")
                Call ShowItem(lst)
                lst.Selected = True
                DoEvents
                lst.EnsureVisible
            End If
        Else
            If bln属于 = True Then
                Call ShowItem(.SelectedItem)
            Else
                '删除
                Dim intIndex As Long
                intIndex = .SelectedItem.Index
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                    .ListItems(intIndex).Selected = True
                    .ListItems(intIndex).EnsureVisible
                End If
            End If
        End If
    End With
    Call frmIllManage.SetMenu
    
    '下次不用再从数据库中提取了
    mstr分类ID = lng分类id
    Save项目 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Function IsChild(ByVal strKey As String, ByVal nod As Node) As Boolean
'判断某个关键字是否属于一个节点或其子节点
    Dim nodTemp As Node
    
    If strKey = Mid(nod.Key, 2) Then
        IsChild = True
        Exit Function
    End If
    Set nodTemp = nod.Child
    Do Until nodTemp Is Nothing
        If IsChild(strKey, nodTemp) = True Then
            IsChild = True
            Exit Function
        End If
        Set nodTemp = nodTemp.Next
    Loop
End Function

Private Sub ShowItem(lst As ListItem)
'重新显示某一行,用于刷新
    Dim rsTemp As New ADODB.Recordset
    Dim lngCol  As Long
    Dim varValue As Variant
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "" & _
    "   Select A.ID,A.编码,附码,A.名称,A.简码 As 拼音码,A.五笔码,A.说明,A.手术类型,A.统计码,A.性别限制," & _
    "       A.疗效限制 as 提醒疗效,decode(A.分娩,'1','录入') 分娩信息,to_char(A.建档时间,'yyyy-mm-dd') as  建档时间, " & _
    "          to_char(A.撤档时间,'yyyy-mm-dd') as 撤档时间" & _
    "   From 疾病编码目录 A  Where A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(lst.Key, 2)))
                
    '根据ListView的列名从数据库取数
    lst.Text = rsTemp("编码")
    With frmIllManage.lvwMain
        For lngCol = 2 To .ColumnHeaders.Count
            varValue = rsTemp(.ColumnHeaders(lngCol).Text).value
            lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function 疾病编辑(ByVal bln分类 As Boolean, ByVal str分类项目 As String, ByVal str分类项目ID As String, _
    ByVal str编码类别 As String, Optional ByVal strID As String = "") As Boolean
'功能:用来与调用的编码类别管理窗口进行通讯的程序
'参数:str分类项目     分类编码类别的名字
'     str分类项目ID   分类编码类别的ID
'     str编码类别     整个编码的类别
'     strID           本编码类别的的ID
'返回值:编辑成功返回True,否则为False
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim intSys As Integer
    Dim j As Integer
    
    mstr编码类别 = str编码类别
    
    cmb性别.AddItem ""
    cmb性别.AddItem "男"
    cmb性别.AddItem "女"
    
    '问题26069 By lesfeng 2009-11-16 增加其他分类 同时注意权限分配
    On Error GoTo ErrHandle
    intSys = 0
    j = 0
    gstrSQL = "SELECT 共享号 FROM zlsystems WHERE 编号=300" ' Int((glngSys)
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!共享号) Then
            intSys = rsTmp!共享号
        End If
    End If
    rsTmp.Close
    If Int(intSys / 100) = Int(glngSys / 100) Then
        gstrSQL = "SELECT 编码,名称,缺省标志 From 治疗结果 "
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        If Not rsTmp.EOF Then
            While Not rsTmp.EOF
                cmb疗效.AddItem IIF(IsNull(rsTmp!名称), "", rsTmp!名称)
                If IsNull(rsTmp!名称) Then cmb疗效.ListIndex = cmb疗效.NewIndex
    '            If IIF(IsNull(rsTmp!缺省标志), 0, rsTmp!缺省标志) = 1 Then
    '                cmb疗效.ListIndex = cmb疗效.NewIndex
    '            End If
                j = 1
                rsTmp.MoveNext
            Wend
        End If
        rsTmp.Close
    End If
    
    If j = 0 Then
        cmb疗效.AddItem ""
        cmb疗效.AddItem "治愈"
        cmb疗效.AddItem "好转"
        cmb疗效.AddItem "未愈"
        cmb疗效.AddItem "死亡"
        cmb疗效.AddItem "无效"
        cmb疗效.AddItem "其他"
    End If
    
    If bln分类 = False Then
        '没有树形类别可选择
        lblEdit(5).Visible = False
        txtEdit(mconInt分类).Visible = False
        cmd分类.Visible = False

'        Frame1.Top = Frame1.Top - 450
'        cmdOk.Top = cmdOk.Top - 450
'        cmdCancel.Top = cmdCancel.Top - 450
'        Height = Height - 450
    End If
    
    mstrID = strID
    
    rsTemp.CursorLocation = adUseClient
    If strID <> "" Then
        
        gstrSQL = "select A.ID,A.编码,A.序号,A.附码,A.名称,A.简码 As 拼音码,A.五笔码,A.性别限制,A.疗效限制,A.分娩,A.手术类型,A.统计码,A.说明 " & _
                ",B.序号 as 统计序号,B.ID as 分类ID,B.名称 as 分类名称,B.编码范围" & _
                " from 疾病编码目录 A,疾病编码分类 B " & _
                " where B.ID(+)=A.分类ID and A.ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
        
        txtEdit(mconInt编码).Text = IIF(IsNull(rsTemp("编码")), "", rsTemp("编码"))
        mstr原编码 = txtEdit(mconInt编码).Text
        mlng序号 = NVL(rsTemp("序号"), 0)
        txtEdit(mconInt附码).Text = IIF(IsNull(rsTemp("附码")), "", rsTemp("附码"))
        txtEdit(mconInt名称).Text = Trim(rsTemp("名称"))
        txtEdit(mconInt拼音码).Text = IIF(IsNull(rsTemp("拼音码")), "", rsTemp("拼音码"))
        txtEdit(mconInt五笔码).Text = IIF(IsNull(rsTemp("五笔码")), "", rsTemp("五笔码"))
        txtEdit(mconInt说明).Text = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
        txtEdit(mconInt统计码).Text = NVL(rsTemp!统计码)
        
        If Not IsNull(rsTemp("性别限制")) Then
            cmb性别.Text = rsTemp("性别限制")
        End If
        If Not IsNull(rsTemp("疗效限制")) Then
            cmb疗效.Text = rsTemp("疗效限制")
        End If
        mstr分类ID = IIF(IsNull(rsTemp("分类ID")), "", rsTemp("分类ID"))
        
        If IsNull(rsTemp("分类名称")) Then
            txtEdit(mconInt分类).Text = "无"
        Else
            txtEdit(mconInt分类).Text = "【" & rsTemp("统计序号") & "】" & Trim(rsTemp("分类名称"))
        End If
        txtEdit(mconInt分类).Tag = IIF(IsNull(rsTemp("编码范围")), "", rsTemp("编码范围"))
        Chk分娩.value = IIF(NVL(rsTemp("分娩"), 0) = 0, 0, 1)
        Call Init手术类型(NVL(rsTemp!手术类型))
    Else
        If mstr编码类别 = "M" Then txtEdit(mconInt编码).Text = "M"
        
        mstr分类ID = str分类项目ID
        txtEdit(mconInt分类).Text = str分类项目
        
        If bln分类 = True Then
            gstrSQL = "select A.编码范围 " & _
                    " from 疾病编码分类 A " & _
                    " where A.ID=[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str分类项目ID))
            
            txtEdit(mconInt分类).Tag = IIF(IsNull(rsTemp("编码范围")), "", rsTemp("编码范围"))
        End If
        mlng序号 = 1
        Call Init手术类型("")
    End If
    Call GetDefineSize
    Call IniDept
    
    '刘兴宏:2007/08/15加入疾病与病种的对照关系
    Call Init病种(Val(strID))
    
    If UCase(mstr编码类别) <> "S" Then
        '刘兴宏:2007/08/14:只有"S"类型的才能编辑手术类型
        cbo手术类型.Enabled = False
    End If
    mblnChange = False
    frmIllItemEdit.Show vbModal
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmd分类_Click()
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim str分类ID As String
    Dim str编码范围 As String
    
    str分类ID = mstr分类ID
    str名称 = txtEdit(mconInt分类).Text
    blnRe = frmClassSel.ShowTree(str分类ID, str名称, str编码范围, mstr编码类别, "", False)
    '成功返回
    If blnRe Then
        '新的本级的宽度
        mstr分类ID = str分类ID
        txtEdit(mconInt分类).Text = str名称
        txtEdit(mconInt分类).Tag = str编码范围
        mblnChange = True
    End If
End Sub

Private Sub Form_Activate()
    Call tabMain_Click
    txtEdit(mconInt编码).SetFocus
    Lvw科室.MultiSelect = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtLocate_KeyPress (vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub

Private Sub tabMain_Click()
    Dim i As Integer
    For i = fra.LBound To fra.UBound
        fra(i).Visible = False
    Next
    i = tabMain.SelectedItem.Index - 1
    fra(i).Visible = True
End Sub


Private Sub txtEdit_Change(Index As Integer)
    If Index = 2 Then
        txtEdit(mconInt拼音码).Text = zlStr.GetCodeByVB(txtEdit(mconInt名称).Text)
        txtEdit(mconInt五笔码).Text = zlStr.GetCodeByORCL(txtEdit(mconInt名称).Text, False, mlng简码长度)
    ElseIf Index = 3 Then
        txtEdit(mconInt拼音码).Text = UCase(txtEdit(mconInt拼音码).Text)
    ElseIf Index = 6 Then
        txtEdit(mconInt五笔码).Text = UCase(txtEdit(mconInt五笔码).Text)
    End If
    mblnChange = True
End Sub

Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSQL = "Select 五笔码 From 疾病编码目录 Where Rownum = 0 "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng简码长度 = rsTmp.Fields("五笔码").DefinedSize
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 2 Or Index = 4 Then
        zlCommFun.OpenIme True
    Else
        zlCommFun.OpenIme False
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
'对于多行文本，最好不要加空格
    If Index = 4 And KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If Index = 0 Or Index = 1 Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
        '只能取这些字母
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/*-+. " & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Init病种(ByVal lng疾病ID As Long)
    '-------------------------------------------------------------------
    '功能:初始病种
    '参数:lng疾病ID-疾病ID(0表示,新增)
    '编制:刘兴宏
    '日期:2007/08/15
    '-------------------------------------------------------------------
    Dim rs病种 As New ADODB.Recordset
    Dim rs险类 As New ADODB.Recordset
    Dim i As Long
    Err = 0: On Error GoTo ErrHand:
    
    gstrSQL = "Select TABLE_NAME from table_privileges where table_name='保险类别'"
    zlDatabase.OpenRecordset rs险类, gstrSQL, Me.Caption
    If rs险类.RecordCount = 0 Then
        '表示没有记录,则不能设置对应关系
        tabMain.Tabs.Remove "K3"
        vs病种.Tag = "DEL"
        Exit Sub
    End If
    vs病种.Tag = ""
    gstrSQL = "" & _
        "   Select b.险类,b.病种ID,c.编码||'-'||c.名称 as 病种 " & _
        "   From 疾病病种对应 b,保险病种 c " & _
        "   where  b.病种id=c.id and b.疾病id=[1]"
    Set rs病种 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng疾病ID)
    
    gstrSQL = "Select 序号,名称 From 保险类别 where 医院编码 is not null"
    Call zlDatabase.OpenRecordset(rs险类, gstrSQL, Me.Caption)
    With vs病种
        If rs险类.RecordCount = 0 Then
            .Rows = 2
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .Cell(flexcpData, 1, i) = ""
                .RowData(1) = 0
            Next
            .Editable = flexEDNone
            Exit Sub
        End If
        .Rows = rs险类.RecordCount + 1
        i = 1
        Do While Not rs险类.EOF
            .RowData(i) = Val(NVL(rs险类!序号))
            .TextMatrix(i, .ColIndex("险类")) = NVL(rs险类!名称)
            rs病种.Filter = "险类=" & Val(NVL(rs险类!序号))
            If rs病种.EOF = True Then
                .TextMatrix(i, .ColIndex("病种")) = ""
                .Cell(flexcpData, i, .ColIndex("病种")) = ""
            Else
                .TextMatrix(i, .ColIndex("病种")) = NVL(rs病种!病种)
                .Cell(flexcpData, i, .ColIndex("病种")) = NVL(rs病种!病种ID)
            End If
            i = i + 1
             rs险类.MoveNext
        Loop
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("病种")) = "..."
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
 

Private Sub txtLocate_GotFocus()
    zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long
    
    If KeyAscii = vbKeyReturn Then
        If txtLocate.Tag <> txtLocate.Text Then
            lblLocate.Tag = ""
            txtLocate.Tag = txtLocate.Text
        End If
        
        lngStart = Val("" & lblLocate.Tag) + 1
        If lngStart >= Lvw科室.ListItems.Count Then lngStart = 1
    
        For i = lngStart To Lvw科室.ListItems.Count
            If Lvw科室.ListItems(i).Text Like "*" & txtLocate.Text & "*" Or Lvw科室.ListItems(i).Tag Like "*" & UCase(txtLocate.Text) & "*" Then
                Call Lvw科室.ListItems(i).EnsureVisible
                Lvw科室.ListItems(i).Selected = True
                lblLocate.Tag = i
                Lvw科室.SetFocus
                Exit For
            End If
        Next
    End If
End Sub


Private Sub vs病种_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vs病种
        Select Case Col
        Case .ColIndex("病种")
             .ColComboList(.ColIndex("病种")) = "..."
        End Select
    End With
End Sub

Private Sub vs病种_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs病种
        Select Case Col
        Case .ColIndex("险类")
             Cancel = True
        End Select
    End With

End Sub

Private Sub vs病种_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case vs病种.ColIndex("病种")
        '选择病种
        Call Select病种(vs病种.RowData(Row), "")
    Case Else
    End Select
End Sub
Private Function Select病种(ByVal lng险类 As Long, ByVal strKey As String)
    '---------------------------------------------------------------------------------
    '功能:选择指定险类的病种
    '参数:lng险类-险类
    '返回:选择成功,返回ture,否则返回False
    '编制:刘兴宏
    '日期:2007/08/15
    '---------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strLeft As String
    Dim blnCancel As Boolean
    
    Dim vRect As RECT

    strLeft = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Err = 0: On Error GoTo ErrHand:
    
    Dim sngX As Single, sngY As Single
    Call CalcPosition(sngX, sngY, vs病种)
     
     If strKey <> "" Then
        strKey = strLeft & strKey & "%"
        gstrSQL = "" & _
          "   Select Id, 编码, 名称, 简码, decode('0','普通病','1','慢性病','2','特种病','') As 类别, 特殊封顶线, 封顶线金额 " & _
          "    From 保险病种 " & _
          "    Where 险类 = [1] And (编码 Like [2] Or 名称 Like [2] Or 简码 Like [3]) " & _
          "    Order by 编码"
    Else
        gstrSQL = "" & _
          "   Select Id, 编码, 名称, 简码, decode('0','普通病','1','慢性病','2','特种病','') As 类别, 特殊封顶线, 封顶线金额 " & _
          "    From 保险病种 " & _
          "    Where 险类 = [1]" & _
          "    Order by 编码"
    End If
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "保险病种选择", False, "", "", False, False, True, sngX, sngY - vs病种.CellHeight, vs病种.CellHeight, blnCancel, False, False, lng险类, strKey, CStr(UCase(strKey)))
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then
        MsgBox "不存在指定的病种,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    With vs病种
        .TextMatrix(.Row, .ColIndex("病种")) = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
        If strKey <> "" Then
            .EditText = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
        End If
        .Cell(flexcpData, .Row, .ColIndex("病种")) = NVL(rsTemp!ID)
    End With
    Select病种 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vs病种_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vs病种_DblClick()
    With vs病种
        Select Case .Col
        Case .ColIndex("病种")
             .ColComboList(.ColIndex("病种")) = ""
        End Select
    End With
End Sub

Private Sub vs病种_EnterCell()
    vs病种.ColComboList(vs病种.ColIndex("病种")) = "..."
End Sub

Private Sub vs病种_GotFocus()
    With vs病种
        .BackColorSel = &H8000000D
        .GridColor = &H0&
        .GridColorFixed = &H0&
    End With
End Sub

Private Sub vs病种_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long
    If vs病种.Col = vs病种.ColIndex("病种") And KeyCode <> vbKeyReturn Then
       vs病种.ColComboList(vs病种.ColIndex("病种")) = ""
    End If
End Sub

Private Sub vs病种_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
        Dim strKey As String
        If KeyCode <> vbKeyReturn Then Exit Sub
        
        With vs病种
            Select Case Col
            Case .ColIndex("病种")
                .Cell(flexcpData, Row, Col) = ""
                If KeyCode <> vbKeyReturn Then Exit Sub
                
                strKey = Trim(vs病种.EditText)
                strKey = Replace(strKey, Chr(vbKeyReturn), "")
                strKey = Replace(strKey, Chr(10), "")
                
                If strKey = "" Then Exit Sub
                
                If Select病种(.RowData(.Row), strKey) = False Then
                    '选择失败
                    
                End If
                
                .Col = 1
                .SetFocus
            End Select
        End With
End Sub

Private Sub vs病种_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vs病种_LostFocus()
    With vs病种
        .BackColorSel = &H8000000C
        .GridColor = &H808080
        .GridColorFixed = &H808080
    End With
End Sub
Private Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hwnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

