VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDiagItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "疾病诊断编辑"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmDiagItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgList 
      Left            =   2520
      Top             =   7440
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
            Picture         =   "frmDiagItem.frx":000C
            Key             =   "CLASS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagItem.frx":05A6
            Key             =   "ITEM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4920
      TabIndex        =   1
      Top             =   7710
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   0
      Top             =   7710
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7695
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   13150
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本项目(&0)"
      TabPicture(0)   =   "frmDiagItem.frx":09F8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraNote(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraNote(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraNote(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tvwClass"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvwList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "对应科室(&1)"
      TabPicture(1)   =   "frmDiagItem.frx":0A14
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtLocate"
      Tab(1).Control(1)=   "chkSelectAll"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "opt应用范围(0)"
      Tab(1).Control(3)=   "opt应用范围(1)"
      Tab(1).Control(4)=   "opt应用范围(2)"
      Tab(1).Control(5)=   "Lvw科室"
      Tab(1).Control(6)=   "lblLocate"
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtLocate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73350
         TabIndex        =   47
         ToolTipText     =   "查找下一个按F3或回车，定位输入框按F4"
         Top             =   442
         Width           =   1905
      End
      Begin MSComctlLib.ListView lvwList 
         Height          =   4050
         Left            =   5880
         TabIndex        =   46
         Top             =   4440
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   7144
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   2310
         Left            =   5880
         TabIndex        =   45
         TabStop         =   0   'False
         Tag             =   "1000"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   4075
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
      Begin VB.CheckBox chkSelectAll 
         Appearance      =   0  'Flat
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -74835
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   480
         Width           =   675
      End
      Begin VB.OptionButton opt应用范围 
         Appearance      =   0  'Flat
         Caption         =   "应用于当前项目"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   -74880
         TabIndex        =   42
         Top             =   6240
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton opt应用范围 
         Appearance      =   0  'Flat
         Caption         =   "应用于同级项目"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   -74880
         TabIndex        =   41
         Top             =   6600
         Width           =   5700
      End
      Begin VB.OptionButton opt应用范围 
         Appearance      =   0  'Flat
         Caption         =   "应用于当前分类"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   -74880
         TabIndex        =   40
         Top             =   6960
         Width           =   5775
      End
      Begin VB.Frame fraNote 
         Height          =   1875
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   5400
         Width           =   5745
         Begin VB.TextBox txtStandard 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   2565
            TabIndex        =   27
            Top             =   1440
            Width           =   1065
         End
         Begin VB.TextBox txtStandard 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   2565
            TabIndex        =   33
            Top             =   285
            Width           =   1065
         End
         Begin VB.TextBox txtStandard 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   2565
            TabIndex        =   31
            Top             =   667
            Width           =   1065
         End
         Begin VB.TextBox txtStandard 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   2565
            TabIndex        =   29
            Top             =   1050
            Width           =   1065
         End
         Begin VB.CheckBox chkStandard 
            Caption         =   "ICD-10疾病码(&1)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   705
            TabIndex        =   34
            Top             =   300
            Width           =   1980
         End
         Begin VB.CheckBox chkStandard 
            Caption         =   "损伤中毒原因码(&2)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   705
            TabIndex        =   32
            Top             =   675
            Width           =   1980
         End
         Begin VB.CheckBox chkStandard 
            Caption         =   "肿瘤形态学编码(&3)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   705
            TabIndex        =   30
            Top             =   1050
            Width           =   1980
         End
         Begin VB.CheckBox chkStandard 
            Caption         =   "中医疾病编码(&4)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   705
            TabIndex        =   28
            Top             =   1440
            Width           =   1980
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   105
            Picture         =   "frmDiagItem.frx":0A30
            Top             =   45
            Width           =   480
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "标准编码：该疾病对应的国际或国家标准编码"
            Height          =   180
            Index           =   2
            Left            =   630
            TabIndex        =   39
            Top             =   0
            Width           =   3600
         End
         Begin VB.Label lblStandard 
            Caption         =   "疾病名称..."
            Height          =   180
            Index           =   0
            Left            =   3690
            TabIndex        =   38
            Top             =   345
            Width           =   1950
         End
         Begin VB.Label lblStandard 
            Caption         =   "疾病名称..."
            Height          =   180
            Index           =   1
            Left            =   3690
            TabIndex        =   37
            Top             =   727
            Width           =   1950
         End
         Begin VB.Label lblStandard 
            Caption         =   "疾病名称..."
            Height          =   180
            Index           =   2
            Left            =   3690
            TabIndex        =   36
            Top             =   1110
            Width           =   1950
         End
         Begin VB.Label lblStandard 
            Caption         =   "疾病名称..."
            Height          =   180
            Index           =   3
            Left            =   3690
            TabIndex        =   35
            Top             =   1500
            Width           =   1950
         End
      End
      Begin VB.Frame fraNote 
         Height          =   1215
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   3960
         Width           =   5745
         Begin VB.CommandButton cmdSelect 
            Caption         =   "&P"
            Height          =   240
            Left            =   5055
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   315
            Width           =   285
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdClass 
            Height          =   780
            Left            =   705
            TabIndex        =   24
            Top             =   300
            Width           =   4650
            _ExtentX        =   8202
            _ExtentY        =   1376
            _Version        =   393216
            Rows            =   3
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollBars      =   0
            SelectionMode   =   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   105
            Picture         =   "frmDiagItem.frx":12FA
            Top             =   45
            Width           =   480
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "所属分类：由于学科交叉性，疾病可同时隶属多个分类"
            Height          =   180
            Index           =   1
            Left            =   630
            TabIndex        =   25
            Top             =   0
            Width           =   4320
         End
      End
      Begin VB.Frame fraNote 
         Height          =   3285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   5745
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   4
            Left            =   1785
            TabIndex        =   13
            Top             =   1386
            Width           =   3615
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   2
            Left            =   1785
            TabIndex        =   12
            Top             =   1019
            Width           =   1215
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   1
            Left            =   1785
            TabIndex        =   11
            Top             =   652
            Width           =   3615
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   0
            Left            =   1785
            TabIndex        =   10
            Top             =   285
            Width           =   1605
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   5
            Left            =   1785
            TabIndex        =   9
            Top             =   1753
            Width           =   3615
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   6
            Left            =   1785
            TabIndex        =   8
            Top             =   2120
            Width           =   1215
         End
         Begin VB.TextBox txtItem 
            Height          =   645
            Index           =   8
            Left            =   1785
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2490
            Width           =   3645
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   3
            Left            =   3630
            TabIndex        =   6
            Top             =   1019
            Width           =   1215
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   7
            Left            =   3630
            TabIndex        =   5
            Top             =   2120
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   105
            Picture         =   "frmDiagItem.frx":1BC4
            Top             =   45
            Width           =   480
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "诊断命名：包括编号、名称、别名和简要说明等"
            Height          =   180
            Index           =   0
            Left            =   645
            TabIndex        =   21
            Top             =   0
            Width           =   3780
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "其他别名(&A)"
            Height          =   180
            Index           =   5
            Left            =   750
            TabIndex        =   20
            Top             =   1830
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "英文名称(&E)"
            Height          =   180
            Index           =   4
            Left            =   735
            TabIndex        =   19
            Top             =   1455
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "名称简码(&S)              (拼音)               (五笔)"
            Height          =   180
            Index           =   2
            Left            =   750
            TabIndex        =   18
            Top             =   1095
            Width           =   4680
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "诊断名称(&N)"
            Height          =   180
            Index           =   1
            Left            =   750
            TabIndex        =   17
            Top             =   720
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "诊断编号(&D)"
            Height          =   180
            Index           =   0
            Left            =   750
            TabIndex        =   16
            Top             =   345
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "别名简码(&B)              (拼音)               (五笔)"
            Height          =   180
            Index           =   6
            Left            =   750
            TabIndex        =   15
            Top             =   2190
            Width           =   4680
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "简要说明(&M)"
            Height          =   180
            Index           =   8
            Left            =   750
            TabIndex        =   14
            Top             =   2565
            Width           =   990
         End
      End
      Begin MSComctlLib.ListView Lvw科室 
         Height          =   5205
         Left            =   -74880
         TabIndex        =   44
         Top             =   840
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   9181
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
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
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
         Left            =   -73800
         TabIndex        =   48
         Top             =   495
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmDiagItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer
Dim mlng分类id As Long

Const con诊断编号 As Integer = 0
Const con诊断名称 As Integer = 1
Const con名称拼音码 As Integer = 2
Const con名称五笔码 As Integer = 3
Const con英文名称 As Integer = 4
Const con其他别名 As Integer = 5
Const con别名拼音码 As Integer = 6
Const con别名五笔码 As Integer = 7
Const con简要说明 As Integer = 8

Const conICD10疾病 As Integer = 0
Const con损伤与中毒 As Integer = 1
Const con肿瘤形态学 As Integer = 2
Const con中医疾病 As Integer = 3

Private Sub IniDept()
    Dim rsTemp As ADODB.Recordset
    
    '设置对应科室
    On Error GoTo errHandle
    gstrSql = " Select A.编码 || '-' || A.名称 科室, A.ID, Nvl(B.科室id, 0) 科室id, A.简码 " & _
            " From 部门表 A, (Select 科室id From 疾病诊断科室 Where 诊断id = [1]) B " & _
            " Where A.ID = B.科室id(+) And " & _
            " A.ID In (Select 部门id From 部门性质说明 Where 工作性质 In ('临床', '检查', '检验', '治疗', '手术', '营养')) And " & _
            " (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By A.编码 || '-' || A.名称 "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "取临床、医技类部门", Val(Me.Tag))
    
    With rsTemp
        If .EOF Then
            MsgBox "没有设置临床或医技类部门！（部门管理）", vbInformation, gstrSysName
            Exit Sub
        End If
        Me.Lvw科室.ListItems.Clear
        Do While Not .EOF
            Me.Lvw科室.ListItems.Add , "_" & !ID, !科室, 1, 1
            Me.Lvw科室.ListItems("_" & !ID).Tag = Nvl(!简码)
            If !科室ID > 0 Then
                Me.Lvw科室.ListItems("_" & !ID).Checked = True
            End If
            .MoveNext
        Loop
    End With
    
    '设置应用范围
    gstrSql = "Select 名称 From 疾病诊断分类 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "取分类名称", mlng分类id)
    
    If Not rsTemp.EOF Then
        opt应用范围(1).Caption = "应用于【" & rsTemp!名称 & "】分类下的所有项目"
    End If
    
    gstrSql = "Select 名称 From 疾病诊断分类 Where 类别 = 1 And 上级id Is Null Start With ID = [1] Connect By ID = Prior 上级id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "取分类名称", mlng分类id)
    
    If Not rsTemp.EOF Then
        opt应用范围(2).Caption = "应用于【" & rsTemp!名称 & "】分类及子分类下的所有项目"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkSelectAll_Click()
    Dim n As Integer
    Dim BlnSelect As Boolean
    
    If chkSelectAll.Value = 2 Then Exit Sub
    
    BlnSelect = (chkSelectAll.Value = 1)
    
    With Lvw科室
        For n = 1 To .ListItems.Count
            .ListItems(n).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub chkStandard_Click(Index As Integer)
    If Me.chkStandard(Index).Value = 1 Then
        Me.txtStandard(Index).Enabled = True
        Me.txtStandard(Index).BackColor = &H80000005
        Me.txtStandard(Index).SetFocus
    Else
        Me.txtStandard(Index).Enabled = False
        Me.txtStandard(Index).BackColor = &H8000000F
    End If
End Sub

Private Sub chkStandard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim lngItemId As Long, StrClass As String, strCollate As String
    Dim strDeptId As String
    Dim n As Integer
    Dim int应用范围 As Integer
    
    If Trim(Me.txtItem(con诊断编号).Text) = "" Then
        MsgBox "编码必须输入", vbExclamation, gstrSysName
        Me.txtItem(con诊断编号).SetFocus
        Exit Sub
    End If
    If Trim(Me.txtItem(con诊断名称).Text) = "" Then
        MsgBox "名称必须输入", vbExclamation, gstrSysName
        Me.txtItem(con诊断名称).SetFocus
        Exit Sub
    End If
    For intCount = Me.txtItem.LBound To Me.txtItem.UBound
        Select Case intCount
        Case con诊断名称, con其他别名, con简要说明
            If LenB(StrConv(Trim(Me.txtItem(intCount).Text), vbFromUnicode)) > Me.txtItem(intCount).MaxLength Then
                MsgBox Me.lblItem(intCount).Caption & "超过" & Me.txtItem(intCount).MaxLength & "的长度限制", vbExclamation, gstrSysName
                Me.txtItem(intCount).SetFocus
                Exit Sub
            End If
        End Select
    Next
    
    StrClass = ""
    With Me.hgdClass
        For intCount = 0 To .Rows - 1
            If .RowData(intCount) <> 0 Then
                StrClass = StrClass & "," & .RowData(intCount)
            End If
        Next
        If StrClass = "" Then
            MsgBox "至少属于一种疾病诊断分类", vbExclamation, gstrSysName
            .SetFocus
            Exit Sub
        Else
            StrClass = Mid(StrClass, 2)
        End If
    End With
    
    strCollate = ""
    For intCount = Me.chkStandard.LBound To Me.chkStandard.UBound
        If Me.chkStandard(intCount).Value = 1 Then
            If Val(Me.txtStandard(intCount).Tag) <> 0 Then
                strCollate = strCollate & "," & Me.txtStandard(intCount).Tag
            End If
        End If
    Next
    If strCollate <> "" Then
        strCollate = Mid(strCollate, 2)
    End If
    
    '对应科室设置
    For n = 1 To Lvw科室.ListItems.Count
        If Lvw科室.ListItems(n).Checked = True Then
            strDeptId = IIf(strDeptId = "", Mid(Lvw科室.ListItems(n).Key, 2), strDeptId & "," & Mid(Lvw科室.ListItems(n).Key, 2))
        End If
    Next
    
    For n = 0 To opt应用范围.UBound
        If opt应用范围(n).Value = True Then
            int应用范围 = n
            Exit For
        End If
    Next
    
    Err = 0: On Error GoTo ErrHand
    If Me.Tag = "增加" Then
        lngItemId = zlDatabase.GetNextId("疾病诊断目录")
        gstrSql = "zl_疾病诊断目录_Insert(" & _
            lngItemId & "," & _
            "'" & Trim(Me.txtItem(con诊断编号).Text) & "'," & _
            "'" & Trim(Me.txtItem(con诊断名称).Text) & "'," & _
            "" & _
            "'" & Trim(Me.txtItem(con名称拼音码).Text) & "'," & _
            "'" & Trim(Me.txtItem(con名称五笔码).Text) & "'," & _
            "'" & Trim(Me.txtItem(con英文名称).Text) & "'," & _
            "'" & Trim(Me.txtItem(con其他别名).Text) & "'," & _
            "'" & Trim(Me.txtItem(con别名拼音码).Text) & "'," & _
            "'" & Trim(Me.txtItem(con别名五笔码).Text) & "'," & _
            "'" & Trim(Me.txtItem(con简要说明).Text) & "'," & _
            IIf(Me.lblNote(0).Tag = "西医", 1, 2) & "," & _
            "'" & StrClass & "','" & strCollate & "'," & _
            mlng分类id & ",'" & strDeptId & "'," & int应用范围 & ")"
    Else
        lngItemId = Me.Tag
        gstrSql = "zl_疾病诊断目录_Update(" & _
            lngItemId & "," & _
            "'" & Trim(Me.txtItem(con诊断编号).Text) & "'," & _
            "'" & Trim(Me.txtItem(con诊断名称).Text) & "'," & _
            "" & _
            "'" & Trim(Me.txtItem(con名称拼音码).Text) & "'," & _
            "'" & Trim(Me.txtItem(con名称五笔码).Text) & "'," & _
            "'" & Trim(Me.txtItem(con英文名称).Text) & "'," & _
            "'" & Trim(Me.txtItem(con其他别名).Text) & "'," & _
            "'" & Trim(Me.txtItem(con别名拼音码).Text) & "'," & _
            "'" & Trim(Me.txtItem(con别名五笔码).Text) & "'," & _
            "'" & Trim(Me.txtItem(con简要说明).Text) & "'," & _
            IIf(Me.lblNote(0).Tag = "西医", 1, 2) & "," & _
            "'" & StrClass & "','" & strCollate & "'," & _
            mlng分类id & ",'" & strDeptId & "'," & int应用范围 & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    With Me.tvwClass
        .Left = Me.fraNote(1).Left + Me.hgdClass.Left + Me.hgdClass.ColWidth(1)
        .Top = Me.fraNote(1).Top + Me.cmdSelect.Top + Me.cmdSelect.Height
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    '屏幕整理调整
    If Me.lblNote(0).Tag = "西医" Then
        Me.lblNote(0).Caption = "西医诊断命名：包括编号、名称、别名和简要说明等"
        Me.chkStandard(3).Visible = False
        Me.txtStandard(3).Visible = False
        Me.lblStandard(3).Visible = False
'        Me.fraNote(2).Height = 1440
'        Me.cmdHelp.Top = 6420
'        Me.cmdOK.Top = 6420
'        Me.cmdCancel.Top = 6420
'        Me.Height = 7305
    Else
        Me.lblNote(0).Caption = "中医诊断命名：包括编号、名称、别名和简要说明等"
        Me.chkStandard(3).Visible = True
        Me.txtStandard(3).Visible = True
        Me.lblStandard(3).Visible = True
'        Me.fraNote(2).Height = 1875
'        Me.cmdHelp.Top = 6855
'        Me.cmdOK.Top = 6855
'        Me.cmdCancel.Top = 6855
'        Me.Height = 7740
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    mlng分类id = Val(Me.txtItem(0).Tag)

    '分类选择树装入
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 疾病诊断分类" & _
            " Where 类别 = [1] " & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.lblNote(0).Tag = "西医", 1, 2))
    
    With rsTemp
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "CLASS")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "CLASS")
            End If
            objNode.Sorted = True
            .MoveNext
        Loop
    End With
    
    '名称等填写
    gstrSql = "select ID,编码,名称,说明" & _
            " From 疾病诊断目录" & _
            " Where ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "增加", -1, Me.Tag))
    
    Me.txtItem(con诊断编号).MaxLength = rsTemp.Fields("编码").DefinedSize
    Me.txtItem(con诊断名称).MaxLength = rsTemp.Fields("名称").DefinedSize
    Me.txtItem(con简要说明).MaxLength = rsTemp.Fields("说明").DefinedSize
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Me.txtItem(con诊断编号).Text = rsTemp!编码
        Me.txtItem(con诊断名称).Text = rsTemp!名称
        Me.txtItem(con简要说明).Text = IIf(IsNull(rsTemp!说明), "", rsTemp!说明)
    Else
        gstrSql = "select nvl(max(编码),'000000') as 编码" & _
                " From 疾病诊断目录" & _
                " Where 类别 = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.lblNote(0).Tag = "西医", 1, 2))
        
        Me.txtItem(con诊断编号).Text = Right(String(10, "0") & Val(rsTemp!编码) + 1, Len(rsTemp!编码))
    End If
    
    '别名填写
    gstrSql = "select nvl(名称,'') as 名称, 性质, nvl(简码,'') as 简码, 码类" & _
            " From 疾病诊断别名" & _
            " Where 诊断id=[1] " & _
            " Order by 性质,码类"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "增加", -1, Me.Tag))
    
    With rsTemp
        Me.txtItem(con英文名称).MaxLength = .Fields("名称").DefinedSize
        Me.txtItem(con其他别名).MaxLength = .Fields("名称").DefinedSize
        Me.txtItem(con名称拼音码).MaxLength = .Fields("简码").DefinedSize
        Me.txtItem(con名称五笔码).MaxLength = .Fields("简码").DefinedSize
        Me.txtItem(con别名拼音码).MaxLength = .Fields("简码").DefinedSize
        Me.txtItem(con别名五笔码).MaxLength = .Fields("简码").DefinedSize
        Do While Not .EOF
            Select Case !性质
            Case 1
                If !码类 = 2 Then
                    Me.txtItem(con名称五笔码).Text = !简码
                Else
                    Me.txtItem(con名称拼音码).Text = !简码
                End If
            Case 2
                Me.txtItem(con英文名称).Text = !名称
            Case 9
                Me.txtItem(con其他别名).Text = !名称
                If !码类 = 2 Then
                    Me.txtItem(con别名五笔码).Text = !简码
                Else
                    Me.txtItem(con别名拼音码).Text = !简码
                End If
            End Select
            .MoveNext
        Loop
    End With
    
    '所属分类填写(至多三项)
    gstrSql = "select I.ID,I.编码,I.名称" & _
            " from 疾病诊断属类 R,疾病诊断分类 I" & _
            " where R.分类ID=I.ID and R.诊断id=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "增加", -1, Me.Tag))
        
    With rsTemp
        Do While Not .EOF
            Me.hgdClass.RowData(.AbsolutePosition - 1) = !ID
            Me.hgdClass.TextMatrix(.AbsolutePosition - 1, 1) = .AbsolutePosition & "."
            Me.hgdClass.TextMatrix(.AbsolutePosition - 1, 2) = "[" & !编码 & "]" & !名称
            If .AbsolutePosition >= 3 Then Exit Do
            .MoveNext
        Loop
        
        '疾病编码填写
        gstrSql = "select distinct 类别 from 疾病编码目录"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Activate")
'        Call SQLTest
        Do While Not rsTemp.EOF
            Select Case rsTemp!类别
            Case "B"
                Me.chkStandard(con中医疾病).Enabled = True
                Me.lblStandard(con中医疾病).Caption = ""
            Case "D"
                Me.chkStandard(conICD10疾病).Enabled = True
                Me.lblStandard(conICD10疾病).Caption = ""
            Case "M"
                Me.chkStandard(con肿瘤形态学).Enabled = True
                Me.lblStandard(con肿瘤形态学).Caption = ""
            Case "Y"
                Me.chkStandard(con损伤与中毒).Enabled = True
                Me.lblStandard(con损伤与中毒).Caption = ""
            End Select
            rsTemp.MoveNext
        Loop
    End With
    
    gstrSql = "select I.类别,I.ID,I.编码,I.名称" & _
            " from 疾病诊断对照 R,疾病编码目录 I" & _
            " where R.疾病ID=I.ID and R.诊断id=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "增加", -1, Me.Tag))
    
    With rsTemp
        Do While Not .EOF
            Select Case !类别
            Case "B"    '中医疾病编码
                Me.chkStandard(con中医疾病).Value = 1
                Me.txtStandard(con中医疾病).BackColor = &H80000005
                Me.txtStandard(con中医疾病).Tag = !ID
                Me.txtStandard(con中医疾病).Text = !编码
                Me.lblStandard(con中医疾病).Caption = TextInLength(!名称, lblStandard(con中医疾病).Width)
            Case "D"    'ICD-10疾病编码
                Me.chkStandard(conICD10疾病).Value = 1
                Me.txtStandard(conICD10疾病).BackColor = &H80000005
                Me.txtStandard(conICD10疾病).Tag = !ID
                Me.txtStandard(conICD10疾病).Text = !编码
                Me.lblStandard(conICD10疾病).Caption = TextInLength(!名称, lblStandard(conICD10疾病).Width)
            Case "M"    '肿瘤形态学编码
                Me.chkStandard(con肿瘤形态学).Value = 1
                Me.txtStandard(con肿瘤形态学).BackColor = &H80000005
                Me.txtStandard(con肿瘤形态学).Tag = !ID
                Me.txtStandard(con肿瘤形态学).Text = !编码
                Me.lblStandard(con肿瘤形态学).Caption = TextInLength(!名称, lblStandard(con肿瘤形态学).Width)
            Case "Y"    '损伤中毒的外部原因
                Me.chkStandard(con损伤与中毒).Value = 1
                Me.txtStandard(con损伤与中毒).BackColor = &H80000005
                Me.txtStandard(con损伤与中毒).Tag = !ID
                Me.txtStandard(con损伤与中毒).Text = !编码
                Me.lblStandard(con损伤与中毒).Caption = TextInLength(!名称, lblStandard(con损伤与中毒).Width)
            End Select
            .MoveNext
        Loop
    End With
    
    Call IniDept
    
    Me.hgdClass.Row = 0
    Call hgdClass_RowColChange
    Me.txtItem(con诊断编号).SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtLocate_KeyPress (vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    With Me.hgdClass
        .Redraw = False
        .ColAlignmentFixed(1) = 7
        .ColAlignmentFixed(2) = 4
        .ColWidth(0) = 0
        .ColWidth(1) = 600
        .ColWidth(2) = .Width - .ColWidth(1) - 15
        .Redraw = True
    End With
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 3000
        .Add , "编码", "编码", 900
    End With
    Me.lvwList.ColumnHeaders("编码").Position = 1
    Me.Lvw科室.MultiSelect = False
    For intCount = Me.lblStandard.LBound To Me.lblStandard.UBound
        Me.chkStandard(intCount).Value = 0
        Me.lblStandard(intCount).Caption = "(未设置该类标准编码)"
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.tvwClass.Visible Or Me.lvwList.Visible Then
        Me.tvwClass.Visible = False
        Me.lvwList.Visible = False
        Cancel = True
    End If
End Sub

Private Sub hgdClass_GotFocus()
    Me.hgdClass.RowSel = Me.hgdClass.Row
    Me.hgdClass.ColSel = Me.hgdClass.Cols - 1
End Sub

Private Sub hgdClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then Exit Sub
    With Me.hgdClass
        For intCount = .Row To .Rows - 2
            .RowData(intCount) = .RowData(intCount + 1)
            .TextMatrix(intCount, 2) = .TextMatrix(intCount + 1, 2)
            If .TextMatrix(intCount, 2) = "" Then
                .TextMatrix(intCount, 1) = ""
            Else
                .TextMatrix(intCount, 1) = intCount + 1 & "."
            End If
        Next
        .RowData(.Rows - 1) = 0
        .TextMatrix(.Rows - 1, 1) = ""
        .TextMatrix(.Rows - 1, 2) = ""
    End With
End Sub

Private Sub hgdClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub hgdClass_RowColChange()
    With Me.cmdSelect
        .Top = Me.hgdClass.Top + Me.hgdClass.RowHeight(0) * Me.hgdClass.Row + 15
        .Left = Me.hgdClass.Left + Me.hgdClass.Width - .Width - 15
    End With
End Sub

Private Sub lblItem_Click(Index As Integer)
    Me.txtItem(Index).SetFocus
End Sub

Private Sub lvwList_DblClick()
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwList
        Me.txtStandard(.Tag).Tag = Mid(.SelectedItem.Key, 2)
        Me.txtStandard(.Tag).Text = .SelectedItem.SubItems(Me.lvwList.ColumnHeaders("编码").Index - 1)
        Me.lblStandard(.Tag).Caption = TextInLength(.SelectedItem.Text, lblStandard(.Tag).Width)
        Me.txtStandard(.Tag).SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        Call lvwList_DblClick
    End Select
End Sub

Private Sub lvwList_LostFocus()
    Me.lvwList.Visible = False
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    With Me.hgdClass
        .RowData(.Row) = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .TextMatrix(.Row, 1) = .Row + 1 & "."
        .TextMatrix(.Row, 2) = Me.tvwClass.SelectedItem.Text
    End With
    Me.hgdClass.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        Call tvwClass_DblClick
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If cmdSelect Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    Select Case Index
    Case con诊断名称, con其他别名, con简要说明
        Call zlCommFun.OpenIme(True)
    End Select
    Me.txtItem(Index).SelStart = 0: Me.txtItem(Index).SelLength = 100
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case con诊断编号
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case vbKeyReturn
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Case Else
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        End Select
        KeyAscii = 0
    Case con诊断名称, con其他别名
        If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case con英文名称, con名称拼音码, con名称五笔码, con别名拼音码, con别名五笔码
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case vbKeyReturn
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii = vbKeySpace Then Exit Sub
        End Select
        KeyAscii = 0
    Case con简要说明
        If InStr("%_'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
End Sub

Private Sub txtItem_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case con诊断名称
        Me.txtItem(con名称拼音码).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), False)
        Me.txtItem(con名称五笔码).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), True)
    Case con其他别名
        Me.txtItem(con别名拼音码).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), False)
        Me.txtItem(con别名五笔码).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), True)
    End Select
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    Select Case Index
    Case con诊断名称, con其他别名, con简要说明
        Call zlCommFun.OpenIme(False)
    End Select
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

Private Sub txtStandard_GotFocus(Index As Integer)
    Me.txtStandard(Index).SelStart = 0: Me.txtStandard(Index).SelLength = 100
End Sub

Private Sub txtStandard_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("%^&*()+|=`'"":,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,编码,名称,简码" & _
            " from 疾病编码目录" & _
            " where 类别='" & _
            Switch(Index = conICD10疾病, "D", _
                   Index = con损伤与中毒, "Y", _
                   Index = con肿瘤形态学, "M", _
                   Index = con中医疾病, "B") & "'" & _
            "   and (编码 like [1] " & _
            "       OR 简码 like [2] " & _
            "       OR 名称 like [2])" & _
            " and (撤档时间 Is Null Or 撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Me.txtStandard(Index).Text) & "%", gstrMatch & Trim(Me.txtStandard(Index).Text) & "%")
    
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "未找到指定标准疾病编码", vbExclamation, gstrSysName
            Me.txtStandard(Index).SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txtStandard(Index).Tag = !ID
            Me.txtStandard(Index).Text = IIf(IsNull(!编码), "", !编码)
            Me.lblStandard(Index).Caption = TextInLength(IIf(IsNull(!名称), "", !名称), lblStandard(Index).Width)
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Exit Sub
        End If
        
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !ID, !名称, "ITEM", "ITEM")
            objItem.SubItems(Me.lvwList.ColumnHeaders("编码").Index - 1) = !编码
            .MoveNext
        Loop
        With Me.lvwList
            .Tag = Index
            .ListItems(1).Selected = True
            .Left = Me.SSTab.Width - Me.lvwList.Width - 50
            .Top = Me.fraNote(2).Top + Me.txtStandard(Index).Top - .Height
            .Visible = True
            .SetFocus
        End With
    End With
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function TextInLength(ByVal strText As String, ByVal DispLen As Single) As String
    Dim iDispNum As Integer
    TextInLength = strText
    
    On Error Resume Next
    If Me.TextWidth(strText) < DispLen Then Exit Function
    iDispNum = CInt((DispLen - Me.TextWidth("...")) / Me.TextWidth(" "))
    If Me.TextWidth(MidB(strText, 1, iDispNum) & "...") > DispLen Then iDispNum = iDispNum - 1
    TextInLength = MidB(strText, 1, iDispNum)
    TextInLength = Mid(TextInLength, 1, Len(TextInLength)) & "..."
End Function
