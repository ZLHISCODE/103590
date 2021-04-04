VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeptSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "部门设置"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmDeptSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4635
      TabIndex        =   12
      Top             =   7230
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5895
      TabIndex        =   13
      Top             =   7230
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   225
      TabIndex        =   35
      Top             =   7230
      Width           =   1100
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   2205
      TabIndex        =   34
      Top             =   7260
      Width           =   1335
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6555
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   6735
      Begin VB.Frame fra基本信息 
         Caption         =   "基本信息"
         Height          =   3975
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   3345
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   840
            MaxLength       =   100
            TabIndex        =   3
            Top             =   1380
            Width           =   1275
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   5
            Left            =   840
            MaxLength       =   50
            TabIndex        =   5
            Top             =   2100
            Width           =   2355
         End
         Begin VB.CommandButton cmd上级 
            Caption         =   "…"
            Height          =   240
            Left            =   2910
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   2490
            Width           =   255
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   840
            TabIndex        =   1
            Top             =   660
            Width           =   2355
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   840
            MaxLength       =   100
            TabIndex        =   2
            Top             =   1020
            Width           =   1275
         End
         Begin VB.TextBox txtEdit 
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   1
            Left            =   960
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "编码"
            Text            =   "111111"
            Top             =   345
            Width           =   1035
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   3540
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cbo环境类别 
            Height          =   300
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3180
            Width           =   2055
         End
         Begin VB.ComboBox cbo负责人 
            Height          =   300
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2820
            Width           =   1815
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   840
            MaxLength       =   3
            TabIndex        =   4
            Top             =   1740
            Width           =   1275
         End
         Begin VB.TextBox txtTemp 
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   840
            MaxLength       =   10
            TabIndex        =   18
            TabStop         =   0   'False
            Tag             =   "编码"
            Text            =   "1111111111"
            Top             =   300
            Width           =   1275
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   840
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   6
            Top             =   2460
            Width           =   2355
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "别名(&B)"
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   38
            Top             =   1440
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "位置(&L)"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   2160
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "编码(&U)"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "名称(&N)"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "简码(&S)"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "上级(&P)"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   2520
            Width           =   630
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "院区(&B)"
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   3600
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "环境类别(&T)"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   3240
            Width           =   1005
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "部门负责人(&D)"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   2880
            Width           =   1170
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "顺序(&R)"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   630
         End
      End
      Begin VB.ComboBox cmb诊疗科目编码 
         Height          =   300
         Left            =   3450
         TabIndex        =   11
         Text            =   "cmb诊疗科目编码"
         Top             =   6165
         Width           =   3105
      End
      Begin VB.Frame fra说明 
         Caption         =   "工作性质说明"
         Height          =   2310
         Left            =   0
         TabIndex        =   15
         Top             =   4200
         Width           =   3345
         Begin VB.Label lbl说明 
            Caption         =   "Label3"
            Height          =   1575
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   3015
         End
      End
      Begin MSComctlLib.ListView lvw性质 
         Height          =   5205
         Left            =   3450
         TabIndex        =   10
         Top             =   480
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   9181
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "工作性质"
            Object.Tag             =   "工作性质"
            Text            =   "工作性质"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "服务对象"
            Object.Tag             =   "服务对象"
            Text            =   "服务对象"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "默认值"
            Object.Tag             =   "默认值"
            Text            =   "默认值"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl工作性质 
         AutoSize        =   -1  'True
         Caption         =   "工作性质(&W)"
         Height          =   180
         Left            =   3480
         TabIndex        =   30
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lbl诊疗科目编码 
         AutoSize        =   -1  'True
         Caption         =   "临床性质的诊疗科室编码(&D)"
         Height          =   180
         Left            =   3450
         TabIndex        =   29
         Top             =   5910
         Width           =   2250
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   7470
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
            Picture         =   "frmDeptSet.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptSet.frx":0326
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6465
      Index           =   1
      Left            =   240
      TabIndex        =   31
      Top             =   390
      Visible         =   0   'False
      Width           =   6735
      Begin MSComctlLib.ListView Lvw科室 
         Height          =   5925
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   10451
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   $"frmDeptSet.frx":0640
         Height          =   360
         Left            =   120
         TabIndex        =   33
         Top             =   45
         Width           =   6480
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   6990
      Left            =   120
      TabIndex        =   36
      Top             =   30
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12330
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本设置"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "科室病区对应"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFind 
      Caption         =   "查找(&F)"
      Height          =   255
      Left            =   1500
      TabIndex        =   37
      Top             =   7305
      Width           =   735
   End
   Begin VB.Menu mnuShort 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPatient 
         Caption         =   "门诊病人(&O)"
         Index           =   0
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "住院病人(&I)"
         Index           =   1
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "门诊和住院病人(&B)"
         Index           =   2
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "不服务于病人(&N)"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmDeptSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const mlng编码长度 As Long = 10

Dim mstr上级部门ID As String     '当前编辑的上级部门ID
Dim mstrID As String         '当前编辑的部门ID
Dim mstr上级编码 As String    '原始的上级编码的值
Dim mstr编码 As String        '原始的本级编码的值
Dim mint编码 As Integer       '修改前包括下级在内的编码最长的长度
Dim mblnItem As Boolean       '是否点击了某项
Dim mblnChange As Boolean     '是否改变了
Dim mbln药店  As Boolean
Dim mint原工作性质 As Integer   '0-不含药库药房性质；1-含药房性质；2-只含药库性质
Dim mint性质 As Integer         '1-临床科室;2-病区;3-临床且病区
'Dim mint服务对象 As Integer     '1-门诊;2-住院;3-门诊或住院
Dim mint服务对象_临床 As Integer
Dim mint服务对象_病区 As Integer
Private mlng配置中心 As Long
Private mstrPrivs As String
Private mint编辑状态 As Integer     '1-新增 2-修改
Private mStr部门id As String           '部门id
Private mint编辑模式 As Integer     '1-按层次显示 2-按性质显示
Private mstr上级码 As String
Private mstr性质 As String          '记录是选择的什么性质的
Private mintInputMethod As Integer   '编码录入方式，0-按上级编码，1-自由录入
Private mblnPACSInterface As Boolean        '启用影像信息系统接口

Private Function Check床位状况(ByVal lng部门ID As Long, ByVal lng病区id As Long, ByVal int性质 As Integer) As Boolean
    'int性质：0-部门性质检查;1-病区科室对应检查
    
    Dim rsTmp As ADODB.Recordset
    
    If lng部门ID = 0 Then Exit Function
    
    On Error GoTo ErrHandle
    If int性质 = 0 Then
        '部门性质检查：临床或护理性质取消时，检查床位状况
        gstrSQL = "Select 1 From 床位状况记录 Where (科室id = [1] Or 病区id = [1]) And Rownum = 1"
    Else
        '病区科室对应检查：对应病区或科室取消时，检查床位记录
        gstrSQL = "Select 1 From 床位状况记录 Where (科室id = [1] And 病区id = [2]) And Rownum = 1"
    End If
        
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "床位状况检查", lng部门ID, lng病区id)
    
    Check床位状况 = (rsTmp.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub IniStationNo()
    Dim lst As ListItem
    Dim rsRecord As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    lblStationNo.Visible = True
    cmbStationNo.Visible = True
     
    frmDeptSet.Height = frmDeptSet.Height + 50
    tabMain.Height = tabMain.Height + 50
    fraMain(0).Height = fraMain(0).Height + 50
    fra基本信息.Height = fra基本信息.Height + 50
    fra说明.Top = fra说明.Top + 50
    lvw性质.Height = lvw性质.Height + 250
    fraMain(1).Height = fraMain(1).Height + 50
    Lvw科室.Height = Lvw科室.Height + 50

    cmdHelp.Top = cmdHelp.Top
    lblFind.Top = cmdHelp.Top + 100
    txtFind.Top = cmdHelp.Top + 25
    cmdOK.Top = cmdHelp.Top
    cmdCancel.Top = cmdHelp.Top
    
    strSQL = "select 编号,名称 from zlnodelist"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "站点查询")
    
    If rsRecord.RecordCount = 0 Then
        lblStationNo.Visible = False
        cmbStationNo.Visible = False
    Else
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!编号 & "-" & rsRecord!名称
                rsRecord.MoveNext
            Loop
        End With
    End If

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Ini病区科室对应(ByVal str科室id As String, Optional ByVal int初始 As Integer = 1)
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strCon As String
    Dim strCon性质 As String
    Dim strCon服务对象 As String
    
    tabMain.Tabs.Clear
    tabMain.Tabs.Add , "_基本信息", "基本信息"
    
    On Error GoTo ErrHandle
    If glngSys \ 100 <> 1 Then Exit Sub
    
    '当前科室是临床或病区，并且服务与门诊或住院病人时才能设置
    If mint性质 = 0 Or (mint服务对象_临床 = 0 And mint服务对象_病区 = 0) Then Exit Sub
    
    tabMain.Tabs.Add , "_科室病区对应", "科室病区对应"
    
    '当前科室的服务性质改变时才更新病区科室列表的数据
    If (mint服务对象_临床 = Val(Mid(Lvw科室.Tag, 1, 1)) And mint服务对象_病区 = Val(Mid(Lvw科室.Tag, 2, 1))) And int初始 <> 1 Then Exit Sub
    
    Lvw科室.Tag = CStr(mint服务对象_临床) & CStr(mint服务对象_病区)
    
    '根据部门性质和服务对象设置条件
    If mint性质 = 1 Then
        '取病区
        If mint服务对象_临床 = 1 Then
            strCon性质 = " 服务对象 IN(1,3) And 工作性质 = '护理' "
        ElseIf mint服务对象_临床 = 2 Then
            strCon性质 = " 服务对象 IN(2,3) And 工作性质 = '护理' "
        Else
            strCon性质 = " 服务对象 IN(1,2,3) And 工作性质 = '护理' "
        End If
    ElseIf mint性质 = 2 Then
        '取临床
        If mint服务对象_病区 = 1 Then
            strCon性质 = " 服务对象 IN(1,3) And 工作性质 = '临床' "
        ElseIf mint服务对象_病区 = 2 Then
            strCon性质 = " 服务对象 IN(2,3) And 工作性质 = '临床' "
        Else
            strCon性质 = " 服务对象 IN(1,2,3) And 工作性质 = '临床' "
        End If
    ElseIf mint性质 = 3 Then
        '取临床和病区
        If mint服务对象_临床 = 1 Then
            strCon性质 = " (服务对象 IN(1,3) And 工作性质 = '护理') "
        ElseIf mint服务对象_临床 = 2 Then
            strCon性质 = " (服务对象 IN(2,3) And 工作性质 = '护理') "
        Else
            strCon性质 = " (服务对象 IN(1,2,3) And 工作性质 = '护理') "
        End If
        
        If mint服务对象_病区 = 1 Then
            strCon性质 = strCon性质 & " Or (服务对象 IN(1,3) And 工作性质 = '临床') "
        ElseIf mint服务对象_病区 = 2 Then
            strCon性质 = strCon性质 & " Or (服务对象 IN(2,3) And 工作性质 = '临床') "
        Else
            strCon性质 = strCon性质 & " Or (服务对象 IN(1,2,3) And 工作性质 = '临床') "
        End If
    End If
    
    mstr性质 = strCon性质
    gstrSQL = " Select Distinct 编码||'-'||名称 科室,ID From 部门表 " & _
         " Where ID in (Select 部门ID From 部门性质说明 Where " & strCon性质 & ")" & _
         " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
         " Order By 编码||'-'||名称 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "所有临床科室或病区")
    
    Lvw科室.ListItems.Clear
    With rsTmp
        Do While Not .EOF
            Lvw科室.ListItems.Add , "_" & !ID, !科室, 1, 1
            .MoveNext
        Loop
    End With
    
    '取病区科室对应关系
    If mstrID <> "" Then
        If mint性质 = 1 Then
            strCon = " 科室id = [1] "
        ElseIf mint性质 = 2 Then
            strCon = " 病区id = [1] "
        ElseIf mint性质 = 3 Then
            strCon = " 科室id = [1] Or 病区id = [1] "
        End If
        
        gstrSQL = "Select Distinct 病区id,科室id From  病区科室对应 Where " & strCon
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "病区科室对应", Val(str科室id))
        
        With rsTmp
            Do While Not .EOF
                For n = 1 To Lvw科室.ListItems.Count
                    If mint性质 = 1 Then
                        If Val(Mid(Lvw科室.ListItems(n).Key, 2)) = !病区ID Then
                            Lvw科室.ListItems(n).Tag = 1
                            Lvw科室.ListItems(n).Checked = True
                        End If
                    ElseIf mint性质 = 2 Then
                        If Val(Mid(Lvw科室.ListItems(n).Key, 2)) = !科室ID Then
                            Lvw科室.ListItems(n).Tag = 1
                            Lvw科室.ListItems(n).Checked = True
                        End If
                    ElseIf mint性质 = 3 Then
                        If Val(Mid(Lvw科室.ListItems(n).Key, 2)) = !病区ID Or Val(Mid(Lvw科室.ListItems(n).Key, 2)) = !科室ID Then
                            Lvw科室.ListItems(n).Tag = 1
                            Lvw科室.ListItems(n).Checked = True
                        End If
                    End If
                Next
                .MoveNext
            Loop
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
'    If gstrNodeNo = "-" Then Exit Sub
    
    If cmbStationNo.ListCount = 0 Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub
Private Sub Set病区科室对应()
    Dim i As Long
    Dim str性质 As String
    Dim bln性质_临床 As Boolean
    Dim bln性质_病区 As Boolean
    Dim str服务对象_临床 As String
    Dim str服务对象_病区 As String
    
    mint性质 = 0
    With lvw性质
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                str性质 = IIF(str性质 = "", "", str性质 & ",") & .ListItems(i).Text
            End If
        Next
        If InStr(1, str性质, "临床") Then bln性质_临床 = True
        If InStr(1, str性质, "护理") Then bln性质_病区 = True
        
        If bln性质_临床 = True And bln性质_病区 = True Then
            mint性质 = 3
        ElseIf bln性质_临床 = True Then
            mint性质 = 1
        ElseIf bln性质_病区 = True Then
            mint性质 = 2
        End If
    End With
    
    mint服务对象_临床 = 0
    mint服务对象_病区 = 0
    With lvw性质
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                If InStr(1, .ListItems(i).Text, "临床") > 0 Then
                    str服务对象_临床 = IIF(str服务对象_临床 = "", "", str服务对象_临床 & ",") & .ListItems(i).SubItems(1)
                End If
                If InStr(1, .ListItems(i).Text, "护理") > 0 Then
                    str服务对象_病区 = IIF(str服务对象_病区 = "", "", str服务对象_病区 & ",") & .ListItems(i).SubItems(1)
                End If
            End If
        Next
        
        If InStr(1, str服务对象_临床, "门诊和住院病人") > 0 Then
            mint服务对象_临床 = 3
        ElseIf InStr(1, str服务对象_临床, "住院病人") > 0 Then
            mint服务对象_临床 = 2
        ElseIf InStr(1, str服务对象_临床, "门诊病人") > 0 Then
            mint服务对象_临床 = 1
        End If
        
        If InStr(1, str服务对象_病区, "门诊和住院病人") > 0 Then
            mint服务对象_病区 = 3
        ElseIf InStr(1, str服务对象_病区, "住院病人") > 0 Then
            mint服务对象_病区 = 2
        ElseIf InStr(1, str服务对象_病区, "门诊病人") > 0 Then
            mint服务对象_病区 = 1
        End If
    End With
    
    Call Ini病区科室对应(mstrID, 0)
End Sub

Private Sub cbo环境类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cbo环境类别.ListIndex = -1
End Sub

Private Sub cmb诊疗科目编码_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim blnTmp As Boolean
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte
    Dim strTemp As String
    
    Select Case KeyAscii
        Case Is = 13
            '检验输入的编码是不是存在
    '        If cmb诊疗科目编码.Enabled Then
    '            If Trim(cmb诊疗科目编码.Text) = "" Then MsgBox "请选择一个诊疗科目编码！", vbInformation, gstrSysName: Exit Sub
    '            blnTmp = False
    '            For i = 0 To cmb诊疗科目编码.ListCount - 1
    '                If cmb诊疗科目编码.List(i) Like Trim(cmb诊疗科目编码.Text) & "* *" Then
    '                    blnTmp = True
    '                    cmb诊疗科目编码.ListIndex = i
    '                    Exit For
    '                ElseIf cmb诊疗科目编码.List(i) Like "* " & Trim(cmb诊疗科目编码.Text) Then
    '                    blnTmp = True
    '                    cmb诊疗科目编码.ListIndex = i
    '                    Exit For
    '                ElseIf Trim(cmb诊疗科目编码.List(i)) = Trim(cmb诊疗科目编码.Text) Then
    '                    blnTmp = True
    '                    cmb诊疗科目编码.ListIndex = i
    '                    Exit For
    '                End If
    '            Next
    '            If blnTmp = False Then
    '                MsgBox "输入的诊疗科目编码不存在，请重新输入！", vbExclamation, gstrSysName
    '                cmb诊疗科目编码.Text = ""
    '                cmb诊疗科目编码.SetFocus
    '                Exit Sub
    '            End If
    '        End If
            
            strTemp = Trim(UCase(cmb诊疗科目编码.Text))
            If strTemp = "" Then Exit Sub
            If mStr部门id = "" Or cmb诊疗科目编码.Enabled = False Then
                '不具有临床性质
                gstrSQL = "Select rownum as id, 编码,名称  From 临床性质 where 编码 like [1] or 名称 like [2] or 简码 like [3] Order By 序号"
            Else
                gstrSQL = "select rownum as id, A.编码,A.名称,B.工作性质 from 临床性质 A,临床部门 B " & _
                    "where A.编码=B.工作性质(+) and b.部门ID(+)=[4] and ( a.编码 like [1] or a.名称 like [2] or a.简码 like [3]) order by A.序号"
            End If
            
            vRect = zlControl.GetControlRect(cmb诊疗科目编码.hwnd)
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "诊疗科目", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, cmb诊疗科目编码.Height, True, False, True, strTemp & gstrLike, strTemp & gstrLike, strTemp & gstrLike, Val(mStr部门id))
                
            If Not rsTemp Is Nothing Then
                cmb诊疗科目编码.Text = Format(rsTemp("编码"), "!@@@@@") & rsTemp("名称")
            End If
    End Select
End Sub

Private Sub cmb诊疗科目编码_Validate(Cancel As Boolean)
    Dim i As Long
    Dim blnTmp As Boolean
    
    '检验输入的编码是不是存在
    If cmb诊疗科目编码.Enabled Then
        If Trim(cmb诊疗科目编码.Text) = "" Then MsgBox "请选择一个诊疗科目编码！", vbInformation, gstrSysName: Cancel = True: Exit Sub
        blnTmp = False
        For i = 0 To cmb诊疗科目编码.ListCount - 1
            If cmb诊疗科目编码.List(i) Like Trim(cmb诊疗科目编码.Text) & "* *" Then
                blnTmp = True
                cmb诊疗科目编码.ListIndex = i
                Exit For
            ElseIf cmb诊疗科目编码.List(i) Like "* " & Trim(cmb诊疗科目编码.Text) Then
                blnTmp = True
                cmb诊疗科目编码.ListIndex = i
                Exit For
            ElseIf Trim(cmb诊疗科目编码.List(i)) = Trim(cmb诊疗科目编码.Text) Then
                blnTmp = True
                cmb诊疗科目编码.ListIndex = i
                Exit For
            End If
        Next
        If blnTmp = False Then
            MsgBox "输入的诊疗科目编码不存在，请重新输入！", vbExclamation, gstrSysName
            cmb诊疗科目编码.Text = ""
            cmb诊疗科目编码.ListIndex = -1
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim lngTmp As Long
    
    If IsValid() = False Then Exit Sub
        
    '检查是否与已删除部门的名称相同
    If mstrID = "" Then
        If CheckSameDept(txtEdit(2).Text, lngTmp) Then
            If MsgBox("当前录入的部门名称与已删除的部门名称相同。" & vbNewLine & "【是】将其恢复，【否】不恢复也不保存。", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            mstrID = lngTmp
        End If
    End If
    
    If Save部门() = False Then Exit Sub
    
    '改变主窗口的显示
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    Else
    
    End If
    '连续增加
    mstrID = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(5).Text = ""
    txtEdit(6).Text = ""
    If mintInputMethod = 0 Then
        '自动生成
        txtEdit(1).Text = GetMaxLocalCode(mstr上级部门ID, "部门表")
    Else
        '自由录入编码
        txtEdit(1).Text = ""
    End If
    
    For i = 1 To lvw性质.ListItems.Count
        lvw性质.ListItems(i).Checked = False
    Next
    lbl诊疗科目编码.Enabled = False
    cmb诊疗科目编码.Enabled = False
    cmb诊疗科目编码.ListIndex = -1
    
    txtTemp.MaxLength = GetLocalCodeLength(mstr上级部门ID, "部门表")
    txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(txtTemp.Text)
    
    Call ShowTab(1)
    txtEdit(2).SetFocus
    
    mblnChange = False
End Sub

Private Function IsValid() As Boolean
    Dim i As Long
    Dim blnTmp As Boolean
    Dim strTemp As String
    Dim int现工作性质 As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    strSQL = "SELECT 别名 FROM 部门表 Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "结算方式编辑")
    
    txtEdit(6).MaxLength = rsTemp.Fields("别名").DefinedSize
    
    For i = 1 To 6
        If i <> 4 Then
            If zlCommFun.StrIsValid(Trim(txtEdit(i).Text), txtEdit(i).MaxLength) = False Then
                Call ShowTab(1)
                txtEdit(i).SetFocus
                zlControl.TxtSelAll txtEdit(i)
                Exit Function
            End If
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)

    If Len(Trim(txtEdit(4).Text)) = 0 And Me.Tag = "恢复" Then
        MsgBox "上级不能为空。", vbExclamation, gstrSysName
        Call ShowTab(1)
        txtEdit(4).SetFocus
        Exit Function
    End If
    
    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(1).Text) = 0 Then
            MsgBox "编码不能为空。", vbExclamation, gstrSysName
            Call ShowTab(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
    Else
        If Len(txtEdit(1).Text) < txtEdit(1).MaxLength Then
            MsgBox "编码的长度不够。", vbExclamation, gstrSysName
            Call ShowTab(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txtEdit(1).Text) Or InStr(txtEdit(1).Text, ",") > 0 Or InStr(txtEdit(1).Text, ".") > 0 Then
        MsgBox "编码应由数字组成。", vbExclamation, gstrSysName
        Call ShowTab(1)
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        Call ShowTab(1)
        txtEdit(2).Text = ""
        txtEdit(2).SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtEdit(2).Text, vbFromUnicode)) > 100 Then
        MsgBox "名称长度不能超过50个汉字或者100个字符，请重新录入！", vbInformation, gstrSysName
        txtEdit(2).SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtEdit(3).Text, vbFromUnicode)) > 100 Then
        MsgBox "编码长度不能超过100个字符，请重新录入！", vbInformation, gstrSysName
        txtEdit(3).SetFocus
        Exit Function
    End If
    
    '重新检验输入的编码是不是存在
    If glngSys \ 100 = 8 Then
        '药店系统不处理
    Else
        '对临床性质的判断
        If cmb诊疗科目编码.Enabled Then
            If Trim(cmb诊疗科目编码.Text) = "" Then
                Call ShowTab(1)
                MsgBox "请为临床部门设置它的诊疗科目编码。", vbExclamation, gstrSysName
                cmb诊疗科目编码.SetFocus
                Exit Function
            End If
            blnTmp = False
            For i = 0 To cmb诊疗科目编码.ListCount - 1
                If cmb诊疗科目编码.List(i) = cmb诊疗科目编码.Text Then
                    blnTmp = True
                    Exit For
                End If
            Next
            If blnTmp = False Then
                MsgBox "输入的诊疗科目编码不存在，请重新输入！", vbExclamation, gstrSysName
                Call ShowTab(1)
                cmb诊疗科目编码.Text = ""
                cmb诊疗科目编码.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '检查部门的工作性质变化（主要是药库药房性质变化），如果有药库药房性质的变换则检查库存，有库存则提示
    On Error Resume Next
    If mstrID <> "" Then
        int现工作性质 = 0
        For i = 1 To lvw性质.ListItems.Count
            If lvw性质.ListItems(i).Checked = True Then
                If int现工作性质 <> 1 Then
                    If InStr(lvw性质.ListItems(i), "药房") > 0 Or lvw性质.ListItems(i) = "制剂室" Then
                        int现工作性质 = 1
                    ElseIf InStr(lvw性质.ListItems(i), "药库") > 0 Then
                        int现工作性质 = 2
                    End If
                End If
            End If
        Next
        If int现工作性质 <> mint原工作性质 Then
            gstrSQL = "select 1 from 药品库存 where 库房ID=[1] and rownum=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
            
            If rsTemp.RecordCount > 0 Then
                If MsgBox("该部门含有的药库或药房性质发生了变化，可能会影响库存药品的分批属性。是否确定？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    IsValid = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save部门() As Boolean
    Dim strSQL As String
    Dim lng部门ID As Long
    Dim str部门性质 As String
    Dim str临床性质 As String
    Dim i As Integer, int对象 As Integer
    Dim nod As Node
    Dim lst As ListItem
    Dim str病区科室 As String
    Dim str站点 As String
    Dim BeginTrans As Boolean
    Dim arrSQL As Variant
    Dim lngType As Long
    
    On Error GoTo ErrHandle
    
    arrSQL = Array()
    '把所有选中的工作性质做成一个串
    For i = 1 To lvw性质.ListItems.Count
        If lvw性质.ListItems(i).Checked = True Then
            str部门性质 = str部门性质 & lvw性质.ListItems(i) & ":"
            
            If mbln药店 = True Then
                '药店只处理门诊病人
                str部门性质 = str部门性质 & "1:"
            Else
                Select Case lvw性质.ListItems(i).SubItems(1)
                     Case "门诊病人"
                        str部门性质 = str部门性质 & "1:"
                     Case "住院病人"
                        str部门性质 = str部门性质 & "2:"
                     Case "门诊和住院病人"
                        str部门性质 = str部门性质 & "3:"
                     Case Else
                        str部门性质 = str部门性质 & "0:"
                End Select
            End If
        End If
    Next
    If cmb诊疗科目编码.Enabled = False Then
        str临床性质 = ""
    Else
        str临床性质 = Trim(Left(cmb诊疗科目编码.List(cmb诊疗科目编码.ListIndex), 4))
    End If
    
    If mstrID = "" Then       '新增一条记录
        If Check重复部门(mstr上级部门ID, Trim(txtEdit(2).Text)) = True Then
            MsgBox "该级下面已有该部门，不能添加相同部门！", vbInformation, gstrSysName
            Exit Function
        End If
        
        lng部门ID = Sys.NextId("部门表")
        lngType = RISBaseItemOper.AddNew
        gstrSQL = "zl_部门表_insert(" & lng部门ID & "," & IIF(mstr上级部门ID = "", "null", mstr上级部门ID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & UCase(txtEdit(3).Text) & "','" & txtEdit(5).Text & "','" & str部门性质 & "','" & str临床性质 & "' "
        
        If cmbStationNo.Text = "" Then
            str站点 = "Null"
        Else
            str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
        End If
        
        gstrSQL = gstrSQL & ",'" & Trim(cbo环境类别.Text) & "'," & IIF(cmbStationNo.Text = "", "Null", str站点)
        gstrSQL = gstrSQL & "," & IIF(txtEdit(0).Text = "", "Null", txtEdit(0).Text)
        gstrSQL = gstrSQL & ",'" & txtEdit(6).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        '修改主界面的内容
        With frmDeptManage.tvwMain_S
            '增加到TreeView中
            If mint编辑模式 = 1 Then
                Set nod = .Nodes.Add(IIF(mstr上级部门ID = "", "Root", "C" & mstr上级部门ID), tvwChild, _
                    "C" & lng部门ID, "【" & txtTemp.Text & txtEdit(1).Text & "】" & txtEdit(2).Text, "Dept", "Dept")
                nod.Sorted = True
            Else
                Set nod = .Nodes.Add(IIF(mstr上级部门ID = "", "Root", "C" & mstr上级码), tvwChild, _
                    "C" & txtEdit(1).Text & "|" & lng部门ID, "【" & txtTemp.Text & txtEdit(1).Text & "】" & txtEdit(2).Text, "Dept", "Dept")
                nod.Sorted = True
            End If
            '增加到ListView中
        End With
        With frmDeptManage.lvwMain
            If frmDeptManage.tvwMain_S.SelectedItem.Key = IIF(mstr上级部门ID = "", "Root", "C" & mstr上级部门ID) Then
                Set lst = .ListItems.Add(, "C" & lng部门ID, txtEdit(2).Text, "Dept", "Dept")
                For i = 2 To .ColumnHeaders.Count
                    Select Case .ColumnHeaders(i).Text
                        Case "编码"
                            lst.SubItems(i - 1) = txtTemp.Text & txtEdit(1).Text
                        Case "名称"
                            lst.SubItems(i - 1) = txtEdit(2).Text
                        Case "简码"
                            lst.SubItems(i - 1) = txtEdit(3).Text
                        Case "位置"
                            lst.SubItems(i - 1) = txtEdit(5).Text
                        Case "建档时间"
                            lst.SubItems(i - 1) = Format(Sys.Currentdate, "yyyy-MM-dd")
                        Case "撤档时间"
                            lst.SubItems(i - 1) = "3000-01-01"
                        Case "上级部门"
                            lst.SubItems(i - 1) = txtEdit(4).Text
                    End Select
                Next
                If .ListItems.Count = 1 Then
                    .ListItems(1).Selected = True
                    Call frmDeptManage.lvwMain_ItemClick(.ListItems(1))
                End If
            End If
        End With
        
    Else
        '修改
        lng部门ID = Val(mstrID)
        lngType = RISBaseItemOper.Modify
        gstrSQL = "zl_部门表_update(" & mstrID & "," & IIF(mstr上级部门ID = "", "null", mstr上级部门ID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & UCase(txtEdit(3).Text) & "','" & txtEdit(5).Text & "'," & Len(mstr编码) + 1 & ",'" & str部门性质 & "','" & str临床性质 & "' "
        
        gstrSQL = gstrSQL & ",'" & Trim(cbo环境类别.Text) & "',"
        If Me.cbo负责人.ListIndex <= 0 Then
            gstrSQL = gstrSQL & "null,"
        Else
            gstrSQL = gstrSQL & Me.cbo负责人.ItemData(Me.cbo负责人.ListIndex) & ","
        End If
        If cmbStationNo.Text = "" Then
            str站点 = "Null"
        Else
            str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
        End If
        
        gstrSQL = gstrSQL & IIF(cmbStationNo.Text = "", "Null", str站点)
        gstrSQL = gstrSQL & "," & IIF(txtEdit(0).Text = "", "Null", txtEdit(0).Text)
        gstrSQL = gstrSQL & ",'" & txtEdit(6).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    If glngSys \ 100 = 1 Then
        With Lvw科室
            For i = 1 To .ListItems.Count
                If .ListItems(i).Checked = True Then
                    str病区科室 = IIF(str病区科室 = "", "", str病区科室 & ",") & Mid(.ListItems(i).Key, 2)
                End If
            Next
        End With
        gstrSQL = "Zl_病区科室对应_Update(" & IIF(mstrID <> "", Val(mstrID), lng部门ID) & "," & mint性质 & "," & IIF(str病区科室 = "", "Null", "'" & str病区科室 & "'") & ")"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    gcnOracle.BeginTrans: BeginTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
    Next
    
    If glngSys \ 100 = 1 And mblnPACSInterface Then
        If Not gobjRIS Is Nothing Then
            If gobjRIS.HISBasicDictTable(10, lngType, lng部门ID) <> 1 Then
                gcnOracle.RollbackTrans
                MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISBasicDictTable)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans
            MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISBasicDictTable)接口，请与系统管理员联系。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    gcnOracle.CommitTrans: BeginTrans = False

    Call frmDeptManage.FillTree
    Save部门 = True
    Exit Function
ErrHandle:
    If BeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ChangeCode(nod As Node, ByVal strOldCode As String, ByVal strNewCode As String)
'功能:改变下级的编码内容
    Dim nodChild As Node
    
    Set nodChild = nod.Child
    Do Until nodChild Is Nothing
        nodChild.Text = strNewCode & Mid(nodChild.Text, Len(strOldCode))
        ChangeCode nodChild, strOldCode, strNewCode
        Set nodChild = nodChild.Next
    Loop
End Sub

Public Sub 编辑部门(ByVal strPrivs As String, strID As String, ByVal int编辑状态 As Integer, ByVal int编辑模式 As Integer, ByVal str上级编码 As String, Optional str上级ID As String)
'    On Error GoTo errHandle
    
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim str性质 As String
    Dim str服务对象_临床 As String
    Dim str服务对象_病区 As String
    
    mstrPrivs = strPrivs
    mint编辑状态 = int编辑状态
    mStr部门id = strID
    mint编辑模式 = int编辑模式
    mstr上级码 = str上级编码
    mstr上级部门ID = str上级ID
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockReadOnly
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    mint性质 = 0
    
    mintInputMethod = IIF(Val(zlDatabase.GetPara("自由录入编码", glngSys, 1001, "0")) = 1, 1, 0)
    tabMain.Tabs.Clear
    tabMain.Tabs.Add , "_基本信息", "基本信息"
    
    mlng配置中心 = Val(zlDatabase.GetPara("配置中心", glngSys, 0))
        
    Call IniStationNo
    
    Call Cbo.SetListHeight(cmb诊疗科目编码, cmb诊疗科目编码.Height * 16)
    
    mbln药店 = (glngSys \ 100 = 8)
    If mbln药店 = True Then
        '药店系统需要特殊处理
        lbl诊疗科目编码.Visible = False
        cmb诊疗科目编码.Visible = False
        
'        lbl工作性质.Top = lbl诊疗科目编码.Top
'        lvw性质.Top = cmb诊疗科目编码.Top
        lvw性质.Height = fra基本信息.Height + fra说明.Height
        
        lvw性质.ColumnHeaders(2).Text = "说明"
    Else
        lbl诊疗科目编码.Top = lvw性质.Top + lvw性质.Height + 40
        cmb诊疗科目编码.Top = lbl诊疗科目编码.Top + lbl诊疗科目编码.Height + 50
        lvw性质.ToolTipText = "当性质选中时双击或按“C”键可改变服务对象"
    End If
    mstrID = strID
    '部门负责人选择
    gstrSQL = "select a.id,'【'||a.编号||'】'||a.姓名 姓名 from 人员表 a, 部门人员 b where a.id=b.人员id and b.部门id=[1] order by 姓名"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrID)
    With cbo负责人
        .Clear
        .AddItem ""
        .ItemData(0) = -1
        For i = 1 To rsTemp.RecordCount
            .AddItem NVL(rsTemp!姓名)
            .ItemData(i) = rsTemp!ID
            rsTemp.MoveNext
        Next
    End With
    
    cbo环境类别.ListIndex = -1
    If strID <> "" Then
        gstrSQL = "select A.编码,A.名称,A.简码,A.位置,A.环境类别,B.编码 as 上级编码,B.名称 as 上级名称,B.ID as 上级ID,A.站点 " _
                & ",A.部门负责人,A.顺序,A.别名 " _
                & "from 部门表 A,部门表 B  where A.上级ID=B.ID(+) and A.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
        mstr上级部门ID = IIF(IsNull(rsTemp("上级ID")), "", rsTemp("上级ID"))
        mstr上级编码 = IIF(IsNull(rsTemp("上级编码")), "", rsTemp("上级编码"))
        
        
        If Mid(mstr上级编码, 1, 1) = "-" Then
            '恢复已删除部门
            Me.Tag = "恢复"
            mstr上级编码 = ""
            mstr上级部门ID = ""
            txtEdit(4).Text = ""
        Else
            txtEdit(4).Text = IIF(IsNull(rsTemp("上级名称")), "无", rsTemp("上级名称"))
            If mintInputMethod = 0 Then '按上下级
                txtTemp.Text = mstr上级编码
                '取得上级编码，本级编码长度等值
                txtTemp.MaxLength = GetLocalCodeLength(mstr上级部门ID, "部门表")
                'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
                txtEdit(1).Text = Mid(rsTemp("编码"), Len(txtTemp.Text) + 1)
                mstr编码 = rsTemp("编码")
                '求出包括子节点在内的最长编码
                mint编码 = GetDownCodeLength(mstrID, "部门表")
                '10 - (mint编码 - Len(mstr编码))这个公式的意思是要为它的孩子的编码留有余地
                txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10 - (mint编码 - Len(mstr编码)), txtTemp.MaxLength) - Len(mstr上级编码)
            Else
                txtTemp.Text = ""
                txtTemp.MaxLength = 0
                txtEdit(1).Text = rsTemp("编码")
                mstr编码 = rsTemp("编码")
                txtEdit(1).MaxLength = 10
            End If
        End If
        
        txtEdit(2).Text = rsTemp("名称")
        txtEdit(3).Text = IIF(IsNull(rsTemp("简码")), "", rsTemp("简码"))
        txtEdit(5).Text = IIF(IsNull(rsTemp("位置")), "", rsTemp("位置"))
        txtEdit(0).Text = IIF(IsNull(rsTemp("顺序")), "", rsTemp("顺序"))
        txtEdit(6).Text = IIF(IsNull(rsTemp("别名")), "", rsTemp("别名"))
        
        With cbo环境类别
            For i = 0 To .ListCount - 1
                If .List(i) = NVL(rsTemp!环境类别) Then
                    .ListIndex = i: Exit For
                End If
            Next
            If Trim(NVL(rsTemp!环境类别)) <> "" And .ListIndex < 0 Then
                .AddItem NVL(rsTemp!环境类别): .ListIndex = .NewIndex
            End If
        End With
        With cbo负责人
            For i = 0 To .ListCount - 1
                If .ItemData(i) = NVL(rsTemp!部门负责人) Then
                    .ListIndex = i: Exit For
                End If
            Next
        End With
        
        SetStationNo (IIF(IsNull(rsTemp("站点")), "", rsTemp("站点")))
    Else
        If mintInputMethod = 0 Then '按上下级
            If str上级ID = "oot" Then
                mstr上级部门ID = ""
                mstr上级编码 = ""
                txtTemp.Text = ""
                txtEdit(4).Text = "无"
                '取得上级编码，本级编码长度等值
                txtTemp.MaxLength = GetLocalCodeLength("", "部门表")
            Else
                gstrSQL = "select 编码 as 上级编码,名称 as 上级名称,ID as 上级ID from 部门表 where ID=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str上级ID))
                            
                mstr上级部门ID = IIF(IsNull(rsTemp("上级ID")), "", rsTemp("上级ID"))
                mstr上级编码 = IIF(IsNull(rsTemp("上级编码")), "", rsTemp("上级编码"))
                txtEdit(4).Text = IIF(IsNull(rsTemp("上级名称")), "无", rsTemp("上级名称"))
                txtTemp.Text = mstr上级编码
                '判断编码是否满了
                If Len(mstr上级编码) = mlng编码长度 Then
                    MsgBox "不能再增加子部门了，编码长度已经用尽。", vbExclamation, gstrSysName
                    Exit Sub
                End If
                '取得上级编码，本级编码长度等值
                txtTemp.MaxLength = GetLocalCodeLength(mstr上级部门ID, "部门表")
                
                'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
            End If
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(mstr上级编码)
            txtEdit(1).Text = GetMaxLocalCode(mstr上级部门ID, "部门表")
            mstr编码 = mstr上级编码 & txtEdit(1).Text
        Else
            '自由录入编码
            If str上级ID = "oot" Then
                mstr上级部门ID = ""
                mstr上级编码 = ""
                txtTemp.Text = ""
                txtEdit(4).Text = "无"
                '取得上级编码，本级编码长度等值
                txtTemp.MaxLength = 0
            Else
                gstrSQL = "select 编码 as 上级编码,名称 as 上级名称,ID as 上级ID from 部门表 where ID=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str上级ID))
                            
                mstr上级部门ID = IIF(IsNull(rsTemp("上级ID")), "", rsTemp("上级ID"))
                mstr上级编码 = IIF(IsNull(rsTemp("上级编码")), "", rsTemp("上级编码"))
                txtEdit(4).Text = IIF(IsNull(rsTemp("上级名称")), "无", rsTemp("上级名称"))
                txtTemp.Text = ""
                '取得上级编码，本级编码长度等值
                txtTemp.MaxLength = 0
                
                'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
            End If
            txtEdit(1).MaxLength = 10
            txtEdit(1).Text = ""
            mstr编码 = mstr上级编码 & txtEdit(1).Text
        End If
    End If
    '显示部门性质
    If rsTemp.State = 1 Then rsTemp.Close
    
    If strID = "" Then
        gstrSQL = "select 名称,服务病人 as 缺省服务,说明,null as 工作性质,null as 服务对象 from 部门性质分类 order by decode(工作性质,null,1,0) ,编码"
    Else
        gstrSQL = "select A.名称,A.服务病人 as 缺省服务,A.说明,B.工作性质,B.服务对象 from 部门性质分类 A,部门性质说明 B where A.名称=B.工作性质(+) and b.部门ID(+)=[1] order by decode(工作性质,null,1,0),A.编码"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    lbl诊疗科目编码.Enabled = False
    cmb诊疗科目编码.Enabled = False
        
    If rsTemp.EOF Then
        mblnChange = False
        frmDeptSet.Show vbModal
        Exit Sub
    End If
        
    Dim lst As ListItem
    Do Until rsTemp.EOF
        If InStr(1, mstrPrivs, ";" & "设置卫材虚拟库房" & ";") = 0 And rsTemp("名称") = "虚拟库房" Then
            rsTemp.MoveNext
        Else
            Select Case IIF(IsNull(rsTemp("工作性质")), rsTemp("缺省服务"), rsTemp("服务对象"))
                 Case 1
                    strTemp = "门诊病人"
                 Case 2
                    strTemp = "住院病人"
                 Case 3
                    strTemp = "门诊和住院病人"
                 Case Else
                    strTemp = "不服务于病人"
            End Select
            Set lst = lvw性质.ListItems.Add(, rsTemp("名称"), rsTemp("名称"))
            If mbln药店 = True Then
                lst.SubItems(1) = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
            Else
                lst.SubItems(1) = strTemp
            End If
            lst.ListSubItems(1).Tag = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
            lst.Tag = rsTemp("缺省服务")
            If Not IsNull(rsTemp("工作性质")) Then
                lst.SubItems(2) = 1
                lst.Checked = True
                If lst.Text = "临床" Then
                    lbl诊疗科目编码.Enabled = True
                    cmb诊疗科目编码.Enabled = True
                End If
            End If
            
            str性质 = IIF(str性质 = "", "", str性质 & ",") & rsTemp("工作性质")
            If rsTemp("工作性质") = "临床" Then
                str服务对象_临床 = IIF(str服务对象_临床 = "", "", str服务对象_临床 & ",") & strTemp
            End If
            
            If rsTemp("工作性质") = "护理" Then
                str服务对象_病区 = IIF(str服务对象_病区 = "", "", str服务对象_病区 & ",") & strTemp
            End If
        
            rsTemp.MoveNext
        End If
    Loop
    
    '记录初始的部门性质和服务对象
    mint性质 = 0
    mint服务对象_临床 = 0
    mint服务对象_病区 = 0
    
    If InStr(1, str性质, "临床") > 0 And InStr(1, str性质, "护理") > 0 Then
        mint性质 = 3
    ElseIf InStr(1, str性质, "护理") > 0 Then
        mint性质 = 2
    ElseIf InStr(1, str性质, "临床") > 0 Then
        mint性质 = 1
    End If
    
    If InStr(1, str服务对象_临床, "门诊和住院病人") > 0 Then
        mint服务对象_临床 = 3
    ElseIf InStr(1, str服务对象_临床, "住院病人") > 0 Then
        mint服务对象_临床 = 2
    ElseIf InStr(1, str服务对象_临床, "门诊病人") > 0 Then
        mint服务对象_临床 = 1
    End If
    
    If InStr(1, str服务对象_病区, "门诊和住院病人") > 0 Then
        mint服务对象_病区 = 3
    ElseIf InStr(1, str服务对象_病区, "住院病人") > 0 Then
        mint服务对象_病区 = 2
    ElseIf InStr(1, str服务对象_病区, "门诊病人") > 0 Then
        mint服务对象_病区 = 1
    End If
    
    Lvw科室.Tag = CStr(mint服务对象_临床) & CStr(mint服务对象_病区)
    
    lvw性质.ListItems(1).Selected = True
    lvw性质_ItemClick lvw性质.ListItems(1)
    '显示部门临床性质
    If rsTemp.State = 1 Then rsTemp.Close
    
    If strID = "" Or cmb诊疗科目编码.Enabled = False Then
        '不具有临床性质
        gstrSQL = "select 编码,名称,null as 工作性质 from 临床性质 order by 序号"
    Else
        gstrSQL = "select A.编码,A.名称,B.工作性质 from 临床性质 A,临床部门 B " & _
            "where A.编码=B.工作性质(+) and b.部门ID(+)=[1] order by A.序号"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    cmb诊疗科目编码.Clear
    Do Until rsTemp.EOF
        cmb诊疗科目编码.AddItem Format(rsTemp("编码"), "!@@@@@") & rsTemp("名称")
        If Not IsNull(rsTemp("工作性质")) Then
            cmb诊疗科目编码.ListIndex = cmb诊疗科目编码.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    
    '记录原来的工作性质
    mint原工作性质 = 0
    For i = 1 To lvw性质.ListItems.Count
        If lvw性质.ListItems(i).Checked = True Then
            If mint原工作性质 <> 1 Then
                If InStr(lvw性质.ListItems(i), "药房") > 0 Or lvw性质.ListItems(i) = "制剂室" Then
                    mint原工作性质 = 1
                ElseIf InStr(lvw性质.ListItems(i), "药库") > 0 Then
                    mint原工作性质 = 2
                End If
            End If
        End If
    Next
    
    '科室病区设置
    Call Ini病区科室对应(mstrID, 1)
    
    '完成初始化
    If rsTemp.State = 1 Then rsTemp.Close
    
    mblnChange = False
    frmDeptSet.Show vbModal
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd上级_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim str编码 As String
    Dim int编码  As Integer
    
    If mstrID <> "" Then
        strSQL = "select id,上级id,名称,编码,简码 from 部门表 where 撤档时间=to_date('3000-01-01','YYYY-MM-DD') and id<>" & mstrID & " start with 上级id is null connect by prior id =上级id And 上级id<>" & mstrID
    Else
        strSQL = "select id,上级id,名称,编码,简码 from 部门表 where 撤档时间=to_date('3000-01-01','YYYY-MM-DD') start with 上级id is null connect by prior id =上级id "
    End If
    strID = mstr上级部门ID
    str名称 = txtEdit(4).Text
    str编码 = txtTemp.Text
    blnRe = frmTreeSel.ShowTree(strSQL, strID, str名称, str编码, mstrID, "部门表", "所有部门", , mstr编码, 0, 0, 0, False)
    '成功返回
    If blnRe Then       '新的本级的宽度
        int编码 = GetLocalCodeLength(strID, "部门表")
        '只有修改才有必要审核
        If mstrID <> "" Then
            If mint编码 - Len(mstr编码) + IIF(int编码 = 0, Len(str编码) + 1, int编码) > 10 Then
                MsgBox "这个上级不合适，因为它的编码太长了。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        mstr上级部门ID = strID
        txtEdit(4).Text = str名称
        txtTemp.MaxLength = int编码
        txtTemp.Text = str编码
        If mstrID <> "" Then
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10 - (mint编码 - Len(mstr编码)), txtTemp.MaxLength) - Len(str编码)
        Else
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(str编码)
        End If
        txtEdit(1).Text = GetMaxLocalCode(mstr上级部门ID, "部门表")
    End If
    mblnChange = True
    '检查该级下部门顺序
    If CheckOrder = True Then
        txtEdit(0).SetFocus
    End If
End Sub

Private Sub Form_Activate()
    txtEdit(2).SetFocus
    lbl说明.Move 130, 260, fra说明.Width - 160, fra说明.Height - 400
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    '刘兴宏:2008/03/18加入
    On Error GoTo ErrHandle
    gstrSQL = "Select 编码,名称 From 部门环境类别 order by 编码"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With cbo环境类别
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!名称)
            rsTemp.MoveNext
        Loop
    End With
    If mint编辑状态 = 1 Then
        cbo负责人.Enabled = False
    Else
        cbo负责人.Enabled = True
    End If
    
    lblFind.Visible = False
    txtFind.Visible = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Lvw科室_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = False And Val(Lvw科室.ListItems(Item.Index).Tag) = 1 Then
        If Check床位状况(Val(mstrID), Val(Mid(Item.Key, 2)), 1) Then
            MsgBox "该部门存在床位记录，不能取消对应关系！", vbInformation, gstrSysName
            Item.Checked = True
            Exit Sub
        End If
    End If
End Sub


Private Sub lvw性质_DblClick()
    If mblnItem = False Then Exit Sub
    Call ChangeServer
    Call Set病区科室对应
    mblnItem = False
End Sub

Private Sub ChangeServer()
    If lvw性质.SelectedItem Is Nothing Then Exit Sub
    If mbln药店 = True Then Exit Sub
    
    With lvw性质.SelectedItem
        If .Checked = False Then Exit Sub
        Select Case .SubItems(1)
             Case "门诊病人"
                .SubItems(1) = "住院病人"
             Case "住院病人"
                .SubItems(1) = "门诊和住院病人"
             Case "门诊和住院病人"
                .SubItems(1) = "不服务于病人"
             Case Else
                If .Tag <> 0 Then .SubItems(1) = "门诊病人"
        End Select
        mblnChange = True
    End With
End Sub

Private Sub lvw性质_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    Dim bln虚拟库房 As Boolean
    
    mblnChange = True
    
    If Item.Text = "临床" Or Item.Text = "护理" Then
        If Item.Checked = False And Val(Item.SubItems(2)) = 1 Then
            If Check床位状况(Val(mstrID), 0, 0) Then
                MsgBox "该部门存在床位记录，不能取消该性质！", vbInformation, gstrSysName
                Item.Checked = True
                Exit Sub
            End If
        End If
    End If
    
    If Item.Text = "临床" Then
        If Item.Checked = False Then
            lbl诊疗科目编码.Enabled = False
            cmb诊疗科目编码.Enabled = False
            cmb诊疗科目编码.ListIndex = -1
        Else
            lbl诊疗科目编码.Enabled = True
            cmb诊疗科目编码.Enabled = True
        End If
    End If
    
    If mlng配置中心 > 0 And mlng配置中心 = Val(mstrID) Then
        If Item.Text = "配制中心" Or Item.Text = "西药房" Then
            If Item.Checked = False Then
                MsgBox "该部门已被启用为医院的输液配置中心，不能改变属性，请在基础参数设置中处理！", vbInformation, gstrSysName
                Item.Checked = True
                Exit Sub
            End If
        End If
    End If
    
    '如果选择了“虚拟库房”属性，则不能选择其他属性
    With lvw性质
        For i = 1 To .ListItems.Count
            If .ListItems(i).Text = "虚拟库房" And .ListItems(i).Checked = True Then
                bln虚拟库房 = True
                Exit For
            End If
        Next
        
        If bln虚拟库房 = True Then
            For i = 1 To .ListItems.Count
                If .ListItems(i).Text <> "虚拟库房" And .ListItems(i).Checked = True Then
                    .ListItems(i).Checked = False
                End If
            Next
        End If
    End With
    
    Call Set病区科室对应
End Sub

Private Sub lvw性质_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lbl说明.Caption = Item.ListSubItems(1).Tag
    mblnItem = True
End Sub

Private Sub lvw性质_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("C") Or KeyAscii = Asc("c") Then Call ChangeServer
End Sub

Private Sub lvw性质_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lvw性质.SelectedItem
        If .Tag = 0 Then
            lvw性质.ToolTipText = "该部门不服务于病人，工作性质不能修改！"
        Else
            lvw性质.ToolTipText = "当性质选中时双击或按“C”键可改变服务对象！"
        End If
    End With
End Sub

Private Sub lvw性质_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 2 Then
        If lvw性质.SelectedItem Is Nothing Then Exit Sub
        If lvw性质.SelectedItem.Checked = False Then Exit Sub
        If lvw性质.SelectedItem.Tag = 0 Then Exit Sub
        
        For i = 0 To 4
            If mnuPatient(i).Caption <> "-" Then
                If lvw性质.SelectedItem.SubItems(1) = Left(mnuPatient(i).Caption, InStr(mnuPatient(i).Caption, "(") - 1) Then
                    mnuPatient(i).Checked = True
                Else
                    mnuPatient(i).Checked = False
                End If
            End If
        Next
        PopupMenu mnuShort
    End If
End Sub

Private Sub mnuPatient_Click(Index As Integer)
    lvw性质.SelectedItem.SubItems(1) = Left(mnuPatient(Index).Caption, InStr(mnuPatient(Index).Caption, "(") - 1)
    mblnChange = True
End Sub



Private Sub tabMain_Click()
    Dim i As Integer
    
    For i = fraMain.LBound To fraMain.UBound
        fraMain(i).Visible = False
    Next
    
    i = tabMain.SelectedItem.Index - 1
    fraMain(i).Visible = True
    fraMain(i).ZOrder 0
    If tabMain.SelectedItem.Index = 1 Then
        lblFind.Visible = False
        txtFind.Visible = False
    Else
        lblFind.Visible = True
        txtFind.Visible = True
    End If
End Sub
Private Sub ShowTab(ByVal intTab As Integer)
    tabMain.Tabs(intTab).Selected = True
    tabMain_Click
End Sub
Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(txtEdit(2).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = 2 Or Index = 5 Then
        OS.OpenIme True
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    ElseIf Index = 2 Or Index = 3 Then
        If LenB(StrConv(txtEdit(2).Text & Chr(KeyAscii), vbFromUnicode)) > 100 And (KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack) Then
            KeyAscii = 0
        End If
    ElseIf Index = 5 Then
        If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
    ElseIf Index = 0 Then
        If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Or Index = 5 Then
        OS.OpenIme False
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        If CheckOrder = True Then
            Cancel = True
        End If
    End If
End Sub

Private Function CheckOrder() As Boolean
    Dim rsTemp As Recordset
    Dim intOrder As Integer
    
    On Error GoTo ErrHandle
    
    If Val(txtEdit(0).Text) = 0 Then Exit Function
    CheckOrder = False
    
    If mstrID = "" Then '新增
        If mstr上级部门ID = "" Then
            gstrSQL = "Select 1 From 部门表 Where 顺序 = [1] And 上级id is Null"
        Else
            gstrSQL = "Select 1 From 部门表 Where 顺序 = [1] And 上级id =[2]"
        End If
    Else
        If mstr上级部门ID = "" Then
            gstrSQL = "Select 1 From 部门表 Where 顺序 = [1] And 上级id is Null And id <> [3]"
        Else
            gstrSQL = "Select 1 From 部门表 Where 顺序 = [1] And 上级id =[2] And id <> [3]"
        End If
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询部门顺序", Val(txtEdit(0).Text), Val(mstr上级部门ID), Val(mstrID))
    
    If Not rsTemp.EOF Then
        If mstr上级部门ID = "" Then
            gstrSQL = "Select Max(Nvl(顺序,0)) As 最大顺序 From 部门表 Where 上级id Is Null"
        Else
            gstrSQL = "Select Max(Nvl(顺序,0)) As 最大顺序 From 部门表 Where 上级id = [1]"
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询最大顺序", Val(mstr上级部门ID))
        
        MsgBox "该级下顺序为‘" & Val(txtEdit(0).Text) & "’的部门已存在，且最大顺序为‘" & rsTemp!最大顺序 & "’" & "，请重新输入部门顺序！", vbInformation, gstrSysName
        CheckOrder = True
    End If
    
    rsTemp.Close
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    OS.OpenIme True
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte
    Dim strTemp As String
    Dim litem As ListItem
    Dim lsItem As ListItem
    Dim i As Integer
    
    If KeyAscii = 13 Then
        gstrSQL = " Select Distinct 编码,名称,ID From 部门表 " & _
         " Where ID in (Select 部门ID From 部门性质说明 Where " & mstr性质 & ")" & _
         " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) and (编码 like [1] or 名称 like [2] or 简码 like [3]) " & _
         " Order by 名称 "
         
         strTemp = UCase(Trim(txtFind.Text))
         vRect = zlControl.GetControlRect(txtFind.hwnd)
         
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "科室病区对应", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtFind.Height, True, False, True, strTemp & gstrLike, strTemp & gstrLike, strTemp & gstrLike)
                
        If Not rsTemp Is Nothing Then
            cmb诊疗科目编码.Text = rsTemp("编码") & "-" & rsTemp("名称")
            For Each litem In Lvw科室.ListItems
                If litem.Text = cmb诊疗科目编码.Text Then
                    Set lsItem = litem
                Else
                    litem.Selected = False
                End If
            Next
            If Not lsItem Is Nothing Then
                lsItem.Selected = True
                txtFind.SetFocus
                txtFind.SelStart = 0
                txtFind.SelLength = Len(txtFind.Text)
                Exit Sub
            End If
        End If
        
        MsgBox "没有找到你想要得科室，请重新输入查询条件！", vbInformation, gstrSysName
        txtFind.Text = ""
        txtFind.SetFocus
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(1).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(1).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

'CheckSameDept(txtEdit(2).Text, lngTmp) Then
Private Function CheckSameDept(ByVal strDept As String, ByRef lngDeptID As Long) As Boolean
'----------------------------------------------
'功能：检查已删除部门中，是否有相同的部门名称
'参数：strDept：当前录入的部门名称；
'      lngDeptID：写入找到已删除部门相同的部门ID
'返值：True：找到相同；False：没有找到。
'----------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID From 部门表 " & _
              "Where Substr(编码, 1, 1) = '-' And 撤档时间 < To_Date('3000-1-1', 'yyyy-mm-dd') And 名称 <> '已删除部门' and 名称 = [1] " & _
              "Order by 撤档时间 desc "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查已删除部门与当前录入部门是否相同", strDept)
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp!ID) Then
            lngDeptID = rsTemp!ID
            CheckSameDept = True
        End If
    End If
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check重复部门(ByVal str上级ID As String, ByVal str药品名称 As String) As Boolean
    '功能：用来检查是否已具有该部门
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select 名称 from 部门表 where 上级id=[1] and 名称=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询是否有重复部门", str上级ID, str药品名称)
    If rsTemp.EOF Then
        Check重复部门 = False
    Else
        Check重复部门 = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


