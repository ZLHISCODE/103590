VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBatchChangeNumNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "门诊批量换号管理"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10740
   Icon            =   "frmBatchChangeNumNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10740
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsf挂号信息 
      Height          =   2805
      Left            =   60
      TabIndex        =   5
      Top             =   2355
      Width           =   10635
      _cx             =   18759
      _cy             =   4948
      Appearance      =   1
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fra操作 
      Height          =   2175
      Left            =   9090
      TabIndex        =   2
      Top             =   150
      Width           =   1620
      Begin VB.CommandButton btn取消 
         Caption         =   "取消"
         Height          =   350
         Left            =   255
         TabIndex        =   4
         Top             =   1065
         Width           =   1100
      End
      Begin VB.CommandButton btn确定 
         Caption         =   "确定"
         Height          =   350
         Left            =   255
         TabIndex        =   3
         Top             =   315
         Width           =   1100
      End
   End
   Begin VB.Frame fra新挂号安排 
      Caption         =   "新挂号安排信息"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2235
      Left            =   4605
      TabIndex        =   1
      Top             =   105
      Width           =   4410
      Begin VB.CommandButton cmd新排班 
         Caption         =   "P"
         Height          =   375
         Left            =   1635
         TabIndex        =   43
         Top             =   315
         Width           =   405
      End
      Begin VB.TextBox txtNewFilter 
         Height          =   375
         Left            =   675
         TabIndex        =   41
         Top             =   315
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   0
         TabIndex        =   40
         Top             =   750
         Width           =   4350
      End
      Begin VB.TextBox txt新号别 
         Height          =   300
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "号别"
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txt新科室 
         Height          =   300
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "科室"
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txt新医生 
         Height          =   300
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "医生"
         Top             =   1350
         Width           =   1380
      End
      Begin VB.TextBox txt新号类 
         Height          =   300
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "号类"
         Top             =   1335
         Width           =   1380
      End
      Begin VB.TextBox txt新限号 
         Height          =   300
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "限号"
         Top             =   1740
         Width           =   1380
      End
      Begin VB.TextBox txt新限约 
         Height          =   300
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "限约"
         Top             =   1725
         Width           =   1365
      End
      Begin VB.TextBox txt新安排ID 
         Height          =   285
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "安排ID"
         Top             =   855
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txt新项目 
         Height          =   285
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "项目"
         Top             =   885
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label12 
         Caption         =   "条件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   42
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label11 
         Caption         =   "号别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   39
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label10 
         Caption         =   "科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2205
         TabIndex        =   38
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label9 
         Caption         =   "医生"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   37
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "号类"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2220
         TabIndex        =   36
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label7 
         Caption         =   "限号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   35
         Top             =   1785
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "限约"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2235
         TabIndex        =   34
         Top             =   1770
         Width           =   480
      End
   End
   Begin VB.Frame fra原挂号安排 
      Caption         =   "原挂号安排信息"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2235
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   4410
      Begin VB.TextBox txt项目 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "项目"
         Top             =   660
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txt安排ID 
         Height          =   285
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "安排ID"
         Top             =   675
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmd原排班 
         Caption         =   "P"
         Height          =   375
         Left            =   3885
         TabIndex        =   23
         Top             =   285
         Width           =   405
      End
      Begin VB.TextBox txt限约 
         Height          =   300
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "限约"
         Top             =   1755
         Width           =   1380
      End
      Begin VB.TextBox txt限号 
         Height          =   300
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "限号"
         Top             =   1755
         Width           =   1380
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   30
         TabIndex        =   18
         Top             =   750
         Width           =   4350
      End
      Begin VB.TextBox txt号类 
         Height          =   300
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "号类"
         Top             =   1350
         Width           =   1380
      End
      Begin VB.TextBox txt医生 
         Height          =   300
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "医生"
         Top             =   1350
         Width           =   1380
      End
      Begin VB.TextBox txt科室 
         Height          =   300
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "科室"
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txt号别 
         Height          =   300
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "号别"
         Top             =   960
         Width           =   1380
      End
      Begin VB.TextBox txtFilter 
         Height          =   375
         Left            =   2910
         TabIndex        =   8
         Top             =   315
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker dtp预约日期 
         Height          =   345
         Left            =   990
         TabIndex        =   6
         Top             =   315
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         _Version        =   393216
         Format          =   103940097
         CurrentDate     =   41128
      End
      Begin VB.Label Label5 
         Caption         =   "限约"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2190
         TabIndex        =   21
         Top             =   1770
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "限号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   19
         Top             =   1785
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "号类"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   16
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "医生"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   14
         Top             =   1410
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2145
         TabIndex        =   12
         Top             =   990
         Width           =   480
      End
      Begin VB.Label lbl号别 
         Caption         =   "号别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   10
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label lbl过滤 
         Caption         =   "条件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2430
         TabIndex        =   9
         Top             =   390
         Width           =   915
      End
      Begin VB.Label lbl预约日期 
         Caption         =   "预约日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   7
         Top             =   390
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmBatchChangeNumNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
'常量定义
Private Const ID_PANE_源号别信息 = 1
Private Const ID_PANE_目标号别信息 = 2
'操作类型
Private Enum mEnum_查询类型
     需换号排班 = 1
     可换号排班 = 2
End Enum
'A.标识,A.号类,A.号码,A.科室,A.项目,A.医生姓名,A.限号数,A.限约数,A.已约数,A.已接收数,A.时间段,A.病案,A.分诊,A.序号控制

Private mrs挂号安排 As Recordset
Private mrs可换号排班 As Recordset
Private mrs挂号信息 As Recordset

Private Sub btn取消_Click()
    Unload Me
End Sub

Private Sub btn确定_Click()
    If CheckValid = False Then Exit Sub
    '保存数据
    If SaveData = True Then
        MsgBox "换号成功！", vbInformation, gstrSysName
        vsf挂号信息.Clear 1
    Else
        MsgBox "换号失败！", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmd原排班_Click()
    Call Show挂号安排信息(mEnum_查询类型.需换号排班)
    If txt安排ID.Text <> "" Then
        Call Show挂号安排信息(mEnum_查询类型.可换号排班)
    Else
        Call Clear新安排信息
    End If
    
    Call Show挂号信息
End Sub

Private Sub cmd新排班_Click()
        Call Show挂号安排信息(mEnum_查询类型.可换号排班)
End Sub

Private Sub DTP预约日期_Change()
     'Call Show挂号安排信息(mEnum_查询类型.需换号排班)
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    Call Clear原安排信息
    Call Clear新安排信息
    Call Init预约时间
    Call SetHeader
    Call DTP预约日期_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrs挂号安排 = Nothing
    Set mrs挂号信息 = Nothing
    SaveWinState Me, App.ProductName
End Sub

Private Sub Show挂号安排信息(enum查询类型 As mEnum_查询类型)
    '----------------------------------------------------------------------------------------------
    '功能:显示挂号安排信息
    '返回:
    '编制:王吉
    '日期:2012/7/30
    '----------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strSQLBody As String
    Dim strSQL当前时段 As String
    Dim strSQL其他可用时段 As String
    Dim strSQL可换号时段 As String
    Dim strSQL可换号排班 As String
    Dim strSQLWhere As String
    Dim dat时间 As Date
    Dim str时间 As String
    
    On Error GoTo errHanl
    
    dat时间 = dtp预约日期.Value
    str时间 = "to_date('" & dat时间 & "','yyyy-MM-dd HH24:mi:ss')"
    
    strSQLBody = "Select a.Id, b.号类, b.号码, c.名称 As 科室, b.科室id, d.名称 As 项目, Nvl(a.替诊医生姓名, a.医生姓名) As 医生姓名, a.医生id, a.限号数, a.限约数, a.已挂数, a.已约数," & vbNewLine & _
                "       a.其中已接收 As 已接收数, a.上班时段 As 时间段, Decode(b.是否建病案, 0, '', '√') As 病案," & vbNewLine & _
                "       Decode(A.分诊方式, 1, '指定', 2, '动态', 3, '平均', '') As 分诊, Decode(a.是否序号控制, 0, '', '√') As 序号控制" & vbNewLine & _
                "From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D" & vbNewLine & _
                "Where a.号源id = b.Id And a.项目id = d.Id And b.科室id = c.Id And (c.站点 Is Null Or c.站点 = [1]) And a.出诊日期 = [2]"

    
    '模糊查找
    '查找范围:项目,医生姓名,号类,号码
    If Trim(txtFilter.Text) <> "" And enum查询类型 = 需换号排班 Then
        strSQLWhere = " Where " & _
        "     项目 Like '%" & Trim(txtFilter.Text) & "%'" & _
        "  Or 医生姓名 Like '%" & Trim(txtFilter.Text) & "%'" & _
        "  Or 号类 Like '%" & Trim(txtFilter.Text) & "%'" & _
        "  Or 号码 Like '%" & Trim(txtFilter.Text) & "%'"
    End If
    
    If enum查询类型 = 可换号排班 Then
        strSQLBody = strSQLBody & " And (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) "
    End If
    
    strSQL = "" & _
    "   Select ID As ID,号类,号码,科室,科室ID,项目,医生姓名,医生ID,限号数,限约数,已挂数,已约数,已接收数,时间段,病案,分诊,序号控制 From(" & strSQLBody & ") " & strSQLWhere & " Order By 号码"
    
    If enum查询类型 = 需换号排班 Then
        Set mrs挂号安排 = zlDatabase.ShowSQLSelect(frmBatchChangeNumNew, strSQL, 0, "原挂号安排信息", False, "", "", False, False, False, txtFilter.Left, txtFilter.Top + txtFilter.Height, 1000, True, False, False, gstrNodeNo, dat时间)
        If mrs挂号安排 Is Nothing Then
           Call Clear原安排信息
        Else
           Call Set原安排信息
        End If
    End If
    
    
    If enum查询类型 = 可换号排班 Then
        If mrs挂号安排 Is Nothing Then Exit Sub
        '获取相同时段的排班信息SQL
        strSQL可换号时段 = "Select a.Id, b.号类, b.号码, c.名称 As 科室, b.科室id, d.名称 As 项目, Nvl(a.替诊医生姓名, a.医生姓名) As 医生, a.医生id, a.限号数, a.限约数, a.已挂数, a.已约数," & vbNewLine & _
                        "       a.其中已接收 As 已接收数, a.上班时段 As 时间段, Decode(b.是否建病案, 0, '', '√') As 病案," & vbNewLine & _
                        "       Decode(a.分诊方式, 1, '指定', 2, '动态', 3, '平均', '') As 分诊, Decode(a.是否序号控制, 0, '', '√') As 序号控制" & vbNewLine & _
                        "From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D" & vbNewLine & _
                        "Where a.号源id = b.Id And a.项目id = d.Id And b.科室id = c.Id And (c.站点 Is Null Or c.站点 = [1]) And a.出诊日期 = [2] And" & vbNewLine & _
                        "      (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间))"

        
        '其他可换号限定条件
        '限号数相同,限约数相同,科室相同,号类相同,收费项目相同,已约数与已接收数为0：表示该排班没有被预约过
        strSQLWhere = "" & _
        "   And  Nvl(A.限号数,'0') >= '" & Val(IIf(Not mrs挂号安排 Is Nothing, Nvl(mrs挂号安排!已挂数, "0"), "0")) & "'" & _
        "   And  Nvl(A.限约数,'0') >= '" & Val(IIf(Not mrs挂号安排 Is Nothing, Nvl(mrs挂号安排!已约数, "0"), "0")) & "'" & _
        "   And  Nvl(A.科室,'无') = '" & IIf(Trim(txt科室.Text) <> "", Trim(txt科室.Text), "无") & "'" & _
        "   And  Nvl(A.号类,'无') = '" & IIf(Trim(txt号类.Text) <> "", Trim(txt号类.Text), "无") & "'" & _
        "   And  Nvl(A.项目,'无') = '" & IIf(Trim(txt项目.Text) <> "", Trim(txt项目.Text), "无") & "'" & _
        "   And  Nvl(A.已约数,0) = 0 And Nvl(A.已接收数,0) = 0 And Nvl(A.已挂数,0) = 0 " & _
        "   " & IIf(Not mrs挂号安排 Is Nothing, "And A.号码 <> '" & mrs挂号安排!号码 & "'", "")
        '模糊查找
        '查找范围:项目,医生姓名,号类,号码
        If Trim(txtNewFilter.Text) <> "" Then
            strSQLWhere = strSQLWhere & " And (" & _
            "     项目 Like '%" & Trim(txtNewFilter.Text) & "%'" & _
            "  Or 医生姓名 Like '%" & Trim(txtNewFilter.Text) & "%'" & _
            "  Or 号类 Like '%" & Trim(txtNewFilter.Text) & "%'" & _
            "  Or 号码 Like '%" & Trim(txtNewFilter.Text) & "%'" & _
            "   )"
        End If

        '获取可换号的排班SQL
        strSQL可换号排班 = "" & _
        "   Select A.ID,A.号类,A.号码,A.科室,A.科室ID,A.项目,A.医生姓名,A.医生ID,A.限号数,A.限约数,A.已约数,A.已接收数,A.时间段,A.病案,A.分诊,A.序号控制 From(" & strSQL & ") A,(" & strSQL可换号时段 & ") B " & _
        "   Where A.ID=B.ID(+)" & _
        "   " & strSQLWhere & _
        "   Order By 号码"
               
        '查询可换号排班信息
        
        Set mrs可换号排班 = zlDatabase.ShowSQLSelect(frmBatchChangeNumNew, strSQL可换号排班, 0, "新挂号安排信息", False, "", "", False, False, False, txtFilter.Left, txtFilter.Top + txtFilter.Height, 1000, True, False, False, gstrNodeNo, dat时间)
        
        If mrs可换号排班 Is Nothing Then
            Call Clear新安排信息
        Else
            If mrs可换号排班.RecordCount = 0 Then
                Call Clear新安排信息
            Else
                Call Set新安排信息(mrs可换号排班)
            End If
        End If
    End If
    Exit Sub
errHanl:
    MsgBox Err.Description
End Sub

Private Sub Show挂号信息()
    '----------------------------------------------------------------------------------------------
    '功能:显示挂号信息
    '返回:
    '编制:王吉
    '日期:2012/7/30
    '----------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim str时间 As String
    
    If Not mrs挂号安排 Is Nothing And Not mrs可换号排班 Is Nothing Then
        If Val(Nvl(mrs挂号安排!已挂数)) <> 0 Then
            If MsgBox("选择换号的记录包含已经付款的记录,是否继续进行批量换号？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Clear原安排信息
                Clear新安排信息
                vsf挂号信息.Clear 1
                Exit Sub
            End If
        End If
    End If
    
    str时间 = "to_date('" & dtp预约日期.Value & "','yyyy-MM-dd HH24:mi:ss')"
    
    '问题号:56164
    strSQL = "" & _
    "      Select A.No,A.号序 As 序号,A.发生时间 As 预约时间,A.姓名,A.性别,A.年龄,B.身份证号,B.联系人电话,Decode(A.记录性质,1,'已接收',2,'已预约','已预留') As 状态" & _
    "      From 病人挂号记录 A,病人信息 B " & _
    "      Where A.记录状态=1 " & _
    "      And A.病人ID=B.病人ID(+) " & _
    "      And A.出诊记录ID=" & Val(txt号别.Tag) & _
    ""
'    "      Union ALL" & _
'    "      Select A.No,A.号序 As 序号,A.发生时间 As 预约时间,A.姓名,A.性别,A.年龄,B.身份证号,B.联系人电话,Decode(A.记录性质,1,'已接收',2,'已预约','已预留') As 状态" & _
'    "      From 病人挂号记录 A,病人信息 B " & _
'    "      Where A.发生时间 Between Trunc(" & str时间 & ") And Trunc(" & str时间 & ") +1 -1/24/60/60" & _
'    "      And A.记录状态=1 And Nvl(预约,0) = 1" & _
'    "      And A.病人ID=B.病人ID(+) " & _
'    "      And Nvl(号别,'号别')= '" & IsNothing(Trim(txt号别.Text), "号别") & "'"
    Set mrs挂号信息 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    '设置挂号信息列表
    Set vsf挂号信息.DataSource = mrs挂号信息
    '设置列头信息
    SetHeader
End Sub
Private Sub SetHeader()
    Dim str挂号安排 As String
    Dim str挂号信息 As String

    'str挂号安排 = "标识,1,850|号类,1,850|号码,1,850|科室,1,1700|项目,1,1800|医生姓名,1,1000|限约数,1,650|已约数,1,650|已接收数,1,850|时间段,1,850|病案,1,650|分诊,1,650|序号控制,1,650"
    
    str挂号信息 = "No,4,0|序号,4,650|预约时间,4,2000|姓名,4,1000|性别,4,650|年龄,4,650|身份证号,4,2000|联系人电话,4,2000|状态,4,650"
    
    '设置挂号安排
    SetVSGrid vsf挂号信息, str挂号信息
End Sub


Private Sub SetVSGrid(vsGrid As VSFlexGrid, str列名 As String)
    '----------------------------------------------------------------------------------------------
    '功能:设置FlexGrid列名
    '返回:
    '编制:王吉
    '日期:2012/8/3
    '----------------------------------------------------------------------------------------------
    Dim strArr() As String
    Dim lngCols As Long
    Dim i As Long
    
    strArr = Split(str列名, "|")
    lngCols = UBound(strArr) + 1
    
    'vsGrid.Clear
    'vsGrid.Rows = 2
    vsGrid.Cols = lngCols
    
    With vsGrid
        .ColWidthMin = 0
        .Redraw = False
         For i = 0 To UBound(strArr)
            .ColKey(i) = Split(strArr(i), ",")(0)
            .TextMatrix(0, i) = Split(strArr(i), ",")(0)
            .ColAlignment(i) = Split(strArr(i), ",")(1)
            .ColWidth(i) = Split(strArr(i), ",")(2)
        Next
        .RowHeight(0) = 320
        .ExtendLastCol = True
        .Redraw = True
    End With
End Sub


Private Sub Init预约时间()
    '----------------------------------------------------------------------------------------------
    '功能:初始化预约时间控件
    '返回:
    '编制:王吉
    '日期:2012/8/6
    '----------------------------------------------------------------------------------------------
    Dim Curdate As Date
    
    Curdate = zlDatabase.Currentdate
    
    dtp预约日期.Value = Format(Curdate + 1, "yyyy-MM-dd ")
    dtp预约日期.MinDate = Format(Curdate + 1, "yyyy-MM-dd ")
End Sub

Private Sub Clear原安排信息()
     txt安排ID.Text = ""
     txt号别.Text = ""
     txt号别.Tag = ""
     txt科室.Text = ""
     txt医生.Text = ""
     txt号类.Text = ""
     txt限号.Text = ""
     txt限约.Text = ""
     txt项目.Text = ""
End Sub

Private Sub Set原安排信息()
     txt安排ID.Text = Nvl(mrs挂号安排!ID, "")
     txt号别.Text = Nvl(mrs挂号安排!号码, "")
     txt号别.Tag = Nvl(mrs挂号安排!ID, "")
     txt科室.Text = Nvl(mrs挂号安排!科室, "")
     txt医生.Text = Nvl(mrs挂号安排!医生姓名, "")
     txt号类.Text = Nvl(mrs挂号安排!号类, "")
     txt限号.Text = Nvl(mrs挂号安排!限号数, "")
     txt限约.Text = Nvl(mrs挂号安排!限约数, "")
     txt项目.Text = Nvl(mrs挂号安排!项目, "")
End Sub

Private Sub Clear新安排信息()
     txt新安排ID.Text = ""
     txt新号别.Text = ""
     txt新号别.Tag = ""
     txt新科室.Text = ""
     txt新医生.Text = ""
     txt新号类.Text = ""
     txt新限号.Text = ""
     txt新限约.Text = ""
     txt新项目.Text = ""
End Sub

Private Sub Set新安排信息(rs新安排 As Recordset)
     txt新安排ID.Text = Nvl(rs新安排!ID, "")
     txt新号别.Text = Nvl(rs新安排!号码, "")
     txt新号别.Tag = Nvl(rs新安排!ID, "")
     txt新科室.Text = Nvl(rs新安排!科室, "")
     txt新医生.Text = Nvl(rs新安排!医生姓名, "")
     txt新号类.Text = Nvl(rs新安排!号类, "")
     txt新限号.Text = Nvl(rs新安排!限号数, "")
     txt新限约.Text = Nvl(rs新安排!限约数, "")
     txt新项目.Text = Nvl(rs新安排!项目, "")
End Sub
Public Function Nvl(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
    '-----------------------------------------------------------------------------------
    '功能:取某字段的值
    '参数:rsObj          被检查的字段
    '     varValue       当rsObj为NULL值时的取新值
    '返回:如果不为空值,返回原来的值,如果为空值,则返回指定的varValue值
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        Nvl = varValue
    Else
        Nvl = rsObj
    End If
End Function

Public Function IsNothing(varValue As Variant, Optional varDefalt As Variant = "") As String
    '-----------------------------------------------------------------------------------
    '功能:判断变量是否为空,为空返回默认值
    '参数:objValue   需要判断的对象
    '     strDefalt  默认值
    '返回:如果不为空值,返回原来的值,如果为空值,则返回指定的strDefalt值
    '-----------------------------------------------------------------------------------
    Dim var返回值 As Variant
    
    var返回值 = IIf(Trim(varValue) <> "", varValue, varDefalt)
    IsNothing = var返回值
End Function

Public Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------
    '功能:保存换号数据
    '参数:
    '返回:
    '编制:王吉
    '日期:2012-08-22
    '-----------------------------------------------------------------------------------
    Dim strNos As String
    Dim strSQL As String

    On Error GoTo Errhand
    
    mrs挂号信息.MoveFirst
    '获取单据号
    While mrs挂号信息.EOF = False
        strNos = strNos & "|" & Nvl(mrs挂号信息!NO, "")
        mrs挂号信息.MoveNext
    Wend
    
    If strNos <> "" Then
        strNos = Mid(strNos, 2)
    End If
    
    strSQL = "zl_病人挂号记录_批量换号("
    '单据号     Nos_In varchar
    strSQL = strSQL & "'" & strNos & "',"
    '新号别     新号别_In 病人挂号记录.号别%Type
    strSQL = strSQL & "'" & mrs可换号排班!号码 & "',"
    '新医生姓名 新医生姓名_In 挂号安排.医生姓名%Type
    strSQL = strSQL & "'" & Nvl(mrs可换号排班!医生姓名, Null) & "',"
    '新医生ID   新医生ID_In 挂号安排.医生ID%Type
    strSQL = strSQL & "'" & Nvl(mrs可换号排班!医生ID, Null) & "',"
    '新科室ID   新科室ID_In 挂号安排.科室ID%Type
    strSQL = strSQL & "'" & Nvl(mrs可换号排班!科室ID, Null) & "',"
    '原医生姓名 原医生姓名_In 挂号安排.医生姓名%Type
    strSQL = strSQL & "'" & Nvl(mrs挂号安排!医生姓名, Null) & "',"
    '原医生ID   原医生ID_In 挂号安排.医生ID%Type
    strSQL = strSQL & "'" & Nvl(mrs挂号安排!医生ID, Null) & "',"
    '原号别     原号别_In   病人挂号记录.号别%Type
    strSQL = strSQL & "'" & mrs挂号安排!号码 & "',"
    '操作员姓名 操作员姓名_In 挂号序号状态.操作员姓名%Type
    strSQL = strSQL & "'" & IIf(UserInfo.姓名 = "", Null, UserInfo.姓名) & "',"
    strSQL = strSQL & "" & Val(mrs挂号安排!ID) & ","
    strSQL = strSQL & "" & Val(mrs可换号排班!ID) & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function CheckValid() As Boolean
    '-----------------------------------------------------------------------------------
    '功能:保存换号数据
    '参数:
    '返回:成功返回True,失败返回False
    '编制:王吉
    '日期:2012-08-22
    '-----------------------------------------------------------------------------------

    If mrs挂号安排 Is Nothing Then
        MsgBox "您还没有选择需要进行换号操作的号别,不能进行换号操作!", vbInformation, gstrSysName
        CheckValid = False
        Exit Function
    End If
    If mrs可换号排班 Is Nothing Then
        MsgBox "您还没有选择需要新的号别,不能进行换号操作!", vbInformation, gstrSysName
        CheckValid = False
        Exit Function
    End If
    If mrs挂号信息 Is Nothing Then
        MsgBox "该号别下没有号被挂出不需要进行换号操作!", vbInformation, gstrSysName
        CheckValid = False
        Exit Function
    ElseIf mrs挂号信息.RecordCount <= 0 Then
        MsgBox "该号别下没有号被挂出不需要进行换号操作!", vbInformation, gstrSysName
        CheckValid = False
        Exit Function
    End If
    
    CheckValid = True
End Function
