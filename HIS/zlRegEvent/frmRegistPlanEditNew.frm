VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlanEditNew 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "挂号安排编辑"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   Icon            =   "frmRegistPlanEditNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBaseBack 
      BorderStyle     =   0  'None
      Height          =   8940
      Left            =   0
      ScaleHeight     =   8940
      ScaleWidth      =   9210
      TabIndex        =   3
      Top             =   0
      Width           =   9210
      Begin VB.Frame Frame2 
         Caption         =   "基本信息"
         Height          =   1500
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   8655
         Begin VB.TextBox txt号别 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            MaxLength       =   5
            TabIndex        =   30
            Top             =   270
            Width           =   960
         End
         Begin VB.ComboBox cboItem 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4275
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   660
            Width           =   2115
         End
         Begin VB.ComboBox cboDoctor 
            Height          =   300
            Left            =   1050
            TabIndex        =   28
            Top             =   1065
            Width           =   2115
         End
         Begin VB.ComboBox cbo科室 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            TabIndex        =   27
            Text            =   "cbo科室"
            Top             =   660
            Width           =   2115
         End
         Begin VB.CheckBox chk病案 
            Caption         =   "挂号时必须建病案"
            Height          =   195
            Left            =   4275
            TabIndex        =   26
            Top             =   1118
            Width           =   1845
         End
         Begin VB.ComboBox cbo号类 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4275
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   270
            Width           =   2115
         End
         Begin VB.CheckBox chk序号控制 
            Caption         =   "序号控制"
            Height          =   255
            Left            =   2130
            TabIndex        =   24
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "号别"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   615
            TabIndex        =   35
            Top             =   330
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "科室"
            Height          =   180
            Left            =   645
            TabIndex        =   34
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "项目"
            Height          =   180
            Left            =   3870
            TabIndex        =   33
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lbl医生 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "院内医生↓"
            Height          =   180
            Left            =   120
            TabIndex        =   32
            Top             =   1125
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "号类"
            Height          =   180
            Left            =   3855
            TabIndex        =   31
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "应诊时间"
         Height          =   2550
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   8655
         Begin VB.OptionButton opt天 
            Caption         =   "每天(&D)"
            Height          =   315
            Left            =   225
            TabIndex        =   16
            Top             =   285
            Width           =   960
         End
         Begin VB.OptionButton opt周 
            Caption         =   "每周(&W)"
            Height          =   315
            Left            =   225
            TabIndex        =   15
            Top             =   630
            Width           =   930
         End
         Begin VB.ComboBox cbo天 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   292
            Width           =   1110
         End
         Begin VB.CheckBox chk有效期 
            Caption         =   "有效期"
            Height          =   195
            Left            =   255
            TabIndex        =   13
            Top             =   2115
            Width           =   855
         End
         Begin VB.TextBox txt限约 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4980
            MaxLength       =   5
            TabIndex        =   12
            Top             =   292
            Width           =   1215
         End
         Begin VB.TextBox txt限号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3045
            MaxLength       =   5
            TabIndex        =   11
            Top             =   292
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   1170
            TabIndex        =   17
            Top             =   2055
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   117112835
            CurrentDate     =   38091
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   3555
            TabIndex        =   18
            Top             =   2055
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   117112835
            CurrentDate     =   38091
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPlan 
            Height          =   1275
            Left            =   1155
            TabIndex        =   19
            Top             =   675
            Width           =   7200
            _cx             =   12700
            _cy             =   2249
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
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmRegistPlanEditNew.frx":000C
            ScrollTrack     =   0   'False
            ScrollBars      =   0
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
            Editable        =   2
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "～"
            Height          =   180
            Left            =   3315
            TabIndex        =   22
            Top             =   2115
            Width           =   180
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "限约"
            Height          =   180
            Left            =   4545
            TabIndex        =   21
            Top             =   345
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "限号"
            Height          =   180
            Left            =   2610
            TabIndex        =   20
            Top             =   352
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "应诊诊室:"
         Height          =   4020
         Left            =   240
         TabIndex        =   4
         Top             =   4560
         Width           =   8640
         Begin VB.OptionButton opt分诊 
            Caption         =   "不分诊"
            Height          =   180
            Index           =   0
            Left            =   1020
            TabIndex        =   8
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "指定诊室"
            Height          =   180
            Index           =   1
            Left            =   2010
            TabIndex        =   7
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "动态分诊"
            Height          =   180
            Index           =   2
            Left            =   3180
            TabIndex        =   6
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "平均分诊"
            Height          =   180
            Index           =   3
            Left            =   4335
            TabIndex        =   5
            Top             =   0
            Width           =   1020
         End
         Begin MSComctlLib.ListView lvwDept 
            Height          =   3480
            Left            =   150
            TabIndex        =   9
            Top             =   300
            Width           =   8220
            _ExtentX        =   14499
            _ExtentY        =   6138
            View            =   2
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
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
            NumItems        =   0
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   9840
      TabIndex        =   2
      Top             =   1590
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   9840
      TabIndex        =   1
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9840
      TabIndex        =   0
      Top             =   1065
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   780
      Left            =   9240
      TabIndex        =   36
      Top             =   3720
      Width           =   1575
      _Version        =   589884
      _ExtentX        =   2778
      _ExtentY        =   1376
      _StockProps     =   64
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "院内医生"
         Index           =   0
      End
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "含外援医生"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmRegistPlanEditNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : frmRegistPlanEditNew
'Description :
'Author      : 李光福
'Date        : 05-November-2012 14:31:24
'Comments    :挂号安排管理,在老版本的基础上修改,集合时段和安排设置为一体

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Option Explicit '要求变量声明


Private Enum mPageIndex
    EM_安排 = 0
    EM_时段 = 1
End Enum

Private Enum mPgIndex
    Pg_计划安排 = 1
    Pg_计划时段 = 2
End Enum


Private mfrmTime As frmResistPlanTimeSet
Private mblnChangeByCode As Boolean '是否是代码控制改变了tabelpage的显示页
Private mrsRegOldData As ADODB.Recordset '本地数据集保存,原始挂号安排
Private mrsRegNewData As ADODB.Recordset '本地数据集保存 重新设置后的安排
Private mrsRegHistory As ADODB.Recordset '历次挂号的数据集

Private mlngModule As Long, mstrPrivs As String, mlngID As Long, mfrmMain As Form, mblnChange As Boolean
Private mrs科室 As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mblnFirst As Boolean
Private mblnSucces As Boolean
Private mlng缺省挂号科室ID  As Long '在挂号安排时，根据主界面中选择的科室进行缺省
Private mrs时间段 As ADODB.Recordset
Private mstr限制修改 As String '在某一天或者多天的安排限制更改
Public Enum RegEditType
    ed_新增 = 0
    ed_修改 = 1
    ed_查阅 = 2
End Enum
Private mEditType As RegEditType
Private mstr科室ID As String
Private mblnCboClick As Boolean     '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
Private mblnOnly院内医生 As Boolean '仅只能输院内医生


Private Type PlanInfo               '安排改变需要对比的信息
    str排班         As String       '排班信息
    str限号         As String       '限号信息
    bln序号         As Boolean      '是否序号控制
    bln时间段       As Boolean      '是否设置了时间段
End Type

Private mPlanInfo     As PlanInfo '原始的安排信息  主要用于安排修改时 相应信息的比较







Private Sub Form_Activate()
    Dim i As Integer
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitData = False Then Unload Me: Exit Sub
    If LoadCard = False Then Unload Me: Exit Sub
    Call cboDoctor_Validate(False)
    For i = 0 To opt分诊.UBound
        If opt分诊(i).Value Then Call opt分诊_Click(i): Exit For
    Next
    txt号别.SetFocus
End Sub

Private Sub Form_Load()
    Dim intTYPE As Integer
     Call InitPage
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    mblnFirst = True
    mblnOnly院内医生 = Val(zlDatabase.GetPara("只允许选院内医生", glngSys, mlngModule, "0", , InStr(1, mstrPrivs, ";参数设置;") > 0, intTYPE)) = 1
    If mblnOnly院内医生 Then
        mnuViewDoctor(0).Checked = True
        mnuViewDoctor(1).Checked = False
    Else
        mnuViewDoctor(0).Checked = False
        mnuViewDoctor(1).Checked = True
    End If
    lbl医生.Tag = IIf(mblnOnly院内医生, "0", "1")
    lbl医生.Caption = IIf(mblnOnly院内医生, "院内医生", "医生") & IIf(lbl医生.Tag = "1", "↓", "")
    lbl医生.ToolTipText = IIf(mblnOnly院内医生, "只能选院内建档医生", "含外援医生(除了可以选择院内医生外，还可以输入外援医生)")
End Sub


Private Sub InitPage()
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:

    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_计划安排, "计划安排", picBaseBack.hWnd, 0)
    ObjItem.Tag = mPgIndex.Pg_计划安排

    Set mfrmTime = New frmResistPlanTimeSet
    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_计划时段, "时段设置", mfrmTime.hWnd, 0)
    ObjItem.Tag = mPgIndex.Pg_计划时段
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With cmdOK
        .Left = ScaleWidth - .Width - 100
        cmdCancel.Left = .Left
        cmdHelp.Left = .Left
    End With

    With tbPage
        .Top = 50
        .Height = ScaleHeight - 100
        .Left = 50
        .Width = cmdOK.Left - .Left - 100
    End With

End Sub

Public Function ShowEdit(ByVal frmMain As Form, ByVal EditType As RegEditType, _
    ByVal lngModule As Long, ByVal strPrivs As String, Optional lngID As Long = 0, _
    Optional lng缺省科室ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用的主窗体
    '     EditType-编辑类型
    '出参:  `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   ``
    '返回:
    '编制:刘兴洪
    '日期:2009-09-15 10:25:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs: mlngID = lngID: mlng缺省挂号科室ID = lng缺省科室ID
    mEditType = EditType: mblnSucces = False
    mblnChange = False
    mstr限制修改 = Get已约限制(lngID)
    Me.Show 1, frmMain
    ShowEdit = mblnSucces

End Function

Private Function Get已约限制(ByVal lng安排ID As Long) As String
    '获取不能修改的安排星期
    Dim strSQL As String
    Dim rsTmp   As ADODB.Recordset
    Dim strTmp  As String
    strSQL = "Select Decode(To_Char(A.预约时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7'," & _
    "                             '周六') As 日期 " & vbCrLf & _
    "          From 病人挂号记录 A,挂号安排　B " & vbCrLf & _
    "        Where  A.号别=B.号码 And A.记录状态 = 1 And b.ID = [1] And A.发生时间 > A.登记时间 And A.预约时间 Is Not Null"

    If gint预约天数 = 0 Then
        strSQL = strSQL & " And A.预约时间 > Sysdate "
    Else
        strSQL = strSQL & " And A.预约时间 Between Sysdate And Sysdate+" & gint预约天数
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
    If rsTmp.EOF Then Exit Function

    Do While Not rsTmp.EOF
        If InStr(strTmp, Nvl(rsTmp!日期)) < 0 Or strTmp = "" Then
            strTmp = strTmp & ";" & Nvl(rsTmp!日期)
        End If
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then strTmp = strTmp & ";"
    Get已约限制 = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '返回:成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2009-09-15 13:14:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, rsTemp As ADODB.Recordset
    Dim bln所属部门 As Boolean

    Err = 0: On Error GoTo Errhand:
    gint号长 = GetMaxLen

    strSQL = "" & _
    "   Select '    ' 时间段 From dual Union All  " & _
    "   Select 时间段 From 时间段"
    Set mrs时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsPlan
        .Clear 1
        .Tag = .BuildComboList(mrs时间段, "时间段")

        .ColComboList(1) = .BuildComboList(mrs时间段, "时间段")
        For i = 2 To .Cols - 1
            .ColComboList(i) = .ColComboList(0)
        Next
    End With
    With cbo天
        Do While Not mrs时间段.EOF
            cbo天.AddItem Nvl(mrs时间段!时间段)
            mrs时间段.MoveNext
        Loop
        .ListIndex = 0
    End With

   '取出门诊临床科室
    Set mrs科室 = GetDepartments("'临床'", "1,3", Not zlStr.IsHavePrivs(mstrPrivs, "所有科室"))
    If mrs科室.RecordCount = 0 Then
        MsgBox "你不具备可用的临床科室信息或你权限不足,请先到部门管理中进行设置或找系统管理员分配权限！", vbInformation, gstrSysName
        Exit Function
    End If

    cbo科室.Clear
    Do While Not mrs科室.EOF
        cbo科室.AddItem mrs科室!名称
        cbo科室.ItemData(cbo科室.NewIndex) = Val(Nvl(mrs科室!ID))
        If mlng缺省挂号科室ID = Val(Nvl(mrs科室!ID)) Then cbo科室.ListIndex = cbo科室.NewIndex  '刘兴洪:增加从主界面中传入的科室
        mrs科室.MoveNext
    Loop

    '挂号项目
    strSQL = "Select ID as 序号,名称 From 收费项目目录 " & _
        " Where 类别='1' And (Sysdate Between 建档时间 And 撤档时间 Or 建档时间<Sysdate And 撤档时间 Is Null)" & _
        " And (站点='" & gstrNodeNo & "' Or 站点 is Null)" & _
        " Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    If rsTemp.RecordCount = 0 Then
        MsgBox "没有可用的挂号项目信息,请先到挂号项目设置中初始！", vbInformation, gstrSysName
        Exit Function
    End If
    cboItem.Clear
    Do While Not rsTemp.EOF
        cboItem.AddItem rsTemp!名称
        cboItem.ItemData(cboItem.NewIndex) = rsTemp!序号
        rsTemp.MoveNext
    Loop

    '号类
    strSQL = "Select 编码,名称,缺省标志 From 号类 Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    cbo号类.Clear
    Do While Not rsTemp.EOF
        cbo号类.AddItem rsTemp!名称
        If IIf(IsNull(rsTemp!缺省标志), 0, rsTemp!缺省标志) = 1 Then
            cbo号类.ListIndex = cbo号类.NewIndex
        End If
        rsTemp.MoveNext
    Loop

    '门诊诊室
    strSQL = "Select 编码,名称　From 门诊诊室 Where (站点='" & gstrNodeNo & "' Or 站点 is Null) Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    lvwDept.ListItems.Clear
    Do While Not rsTemp.EOF
        lvwDept.ListItems.Add , "D" & rsTemp!编码, rsTemp!名称
        rsTemp.MoveNext
    Loop
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.ActiveControl Is cbo科室 Then Exit Sub
    If Me.ActiveControl Is cboDoctor Then Exit Sub
    If Me.ActiveControl Is vsPlan Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mstr限制修改 = ""
End Sub



Private Sub lbl医生_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 0 Then Exit Sub
        If Val(lbl医生.Tag) = 0 Then Exit Sub

        PopupMenu mnuPopu, 2
End Sub



Private Sub lvwDept_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    If opt分诊(1).Value Then
        For i = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(i).Key <> Item.Key Then
                lvwDept.ListItems(i).Checked = False
            End If
        Next
    End If
    Set lvwDept.SelectedItem = Item
End Sub


Private Sub mnuViewDoctor_Click(Index As Integer)
        mnuViewDoctor(Index).Checked = True
        If Index = 0 Then
            mnuViewDoctor(1).Checked = False: mblnOnly院内医生 = True
        Else
            mnuViewDoctor(0).Checked = False: mblnOnly院内医生 = False
        End If

        lbl医生.Caption = IIf(mblnOnly院内医生, "院内医生", "医生") & "↓"
        lbl医生.ToolTipText = IIf(mblnOnly院内医生, "只能选择院内建档医生", "含外援医生(除了可以选择院内医生外，还可以输入外援医生)")
End Sub





'
'Private Sub vsPlan_EnterCell(Row As Long, Col As Long)
'    vsPlan.Active = opt周.Value
'End Sub

Private Sub opt分诊_Click(Index As Integer)
    Dim i As Integer, strKey As String
    If opt分诊(1).Value Then
        For i = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(i).Checked Then
                If strKey = "" Then
                    strKey = lvwDept.ListItems(i).Key
                Else
                    lvwDept.ListItems(i).Checked = False
                End If
            End If
        Next
        If strKey <> "" Then
            Set lvwDept.SelectedItem = lvwDept.ListItems(strKey)
            lvwDept.SelectedItem.EnsureVisible
        End If
    End If
End Sub

Private Sub opt天_Click()
    Dim i As Integer
    Dim strPlan As String

    For i = 0 To vsPlan.Cols - 1
        If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
            If strPlan = "" Then
                strPlan = vsPlan.TextMatrix(1, i)
            Else
                If vsPlan.TextMatrix(1, i) <> strPlan Then
                    strPlan = "": Exit For
                End If
            End If
        End If
    Next

    opt天.Value = -True: txt限号.Enabled = True: txt限约.Enabled = True
    cbo天.Enabled = True

    opt周.Value = False
    With vsPlan
        .Enabled = False: .TabStop = False
        For i = 1 To 7
             .TextMatrix(1, i) = ""
             .TextMatrix(2, i) = ""
             .TextMatrix(3, i) = ""
        Next
    End With

    cbo天.ListIndex = cbo.FindIndex(cbo天, strPlan, True)
    cbo天.SetFocus
End Sub

Private Sub opt周_Click()
    Dim i As Integer

    If Trim(cbo天.Text) <> "" Then
        For i = 1 To vsPlan.Cols - 1
            vsPlan.TextMatrix(1, i) = cbo天.Text
            vsPlan.TextMatrix(2, i) = txt限号.Text
            vsPlan.TextMatrix(3, i) = txt限约.Text
        Next
    End If

    opt天.Value = False
    cbo天.Enabled = False: txt限号.Enabled = False: txt限约.Enabled = False
    cbo天.ListIndex = -1

    opt周.Value = True
    vsPlan.Enabled = True: vsPlan.TabStop = True
    vsPlan.Col = 1: vsPlan.SetFocus
End Sub






Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

    If mblnChangeByCode Then Exit Sub
    PageChange Item
End Sub

Private Sub PageChange(ByVal Item As XtremeSuiteControls.ITabControlItem)

    If mblnChangeByCode Then Exit Sub

    If Item.Index = mPageIndex.EM_时段 Then
       mblnChangeByCode = True
       tbPage.Item(mPageIndex.EM_安排).Selected = True
        If isValied() = False Then
            mblnChangeByCode = False
            Exit Sub
        End If
        tbPage.Item(mPageIndex.EM_时段).Selected = True
        mblnChangeByCode = False
        Call LoadTimePlan
    Else
        If mfrmTime.mblnChange = False Then Exit Sub
        If mfrmTime.zl_CheckMoveAssign() = False Then
             mblnChangeByCode = True
            tbPage.Item(mPageIndex.EM_时段).Selected = True
             mblnChangeByCode = False
        End If
         
    End If
End Sub
Private Sub LoadTimePlan(Optional ByVal blnSaveBeforCheck As Boolean = False)
    Dim i As Long
    Dim lng限号数 As Long
    Dim lng限约数 As Long
    Dim strTemp As String
    Dim str安排 As String
    Dim str排班 As String

    If Not mrsRegNewData Is Nothing Then Set mrsRegNewData = Nothing

    If mrsRegNewData Is Nothing Then
        Set mrsRegNewData = New ADODB.Recordset
        mrsRegNewData.Fields.Append "ID", adBigInt, 18
        mrsRegNewData.Fields.Append "限制项目", adVarChar, 20
        mrsRegNewData.Fields.Append "排班", adVarChar, 20
        mrsRegNewData.Fields.Append "限号数", adBigInt, 10
        mrsRegNewData.Fields.Append "限约数", adBigInt, 18
        mrsRegNewData.Fields.Append "序号控制", adBigInt, 18
        mrsRegNewData.CursorLocation = adUseClient
        mrsRegNewData.LockType = adLockOptimistic
        mrsRegNewData.CursorType = adOpenStatic
        mrsRegNewData.Open
     End If

     If opt天.Value = True Then
          lng限号数 = Val(txt限号.Text)
          lng限约数 = Val(txt限约.Text)
          str排班 = Me.cbo天.Text
          For i = 0 To 6
            strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
            '周一,限号数,限约数|周二,限号数,限约数|....
            str安排 = str安排 & "|" & strTemp & "," & lng限号数 & "," & lng限约数
             With mrsRegNewData
                .AddNew
                !ID = 0
                !限制项目 = strTemp
                !排班 = str排班
                !限号数 = lng限号数
                !限约数 = lng限约数
                !序号控制 = Me.chk序号控制.Value
                .Update
            End With
          Next

        Else

           With vsPlan
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTemp = Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
                    lng限号数 = Val(Trim(vsPlan.TextMatrix(2, i)))
                    lng限约数 = Val(Trim(vsPlan.TextMatrix(3, i)))
                    str排班 = Trim(vsPlan.TextMatrix(1, i))
                    str安排 = str安排 & "|" & strTemp & "," & lng限号数 & "," & lng限约数
                    With mrsRegNewData
                        .AddNew
                        !ID = Val(mlngID)
                        !限制项目 = strTemp
                        !排班 = str排班
                        !限号数 = lng限号数
                        !限约数 = lng限约数
                        !序号控制 = Me.chk序号控制.Value
                        .Update
                    End With
                End If
            Next
        End With
     End If
     If str安排 <> "" Then str安排 = Mid(str安排, 2)
'Public Enum mRegEditType
'Ed_计划安排 = 0
'Ed_安排修改 = 1
'Ed_安排删除 = 2
'Ed_安排审核 = 3
'Ed_安排取消 = 4
'Ed_安排查阅 = 5
'End Enum

     mfrmTime.zlShowPagePlan str安排, mrsRegNewData, mrsRegHistory, chk序号控制.Value = 1, Switch(mEditType = ed_计划安排, EM_安排_增加, mEditType = Ed_安排修改, EM_安排_修改, True, EM_安排_查阅), mlngID, Val(0), blnSaveBeforCheck
End Sub

Private Sub txt号别_GotFocus()
    Call zlControl.TxtSelAll(txt号别)
End Sub

Private Sub txt号别_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt限号_GotFocus()
    Call zlControl.TxtSelAll(txt限号)
End Sub

Private Sub txt限号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt限号_Validate(Cancel As Boolean)
    If Trim(txt限号.Text) = "" And Trim(txt限约.Text) <> "" Then
        MsgBox "限约必须限号!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If

    If Trim(txt限号.Text) <> "" And Trim(txt限约.Text) <> "" And Val(txt限号.Text) < Val(txt限约.Text) Then
        MsgBox "限号数应大于限约数!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub txt限约_GotFocus()
    Call zlControl.TxtSelAll(txt限约)
End Sub

Private Sub txt限约_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If Val(txt限号.Text) = 0 Then KeyAscii = 0
End Sub

Private Sub txt限约_Validate(Cancel As Boolean)
    If Val(txt限号.Text) < Val(txt限约.Text) And _
        Trim(txt限号.Text) <> "" And Trim(txt限约.Text) <> "" Then
        MsgBox "限约数应小于限号数!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub
Private Function zlCheckRegistPlanIsValied(ByRef blnMulitNumPlan As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前所输入的号码是否合法
    '出参:blnMulitNumPlan-返回是否有多个相同(同一项目,同一科室,同一人,不同号)的安排
    '返回:合法返回,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-29 10:26:45
    '检查规则（同一项目,同一科室,同一人,不同号）:
    '     1.同天内不能有交叉的安排
    '问题目:35057
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str医生 As String
    Dim lng项目id As Long, lng科室ID As Long, lng医生ID As Long
    Dim str号别 As String, strTemp As String, strTemp1 As String
    Dim i As Long
    On Error GoTo errHandle
    lng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    lng项目id = cboItem.ItemData(cboItem.ListIndex)
    lng医生ID = 0: str医生 = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    strSQL = "" & _
        "   Select 号码,序号,周日 D0,周一 D1,周二 D2,周三 D3,周四 D4,周五 D5,周六 D6," & _
        "           To_Char(开始时间,'YYYY-MM-DD HH24:MI:SS') 开始时间,To_Char(终止时间,'YYYY-MM-DD HH24:MI:SS') 终止时间" & _
        "   From 挂号安排  "

    If lng医生ID = 0 Then
        strSQL = strSQL & _
            "   Where 科室id=[1] and  项目ID =[2] and 医生姓名=[3] and nvl(医生ID,0)=0 and ID<>" & mlngID & " Order by 序号"
    Else
        strSQL = strSQL & _
        "   Where 科室id=[1] and  项目ID =[2] and  医生ID=[4] and ID<>" & mlngID & " Order by 序号"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID, lng项目id, str医生, lng医生ID)
    blnMulitNumPlan = Not rsTemp.EOF
    If blnMulitNumPlan = False Then zlCheckRegistPlanIsValied = True: Exit Function
    str号别 = ""
    Do While Not rsTemp.EOF
        str号别 = str号别 & "," & Nvl(rsTemp!号码)
        If opt天.Value Then
            If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & " 周日:" & Nvl(rsTemp!D0)
            If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & " 周一:" & Nvl(rsTemp!D1)
            If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & " 周二:" & Nvl(rsTemp!D2)
            If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & " 周三:" & Nvl(rsTemp!D3)
            If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & " 周四:" & Nvl(rsTemp!D4)
            If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & " 周五:" & Nvl(rsTemp!D5)
            If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & " 周六:" & Nvl(rsTemp!D6)
            If strTemp <> "" Then
                strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下安排:" & vbCrLf & "        " & Mid(strTemp, 2)
                Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                zlCheckRegistPlanIsValied = False: Exit Function
            End If
        Else
            With vsPlan
                For i = 0 To 6
                    strTemp1 = "周" & Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
                    If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                        '存在,肯定重复了
                        strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                    End If
                Next
            End With
            If strTemp <> "" Then
                strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下安排:" & vbCrLf & "        " & Mid(strTemp, 2)
                Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                zlCheckRegistPlanIsValied = False: Exit Function
            End If
        End If
        rsTemp.MoveNext
    Loop
    If str号别 <> "" Then str号别 = Mid(str号别, 2)
    If MsgBox("注意:" & vbCrLf & "   发现『" & cboDoctor.Text & "』医生已经存在如下安排:" & vbCrLf & "    " & str号别 & vbCrLf & "   是否继续对该医生进行安排?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        zlCheckRegistPlanIsValied = True: Exit Function
    End If
    zlCheckRegistPlanIsValied = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Function zlCheckPlanArrageIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查计划安排是否有效
    '返回:检查计划安排是否存在相关的安排,如果有相关的安排,则返回False,否则返回true
    '编制:刘兴洪
    '日期:2010-12-29 19:53:56
    '问题目:35057
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str医生 As String
    Dim lng项目id As Long, lng科室ID As Long, lng医生ID As Long
    Dim str号别 As String, strTemp As String, strTemp1 As String
    Dim blnCheck As Boolean
    Dim i As Long
    On Error GoTo errHandle
    lng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    lng项目id = cboItem.ItemData(cboItem.ListIndex)
    lng医生ID = 0: str医生 = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)

    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  distinct A.号码,A.周日 D0,A.周一 D1,A.周二 D2,A.周三 D3,A.周四 D4,A.周五 D5,A.周六 D6," & _
    "           To_Char(生效时间,'YYYY-MM-DD HH24:MI:SS') 生效时间,To_Char(失效时间,'YYYY-MM-DD HH24:MI:SS') 失效时间" & _
    "   From 挂号安排计划 A, 挂号安排 B " & _
    "   Where A.安排ID=B.ID    " & _
    "      and   B.科室id=[1] and  B.项目ID =[2] and B.医生姓名=[3] and nvl(B.医生ID,0)=[4] and B.ID<>" & mlngID & _
    "   Order by 号码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID, lng项目id, str医生, lng医生ID)
    If rsTemp.EOF Then
        zlCheckPlanArrageIsValied = True: Exit Function
    End If
    Do While Not rsTemp.EOF
        str号别 = str号别 & "," & Nvl(rsTemp!号码)
        blnCheck = chk有效期.Value = 0
        If chk有效期.Value = 1 Then
            blnCheck = Nvl(rsTemp!生效时间) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!生效时间) < Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
            blnCheck = blnCheck Or Nvl(rsTemp!失效时间) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!失效时间) < Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
            blnCheck = blnCheck Or Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!生效时间) And Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!失效时间)
            blnCheck = blnCheck Or Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!生效时间) And Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!失效时间)

        End If
        If blnCheck Then
            If opt天.Value Then
                If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & " 周日:" & Nvl(rsTemp!D0)
                If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & " 周一:" & Nvl(rsTemp!D1)
                If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & " 周二:" & Nvl(rsTemp!D2)
                If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & " 周三:" & Nvl(rsTemp!D3)
                If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & " 周四:" & Nvl(rsTemp!D4)
                If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & " 周五:" & Nvl(rsTemp!D5)
                If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & " 周六:" & Nvl(rsTemp!D6)
                If strTemp <> "" Then
                    strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下计划安排:" & vbCrLf & "        " & Mid(strTemp, 2)
                    Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckPlanArrageIsValied = False: Exit Function
                End If
            Else
                With vsPlan
                    For i = 0 To 6
                        strTemp1 = "周" & Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
                        If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                            '存在,肯定重复了
                            strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                        End If
                    Next
                End With
                If strTemp <> "" Then
                    strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下计划安排:" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & "  生效时间:" & IIf(Nvl(rsTemp!生效时间) = "1901-01-01", "无限", Nvl(rsTemp!生效时间) & "-" & Nvl(rsTemp!失效时间)) & vbCrLf
                    Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckPlanArrageIsValied = False: Exit Function
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
    zlCheckPlanArrageIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Sub vsPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If mEditType = edt_查阅 Then Cancel = True: Exit Sub
        If Not opt周.Value = True Then Cancel = True: Exit Sub
    End With
End Sub


Private Sub vsPlan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置相关的格式
    '编制:刘兴洪
    '日期:2011-11-11 11:33:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsPlan
       If Row = 1 Then
              If Trim(.EditText) = "" Then
               .TextMatrix(2, Col) = ""
               .TextMatrix(3, Col) = ""
            End If
            Exit Sub
        End If
        .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), "###;;;")
    End With
    Exit Sub
End Sub
Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String
    Call zl_VsGridRowChange(vsPlan, OldRow, NewRow, OldCol, NewCol)
    vsPlan.ColComboList(NewCol) = ""

    If mstr限制修改 <> "" Then
        strTmp = ";周" & vsPlan.TextMatrix(0, NewCol) & ";"
        vsPlan.Editable = flexEDKbdMouse
        'If InStr(mstr限制修改, strTmp) > 0 Then vsPlan.Editable = flexEDNone
    End If

    If OldRow = 1 And Trim(vsPlan.TextMatrix(1, OldCol)) = "" Then
        vsPlan.TextMatrix(2, OldCol) = ""
        vsPlan.TextMatrix(3, OldCol) = ""
    End If
    If NewRow <> 1 Then Exit Sub
    vsPlan.ColComboList(NewCol) = vsPlan.Tag
End Sub
Private Sub vsPlan_GotFocus()
    Call zl_VsGridGotFocus(vsPlan)
End Sub
Private Sub vsPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    With vsPlan
        If KeyCode = vbKeyDelete Then
            .TextMatrix(.Row, .Col) = ""
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub

    With vsPlan
        If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If .Row < 3 Then
            .Row = .Row + 1
        Else
            .Row = 1
            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
         End If
    End With
End Sub

Private Sub vsPlan_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long

    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPlan
            If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If .Row < 3 Then
            .Row = .Row + 1
        Else
            .Row = 1
            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
         End If
    End With
End Sub
Private Sub vsPlan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub vsPlan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsPlan
        If Row <= 1 Then Exit Sub
        VsFlxGridCheckKeyPress vsPlan, Row, Col, KeyAscii, m数字式
    End With
End Sub
Private Sub vsPlan_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsPlan)
End Sub

Private Sub vsPlan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, strTemp As String, strTmp As String
    '数据验证
    With vsPlan
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        If .Row <= 1 Then Exit Sub
        If zlCommFun.DblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
            Cancel = True: Exit Sub
        End If
        strKey = Format(Abs(Val(strKey)), "####;;;")
         If mstr限制修改 <> "" Then
               strTmp = "周" & vsPlan.TextMatrix(0, Col)
               'vsPlan.Editable = flexEDKbdMouse
               If InStr(mstr限制修改, ";" & strTmp & ";") > 0 Then
                   Cancel = Val(strKey) < Val(.TextMatrix(Row, Col))
               End If
        End If
        If Cancel Then Exit Sub
        If Row = 2 Then
            If Val(strKey) < Val(.TextMatrix(3, Col)) Then
                If MsgBox("限号数小于了限约数,是否清空限约数?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True: Exit Sub
                .TextMatrix(3, Col) = ""
            End If
        ElseIf Row = 3 Then
            If Val(strKey) > Val(.TextMatrix(2, Col)) Then
                Call MsgBox("限号数小于了限约数,不能继续", vbOKOnly, gstrSysName)
                Cancel = True: Exit Sub
            End If
        End If

        .EditText = strKey
    End With
End Sub


Private Function Check时段() As Boolean
    '----------------------------------
    '判断是否分时段
    '----------------------------------
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset

    If mEditType = edt_查阅 Or mEditType = edt_新增 Then Exit Function

    On Error GoTo Hd
    strSQL = _
    "   Select 1 As Hdata From 挂号安排时段 Where 安排id =[1] And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
     Check时段 = Not rsTmp.EOF
    Set rsTmp = Nothing
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Private Function LoadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '返回:加载成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2009-09-15 12:14:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL          As String
    Dim rsTemp          As New ADODB.Recordset
    Dim i               As Long
    Dim strTemp         As String
    Dim rs限号          As ADODB.Recordset
    Dim bln每周         As Boolean
    Dim bln限号         As Boolean
    Dim str限号         As String
    Dim bln限约         As Boolean
    Dim str限约         As String
    Err = 0: On Error GoTo Errhand:

    For i = 1 To lvwDept.ListItems.Count
        lvwDept.ListItems(i).Checked = False
    Next

    If mEditType = edt_新增 Then
        txt号别.Text = GetNext号别
        txt限号.Text = ""
        txt限约.Text = ""
        chk病案.Value = 0

        If cbo科室.ListIndex >= 0 Then
            If mlng缺省挂号科室ID <> cbo科室.ItemData(cbo科室.ListIndex) Then
                cbo科室.ListIndex = -1
                cboItem.ListIndex = -1
                cboDoctor.Text = ""
            End If
        Else
            cbo科室.ListIndex = -1
            cboItem.ListIndex = -1
            cboDoctor.Text = ""
        End If
        dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = CDate("3000-01-01")

        opt天.Value = True
        cbo天.Enabled = True
        cbo天.ListIndex = cbo.FindIndex(cbo天, "全日", True)
        If cbo天.ListIndex = -1 Then cbo天.ListIndex = 0
        opt周.Value = False
        vsPlan.Enabled = False
        LoadCard = True
        opt分诊(0).Value = True
        Exit Function
    End If
    '修改或查看
    strSQL = " " & _
    "   Select A.Id as 安排ID,0 as 计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id," & _
    "          A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六,A.默认时段间隔, " & _
    "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室 " & _
    "   From 挂号安排 A,收费项目目录 B,部门表 D " & _
    "   Where A.项目id=b.Id(+) And A.科室id =d.Id(+) " & _
    "         And A.Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)

    If rsTemp.EOF Then
        ShowMsgbox "未找到指定的号别,请检查!"
        Exit Function
    End If
    strSQL = "Select 限制项目,限号数,  限约数 From  挂号安排限制 where 安排ID=[1]       "
   Set rs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)

    cbo号类.ListIndex = cbo.FindIndex(cbo号类, Nvl(rsTemp!号类), True)
    txt号别.Text = Nvl(rsTemp!号码)

    cbo科室.ListIndex = cbo.FindIndex(cbo科室, Nvl(rsTemp!科室), True)
    cboItem.ListIndex = cbo.FindIndex(cboItem, Nvl(rsTemp!项目), True)

    cboDoctor.ListIndex = cbo.FindIndex(cboDoctor, Nvl(rsTemp!医生姓名), True)
    If cboDoctor.ListIndex = -1 Then cboDoctor.Text = Nvl(rsTemp!医生姓名)


    chk病案.Value = IIf(Val(Nvl(rsTemp!病案必须)) = 1, 1, 0)

    chk序号控制.Value = IIf(Val(Nvl(rsTemp!序号控制)) = 1, 1, 0):     chk序号控制.Tag = chk序号控制.Value
    '获取修改前的安排是否序号控制
    mPlanInfo.bln序号 = IIf(Val(Nvl(rsTemp!序号控制)) = 1, True, False)
    '有效时间范围
    dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = CDate("3000-01-01")
    If Not IsNull(rsTemp!开始时间) Then
        chk有效期.Value = 1
        dtpBegin.Value = CDate(Format(rsTemp!开始时间, "yyyy-mm-dd HH:MM:SS"))
        If Not IsNull(rsTemp!终止时间) Then
            dtpEnd.Value = CDate(Format(rsTemp!终止时间, "yyyy-mm-dd HH:MM:SS"))
        End If
    End If

     '加载原始数据到数据集
     With mrsRegOldData
        Set mrsRegOldData = New ADODB.Recordset
        mrsRegOldData.Fields.Append "ID", adBigInt, 18
        mrsRegOldData.Fields.Append "限制项目", adVarChar, 20
        mrsRegOldData.Fields.Append "限号数", adBigInt, 10
        mrsRegOldData.Fields.Append "限约数", adBigInt, 18
        mrsRegOldData.Fields.Append "序号控制", adBigInt, 18
        mrsRegOldData.CursorLocation = adUseClient
        mrsRegOldData.LockType = adLockOptimistic
        mrsRegOldData.CursorType = adOpenStatic
        mrsRegOldData.Open


        rs限号.Filter = 0
        If rs限号.RecordCount > 0 Then rs限号.MoveFirst
        Do While Not rs限号.EOF
            With mrsRegOldData
                .AddNew
                !ID = mlngID
                !限制项目 = Nvl(rs限号!限制项目)
                !限号数 = Val(Nvl(rs限号!限号数))
                !限约数 = Val(Nvl(rs限号!限约数))
                !序号控制 = Val(Nvl(rsTemp!序号控制))
                .Update
            End With
            rs限号.MoveNext
        Loop
    End With

    Call LoadRegHistory

    '---------------------------------------------------
    '判断 每日安排 限号数 限约数 等是否一致
    '---------------------------------------------------
    bln每周 = Nvl(rsTemp!周日) <> Nvl(rsTemp!周一) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周二) _
        Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周三) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周四) _
        Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周五) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周六)

    If bln每周 = False Then
             rs限号.Filter = "限制项目='周日'"
             If Not rs限号.EOF Then
                str限号 = Nvl(rs限号!限号数)
                str限约 = Nvl(rs限号!限约数)
             End If
            For i = 1 To 6
                strTemp = Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
                rs限号.Filter = "限制项目='" & "周" & strTemp & "'"
                If Not rs限号.EOF Then
                    bln限号 = Nvl(rs限号!限号数) = str限号
                    bln限约 = Nvl(rs限号!限约数) = str限约
                    If bln限约 = False Or bln限号 = False Then Exit For
                End If
            Next
          bln每周 = True
         If bln限号 And bln限约 Then bln每周 = False

    End If

   If bln每周 Or mrsRegHistory.RecordCount > 0 Then
        '每周
        opt周.Value = True
        With vsPlan
            For i = 1 To .Cols - 1
                strTemp = Switch(i - 1 = 0, "日", i - 1 = 1, "一", i - 1 = 2, "二", i - 1 = 3, "三", i - 1 = 4, "四", i - 1 = 5, "五", True, "六")
                .TextMatrix(1, i) = Nvl(rsTemp.Fields("周" & strTemp))
                rs限号.Filter = "限制项目='" & "周" & strTemp & "'"
                If Not rs限号.EOF Then
                    .TextMatrix(2, i) = Nvl(rs限号!限号数)
                    .TextMatrix(3, i) = Nvl(rs限号!限约数)
                End If
                If InStr(mstr限制修改, ";周" & strTemp & ";") > 0 Then
                    .Cell(flexcpForeColor, 2, i, 3, i) = vbBlue
                End If
            Next
        End With
        opt天.Value = False: cbo天.Enabled = False: txt限号.Enabled = False: txt限约.Enabled = False
        vsPlan.Enabled = True: chk序号控制.Enabled = mstr限制修改 = ""
    Else
        '每天
        opt天.Value = True:  cbo天.ListIndex = cbo.FindIndex(cbo天, Nvl(rsTemp!周日), True)
        If cbo天.ListIndex = -1 Then cbo天.ListIndex = 0:
        opt周.Value = False: vsPlan.Enabled = False
        If rs限号.RecordCount <> 0 Then rs限号.MoveFirst
        If rs限号.EOF = False Then
            txt限号.Text = Nvl(rs限号!限号数)
            txt限约.Text = Nvl(rs限号!限约数)
        End If
    End If

    '------------------------------
    '获取修改前的 时间段和 限号数
    '用于在保存时 对比限号限约以及时间段是否发生了变化
    '如果发生了变化则需要提示  操作员重新设置时段信息
    '------------------------------
   mPlanInfo.str排班 = ""
   mPlanInfo.str限号 = ""

    If bln每周 Or mrsRegHistory.RecordCount > 0 Then
         For i = 1 To vsPlan.Cols - 1
            mPlanInfo.str排班 = mPlanInfo.str排班 & "'" & Trim(vsPlan.TextMatrix(1, i)) & "',"

                mPlanInfo.str限号 = mPlanInfo.str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                     mPlanInfo.str限号 = mPlanInfo.str限号 & ",0,0"
                Else
                     mPlanInfo.str限号 = mPlanInfo.str限号 & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Val(Trim(vsPlan.TextMatrix(3, i)))
                End If
        Next


    Else
         For i = 1 To 7
             mPlanInfo.str排班 = mPlanInfo.str排班 & "'" & Trim(cbo天.Text) & "',"
             mPlanInfo.str限号 = mPlanInfo.str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
             mPlanInfo.str限号 = mPlanInfo.str限号 & "," & Val(txt限号.Text) & "," & Val(txt限约.Text)
        Next
    End If
    If mPlanInfo.str限号 <> "" Then mPlanInfo.str限号 = Mid(mPlanInfo.str限号, 2)
    '-------------------------------

     Select Case Val(Nvl(rsTemp!分诊方式))     '0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
        Case 0  '"不分诊"
            opt分诊(0).Value = True
        Case 1  ' "指定诊室"
            opt分诊(1).Value = True
        Case 2 '"动态分诊"
            opt分诊(2).Value = True
        Case 3 ' "平均分诊"
            opt分诊(3).Value = True
    End Select

    strSQL = "Select 号表ID,门诊诊室　From 挂号安排诊室 Where 号表ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    Do While Not rsTemp.EOF
        For i = 1 To lvwDept.ListItems.Count
            If rsTemp!门诊诊室 = lvwDept.ListItems(i).Text Then
                lvwDept.ListItems(i).Checked = True
            End If
        Next
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If mstr限制修改 <> "" Then opt天.Enabled = False
    '如果是修改时 获取原来的安排是否已经安排了时段
    If mEditType = edt_修改 Then mPlanInfo.bln时间段 = Check时段
    If mrsRegHistory.RecordCount > 0 Then opt天.Enabled = False
    LoadCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Sub cboDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        cboDoctor.ListIndex = GetCboIndex(cboDoctor, cboDoctor)
'    End If
End Sub

Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    If KeyAscii <> 13 Then Exit Sub
    If cboDoctor.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If mrsDoctor Is Nothing Then Exit Sub
    If Trim(cboDoctor.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub

    If zlPersonSelect(Me, mlngModule, cboDoctor, mrsDoctor, cboDoctor.Text, True, "") = False Then
        If mblnOnly院内医生 = False Then
                zlCommFun.PressKey vbKeyTab
        End If
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cboDoctor_Validate(Cancel As Boolean)
      If mblnOnly院内医生 Then
           If cboDoctor.ListIndex < 0 Then cboDoctor.Text = ""
      End If

    '指定医生时不能指定多个科室
    If Trim(cboDoctor.Text) <> "" Then
        opt分诊(2).Enabled = False
        opt分诊(3).Enabled = False
        If opt分诊(2).Value Or opt分诊(3).Value Then opt分诊(0).Value = True
    Else
        opt分诊(2).Enabled = True
        opt分诊(3).Enabled = True
    End If
End Sub

Private Sub cbo科室_Click()
    mblnCboClick = True
    If cbo科室.ListIndex = -1 Then Exit Sub
    Call LoadDoctor
End Sub

Private Sub LoadDoctor()
    Set mrsDoctor = GetDoctor(Val(cbo科室.ItemData(cbo科室.ListIndex)), "")
    cboDoctor.Clear
    Do While Not mrsDoctor.EOF
        cboDoctor.AddItem mrsDoctor!姓名
        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
        mrsDoctor.MoveNext
    Loop
End Sub

Private Sub cbo科室_GotFocus()
    zlControl.TxtSelAll cbo科室
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo科室.Text = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If cbo科室.ListIndex >= 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        mblnCboClick = True
        If Select科室(Me, mlngModule, mrs科室, cbo科室, cbo科室.Text) = True Then
            mblnCboClick = False
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        If cbo科室.Enabled Then cbo科室.SetFocus
        mblnCboClick = False
        zlControl.TxtSelAll cbo科室
    Else
       ' Call zlControl.CboSetIndex(cbo科室.hWnd, zlControl.CboMatchIndex(cbo科室.hWnd, KeyAscii))
    End If
End Sub

Private Sub cbo科室_Validate(Cancel As Boolean)
 '如果在cbo的keypress事件中用了弹出列表的的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
    If Not mblnCboClick Then cbo科室_Click
    mblnCboClick = False
End Sub

Private Sub chk有效期_Click()
    dtpBegin.Enabled = chk有效期.Value = 1
    dtpEnd.Enabled = chk有效期.Value = 1

    If Visible And dtpBegin.Enabled Then
        dtpBegin.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Function GetDoctorPlan(lng医生ID As Long, str医生姓名 As String) As ADODB.Recordset
'功能:返回指定医生ID或姓名的已有号别的时间信息
'   用于检查新增或修改的号别是否与现有的号别在时间上重复
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select 号码,周日 D0,周一 D1,周二 D2,周三 D3,周四 D4,周五 D5,周六 D6," & _
            " To_Char(开始时间,'YYYY-MM-DD HH24:MI:SS') 开始时间,To_Char(终止时间,'YYYY-MM-DD HH24:MI:SS') 终止时间" & _
            " From 挂号安排 Where (终止时间 is null or 终止时间>sysdate) And " & IIf(lng医生ID <> 0, " 医生ID=[1]", " 医生姓名=[1]") & _
            IIf(mEditType = edt_新增, "", " And ID<>[2]")
    Set GetDoctorPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(lng医生ID <> 0, lng医生ID, str医生姓名), mlngID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckExistsBooking() As Boolean
'功能:检查当前时间段之外是否存在预约挂号单
    Dim rsTemp As ADODB.Recordset, rsBooking As ADODB.Recordset, strSQL As String
    Dim i As Long, str时间段 As String

    On Error GoTo errH
    If opt天.Value Then
        str时间段 = _
               "Select 1 From 时间段 b Where b.时间段 = [2] And (" & _
               " ('3000-01-10 '||To_Char(a.发生时间,'HH24:MI:SS')" & _
               " Between" & _
               " Decode(Sign(b.开始时间-b.终止时间),1,'3000-01-09 '||To_Char(b.开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.开始时间,'HH24:MI:SS'))" & _
               " And" & _
               " '3000-01-10 '||To_Char(b.终止时间,'HH24:MI:SS'))" & _
               " Or" & _
               " ('3000-01-10 '||To_Char(a.发生时间,'HH24:MI:SS')" & _
               " Between" & _
               " '3000-01-10 '||To_Char(b.开始时间,'HH24:MI:SS')" & _
               " And" & _
               " Decode(Sign(b.开始时间-b.终止时间),1,'3000-01-11 '||To_Char(b.终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.终止时间,'HH24:MI:SS'))))"

        strSQL = "Select  /*+ Rule*/ Min(发生时间) 时间" & vbNewLine & _
            "From 门诊费用记录 a" & vbNewLine & _
            "Where 记录性质 = 4 And 记录状态 In (0, 1) And 计算单位 = [1] And 发生时间 > 登记时间"
        If gint预约天数 = 0 Then
            strSQL = strSQL & " And 发生时间 > Sysdate"
        Else
            strSQL = strSQL & " And 发生时间 Between Sysdate And Sysdate+" & gint预约天数
        End If
        strSQL = strSQL & " And Not Exists (" & str时间段 & ")"

        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt号别.Text, Trim(cbo天.Text))
        CheckExistsBooking = Not IsNull(rsTemp!时间)
    Else
        strSQL = "Select /*+ Rule*/ 发生时间,To_Char(发生时间,'D') 星期 From 门诊费用记录 a Where 记录性质 = 4 and 记录状态 In(0,1) And 计算单位 = [1] And 发生时间 > 登记时间"
        If gint预约天数 = 0 Then
            strSQL = strSQL & " And 发生时间 > Sysdate"
        Else
            strSQL = strSQL & " And 发生时间 Between Sysdate And Sysdate+" & gint预约天数
        End If

        Set rsBooking = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt号别.Text)
        For i = 1 To rsBooking.RecordCount
            str时间段 = Trim(vsPlan.TextMatrix(1, rsBooking!星期 - 1))
            If str时间段 = "" Then
                CheckExistsBooking = True
            Else
               strSQL = _
                    "Select Count(*) cnt From 时间段 b Where b.时间段 = [2] And (" & _
                    " ('3000-01-10 '||To_Char([1],'HH24:MI:SS')" & _
                    " Between" & _
                    " Decode(Sign(b.开始时间-b.终止时间),1,'3000-01-09 '||To_Char(b.开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.开始时间,'HH24:MI:SS'))" & _
                    " And" & _
                    " '3000-01-10 '||To_Char(b.终止时间,'HH24:MI:SS'))" & _
                    " Or" & _
                    " ('3000-01-10 '||To_Char([1],'HH24:MI:SS')" & _
                    " Between" & _
                    " '3000-01-10 '||To_Char(b.开始时间,'HH24:MI:SS')" & _
                    " And" & _
                    " Decode(Sign(b.开始时间-b.终止时间),1,'3000-01-11 '||To_Char(b.终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.终止时间,'HH24:MI:SS'))))"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(rsBooking!发生时间), str时间段)
                CheckExistsBooking = rsTemp!cnt = 0
            End If

            If CheckExistsBooking Then Exit Function
            rsBooking.MoveNext
        Next
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function isValied() As Boolean
     Dim i As Integer, intCount As Integer, j As Integer
    Dim str时间段 As String, str诊室 As String, str限号 As String
    Dim lngNextID As Long, lng医生ID As Long
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strInfo As String, strTmp As String, strOld As String, strNew As String
    Dim str号别 As String
    Dim rsDoctorPlan As ADODB.Recordset
    Dim rsNewDate As ADODB.Recordset
    Dim rsOldDate As ADODB.Recordset
    Dim rsSNState As ADODB.Recordset
    Dim blnMulitNumPlan As Boolean  '是否多次安排
    Dim blnChange       As Boolean '是否改变了 时间安排
    Dim strMsg          As String

    If opt天.Value Then
        If cbo天.ListIndex = -1 Then
            MsgBox "该号别每天的应诊时间未设置！", vbInformation, gstrSysName
            cbo天.SetFocus: Exit Function
        End If

        If Val(txt限号.Text) = 0 And Val(txt限约.Text) = 0 Then
            MsgBox "安排设置时段时,必须设置限号或限约数！", vbInformation, gstrSysName
            txt限号.SetFocus: Exit Function
        End If
        '限号限约规则
        If Trim(txt限号.Text) <> "" Then
            If Trim(txt限约.Text) <> "" And Val(txt限号.Text) < Val(txt限约.Text) Then
                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
                txt限约.SetFocus: Exit Function
            End If
        ElseIf Trim(txt限约.Text) <> "" Then
            MsgBox "限约必须限号！", vbInformation, gstrSysName
            txt限号.SetFocus: Exit Function
        End If
    Else
        With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))

                        If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                            MsgBox "安排设置时段时,必须设置限号或限约数！", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Function
                        End If

                        '限号限约规则
                        If Val(.TextMatrix(2, i)) <> 0 Then
                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
                                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
                                .Row = 2: .Col = i
                                .SetFocus: Exit Function
                            End If
                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
                            
                            MsgBox "限约必须限号！", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Function
                        End If
                End If
            Next
            If strTmp = "" Then
                MsgBox "该号别每周的应诊时间未设置！", vbInformation, gstrSysName
                vsPlan.SetFocus: Exit Function
            End If
        End With
    End If
    isValied = True
End Function
Private Sub cmdOK_Click()
    Dim i As Integer, intCount As Integer, j As Integer
    Dim str时间段 As String, str诊室 As String, str限号 As String
    Dim lngNextID As Long, lng医生ID As Long
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strInfo As String, strTmp As String, strOld As String, strNew As String
    Dim cllPro As Collection
    Dim str号别 As String
    Dim rsDoctorPlan As ADODB.Recordset
    Dim rsNewDate As ADODB.Recordset
    Dim rsOldDate As ADODB.Recordset
    Dim rsSNState As ADODB.Recordset
    Dim blnMulitNumPlan As Boolean  '是否多次安排
    Dim blnChange       As Boolean '是否改变了 时间安排
    Dim strMsg          As String
    If mEditType = edt_查阅 Then Unload Me: Exit Sub
    If Me.tbPage.Item(mPageIndex.EM_安排).Selected = False Then
        mblnChangeByCode = True
        tbPage.Item(mPageIndex.EM_安排).Selected = True
        mblnChangeByCode = False
    End If
    If mblnOnly院内医生 Then
        If cboDoctor.ListIndex < 0 And cboDoctor.Text <> "" Then
                MsgBox "你选择的医生不存在,请重新输入医生!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                If cboDoctor.Enabled Then cboDoctor.SetFocus
                Exit Sub
        End If
    End If
    '完整性检查
    If Trim(txt号别) = "" Then
        MsgBox "号别不能为空！", vbInformation, gstrSysName
        txt号别.SetFocus: Exit Sub
    End If
    If cbo科室.ListIndex = -1 Then
        MsgBox "未设置号别所对应的科室！", vbInformation, gstrSysName
        cbo科室.SetFocus: Exit Sub
    End If
    If cboItem.ListIndex = -1 Then
        MsgBox "未设置号别所对应的挂号项目！", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Sub
    End If

    If dtpBegin.Enabled And dtpEnd.Enabled Then
        If dtpBegin.Value >= dtpEnd.Value Then
            MsgBox "开始时间应该小于结束时间。", vbInformation, gstrSysName
            dtpBegin.SetFocus: Exit Sub
        End If
    End If

    If opt天.Value Then
        If cbo天.ListIndex = -1 Then
            MsgBox "该号别每天的应诊时间未设置！", vbInformation, gstrSysName
            cbo天.SetFocus: Exit Sub
        End If
        If chk序号控制.Value = 1 Then
            If Val(txt限号.Text) = 0 And Val(txt限约.Text) = 0 Then
                MsgBox "使用序号控制时,必须设置限号或限约数！", vbInformation, gstrSysName
                txt限号.SetFocus: Exit Sub
            End If
        End If
        '限号限约规则
        If Trim(txt限号.Text) <> "" Then
            If Trim(txt限约.Text) <> "" And Val(txt限号.Text) < Val(txt限约.Text) Then
                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
                txt限约.SetFocus: Exit Sub
            End If
        ElseIf Trim(txt限约.Text) <> "" Then
            MsgBox "限约必须限号！", vbInformation, gstrSysName
            txt限号.SetFocus: Exit Sub
        End If
    Else
        With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))
                    If chk序号控制.Value = 1 Then
                          If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                              MsgBox "使用序号控制时,必须设置限号或限约数！", vbInformation, gstrSysName
                              .Row = 2: .Col = i
                              .SetFocus: Exit Sub
                          End If
                      End If
                        '限号限约规则
                        If Val(.TextMatrix(2, i)) <> 0 Then
                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
                                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
                                .Row = 2: .Col = i
                                .SetFocus: Exit Sub
                            End If
                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
                            MsgBox "限约必须限号！", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Sub
                        End If
                End If
            Next
            If strTmp = "" Then
                MsgBox "该号别每周的应诊时间未设置！", vbInformation, gstrSysName
                vsPlan.SetFocus: Exit Sub
            End If
        End With
    End If
    '诊室判断
    If opt分诊(1).Value Or opt分诊(2).Value Or opt分诊(3).Value Then
        intCount = 0
        For i = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(i).Checked Then intCount = intCount + 1
        Next
        If opt分诊(1).Value Then
            If intCount = 0 Then
                MsgBox "指定诊室时必须选择一个对应的门诊诊室！", vbInformation, gstrSysName
                lvwDept.SetFocus: Exit Sub
            ElseIf intCount > 1 Then
                MsgBox "指定诊室时只能选择一个对应的门诊诊室！", vbInformation, gstrSysName
                lvwDept.SetFocus: Exit Sub
            End If
        ElseIf opt分诊(2).Value Or opt分诊(3).Value Then
            If intCount < 2 Then
                MsgBox "动态分诊或平均分诊时至少要选择两个对应的门诊诊室！", vbInformation, gstrSysName
                lvwDept.SetFocus: Exit Sub
            End If
        End If
    End If

    '项目价格判断
    If ReadRegistPrice(cboItem.ItemData(cboItem.ListIndex), False, False) = 0 Then
        MsgBox "项目""" & cboItem.Text & """未设置有效价格,请先到收费项目管理中设置！", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Sub
    End If

    '取医生ID
    If cboDoctor.ListIndex <> -1 Then lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
'    '问题:现在一个医生可以加入重复号了
'    If zlCheckPlanArrageIsValied = False Then
'        If cboDoctor.Enabled Then cboDoctor.SetFocus
'        Exit Sub
'    End If
'
'    If zlCheckRegistPlanIsValied(blnMulitNumPlan) = False Then
'        If cboDoctor.Enabled Then cboDoctor.SetFocus
'        Exit Sub
'    End If
    '是否同一医生的安排时间段是否重复或交叉
    If Trim(cboDoctor.Text) <> "" Then
        Set rsDoctorPlan = GetDoctorPlan(lng医生ID, cboDoctor.Text)
        If rsDoctorPlan.RecordCount > 0 Then
            strSQL = "Select 时间段, 开始时间, Decode(Sign(终止时间 - 开始时间), 1, 终止时间 , 终止时间+ 1) 终止时间 From 时间段"
            Set rsNewDate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            Set rsOldDate = rsNewDate.Clone
        End If

        strInfo = ""
        For j = 1 To rsDoctorPlan.RecordCount
            strTmp = ""
            For i = 0 To IIf(opt天.Value, 6, vsPlan.Cols - 2)
               strOld = "" & rsDoctorPlan.Fields("D" & i).Value
               If opt天.Value Then
                   strNew = cbo天.Text
               Else
                   strNew = Trim(vsPlan.TextMatrix(1, i + 1))
               End If

               rsNewDate.Filter = "时间段='" & strNew & "'"
               rsOldDate.Filter = "时间段='" & strOld & "'"
               If rsNewDate.RecordCount > 0 And rsOldDate.RecordCount > 0 Then
                    If rsNewDate!开始时间 >= rsOldDate!开始时间 And rsNewDate!开始时间 <= rsOldDate!终止时间 Or rsNewDate!终止时间 >= rsOldDate!开始时间 And rsNewDate!终止时间 <= rsOldDate!终止时间 Or rsNewDate!开始时间 <= rsOldDate!开始时间 And rsNewDate!终止时间 >= rsOldDate!终止时间 Then
                    '时间交叉,再判断效期是否交叉
                         If chk有效期.Value = 0 Then
                             strTmp = strTmp & "," & "星期" & Choose(i + 1, "日", "一", "二", "三", "四", "五", "六") & ":" & strOld
                         Else
                             '为简化判断,假定数据按规范保存,开始时间和结束时间,要么都有,要么都没有,所以仅以开始时间来判断有无
                             If IsNull(rsDoctorPlan!开始时间) Then
                                 strTmp = strTmp & "," & "星期" & Choose(i + 1, "日", "一", "二", "三", "四", "五", "六") & ":" & strOld
                             Else
                                 If dtpBegin.Value >= CDate(rsDoctorPlan!开始时间) And dtpBegin.Value <= CDate(Nvl(rsDoctorPlan!终止时间, "3000-01-01")) Or dtpEnd.Value >= CDate(rsDoctorPlan!开始时间) And dtpEnd.Value <= CDate(Nvl(rsDoctorPlan!终止时间, "3000-01-01")) Or dtpBegin.Value <= CDate(rsDoctorPlan!开始时间) And dtpEnd.Value >= CDate(Nvl(rsDoctorPlan!终止时间, "3000-01-01")) Then
                                    strTmp = strTmp & "," & "星期" & Choose(i + 1, "日", "一", "二", "三", "四", "五", "六") & ":" & strOld
                                 End If
                             End If
                         End If
                    End If
               End If
            Next
            If strTmp <> "" Then
                strInfo = strInfo & vbCrLf & "在号别 [" & rsDoctorPlan!号码 & "] 中已有如下安排:" & vbCrLf & "        " & Mid(strTmp, 2)
                If Not IsNull(rsDoctorPlan!开始时间) Then
                    strInfo = strInfo & vbCrLf & "        有效期:" & rsDoctorPlan!开始时间 & "~" & rsDoctorPlan!终止时间
                Else
                    strInfo = strInfo & vbCrLf & "        有效期:不限"
                End If
            End If
            rsDoctorPlan.MoveNext
        Next
        If strInfo <> "" Then
            If blnMulitNumPlan Then
                '多次安排时,不能存在交叉
                Call MsgBox("发现" & cboDoctor.Text & "医生存在与当前号别重复或交叉的挂号安排" & vbCrLf & strInfo & vbCrLf & vbCrLf & "不能安排!", vbInformation + vbOKOnly, gstrSysName)
                Exit Sub
            Else
                If MsgBox("发现" & cboDoctor.Text & "医生存在与当前号别重复或交叉的挂号安排" & vbCrLf & strInfo & vbCrLf & vbCrLf & "确实要保存当前号别吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If

    If Not mEditType = edt_新增 Then
        If CheckExistsBooking() Then
            If MsgBox("该号别当前安排的时间段之外存在预约挂号单,是否要继续?", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    '先检查
    '取时间段
    str限号 = ""
    If opt天.Value Then '每天
        For i = 1 To 7
            str时间段 = str时间段 & "'" & Trim(cbo天.Text) & "',"
            str限号 = str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
            str限号 = str限号 & "," & Val(txt限号.Text) & "," & Val(txt限约.Text)
        Next
    Else
        For i = 1 To vsPlan.Cols - 1
            str时间段 = str时间段 & "'" & Trim(vsPlan.TextMatrix(1, i)) & "',"

                str限号 = str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                    str限号 = str限号 & ",0,0"
                Else
                    str限号 = str限号 & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Val(Trim(vsPlan.TextMatrix(3, i)))
                End If
        Next
    End If
    If str限号 <> "" Then str限号 = Mid(str限号, 2)


    '取挂号诊室
    For i = 1 To lvwDept.ListItems.Count
        If lvwDept.ListItems(i).Checked Then
            str诊室 = str诊室 & ";" & lvwDept.ListItems(i).Text
        End If
    Next
    str诊室 = Mid(str诊室, 2)


    '取分诊方式
    intCount = 0
    For i = 0 To opt分诊.UBound
        If opt分诊(i).Value Then intCount = i: Exit For
    Next

    '取开始时间范围
    strBegin = "NULL": strEnd = "NULL"
    If chk有效期.Value = 1 Then
        strBegin = "To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strEnd = "To_Date('" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    End If

      '查看是否改变了排班或者 改变了 限号数 限约数 或者序号控制
    blnChange = (str限号 <> mPlanInfo.str限号) Or (str时间段 <> mPlanInfo.str排班)
    blnChange = blnChange Or (chk序号控制.Value <> IIf(mPlanInfo.bln序号, 1, 0))
    str限号 = "'" & str限号 & "',"
    Set cllPro = New Collection
    '取ID
    If mEditType = edt_新增 Then

        '新增
        lngNextID = zlDatabase.GetNextId("挂号安排")

        strSQL = "zl_挂号安排_INSERT(" & _
            lngNextID & ",'" & Trim(txt号别.Text) & "','" & cbo号类.Text & "'," & _
            cbo科室.ItemData(cbo科室.ListIndex) & "," & _
            cboItem.ItemData(cboItem.ListIndex) & ",'" & Trim(cboDoctor.Text) & "'," & _
            lng医生ID & "," & _
            chk病案.Value & "," & str时间段 & str限号 & intCount & "," & _
            "'" & str诊室 & "'," & strBegin & "," & strEnd & ",1," & chk序号控制.Value & ",0," & 5 & ")"
    Else
'
' Zl_挂号安排_Insert
'(
'  Id_In       挂号安排.ID%Type,
'  号码_In     挂号安排.号码%Type,
'  号类_In     挂号安排.号类%Type,
'  科室id_In   挂号安排.科室id%Type,
'  项目id_In   挂号安排.项目id%Type,
'  医生_In     挂号安排.医生姓名%Type,
'  医生id_In   挂号安排.医生id%Type,
'  病案必须_In 挂号安排.病案必须%Type,
'  周日_In     挂号安排.周日%Type,
'  周一_In     挂号安排.周一%Type,
'  周二_In     挂号安排.周二%Type,
'  周三_In     挂号安排.周三%Type,
'  周四_In     挂号安排.周四%Type,
'  周五_In     挂号安排.周五%Type,
'  周六_In     挂号安排.周六%Type,
'  限号控制_In Varchar2,
'  分诊方式_In 挂号安排.分诊方式%Type,
'  诊室_In     Varchar2,
'  开始时间_In 挂号安排.开始时间%Type,
'  终止时间_In 挂号安排.终止时间%Type,
'  新增_In     Number,
'  序号控制_In 挂号安排.序号控制%Type,
'  处理类型_In Number:=0,
'  默认时段间隔_In 挂号安排.默认时段间隔%Type
') As
'  -----------------------------------------------------------
'  --参数：
'  --  诊室_IN=以';'号分隔的多个诊室名称
'  --  限号控制_IN:|周一,22(限号),13(限约)|周二,20(限号),11(限约)....
'  --  处理类型_IN:修改安排时 对时段数据的处理 0--不处理 1--删除时段信息
        '修改

        lngNextID = mlngID
        strSQL = "    " & vbNewLine & "zl_挂号安排_INSERT("
        strSQL = strSQL & vbNewLine & lngNextID
        strSQL = strSQL & vbNewLine & ",'" & (txt号别.Text) & "','" & cbo号类.Text & "',"
        strSQL = strSQL & vbNewLine & cbo科室.ItemData(cbo科室.ListIndex) & ","
        strSQL = strSQL & vbNewLine & cboItem.ItemData(cboItem.ListIndex) & ",'" & Trim(cboDoctor.Text) & "',"
        strSQL = strSQL & vbNewLine & lng医生ID & "," & chk病案.Value & ","
        strSQL = strSQL & vbNewLine & str时间段 & str限号 & intCount & ","
        strSQL = strSQL & vbNewLine & "'" & str诊室 & "'," & strBegin & "," & strEnd & ",0," & chk序号控制.Value & ","
        strSQL = strSQL & vbNewLine & 0 & ","
        strSQL = strSQL & vbNewLine & 5 & ")"


    End If

    On Error GoTo errH
    zlAddArray cllPro, strSQL

   ' Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    LoadTimePlan True

    mfrmTime.zlSaveData lngNextID, cllPro
    zlExecuteProcedureArrAy cllPro, Me.Caption
    On Error GoTo 0
    mblnSucces = True

    If mEditType <> edt_新增 Then Unload Me: Exit Sub
    Call LoadCard
    mblnChangeByCode = True
    tbPage.Item(mPageIndex.EM_安排).Selected = True
    mblnChangeByCode = False
    Call mfrmTime.ClearCustomData
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadRegHistory() As Boolean
    Dim strSQL As String
    strSQL = " Select Decode(To_Char(a.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',"
    strSQL = strSQL & vbCrLf & "                       '7', '周六') As 限制项目, Max(Nvl(a.号序, 0)) As 最大序号, Count(1) As 统计,to_char(Max(发生时间),'hh24:mi:ss') as 发生时间"
    strSQL = strSQL & vbCrLf & " From 病人挂号记录 a, 挂号安排 b"
    strSQL = strSQL & vbCrLf & " Where a.记录状态 = 1 And a.发生时间 Between Sysdate And Sysdate + " & IIf(gint预约天数 = 0, 15, gint预约天数) & " And a.号别 = b.号码 And b.Id=[1]"
    strSQL = strSQL & vbCrLf & " Group By Decode(To_Char(a.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',"
    strSQL = strSQL & vbCrLf & "                             '7', '周六')"

    On Error GoTo Hd:
    Set mrsRegHistory = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    LoadRegHistory = True
Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
