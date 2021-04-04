VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmResistPlanTimeSet 
   BorderStyle     =   0  'None
   Caption         =   "分时段设置"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdOther 
      Caption         =   "其他辅助计算(&T)"
      Height          =   350
      Left            =   3840
      TabIndex        =   14
      ToolTipText     =   "点击重新计算时段"
      Top             =   0
      Width           =   1515
   End
   Begin VB.Frame fra应用于 
      Caption         =   "应用于…"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   7320
      Width           =   7755
      Begin VB.OptionButton opt应用于 
         Caption         =   "本医生(张三)"
         Height          =   255
         Index           =   1
         Left            =   2115
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "本号码"
         Height          =   255
         Index           =   0
         Left            =   795
         TabIndex        =   12
         Top             =   255
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton opt科室 
         Caption         =   "本科室(内科)"
         Height          =   255
         Left            =   3870
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt所有 
         Caption         =   "所有号别"
         Height          =   255
         Left            =   5685
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3540
      Index           =   0
      Left            =   795
      ScaleHeight     =   3540
      ScaleWidth      =   2535
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   4875
      Left            =   525
      TabIndex        =   6
      Top             =   2010
      Width           =   2535
      _Version        =   589884
      _ExtentX        =   4471
      _ExtentY        =   8599
      _StockProps     =   64
   End
   Begin VB.CommandButton cmd设置时段 
      Caption         =   "辅助计算(&F)"
      Height          =   350
      Left            =   2385
      TabIndex        =   0
      ToolTipText     =   "点击重新计算时段"
      Top             =   0
      Width           =   1150
   End
   Begin MSComCtl2.UpDown udTime 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   30
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtTimeOut"
      BuddyDispid     =   196618
      OrigLeft        =   2025
      OrigTop         =   3
      OrigRight       =   2280
      OrigBottom      =   348
      Max             =   1440
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTime 
      Height          =   7545
      Index           =   0
      Left            =   3465
      TabIndex        =   2
      Top             =   900
      Width           =   5100
      _cx             =   8996
      _cy             =   13309
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmResistPlanTimeSet.frx":0000
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
      Begin VB.CommandButton cmd删除 
         Caption         =   "删"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd预约 
         Caption         =   "预"
         Height          =   255
         Index           =   0
         Left            =   2685
         TabIndex        =   3
         Top             =   2535
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   300
      Left            =   1335
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "10"
      Top             =   30
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "时间间隔(分)"
      Height          =   180
      Left            =   225
      TabIndex        =   4
      Top             =   85
      Width           =   1080
   End
End
Attribute VB_Name = "frmResistPlanTimeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit '要求变量声明
'Public Enum gPlanEditType
'    EM_安排_增加 = 0
'    EM_安排_修改
'    EM_安排_查阅
'    EM_计划_增加 = 11
'    EM_计划_修改
'    EM_计划_查阅
'End Enum
Private mEditType         As gPlanEditType
Private mstr安排 As String  '周一,限号数,限约数|周二,限号数,限约数|....
Private mbln序号控制 As Boolean
Private mlngSelIndex As Long '正处于编辑状态的索引
Private mblnOnChange As Boolean  '是否代码触发tbpage的SelectedChanged事件
Private mlng安排ID As Long '
Private mlng计划Id As Long '
Private mblnInit As Boolean  '是否进行了初始化的调用
'Private mrsTime          As ADODB.Recordset
Private mrs限号          As ADODB.Recordset
Private mrs上班时间段    As ADODB.Recordset
Private mrs安排          As ADODB.Recordset
Private mrsRegPlan       As ADODB.Recordset ' 修改
Private mrsAssign        As ADODB.Recordset '已分配序号
Private mblnCellChange   As Boolean
Private mstrKey         As String
Public mblnChange       As Boolean '是否改变了内容
Private mblnReload      As Boolean '在挂号安排管理页面调用 ShowMe以后 是否需要刷新
Private mstr限制修改 As String '在某一天或者多天的安排限制更改
Private mrsHistory As ADODB.Recordset '预约挂号统计信息
Private WithEvents mfrmOtherCalc As frmRegistPlanTimeOther  '
Attribute mfrmOtherCalc.VB_VarHelpID = -1
'对外上班时间
Private Type t_上班时间
  dat_上午上班 As Date
  dat_上午下班 As Date
  dat_下午上班 As Date
  dat_下午下班 As Date
End Type
Private t_时间 As t_上班时间
Private Const strMaskKey As String = "09:00-09:00"
Private mstr应诊时段 As String
Public Event zlSaveTimePageSelected(ByVal str星期 As String)
Private mblnNotBrush As Boolean '不是进行刷新操作

Private Sub LoadPageControl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载页控件
    '编制:刘兴洪
    '日期:2012-06-15 13:33:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    For i = 1 To 6
        Load picPage(i): Load vsTime(i)
        Load cmd预约(i): Load cmd删除(i)
       ' cmd预约(i).Visible = True
        Set cmd预约(i).Container = vsTime(i)
        Set cmd删除(i).Container = vsTime(i)
        'cmd删除(i).Visible = True
        picPage(i).Visible = True: vsTime(i).Visible = True
        Set vsTime(i).Container = picPage(i)
    Next
    Set vsTime(0).Container = picPage(0)
    Call LoadPage
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : ClearCustomData
'Description : 清空变量信息
'Author      : 李光福
'Date        : 05-November-2012 14:58:54
'Input       :
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Public Sub ClearCustomData()
     mstr安排 = ""
     mbln序号控制 = False
     mlngSelIndex = 0
     mblnOnChange = False
     mlng安排ID = 0
     mlng计划Id = 0
     mblnInit = False
     Set mrs限号 = Nothing
     Set mrsRegPlan = Nothing
     Set mrsAssign = Nothing
     mstrKey = ""
     mblnChange = False
     mstr限制修改 = ""
     Set mrsHistory = Nothing
End Sub

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : LoadPage
'Description : 加载页面信息
'Author      : 李光福
'Date        : 05-November-2012 14:59:21
'Input       :
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-

Private Function LoadPage() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载页
    '编制:刘兴洪
    '日期:2012-06-15 13:37:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Dim strTemp As String
    On Error GoTo errHandle
    
    tbPage.RemoveAll
    For i = 0 To 6
        strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
        Set ObjItem = tbPage.InsertItem(i + 1, strTemp, picPage(i).hwnd, 0)
        ObjItem.Tag = strTemp
    Next
     With tbPage
         
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionBottom
    End With
    LoadPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
 Private Sub ShowPage()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示页面
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-26 15:21:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, i As Long
    Dim j As Long, lngIndex As Long, p As Long, strTemp As String
    
    For j = 0 To tbPage.ItemCount - 1
         tbPage(j).Visible = False: tbPage(j).Enabled = False
         tbPage(j).Selected = False
    Next
    
    On Error GoTo errHandle
    varData = Split(mstr安排, "|")
    lngIndex = -1: mlngSelIndex = -1
    For i = 0 To UBound(varData)
        ''周一,限号数,限约数|周二,限号数,限约数|....
        varTemp = Split(varData(i) & ",,,,", ",")
        If varTemp(0) <> "" Then
            For j = 0 To tbPage.ItemCount - 1
                If tbPage(j).Tag = varTemp(0) Then
                    If lngIndex < 0 Then lngIndex = j
                    tbPage(j).Visible = True: tbPage(j).Enabled = True
                    p = GetVsGridIndex(varTemp(0))
                    vsTime(p).Tag = varTemp(1) & "," & varTemp(2)
                    If mlngSelIndex = -1 Then mlngSelIndex = j: tbPage(j).Selected = True
                End If
            Next
        End If
    Next
    If mlngSelIndex = -1 Then mlngSelIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub
 
 Private Function GetVsGridIndex(ByVal str星期 As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取相关索引
    '编制:刘兴洪
    '日期:2012-06-15 14:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    str星期 = Switch(str星期 = "周日", 0, str星期 = "周一", 1, str星期 = "周二", 2, str星期 = "周三", 3, str星期 = "周四", 4, str星期 = "周五", 5, str星期 = "周六", 6, True, 0)
    GetVsGridIndex = Val(str星期)
 End Function
 
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : GetVsGridCaption
'Description : 根据索引获取限制项目
'Author      : 李光福
'Date        : 05-November-2012 15:02:14
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'nIndex            Integer           ByVal                .索引值
'Output      :     对应的星期
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
 
 Private Function GetVsGridCaption(ByVal nIndex As Integer) As String
    '功能:根据索引获取限制项目
    Dim str星期 As String
    str星期 = Switch(nIndex = 0, "周日", nIndex = 1, "周一", nIndex = 2, "周二", nIndex = 3, "周三", nIndex = 4, "周四", nIndex = 5, "周五", nIndex = 6, "周六", True, "")
    GetVsGridCaption = str星期
 End Function
 
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : zlShowPagePlan
'Description : 显示页面
'Author      : 李光福
'Date        : 05-November-2012 15:03:02
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'str安排类别       String            ByVal                .
'rsRegPlan         ADODB.Recor...    ByVal                .
'rsHistory         ADODB.Recor...    ByRef                .
'bln序号控制       Boolean           ByVal                .
'bytType           gPlanEditType     ByVal                .
'lng安排ID         Long              ByVal                .
'lng计划ID         Long              ByVal                .
'blnBeforCheck     Boolean = F...    ByVal                .
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
 Public Sub zlShowPagePlan(ByVal str安排类别 As String, ByVal rsRegPlan As ADODB.Recordset, ByRef rsHistory As ADODB.Recordset, _
                        ByVal bln序号控制 As Boolean, ByVal BytType As gPlanEditType, Optional ByVal lng安排ID As Long, _
                        Optional ByVal lng计划ID As Long, Optional ByVal blnBeforCheck As Boolean = False, Optional ByVal str应诊时段 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示页面
    '编制:刘兴洪
    '日期:2012-06-15 13:49:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    mstr安排 = str安排类别
    mstr应诊时段 = str应诊时段
                                                                          
    Set mrsRegPlan = rsRegPlan
    If bln序号控制 <> mbln序号控制 And Not mrsAssign Is Nothing Then
         mrsAssign.Filter = 0
         Do While Not mrsAssign.EOF
            mrsAssign.Delete
            mrsAssign.MoveNext
         Loop
         If blnBeforCheck Then Exit Sub
    End If
'    mlngSelIndex = -1
     mEditType = BytType: mlng安排ID = lng安排ID: mlng计划Id = lng计划ID
    Set mrsHistory = rsHistory
    If Not blnBeforCheck Then Call ShowPage
    If mblnInit Then
        Call AssignManage
    End If
    mblnInit = True
    Call InitRs(mbln序号控制 = bln序号控制)
    mbln序号控制 = bln序号控制
    If blnBeforCheck Then Exit Sub
    For i = 0 To 6
       If tbPage.Item(i).Selected Then
            Call tbPage_SelectedChanged(tbPage.Item(i))
            Exit For
       End If
    Next
 End Sub
 
 
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : SavePlanData
'Description : 安排时段数据保存
'Author      : 李光福
'Date        : 05-November-2012 15:04:05
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'lngID             Long              ByVal                .安排ID
'cllPro            Collection        ByRef                .返回相关保存数据的SQL
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
 Private Function SavePlanData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
    Dim i As Long, str星期 As String, lng序号 As String, strSQL As String
    Dim str序号s As String, BytType As Byte '应用于
    Dim bytRowStep As Byte, bytStepCol As Byte
    Dim intPage As Integer, cllPage As Collection
    Dim str时段 As String
    Dim strProc As String
    Dim strTmp As String
    Dim strTemp As String
    Dim p As Integer, j As Long
   
    On Error GoTo errHandle
    
    Call AssignManage  '序号分配处理
    If cllPro Is Nothing Then
        Set cllPro = New Collection
    End If
    strSQL = "Zl_挂号计划时段_Delete(" & lngID & ")"
    zlAddArray cllPro, strSQL
    For i = 0 To 6
        strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
        mrsAssign.Filter = "限制项目='" & strTemp & "'"
        If mrsAssign.RecordCount > 0 Then
            Do While Not mrsAssign.EOF
    '            序号,开始时间,结束时间,限制数量,预约标志|...
                strTmp = mrsAssign!序号
                strTmp = strTmp & "," & mrsAssign!开始时间 & "," & mrsAssign!结束时间 & "," & mrsAssign!限制数量 & "," & mrsAssign!是否预约
                If Len(str时段 & "|" & strTmp) > 4000 Then
                    str时段 = Mid(str时段, 2)
                    strSQL = "  Zl_挂号计划时段_Insert("
                    '  安排id_In 挂号安排时段.安排id%Type,
                    strSQL = strSQL & lngID & ","
                    '  星期_In   挂号安排时段.星期%Type,
                    strSQL = strSQL & "'" & strTemp & "',"
                    '  时段_In   Varchar2,
                    strSQL = strSQL & "'" & str时段 & "'"
                    strSQL = strSQL & "" & ")"
                    zlAddArray cllPro, strSQL
                    str时段 = ""
                End If
                str时段 = str时段 & "|" & strTmp
                mrsAssign.MoveNext
            Loop
            If str时段 <> "" Then
                 
                str时段 = Mid(str时段, 2)
                strSQL = "  Zl_挂号计划时段_Insert("
                '  安排id_In 挂号安排时段.安排id%Type,
                strSQL = strSQL & lngID & ","
                '  星期_In   挂号安排时段.星期%Type,
                strSQL = strSQL & "'" & strTemp & "',"
                '  时段_In   Varchar2,
                strSQL = strSQL & "'" & str时段 & "'"
                strSQL = strSQL & "" & ")"
                zlAddArray cllPro, strSQL
                str时段 = ""
            End If
        
        End If
    Next
    SavePlanData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

 End Function

Private Function SaveData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:安排时段数据保存
    '入参:lngID-安排ID
    '出参:cllPro-返回相关保存数据的SQL
    '编制:刘兴洪
    '日期:2012-06-15 13:18:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str星期 As String, lng序号 As String, strSQL As String
    Dim str序号s As String, BytType As Byte '应用于
    Dim bytRowStep As Byte, bytStepCol As Byte
    Dim intPage As Integer
    Dim str时段 As String
    Dim strProc As String
    Dim strTmp As String
    Dim strTemp As String
    Dim p As Integer, j As Long
     
   
    On Error GoTo errHandle
      
    Call AssignManage  '序号分配处理
    If cllPro Is Nothing Then
        Set cllPro = New Collection
    End If
    strSQL = "Zl_挂号安排时段_Delete(" & lngID & ")"
    zlAddArray cllPro, strSQL
    For i = 0 To 6
        strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
        mrsAssign.Filter = "限制项目='" & strTemp & "'"
        If mrsAssign.RecordCount > 0 Then
            Do While Not mrsAssign.EOF
    '            序号,开始时间,结束时间,限制数量,预约标志|...
                strTmp = mrsAssign!序号
                strTmp = strTmp & "," & mrsAssign!开始时间 & "," & mrsAssign!结束时间 & "," & mrsAssign!限制数量 & "," & mrsAssign!是否预约
                If Len(str时段 & "|" & strTmp) > 4000 Then
                    str时段 = Mid(str时段, 2)
                    strSQL = "  Zl_挂号安排时段_Insert("
                    '  安排id_In 挂号安排时段.安排id%Type,
                    strSQL = strSQL & lngID & ","
                    '  星期_In   挂号安排时段.星期%Type,
                    strSQL = strSQL & "'" & strTemp & "',"
                    '  时段_In   Varchar2,
                    strSQL = strSQL & "'" & str时段 & "'"
                    strSQL = strSQL & "" & ")"
                    zlAddArray cllPro, strSQL
                    str时段 = ""
                End If
                str时段 = str时段 & "|" & strTmp
                mrsAssign.MoveNext
            Loop
            If str时段 <> "" Then
                 
                str时段 = Mid(str时段, 2)
                strSQL = "  Zl_挂号安排时段_Insert("
                '  安排id_In 挂号安排时段.安排id%Type,
                strSQL = strSQL & lngID & ","
                '  星期_In   挂号安排时段.星期%Type,
                strSQL = strSQL & "'" & strTemp & "',"
                '  时段_In   Varchar2,
                strSQL = strSQL & "'" & str时段 & "'"
                strSQL = strSQL & "" & ")"
                zlAddArray cllPro, strSQL
                str时段 = ""
            End If
        
        End If
    Next
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Public Function zlSaveData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '入参:
    '出参:cllPro-返回相关保存数据的SQL
    '编制:刘兴洪
    '日期:2012-06-15 13:18:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mblnInit Then Exit Function
    If zl_CheckMoveAssign() = False Then Exit Function
    If VsTimeValidate(-1) = False Then Exit Function
    
    If mEditType = EM_安排_修改 Or mEditType = EM_安排_增加 Then
        If SaveData(lngID, cllPro) = False Then Exit Function
    Else
        If SavePlanData(lngID, cllPro) = False Then Exit Function
    End If
    
    zlSaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
'
'
'
'Private Sub cmd设置时段_Click()
''对挂号安排时段进行设置
'    Dim str星期         As String
'
'    cmd预约.Visible = False
'    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
'    str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
'    mrsTime.Filter = "星期='" & str星期 & "'"
'    If mrsTime.RecordCount > 0 Then
'      '****************************************************************
'      '在已有挂号安排时段的情况下
'      '提示操作员 是否需要重新计算时段
'      '****************************************************************
'        If MsgBox("此安排在" & str星期 & "已经存在时段 " & vbCrLf & "是否重新计算时段?", vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
'            mrsTime.Filter = 0
'            Exit Sub
'        End If
'    End If
'    Select Case chk序号控制.Value = 1
'    Case True:
'        Set专家号时段
'        setVsFlexBgColor (True)
'    Case False:
'        Set普通号时段
'        setVsFlexBgColor
'    End Select
'
'    mblnChange = True
'End Sub
'Private Sub Set普通号时段()
'    Dim strSQL      As String
'    Dim str星期     As String
'    Dim str时段     As String
'    Dim lng限号     As Long
'    Dim lng限约     As Long
'    Dim lng间隔     As Long
'    Dim dblDatCount As Long '总时间间隔
'    Dim dat时点     As Date '每个时间段的
'    Dim bln全天     As Boolean  '是否是全天都允许挂号 如果是全天则分为上午和下午
'    Dim datStart    As Date
'    Dim datEnd      As Date
'    Dim i           As Long
'    Dim j           As Long
'    Dim lngRow      As Long
'    Dim lngCol      As Long
'    Dim strData     As String
'    Dim strTime     As String
'    Dim strList()   As String
'    Dim blnExit     As Boolean
'    Dim lngIndex    As Long
'    Dim lngStart    As Long
'    On Error GoTo Hd
'    If mrs上班时间段 Is Nothing Then Exit Sub
'    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
'    str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
'    mrs限号.Filter = "星期='" & str星期 & "'"
'    If mrs限号.RecordCount = 0 Then
'        MsgBox "当前号别在" & str星期 & ",没有对应的挂号安排限制" & vbCrLf & "请到挂号安排中设置!", vbOKOnly, Me.Caption
'        Exit Sub '如果挂号安排中没有设置此天的信息 就不允许设置
'    End If
'    lng限号 = Nvl(mrs限号!限号数, 0): lng限约 = Nvl(mrs限号!限约数, 0)
'    If lng限号 = 0 Then
'        MsgBox "当前号别在" & str星期 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
'        Exit Sub
'    End If
'    Me.txt限号.Text = lng限号
'    Me.txt限约.Text = lng限约
'    If lng限约 = 0 Then lng限约 = lng限号 '如果对预约没有限制则认为最大限约数和限号数相同
'    str时段 = Nvl(mrs安排(str星期).Value)
'    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
'
'    '*********************************
'    '分时段具体处理 分为全天和非全天
'    '全天分为上午和下午
'    '*********************************
'
'    lng间隔 = Val(txtTimeOut.Text)
'
'    With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 400: .RowHeightMin = 400
'        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'    End With
'    '*************************************
'    '普通号
'    '*************************************
'    With vsTime
'        .Cols = 8: .FixedCols = 0
'        .Rows = 1: .FixedRows = 1
'        For i = 0 To .Cols - 1 Step 2
'           .TextMatrix(0, i) = "时间段"
'        Next
'        For i = 1 To .Cols - 1 Step 2
'           .TextMatrix(0, i) = "预约人数"
'        Next
'        lngRow = 1: lngCol = -1
'        j = 1: lngStart = 1
'        Do While Not mrs上班时间段.EOF
'            If blnExit Then Exit Do
'            dat时点 = CDate(Nvl(mrs上班时间段!上班, "00:00:00"))
'            For i = j To lng限号
'                If lngStart > lng限号 Then
'                    blnExit = True
'                    Exit For
'                End If
'
'                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
'                    j = i
'                    Exit For
'                End If
'
'                lngCol = lngCol + 1
'                If lngCol * 2 > .Cols - 2 Then lngRow = lngRow + 1: lngCol = 0
'                strData = IIf(lng限约 >= i, 1, 0)
'                strTime = Format(dat时点, "HH:mm") & "-" & _
'                      IIf(Format(DateAdd("n", lng间隔, dat时点), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
'                      Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng间隔, dat时点), "HH:mm"))
'
'                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(lngRow, lngCol * 2) = strTime
'                .TextMatrix(lngRow, lngCol * 2 + 1) = strData
'                lngStart = lngStart + 1
'                dat时点 = DateAdd("n", lng间隔, dat时点)
'            Next
'            mrs上班时间段.MoveNext
'        Loop
'
'
'         For i = 0 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'         Next
'         .Redraw = flexRDBuffered
'    End With
'
'Exit Sub
'Hd:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Sub
'Private Sub Set专家号时段()
'    Dim strSQL      As String
'    Dim str星期     As String
'    Dim str时段     As String
'    Dim lng限号     As Long
'    Dim lng限约     As Long
'    Dim lng间隔     As Long
'    Dim dblDatCount As Long '总时间间隔
'    Dim dat时点     As Date '每个时间段的
'    Dim str时点     As String
'    Dim bln全天     As Boolean  '是否是全天都允许挂号 如果是全天则分为上午和下午
'    Dim datStart    As Date
'    Dim datEnd      As Date
'    Dim i           As Long
'    Dim j           As Long
'    Dim lngRow      As Long
'    Dim lngCol      As Long
'    Dim strData     As String
'    Dim strTime     As String
'    Dim strList()   As String
'    Dim blnExit     As Boolean
'    Dim lngIndex    As Long
'    Dim lngStart    As Long
'    On Error GoTo Hd
'    If mrs上班时间段 Is Nothing Then Exit Sub
'    If mrs限号 Is Nothing Then
'        strSQL = _
'        "Select 安排id, 限制项目 as 星期 , 限号数, 限约数 From 挂号安排限制 Where 安排id = [1]"
'        Set mrs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(txt号别.Tag))
'        If mrsTime.RecordCount = 0 Then
'        MsgBox "当前号别没有对应的挂号安排限制" & vbCrLf & "请到挂号安排中设置!", vbOKOnly, Me.Caption
'        Set mrs限号 = Nothing
'        Exit Sub '如果挂号安排中没有设置此天的信息 就不允许设置
'    End If
'    End If
'    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
'    str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
'    mrs限号.Filter = "星期='" & str星期 & "'"
'    If mrs限号.RecordCount = 0 Then
'        MsgBox "当前号别在" & str星期 & ",没有对应的挂号安排限制" & vbCrLf & "请到挂号安排中设置!", vbOKOnly, Me.Caption
'        Exit Sub '如果挂号安排中没有设置此天的信息 就不允许设置
'    End If
'    lng限号 = Nvl(mrs限号!限号数, 0): lng限约 = Nvl(mrs限号!限约数, 0)
'    If lng限号 = 0 Then
'        MsgBox "当前号别在" & str星期 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
'        Exit Sub
'    End If
'    Me.txt限号.Text = lng限号
'    Me.txt限约.Text = lng限约
'    lng限约 = lng限号
'    str时段 = Nvl(mrs安排(str星期).Value)
'    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
'
''*************************************************************
''时间间隔根据 设置的间隔
''*************************************************************
'      lng间隔 = Val(Me.txtTimeOut.Text)
'     ' dat时点 = CDate(Nvl(mrs上班时间段!开始时间, "00:00:00"))
'
'      With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 400: .RowHeightMin = 400
'        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'      End With
'    '*************************************
'    '专家号
'    '序号填充规则
'    '根据 时间段表中的 上下班时间来判断
'    '其中 全天这种情况  分为上午和下午
'    '*************************************
'
'    With vsTime
'         .Cols = 2
'         lngRow = -1: lngCol = 0
'         j = 1
'         lngStart = 1
'         Do While Not mrs上班时间段.EOF
'            If blnExit Then Exit Do
'
'            dat时点 = CDate(Nvl(mrs上班时间段!上班, "00:00:00"))
'             For i = j To lng限约
'                If lngStart > lng限约 Then
'                    blnExit = True
'                    Exit For
'                End If
'
'                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
'                    j = i
'                    Exit For
'                 End If
'                lngCol = lngCol + 1
'                If str时点 <> Format(dat时点, "HH") & ":00" Then lngRow = lngRow + 2: lngCol = 1
'                If lngCol = 1 Then
'                     If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'                     str时点 = Format(dat时点, "HH") & ":00"
'                     vsTime.TextMatrix(lngRow - 1, 0) = str时点
'                     vsTime.TextMatrix(lngRow, 0) = str时点
'
'                End If
'                strData = lngStart
'                lngStart = lngStart + 1
'                strTime = Format(dat时点, "HH:mm") & "-" & _
'                           IIf(Format(DateAdd("n", lng间隔, dat时点), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
'                           Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng间隔, dat时点), "HH:mm"))
'
'                If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
'                vsTime.TextMatrix(lngRow - 1, lngCol) = strData
'                vsTime.TextMatrix(lngRow, lngCol) = strTime
'                '是第一项时 填写 开始时间到首行
'
'                dat时点 = DateAdd("n", lng间隔, dat时点)
'             Next
'             mrs上班时间段.MoveNext
'         Loop
''         '***********************
''         '序号填充
''         '**********************
''         For i = 1 To lng限约
''            If Format(dat时点, "dd:mm:ss") >= Format(CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), "dd:mm:ss") Then Exit For
''            lngCol = lngCol + 1
''            If str时点 <> Format(dat时点, "HH") & ":00" Then lngRow = lngRow + 2: lngCol = 1
''            If lngCol = 1 Then
''                 If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
''                 str时点 = Format(dat时点, "HH") & ":00"
''                 vsTime.TextMatrix(lngRow - 1, 0) = str时点
''                 vsTime.TextMatrix(lngRow, 0) = str时点
''
''            End If
''            strData = i
''            strTime = Format(dat时点, "HH:mm") & "-" & _
''                       IIf(DateAdd("n", lng间隔, dat时点) > CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), _
''                       Format(CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng间隔, dat时点), "HH:mm"))
''
''            If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
''            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
''            vsTime.TextMatrix(lngRow, lngCol) = strTime
''            '是第一项时 填写 开始时间到首行
''
''            dat时点 = DateAdd("n", lng间隔, dat时点)
''         Next
''         If bln全天 Then
''             mrs上班时间段.Filter = "时间段='下午'"
''            dat时点 = CDate(Nvl(mrs上班时间段!开始时间, "00:00:00"))
''         End If
''         j = i
''         For i = j To lng限约
''            If Format(dat时点, "dd:mm:ss") >= CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")) Then Exit For
''            lngCol = lngCol + 1
''            If lngCol > vsTime.Cols - 1 Then lngRow = lngRow + 2: lngCol = 1
''            strData = i
''            strTime = Format(dat时点, "HH:mm") & "-" & _
''                       IIf(DateAdd("n", lng间隔, dat时点) > CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), _
''                       Format(CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng间隔, dat时点), "HH:mm"))
''            If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
''            If lngRow < 0 Then vsTime.Rows = vsTime.Rows + 2: lngRow = lngRow + 2
''            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
''            vsTime.TextMatrix(lngRow, lngCol) = strTime
''
''            '是第一项时 填写 开始时间到首行
''            If lngCol = 1 Then
''                 vsTime.TextMatrix(lngRow - 1, 0) = Format(dat时点, "HH:mm")
''                 vsTime.TextMatrix(lngRow, 0) = Format(dat时点, "HH:mm")
''            End If
''            dat时点 = DateAdd("n", lng间隔, dat时点)
''         Next
'         For i = 1 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'         Next
'         .ColWidth(0) = 1200
'         .FixedAlignment(0) = flexAlignRightTop
'         .ColAlignment(0) = flexAlignRightTop
'         If .Rows > 0 Then
'            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
'            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
'         End If
'         .Redraw = flexRDBuffered
'    End With
'
'Exit Sub
'Hd:
'    If ErrCenter() = 1 Then
'         Resume
'    End If
'    SaveErrLog
'End Sub
'
'Private Sub cmd预约_Click()
'    '对时间段能否预约进行设置
'    If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Then Exit Sub
'    If mViewMode = ViewMode.ViewItem Or vsTime.TextMatrix(vsTime.MouseRow, vsTime.MouseCol) = "" Then Exit Sub
'    With vsTime
'        If .CellForeColor = vbBlue Then
'            .Cell(flexcpForeColor, .Row, .Col, .Row + 1, .Col) = &H80000008
'            .Cell(flexcpFontBold, .Row, .Col, .Row + 1, .Col) = False
'         Else
'            .Cell(flexcpForeColor, .Row, .Col, .Row + 1, .Col) = vbBlue
'            .Cell(flexcpFontBold, .Row, .Col, .Row + 1, .Col) = True
'        End If
'    End With
'    mblnChange = True
'End Sub
'
'Private Sub Form_Activate()
'    Me.Icon = frmRegistPlan.Icon
'End Sub
'
'Private Sub Form_Load()
'    Init时间段
'End Sub
'
'Private Sub Form_Resize()
'  On Error Resume Next
'  '********************************************
'  '首先设置 窗体的最小宽度和最小高度
'  '********************************************
'  If Me.Width < 701 * Screen.TwipsPerPixelX Then Me.Width = 701 * Screen.TwipsPerPixelX
'  If Me.Height < 511 * Screen.TwipsPerPixelY Then Me.Height = 511 * Screen.TwipsPerPixelY
'  '********************************************
'  '挂号安排基本信息 位置不移动移动
'  '仅移动 时段设置
'  '********************************************
'  With fraDate
'     .Width = Me.ScaleWidth - 2 * .Left
'     .Height = Me.ScaleHeight - Me.fraInfo.Top - Me.fraInfo.Height - 65 * Screen.TwipsPerPixelY
'  End With
'
'  With picTime
'     .Width = fraDate.Width - 2 * .Left
'     .Height = fraDate.Height - .Top * 2
'  End With
'  With Me.tbWeekTime
'    .Width = picTime.ScaleWidth - 2 * .Left
'  End With
'  With Me.vsTime
'    .Width = picTime.ScaleWidth - 2 * .Left
'    .Height = picTime.ScaleHeight - .Top - cmd设置时段.Top
'  End With
'  '-------------------------------------------
'  '应用于 位置的调整
'  '-------------------------------------------
'  With Me.fra应用于
'       .Left = .Left
'       .Top = Me.fraDate.Top + Me.fraDate.Height + 5 * Screen.TwipsPerPixelY
'
'  End With
'
'  '********************************************
'  '确定按钮和取消按钮的移动
'  '********************************************
'
'  With Me.cmdCancel
'       .Left = Me.ScaleWidth - 40 * Screen.TwipsPerPixelX - .Width
'       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
'  End With
'  With Me.cmdOK
'       .Left = cmdCancel.Left - 20 * Screen.TwipsPerPixelX - .Width
'       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
'  End With
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'     mlngPre安排ID = -1
'     mblnChange = False
'     Set mrsTime = Nothing
'     mstr限制修改 = ""
'     Set mrs限号 = Nothing
'     Set mrs上班时间段 = Nothing
'     Set mrs安排 = Nothing
'End Sub
'
'
'
'
'
'Private Sub tbWeekTime_Click()
'    Dim i       As Integer
'    If mblnChange Then
'        mblnChange = False
'        If MsgBox("当前挂号安排在" & mstrKey & "的时段已改变!是否保存?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption) = vbYes Then
'            cmdOK_Click
'         For i = 1 To tbWeekTime.Tabs.Count
'            If tbWeekTime.Tabs(i).Key = "K" & mstrKey Then
'                tbWeekTime.Tabs(i).Selected = True
'                Exit For
'            End If
'         Next
'        End If
'    End If
'    mstrKey = Mid(tbWeekTime.SelectedItem.Key, 2)
'     If mstr限制修改 <> "" Then
'        vsTime.Editable = flexEDKbdMouse: cmd设置时段.Enabled = True
'        If InStr(mstr限制修改, ";" & mstrKey & ";") > 0 Then vsTime.Editable = flexEDNone: cmd设置时段.Enabled = False
'    End If
'    Select Case mViewMode
'        Case ViewMode.ViewItem:
'             Call LoadTimePlan(mlng安排ID, Me.chk序号控制.Value = 1)
'        Case ViewMode.Edit:
'            cmd预约.Visible = False
'            Call LoadEditTimePlan(mlng安排ID, Me.chk序号控制.Value = 1)
'    End Select
'     setVsFlexBgColor (Me.chk序号控制.Value = 1)
'End Sub
'
'
'
'
'Private Sub txtTimeOut_KeyPress(KeyAscii As Integer)
'
'    '限制非数字输入
'    If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
'    If txtTimeOut.Text = "" And KeyAscii = Asc(0) Then KeyAscii = 0
'End Sub
'
'Private Sub txtTimeOut_Validate(Cancel As Boolean)
'    If Val(txtTimeOut.Text) < 1 Then Cancel = True
'End Sub
'
'
'
'Private Sub udTime_DownClick()
'    If Val(txtTimeOut.Text) < 2 Then Exit Sub
'    txtTimeOut.Text = Val(txtTimeOut.Text) - 1
'End Sub
'
'Private Sub udTime_UpClick()
'  txtTimeOut.Text = Val(txtTimeOut.Text) + 1
'End Sub
'
'
'
'
''Private Sub vsTime_Click()
''  Select Case mViewMode
''    Case ViewMode.Edit, ViewMode.NewItem:
''       If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Or (chk序号控制.Value = 0 And vsTime.MouseRow < 1) Then Exit Sub
''       Select Case chk序号控制.Value = 1
''            Case True:
''            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
''            Case False:
''            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
''       End Select
''        If vsTime.MouseRow < 0 Or vsTime.MouseCol < 1 Then Exit Sub
''
''        If chk序号控制.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
''            cmd预约.Left = vsTime.MouseCol * 1200 + 20
''            cmd预约.Top = vsTime.MouseRow * 400 + 20
''            cmd预约.Visible = True
''        End If
''
''    Case ViewMode.ViewItem:
''         vsTime.Editable = flexEDNone
''  End Select
''End Sub
'
'Public Function ShowMe(lng安排ID As Long, mode As ViewMode) As Boolean
'    mViewMode = mode: mlng安排ID = lng安排ID
'    If InitData() = False Then
'        '加载挂号安排基本信息
'         Exit Function
'    End If
'    Select Case mViewMode
'         Case ViewMode.ViewItem:
'                vsTime.Editable = flexEDNone
'                Me.txtTimeOut.Enabled = False
'                Me.cmd设置时段.Enabled = False
'               '查看
'              Call LoadTimePlan(mlng安排ID, chk序号控制.Value = 1, False)
'         Case ViewMode.Edit
'              If LoadEditTimePlan(mlng安排ID, chk序号控制.Value = 1, False) = False Then
'               Exit Function
'              End If
'    End Select
'    setVsFlexBgColor (chk序号控制.Value = 1)
'    Me.Show 1
'    ShowMe = mblnReload
'End Function
''------------------------------------------------------------------------
''页面调用过程与方法
''------------------------------------------------------------------------
'Public Function InitData() As Boolean
'    Dim strSQL          As String
'    Dim lng安排ID       As Long
'    If mlng安排ID = -1 Then Exit Function
'     lng安排ID = mlng安排ID
'     On Error GoTo Hd
'     strSQL = " " & _
'        "   Select A.Id as 安排ID,0 as 计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id," & _
'        "          A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六,nvl(A.默认时段间隔,5) As 默认时段间隔, " & _
'        "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室 " & _
'        "   From 挂号安排 A,收费项目目录 B,部门表 D " & _
'        "   Where A.项目id=b.Id(+) And A.科室id =d.Id(+) " & _
'        "         And A.Id=[1]"
'         Set mrs安排 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
'
'         If mrs安排.EOF Then
'              ShowMsgbox "未找到指定的号别,请检查!"
'             Exit Function
'        End If
'        strSQL = "Select 限制项目,限号数,  限约数,限制项目 as 星期 From  挂号安排限制 where 安排ID=[1]  Order BY 限制项目      "
'        Set mrs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
'        cbo号类.Text = Nvl(mrs安排!号类)
'        txt号别.Tag = Nvl(mrs安排!安排id)
'        txtTimeOut.Tag = Val(Nvl(mrs安排!默认时段间隔, 0))
'        txtTimeOut.Text = txtTimeOut.Tag
'        txt号别.Text = Nvl(mrs安排!号码)
'        cbo科室.Text = Nvl(mrs安排!科室)
'        cboItem.Text = Nvl(mrs安排!项目)
'        cboDoctor.Text = Nvl(mrs安排!医生姓名)
'        chk病案.Value = IIf(Val(Nvl(mrs安排!病案必须)) = 1, 1, 0)
'       chk序号控制.Value = IIf(Val(Nvl(mrs安排!序号控制)) = 1, 1, 0):  chk序号控制.Tag = chk序号控制.Value
'        strSQL = "" & _
'        "   Select decode(星期,'周日',1,'周一',2,'周二',3,'周三',4,'周四',5,'周五',6,7) as 排序,星期,to_char(开始时间,'HH24')||':00' as 时点,序号,to_char(开始时间,'hh24:mi')||'-' ||to_char(结束时间,'hh24:mi') as 时间范围, " & _
'        "               限制数量,是否预约" & _
'        "   From  挂号安排时段 " & _
'        "   Where 安排ID=[1]" & _
'        "   Order by 排序,时点,序号"
'        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
'        mstr限制修改 = Get已约限制(mlng安排ID)
'       InitData = True
'Exit Function
'Hd:
'     If ErrCenter() = 1 Then Resume
'     SaveErrLog
'End Function
'
'
'Private Function LoadEditTimePlan(ByVal lng安排ID As Long, ByVal bln序号控制 As Boolean, _
'    Optional bln计划 As Boolean = False) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:
'    '入参:
'    '编制:
'    '日期:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL           As String
'    Dim rsTemp           As ADODB.Recordset
'    Dim str星期          As String
'    Dim i                As Long
'    Dim r                As Integer
'    Dim lngRow           As Long
'    Dim lngCol           As Integer
'    Dim str时点          As String
'    Dim strTime          As String
'    Dim strData          As String
'    Dim strKey           As String
'
'    On Error GoTo errHandle
'    '加载该挂号项目的的停用时间信息
'    If mrsTime Is Nothing Then
'        mlngPre安排ID = -1
'    ElseIf mrsTime.State <> 1 Then
'         mlngPre安排ID = -1
'    End If
'    If mlngPre安排ID <> lng安排ID Then
'        mlngPre安排ID = lng安排ID
'        tbWeekTime.Tabs.Clear
'        With tbWeekTime
'            If Not mrs限号.EOF Then
'                mrs限号.Filter = "星期='周一'"
'                If mrs限号.RecordCount > 0 Then
'                '限号数,  限约数,限制项目
'                    If Nvl(mrs限号!限号数, 0) > 0 Then
'                        tbWeekTime.Tabs.Add , _
'                            "K周一", "周一" & IIf(Nvl(mrs安排!周一) = "", "", "(" & Nvl(mrs安排!周一) & ")")
'                    End If
'                End If
'                mrs限号.Filter = "星期='周二'"
'                If mrs限号.RecordCount > 0 Then
'                   If Nvl(mrs限号!限号数, 0) > 0 Then
'                    tbWeekTime.Tabs.Add , _
'                        "K周二", "周二" & IIf(Nvl(mrs安排!周二) = "", "", "(" & Nvl(mrs安排!周二) & ")")
'                    End If
'                End If
'                mrs限号.Filter = "星期='周三'"
'                If mrs限号.RecordCount > 0 Then
'                     If Nvl(mrs限号!限号数, 0) > 0 Then
'                    tbWeekTime.Tabs.Add , _
'                        "K周三", "周三" & IIf(Nvl(mrs安排!周三) = "", "", "(" & Nvl(mrs安排!周三) & ")")
'                    End If
'                 End If
'
'                mrs限号.Filter = "星期='周四'"
'                If mrs限号.RecordCount > 0 Then
'                  If Nvl(mrs限号!限号数, 0) > 0 Then
'                    tbWeekTime.Tabs.Add , _
'                      "K周四", "周四" & IIf(Nvl(mrs安排!周四) = "", "", "(" & Nvl(mrs安排!周四) & ")")
'                  End If
'                End If
'                mrs限号.Filter = "星期='周五'"
'                If mrs限号.RecordCount > 0 Then
'                     If Nvl(mrs限号!限号数, 0) > 0 Then
'                        tbWeekTime.Tabs.Add , _
'                            "K周五", "周五" & IIf(Nvl(mrs安排!周五) = "", "", "(" & Nvl(mrs安排!周五) & ")")
'                     End If
'                End If
'
'                mrs限号.Filter = "星期='周六'"
'                If mrs限号.RecordCount > 0 Then
'                   If Nvl(mrs限号!限号数, 0) > 0 Then
'                        tbWeekTime.Tabs.Add , _
'                          "K周六", "周六" & IIf(Nvl(mrs安排!周六) = "", "", "(" & Nvl(mrs安排!周六) & ")")
'                   End If
'                End If
'                mrs限号.Filter = "星期='周日'"
'                If mrs限号.RecordCount > 0 Then
'                    If Nvl(mrs限号!限号数, 0) > 0 Then
'                        tbWeekTime.Tabs.Add , _
'                            "K周日", "周日" & IIf(Nvl(mrs安排!周日) = "", "", "(" & Nvl(mrs安排!周日) & ")")
'                    End If
'                End If
'                mrs限号.Filter = 0
'            End If
'            .Visible = tbWeekTime.Tabs.Count <> 0
'            If .Tabs.Count > 0 Then
'                .Tabs(1).Selected = True
'            Else
'                MsgBox "该安排没有设置对应的限号数和限约数,请检查!", vbOKOnly, Me.Caption
'                Exit Function
'            End If
'
'        End With
'    End If
'    str星期 = "": strTime = ""
'    If Not tbWeekTime.SelectedItem Is Nothing Then
'        str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
'    End If
'    mrsTime.Filter = "星期='" & str星期 & "'"
'    mrs限号.Filter = "星期='" & str星期 & "'"
'    txt限号.Text = ""
'    txt限约.Text = ""
'    If mrs限号.RecordCount <> 0 Then
'        Me.txt限号.Text = Nvl(mrs限号!限号数, 0)
'        Me.txt限约.Text = Nvl(mrs限号!限约数, 0)
'    End If
'     str时点 = ""
'    With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 400: .RowHeightMin = 400
'        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'        If Not bln序号控制 Then
'             .Cols = 8: .FixedCols = 0
'             .Rows = 1: .FixedRows = 1
'             For i = 0 To .Cols - 1 Step 2
'                .TextMatrix(0, i) = "时间段"
'             Next
'             For i = 1 To .Cols - 1 Step 2
'                .TextMatrix(0, i) = "预约人数"
'             Next
'
'             r = 1: i = -1
'            Do While Not mrsTime.EOF
'                i = i + 1
'                If i * 2 > .Cols - 2 Then r = r + 1: i = 0
'                strData = Val(Nvl(mrsTime!限制数量))
'                strTime = mrsTime!时间范围
'                If r > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(r, i * 2) = strTime
'                .TextMatrix(r, i * 2 + 1) = strData
'                mrsTime.MoveNext
'            Loop
'            For i = 0 To .Cols - 1
'                .ColAlignment(i) = flexAlignCenterCenter
'                .ColWidth(i) = 1200
'            Next
'            .Redraw = flexRDBuffered
'            LoadEditTimePlan = True
'            Exit Function
'        End If
'        .Cols = 7: .FixedCols = 1
'        .Rows = 0: .FixedRows = 0
'        i = 1: r = -1
'        lngRow = -1: lngCol = 1
'        '******************************************
'        With vsTime
'         .Cols = 2
'         lngRow = -1: lngCol = 0
'         '***********************
'         '序号填充
'         '**********************
'         r = mrsTime.RecordCount
'         For i = 1 To r
'            If mrsTime.EOF Then Exit For
'            lngCol = lngCol + 1
'            If str时点 <> Nvl(mrsTime!时点) Then lngRow = lngRow + 2: lngCol = 1
'             If lngCol = 1 Then
'                str时点 = Nvl(mrsTime!时点)
'                If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'                vsTime.TextMatrix(lngRow - 1, 0) = str时点
'                vsTime.TextMatrix(lngRow, 0) = str时点
'             End If
'            strData = mrsTime!序号
'            strTime = mrsTime!时间范围
'            If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
'            'If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
'            vsTime.TextMatrix(lngRow, lngCol) = strTime
'            '是第一项时 填写 开始时间到首行
'            If lngCol = 1 Then
'            End If
'            If Val(Nvl(mrsTime!是否预约)) = 1 Then
'                .Cell(flexcpForeColor, lngRow - 1, lngCol, lngRow, lngCol) = vbBlue
'                .Cell(flexcpFontBold, lngRow - 1, lngCol, lngRow, lngCol) = True
'            End If
'            mrsTime.MoveNext
'         Next
'
'         End With
'        '******************************************
''        Do While Not mrsTime.EOF
''            If i = 1 Then
''                r = r + 2
''                str时点 = Nvl(mrsTime!时点)
''                If r > .Rows - 1 Then .Rows = .Rows + 2
''                .TextMatrix(r, 0) = str时点
''                .TextMatrix(r - 1, 0) = str时点
''            End If
''            i = i + 1
''            strData = mrsTime!序号
''            strTime = mrsTime!时间范围
''            If i >= .Cols - 1 Then i = 1
''            If r > .Rows - 1 Then .Rows = .Rows + 2
''            .TextMatrix(r, i) = strTime
''            .TextMatrix(r - 1, i) = strData
''
''        Loop
'
'
'        For i = 1 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'        Next
'        .ColWidth(0) = 1200
'        .FixedAlignment(0) = flexAlignRightTop
'        .ColAlignment(0) = flexAlignRightTop
'        If .Rows > 0 Then
'            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
'            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
'        End If
'        .MergeCellsFixed = flexMergeRestrictColumns
'        .MergeCol(0) = True
'        .Redraw = flexRDBuffered
'    End With
'    LoadEditTimePlan = True
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

'Private Sub LoadEditTimePlantext(ByVal lng安排ID As Long, ByVal bln序号控制 As Boolean, _
'    Optional bln计划 As Boolean = False)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:
'    '入参:
'    '编制:
'    '日期:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL           As String
'    Dim rsTemp           As ADODB.Recordset
'    Dim str星期          As String
'    Dim i                As Long
'    Dim r                As Integer
'    Dim str时点          As String
'    Dim strTime          As String
'    Dim strData          As String
'    Dim strKey           As String
'
'    On Error GoTo errHandle
'    '加载该挂号项目的的停用时间信息
'    If mrsTime Is Nothing Then
'        mlngPre安排ID = -1
'    ElseIf mrsTime.State <> 1 Then
'         mlngPre安排ID = -1
'    End If
'    If mlngPre安排ID <> lng安排ID Then
'        mlngPre安排ID = lng安排ID
'        tbWeekTime.Tabs.Clear
'        With mrsTime
'            strTime = ""
'            Do While Not .EOF
'                If strTime <> Nvl(mrsTime!星期) Then
'                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!星期), Nvl(mrsTime!星期)
'                    strTime = Nvl(mrsTime!星期)
'                End If
'                .MoveNext
'            Loop
'            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
'            If tbWeekTime.Tabs.Count > 0 Then
'                tbWeekTime.Tabs(1).Selected = True
'            End If
'            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
'        End With
'    End If
'    str星期 = "": strTime = ""
'    If Not tbWeekTime.SelectedItem Is Nothing Then
'        str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
'    End If
'    mrsTime.Filter = "星期='" & str星期 & "'"
'    mrs限号.Filter = "星期='" & str星期 & "'"
'    txt限号.Text = ""
'    txt限约.Text = ""
'    If mrs限号.RecordCount <> 0 Then
'        Me.txt限号.Text = Nvl(mrs限号!限号数, 0)
'        Me.txt限约.Text = Nvl(mrs限号!限约数, 0)
'    End If
'     str时点 = ""
'    With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 400: .RowHeightMin = 400
'        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'        If Not bln序号控制 Then
'             .Cols = 8: .FixedCols = 0
'             .Rows = 1: .FixedRows = 1
'             For i = 0 To .Cols - 1 Step 2
'                .TextMatrix(0, i) = "时间段"
'             Next
'             For i = 1 To .Cols - 1 Step 2
'                .TextMatrix(0, i) = "预约人数"
'             Next
'
'             r = 1: i = -1
'            Do While Not mrsTime.EOF
'                If i * 2 > .Cols - 2 Then r = r + 1: i = -1
'                i = i + 1
'                strData = Val(Nvl(mrsTime!限制数量))
'                strTime = mrsTime!时间范围
'                If r > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(r, i * 2) = strTime
'                .TextMatrix(r, i * 2 + 1) = strData
'                mrsTime.MoveNext
'            Loop
'            For i = 0 To .Cols - 1
'                .ColAlignment(i) = flexAlignCenterCenter
'                .ColWidth(i) = 1200
'            Next
'            .Redraw = flexRDBuffered
'             Exit Sub
'        End If
'        Do While Not mrsTime.EOF
'            If str时点 <> Nvl(mrsTime!时点) Then
'                r = r + 2
'                str时点 = Nvl(mrsTime!时点)
'                If r > .Rows - 1 Then .Rows = .Rows + 2
'                .TextMatrix(r, 0) = str时点
'                .TextMatrix(r - 1, 0) = str时点
'                i = 0
'            End If
'            i = i + 1
'            strData = mrsTime!序号
'            strTime = mrsTime!时间范围
'            If i > .Cols - 1 Then .Cols = .Cols + 1
'            If r > .Rows - 1 Then .Rows = .Rows + 1
'            .TextMatrix(r, i) = strTime
'            .TextMatrix(r - 1, i) = strData
'            If Val(Nvl(mrsTime!是否预约)) = 1 Then
'
'                .Cell(flexcpForeColor, r - 1, i, r, i) = vbBlue
'                .Cell(flexcpFontBold, r - 1, i, r, i) = True
'            End If
'            mrsTime.MoveNext
'        Loop
'        For i = 1 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'        Next
'        .ColWidth(0) = 1200
'        .FixedAlignment(0) = flexAlignRightTop
'        .ColAlignment(0) = flexAlignRightTop
'        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
'        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
'        .MergeCellsFixed = flexMergeRestrictColumns
'        .MergeCol(0) = True
'        .Redraw = flexRDBuffered
'    End With
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Sub
'
'
'Private Sub LoadTimePlan(ByVal lng安排ID As Long, ByVal bln序号控制 As Boolean, _
'    Optional bln计划 As Boolean = False)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:
'    '入参:
'    '编制:
'    '日期:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL           As String
'    Dim rsTemp           As ADODB.Recordset
'    Dim str星期          As String
'    Dim i                As Long
'    Dim r                As Integer
'    Dim str时点          As String
'    Dim strTime          As String
'    Dim strKey           As String
'    On Error GoTo errHandle
'    '加载该挂号项目的的停用时间信息
'    If mrsTime Is Nothing Then
'         mlngPre安排ID = -1
'    ElseIf mrsTime.State <> 1 Then
'         mlngPre安排ID = -1
'    End If
'    If mlngPre安排ID <> lng安排ID Then
'        mlngPre安排ID = lng安排ID
''        strSQL = "" & _
''        "   Select decode(星期,'周日',1,'周一',2,'周二',3,'周三',4,'周四',5,'周五',6,7) as 排序,星期,to_char(开始时间,'HH24')||':00' as 时点,序号,to_char(开始时间,'hh24:mi')||'-' ||to_char(结束时间,'hh24:mi') as 时间范围, " & _
''        "               限制数量,是否预约" & _
''        "   From  挂号安排时段 " & _
''        "   Where 安排ID=[1]" & _
''        "   Order by 排序,时点,序号"
''        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
'        tbWeekTime.Tabs.Clear
'        With mrsTime
'            strTime = ""
'            Do While Not .EOF
'                If strTime <> Nvl(mrsTime!星期) Then
'                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!星期), Nvl(mrsTime!星期)
'                    strTime = Nvl(mrsTime!星期)
'                End If
'                .MoveNext
'            Loop
'            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
'            If tbWeekTime.Tabs.Count > 0 Then
'                tbWeekTime.Tabs(1).Selected = True
'            End If
'            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
'        End With
'        If tbWeekTime.Tabs.Count = 0 Then
'            MsgBox "该安排没有设置对应的时段,请检查!"
'        End If
'    End If
'    str星期 = "": strTime = ""
'    If Not tbWeekTime.SelectedItem Is Nothing Then
'        str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
'    End If
'    mrsTime.Filter = "星期='" & str星期 & "'"
'    mrs限号.Filter = "星期='" & str星期 & "'"
'    txt限号.Text = ""
'    txt限约.Text = ""
'    If mrs限号.RecordCount <> 0 Then
'        Me.txt限号.Text = Nvl(mrs限号!限号数, 0)
'        Me.txt限约.Text = Nvl(mrs限号!限约数, 0)
'    End If
'     str时点 = ""
'    With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 800: .RowHeightMin = 800
'        .Rows = 1: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'        If Not bln序号控制 Then
'             .Cols = 8: .FixedCols = 0
'             r = 0: i = 0
'            Do While Not mrsTime.EOF
'                i = i + 1
'                If i > .Cols - 1 Then r = r + 1: i = 0
'                strTime = "预约" & Val(Nvl(mrsTime!限制数量)) & "人" & vbCrLf & vbCrLf
'                strTime = strTime & mrsTime!时间范围
'                If r > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(r, i) = strTime
'                mrsTime.MoveNext
'            Loop
'            For i = 0 To .Cols - 1
'                .ColAlignment(i) = flexAlignCenterCenter
'                .ColWidth(i) = 1200
'            Next
'            .Redraw = flexRDBuffered
'             Exit Sub
'        End If
'        Do While Not mrsTime.EOF
'            If str时点 <> Nvl(mrsTime!时点) Then
'                r = r + 1
'                str时点 = Nvl(mrsTime!时点)
'                If r > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(r, 0) = str时点
'                i = 0
'            End If
'            i = i + 1
'            strTime = mrsTime!序号 & vbCrLf & vbCrLf
'            strTime = strTime & mrsTime!时间范围
'            If i > .Cols - 1 Then .Cols = .Cols + 1
'            If r > .Rows - 1 Then .Rows = .Rows + 1
'            .TextMatrix(r, i) = strTime
'            If Val(Nvl(mrsTime!是否预约)) = 1 Then
'                .Cell(flexcpForeColor, r, i, r, i) = vbBlue
'                .Cell(flexcpFontBold, r, i, r, i) = True
'            End If
'            mrsTime.MoveNext
'        Loop
'        For i = 1 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'        Next
'        .ColWidth(0) = 1200
'        .FixedAlignment(0) = flexAlignRightTop
'        .ColAlignment(0) = flexAlignRightTop
'        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
'        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
'        .Redraw = flexRDBuffered
'    End With
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Sub
'
'Private Sub vsTime_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
' If vsTime.Row < 0 Or vsTime.Col < 0 Or (chk序号控制.Value = 0 And vsTime.Row < 1) Then cmd预约.Visible = False: mblnCellChange = False: Exit Sub
'    cmd预约.Visible = False
'    Select Case mViewMode
'    Case ViewMode.Edit, ViewMode.NewItem:
'       Select Case chk序号控制.Value = 1
'            Case True:
'            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'            '******************************************
'            '设置日期掩码格式
'            '******************************************
'            If vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
'            Case False:
'            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'            '******************************************
'            '设置日期掩码格式
'            '******************************************
'            If NewCol Mod 2 = 0 And vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
'       End Select
'        If vsTime.Row < 0 Or vsTime.Col < 1 Then Exit Sub
'
'        If chk序号控制.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
'            mblnCellChange = True
'        Else
'           mblnCellChange = False
'        End If
'
'    Case ViewMode.ViewItem:
'         mblnCellChange = False
'         vsTime.Editable = flexEDNone
'  End Select
'   If mstr限制修改 <> "" Then
'        vsTime.Editable = flexEDKbdMouse
'        If InStr(mstr限制修改, ";" & mstrKey & ";") > 0 Then vsTime.Editable = flexEDNone
'
'   End If
'End Sub
'
'Private Sub vsTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button <> 1 Or Not mblnCellChange Then Exit Sub
'    If Nvl(txt限约.Text) = 0 Then Exit Sub
'    If InStr(mstr限制修改, mstrKey) > 0 Then Exit Sub
'    cmd预约.Visible = True
'    cmd预约.Left = X - X Mod 1200 + 20
'    cmd预约.Top = Y - Y Mod 400 + 20
'    mblnCellChange = False
'End Sub
'
'Private Sub vsTime_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    '**************************************************************
'    '当操作员 拖动滚动条时 把 预约按钮 隐藏
'    '**************************************************************
'    Me.cmd预约.Visible = False
'End Sub
'
'Private Sub vsTime_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'    If mViewMode = ViewItem Then Exit Sub
'    Select Case chk序号控制.Value = 1
'        Case True:
'            '******************************************
'            '专家号时 控制输入
'            '******************************************
'            If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
'               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
'        Case False:
'            '******************************************
'            '普通号时 控制输入
'            '******************************************
'            If Col Mod 2 = 0 Then
'                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
'               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
'            Else
'                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
'               Or KeyAscii = 13) Then KeyAscii = 0: Exit Sub
'            End If
'    End Select
'End Sub
'
'Private Function isValied() As Boolean
'    '***************************************
'    '验证用户对挂号安排时段的修改
'    '***************************************
'     Dim i          As Long
'     Dim j          As Long
'     Dim lng预约    As Long
'     Dim lng限约    As Long
'     Dim lng限号    As Long
'     Dim str星期    As String
'     If tbWeekTime.SelectedItem Is Nothing Then Exit Function
'      str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
'     lng限号 = Val(txt限号.Text)
'     lng限约 = Val(txt限约.Text)
'     If lng限约 = 0 Then lng限约 = lng限号
'     Select Case chk序号控制.Value = 1
'     Case True:
'     '*************************************
'     '专家号检查限约数是否大于限号数
'     '*************************************
'        With vsTime
'            For i = 0 To .Rows - 1 Step 2
'                For j = 1 To .Cols - 1
'                    If .Cell(flexcpForeColor, i, j, i, j) = vbBlue And .TextMatrix(i, j) <> "" Then
'                        lng预约 = lng预约 + 1
'                    End If
'                Next
'            Next
'        End With
'     Case False:
'     '*************************************
'     '普通号检查限约数是否大于限号数
'     '*************************************
'        With vsTime
'            For i = 1 To .Rows - 1
'                For j = 1 To .Cols - 1 Step 2
'                    If .TextMatrix(i, j) <> "" Then
'                        lng预约 = lng预约 + Val(.TextMatrix(i, j))
'                    End If
'                Next
'            Next
'        End With
'     End Select
'     If lng预约 > lng限约 Then
'        MsgBox "在" & str星期 & "设置的预约数" & lng预约 & "大于了" & IIf(lng限号 = lng限约, "限号数" & lng限约, "限约数" & lng限约) & ",请检查!", vbOKOnly, Me.Caption
'        Exit Function
'     End If
'    isValied = True
'    Exit Function
'End Function
'
'Private Sub vsTime_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'  Dim i         As Long
'  Dim j         As Long
'  Dim lng限号   As Long
'  Dim lng限约   As Long
'  Dim lng预约数 As Long
'  If mViewMode = ViewItem Then Exit Sub
'
'  '*************************************
'  '时间进行验证 输入了时间范围
'  '**************************************
'  If vsTime.Editable = flexEDKbdMouse And vsTime.ColEditMask(vsTime.Col) = strMaskKey Then
'    Validate时段 Row, Col, Cancel
'    If Not Cancel Then mblnChange = True
'    Exit Sub
'  End If
'  '****************************************
'  '在普通号 分时段 对输入的限制预约数进行限制
'  '****************************************
'   If chk序号控制.Value = 0 And vsTime.ColEditMask(vsTime.Col) <> strMaskKey And vsTime.Editable = flexEDKbdMouse Then
'        If vsTime.EditText = "" Then vsTime.EditText = "0"
'        mblnChange = True
'   End If
'End Sub
'
'Private Sub Validate时段(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'  Dim i         As Long
'  Dim j         As Long
'  Dim lng限号   As Long
'  Dim lng限约   As Long
'  Dim lng预约数 As Long
'
'  Dim str时段()  As String
'  If mViewMode = ViewItem Then Exit Sub
'
'  '*************************************
'  '验证时段
'  '**************************************
'  str时段 = Split(vsTime.EditText, "-")
'  If UBound(str时段) <> 1 Then Cancel = True: Exit Sub
'   If Not IsDate(str时段(0)) Then Cancel = True: Exit Sub
'   If Not IsDate(str时段(1)) Then Cancel = True: Exit Sub
'   If CDate(str时段(0)) >= CDate(str时段(1)) Then
'        MsgBox "开始时间必须小于结束时间!请检查!", vbOKOnly, Me.Caption
'        Cancel = True
'   End If
'
'End Sub
'
'Private Sub setVsFlexBgColor(Optional ByVal bln序号控制 As Boolean = False)
'    '**************************************************************
'    '对时间段设置间隔背景
'    '**************************************************************
'     Dim i           As Long
'     If (bln序号控制 And vsTime.Rows = 0) Or (bln序号控制 = False And vsTime.Rows = 1) Then Exit Sub
'     For i = IIf(bln序号控制, 0, 1) To vsTime.Rows - 1 Step 2
'            vsTime.Cell(flexcpBackColor, i, IIf(bln序号控制, 1, 0), i, vsTime.Cols - 1) = &HE0E0D3
'     Next
'End Sub
'



'Private Sub Init时间段()
'  '--------------------------------
'  '功能:获取上下班时间段
'  '--------------------------------
'    Dim strTmp      As String
'    Dim strSQL      As String
'    Dim rsTmp       As ADODB.Recordset
'    Dim strDat      As String
'    On Error GoTo Hd
'    strTmp = zlDatabase.GetPara("上午上下班时间", glngSys, , "07:00:00 AND 12:00:00")
'    strDat = Split(strTmp, "AND")(0)
'    If IsDate(strDat) Then
'        t_时间.dat_上午上班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
'    Else
'        t_时间.dat_上午上班 = CDate("08:00:00")
'    End If
'
'    strDat = Split(strTmp, "AND")(1)
'    If IsDate(strDat) Then
'        t_时间.dat_上午下班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
'    Else
'        t_时间.dat_上午下班 = CDate("1900-01-01 12:00:00")
'    End If
'    strTmp = zlDatabase.GetPara("下午上下班时间", glngSys, , "14:00:00 AND 18:00:00")
'
'     strDat = Split(strTmp, "AND")(0)
'    If IsDate(strDat) Then
'        t_时间.dat_下午上班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
'    Else
'        t_时间.dat_下午上班 = CDate("1900-01-01 14:00:00")
'    End If
'    strDat = Split(strTmp, "AND")(1)
'    If IsDate(strDat) Then
'        t_时间.dat_下午下班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
'    Else
'        t_时间.dat_下午下班 = CDate("1900-01-01 18:00:00")
'    End If
'    With t_时间
'         If .dat_上午上班 > .dat_上午下班 Then
'            .dat_上午下班 = DateAdd("d", 1, .dat_上午下班)
'         End If
'         If .dat_上午上班 > .dat_上午下班 Then
'            .dat_上午下班 = DateAdd("d", 1, .dat_上午下班)
'         End If
'    End With
'    strSQL = _
'    "       Select 时间段, 上班, 下班 " & vbNewLine & _
'    "       From (" & vbNewLine & _
'    "           With Tb As (Select 时间段,To_Date('1900-01-01 ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As 开始时间," & vbNewLine & _
'    "                               To_Date(Decode(Sign(开始时间 - 终止时间), -1, '1900-01-01 ', '1900-01-02 ') ||To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As 终止时间," & _
'    "                               Sign(开始时间 - 终止时间) As 隔天, " & vbNewLine & _
'    "                                To_Date('" & Format(t_时间.dat_上午上班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 上午上班时间, " & vbNewLine & _
'    "                                To_Date('" & Format(t_时间.dat_上午下班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 上午下班时间, " & vbNewLine & _
'    "                                 To_Date('" & Format(t_时间.dat_下午上班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 下午上班时间," & vbNewLine & _
'    "                                 To_Date('" & Format(t_时间.dat_下午下班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 下午下班时间"
'    strSQL = strSQL & vbNewLine & _
'    "                       From 时间段 )" & vbNewLine & _
'    "           Select 时间段, '无' As 标签, 0 As 标志, 开始时间 As 上班, 终止时间 As 下班, 开始时间, 终止时间," & _
'    "                  上午上班时间 As 上班时间, 上午下班时间 As 下班时间" & vbNewLine & _
'    "            From Tb  Where (开始时间 >= 上午下班时间 Or 终止时间 <= 上午上班时间) And " & _
'    "                      (开始时间 >= 下午下班时间 Or 终止时间 <= 下午上班时间) " & vbNewLine & _
'    "           Union All" & vbNewLine & _
'    "           Select 时间段, '有-上午' As 标签, 1 As 标志, Decode(Sign(上午上班时间 - 开始时间), 1, 上午上班时间, 开始时间) As 上班, " & vbNewLine & _
'    "                        Decode(Sign(终止时间 - 上午下班时间), 1, 上午下班时间, 终止时间) As 下班, 开始时间, 终止时间, " & _
'    "                        上午上班时间 As 上班时间, 上午下班时间 As 下班时间 " & vbNewLine & _
'    "           From Tb a Where 时间段 Not In (Select 时间段 From Tb Where 开始时间 >= 上午下班时间 Or 终止时间 <= 上午上班时间) " & vbNewLine & _
'    "           Union All " & vbNewLine & _
'    "            Select 时间段, '有-下午' As 标签, 1 As 标志, Decode(Sign(下午上班时间 - 开始时间), 1, 下午上班时间, 开始时间) As 上班, " & _
'    "                   Decode(Sign(终止时间 - 下午下班时间), 1, 下午下班时间, 终止时间) As 下班, 开始时间, 终止时间, 下午上班时间 As 上班时间, 下午下班时间 As 下班时间 " & vbNewLine & _
'    "         From Tb a   Where 时间段 Not In (Select 时间段 From Tb Where 开始时间 >= 下午下班时间 Or 终止时间 <= 下午上班时间)" & vbNewLine & _
'    "            ) b" & vbNewLine & _
'    "         Order By 时间段,上班"
'     Set mrs上班时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
'    Exit Sub
'Hd:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Sub
'Private Function Get已约限制(ByVal lng安排ID As Long) As String
'    '获取不能修改的安排星期
'    Dim strSQL As String
'    Dim rsTmp   As ADODB.Recordset
'    Dim strTmp  As String
'    strSQL = "Select Decode(To_Char(A.预约时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7'," & _
'    "                             '周六') As 日期 " & vbCrLf & _
'    "          From 病人挂号记录 A,挂号安排　B " & vbCrLf & _
'    "        Where  A.号别=B.号码 And A.记录状态 = 1 And b.ID = [1] And A.发生时间 > A.登记时间 And A.预约时间 Is Not Null"
'
'    If gint预约天数 = 0 Then
'        strSQL = strSQL & " And A.预约时间 > Sysdate "
'    Else
'        strSQL = strSQL & " And A.预约时间 Between Sysdate And Sysdate+" & gint预约天数
'    End If
'    On Error GoTo errH
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
'    If rsTmp.EOF Then Exit Function
'
'    Do While Not rsTmp.EOF
'        If InStr(strTmp, Nvl(rsTmp!日期)) < 0 Or strTmp = "" Then
'            strTmp = strTmp & ";" & Nvl(rsTmp!日期)
'        End If
'        rsTmp.MoveNext
'    Loop
'    If strTmp <> "" Then strTmp = strTmp & ";"
'    Get已约限制 = strTmp
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Function
 
 

Private Sub cmdOther_Click()
    Dim str安排 As String
    
    If Not mbln序号控制 Then Exit Sub
    Set mfrmOtherCalc = New frmRegistPlanTimeOther
    Call mfrmOtherCalc.zlShowMe(Me, tbPage.Item(mlngSelIndex).Caption, Val(txtTimeOut.Text))
    If Not mfrmOtherCalc Is Nothing Then Unload mfrmOtherCalc
    Set mfrmOtherCalc = Nothing '
End Sub

Private Sub mfrmOtherCalc_zlRefreshCon(ByVal VarTimes As Variant)
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng限号 As Long
    Dim lng限约 As Long
    Dim dat开始时间 As Date
    Dim dat结束时间 As Date
    Dim lng序号 As Long
    Dim strTmp As String
    Dim str时段 As String
    Dim str限制时间 As String
    Dim lng默认间隔 As Long
    Dim lng分配个数 As Long
    Dim lng固定数量 As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim dat时点 As Date
    Dim str分段间隔 As String
    Dim str限制项目 As String
    Dim cllPro As Collection
    Dim varTemp As Variant
    Dim strStart As String
    Dim strEnd As String
    Dim int分钟 As Integer
    Dim str时点 As String
    Dim lng时间间隔 As Long
    Dim varData As Variant
    If Not mbln序号控制 Then Exit Sub
    If VarTimes Is Nothing Then Exit Sub
    If VarTimes("时间间隔") <> "" Then
        txtTimeOut.Text = Val(VarTimes("时间间隔"))
        Call cmd设置时段_Click
        Exit Sub
    End If

    str分段间隔 = VarTimes("分段间隔")
    If Trim(str分段间隔) = "" Then Exit Sub


    If mrs上班时间段 Is Nothing Then
        Call Init时间段
    End If
    str限制项目 = GetVsGridCaption(mlngSelIndex)


    If mrs上班时间段 Is Nothing Then Exit Sub
    mrsRegPlan.Filter = "限制项目='" & str限制项目 & "'"
    If mrsRegPlan.RecordCount = 0 Then mrsRegPlan.Filter = 0: Exit Sub
    lng限号 = Nvl(mrsRegPlan!限号数, 0): lng限约 = Nvl(mrsRegPlan!限约数, 0)
    If lng限约 = 0 Then lng限约 = lng限号
    If lng限号 = 0 Then
        MsgBox "当前号别在" & str限制项目 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
        Exit Sub
    End If


    str时段 = mrsRegPlan!排班
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    If mrs上班时间段.RecordCount = 0 Then
        MsgBox "不存在时段为[" & str时段 & "]的上下班时段,请检查!", vbOKOnly, Me.Caption
        Exit Sub

    End If

    Set cllPro = New Collection
    varData = Split(str分段间隔, ";")

    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ",")
        int分钟 = Val(varTemp(1))
        varTemp = Split(varTemp(0), "～")
        strStart = varTemp(0)
        strEnd = varTemp(1)
        cllPro.Add int分钟, "K" & Replace(strStart, ":", "_")
        cllPro.Add strStart, "K" & Replace(strStart, ":", "_") & "_Start"
        cllPro.Add strEnd, "K" & Replace(strStart, ":", "_") & "_End"
    Next

    mrsAssign.Filter = "限制项目='" & str限制项目 & "' And 已使用=0"
    Do While Not mrsAssign.EOF
        mrsAssign.Delete adAffectCurrent
        mrsAssign.MoveNext
    Loop
    mrsAssign.Filter = "限制项目='" & str限制项目 & "'"
    If mrsAssign.RecordCount <> 0 Then
        lng固定数量 = mrsAssign.RecordCount
        lng默认间隔 = Val(Nvl(mrsAssign!时间间隔, lng时间间隔))
        lng时间间隔 = lng默认间隔
        Do While Not mrsAssign.EOF
            lng分配个数 = lng分配个数 + Val(Nvl(mrsAssign!限制数量))
            mrsAssign.MoveNext
        Loop
    End If
    mrsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs上班时间段.EOF
        dat开始时间 = CDate("1900-01-01 " & Format(mrs上班时间段!上班, "hh:mm:ss"))
        If Format(mrs上班时间段!上班, "hh:mm:ss") > Format(mrs上班时间段!下班, "hh:mm:ss") Then
            dat结束时间 = CDate("1900-01-02 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        Else
            dat结束时间 = CDate("1900-01-01 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        End If



        If blnExit Then Exit Do
        dat时点 = dat开始时间
        mrs上班时间段.MoveNext



        For i = j To lng限号
            ' If lngStart > lng限约 Then blnExit = True: Exit For
            If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                j = i
                Exit For
            End If
            If str时点 <> Format(dat时点, "HH:00") Then
                str时点 = Format(dat时点, "HH:00")

                If InStr("," & str分段间隔, str时点) > 0 Then
                    lng时间间隔 = Val(cllPro("K" & Replace(str时点, ":", "_")))
                Else
                    lng时间间隔 = lng默认间隔
                End If
            End If

            If i > lng固定数量 Then
                With mrsAssign
                    .AddNew
                    !限制项目 = str限制项目
                    !开始时间 = Format(dat时点, "hh:mm:00")
                    !时点 = Format(dat时点, "hh:00:00")
                    !结束时间 = Format(DateAdd("n", lng时间间隔, dat时点), "hh:mm:00")
                    !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng时间间隔, dat时点), "hh:mm")
                    !时间间隔 = lng时间间隔
                    !限制数量 = IIf(lng分配个数 >= lng限号, 0, 1)
                    !是否预约 = 0
                    !序号 = i
                    !已使用 = 0
                    .Update
                    lng分配个数 = lng分配个数 + IIf(lng分配个数 >= lng限号, 0, 1)
                End With
            Else
                mrsAssign.Filter = "序号=" & i
                If mrsAssign.RecordCount > 0 Then
                    lng默认间隔 = Nvl(mrsAssign!时间间隔, lng默认间隔)
                Else
                    lng默认间隔 = lng时间间隔
                End If
            End If
            dat时点 = DateAdd("n", IIf(i > lng固定数量, lng时间间隔, lng默认间隔), dat时点)
        Next


        If i > lng限号 And mbln序号控制 Then
            blnExit = True
        End If
    Loop


    Call tbPage_SelectedChanged(tbPage(mlngSelIndex))





End Sub
Private Sub cmd删除_Click(Index As Integer)
    Dim blnDel As Boolean
    Dim lngSelX As Long
    Dim lngSelY As Long
    Dim i As Long, j As Long
    Dim lngCurrSn As Long
    Dim lngStartCol As Long
    With vsTime(Index)
        If .Col < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(.Row, .Col + 1)) = ""
        Else
                blnDel = True
        End If
        blnDel = blnDel And Trim(.TextMatrix(.Row, .Col)) <> "" And Not .Cell(flexcpFontUnderline, .Row, .Col)
        If Not blnDel Then Exit Sub
        If mbln序号控制 Then
          lngSelX = .Row - (.Row Mod 2): lngSelY = .Col
          lngCurrSn = Val(.TextMatrix(lngSelX, lngSelY))
          .TextMatrix(lngSelX, lngSelY) = ""
          .TextMatrix(lngSelX + 1, lngSelY) = ""
          
          For i = lngSelX To .Rows - 1 Step 2
            lngStartCol = 1
            If i = lngSelX Then lngStartCol = lngSelY
            For j = lngStartCol To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    .TextMatrix(i, j) = lngCurrSn
                     lngCurrSn = lngCurrSn + 1
                End If
            Next
         Next
        End If
        cmd删除(Index).Visible = False
        cmd预约(Index).Visible = False
        .SetFocus
    End With
End Sub

Private Sub cmd设置时段_Click()
    If AssignReapportion(Val(Me.txtTimeOut.Text), tbPage.Item(mlngSelIndex).Caption) = False Then Exit Sub
    Call tbPage_SelectedChanged(tbPage.Item(mlngSelIndex))
End Sub

Private Sub cmd预约_Click(Index As Integer)
    If Not mbln序号控制 Or mlngSelIndex < 0 Then Exit Sub
    If mlngSelIndex <> Index Then Exit Sub
    With vsTime(mlngSelIndex)
        If .MouseRow < 0 Or .MouseCol < 0 Then Exit Sub
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If .Cell(flexcpForeColor, .Row, .Col) = vbBlue Then
           .Cell(flexcpForeColor, .Row - (.Row Mod 2), .Col, .Row + (.Row + 1) Mod 2, .Col) = &H80000008
            .Cell(flexcpFontBold, .Row - (.Row Mod 2), .Col, .Row + (.Row + 1) Mod 2, .Col) = False
        Else
            .Cell(flexcpForeColor, .Row - (.Row Mod 2), .Col, .Row + (.Row + 1) Mod 2, .Col) = vbBlue
            .Cell(flexcpFontBold, .Row - (.Row Mod 2), .Col, .Row + (.Row + 1) Mod 2, .Col) = True
        End If
        mblnChange = True
        .SetFocus
    End With
End Sub

Private Sub Form_Load()
    Call LoadPageControl
    Call LoadPage
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With fra应用于
        .Top = ScaleHeight - .Height - 50
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Visible = True
    End With
    With tbPage
        .Top = txtTimeOut.Top + txtTimeOut.Height + 50
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Height = fra应用于.Top - .Top - 100
    End With
End Sub

Private Sub picPage_Resize(Index As Integer)
    Err = 0: On Error Resume Next

    With picPage(Index)
        vsTime(Index).Left = .ScaleLeft
        vsTime(Index).Top = .ScaleTop
        vsTime(Index).Width = .ScaleWidth
        vsTime(Index).Height = .ScaleHeight
    End With
End Sub

Private Sub InitRs(Optional ByVal blnInitRs As Boolean = True)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    If Not mrsAssign Is Nothing Then Exit Sub
    With mrsAssign
        Set mrsAssign = New ADODB.Recordset
        mrsAssign.Fields.Append "限制项目", adVarChar, 20
        mrsAssign.Fields.Append "开始时间", adVarChar, 20
        mrsAssign.Fields.Append "时点", adVarChar, 20
        mrsAssign.Fields.Append "结束时间", adVarChar, 20
        mrsAssign.Fields.Append "时间段", adVarChar, 50
        mrsAssign.Fields.Append "时间间隔", adBigInt, 4
        mrsAssign.Fields.Append "限制数量", adBigInt, 10
        mrsAssign.Fields.Append "是否预约", adBigInt, 18
        mrsAssign.Fields.Append "序号", adBigInt, 18
        mrsAssign.Fields.Append "已使用", adBigInt, 2
        mrsAssign.CursorLocation = adUseClient
        mrsAssign.LockType = adLockOptimistic
        mrsAssign.CursorType = adOpenStatic
        mrsAssign.Open
    End With
    If blnInitRs Then Call InitAssignRs
End Sub

Private Function InitAssignRs() As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng固定 As Long  '固定的序号不允许更改
    Dim i As Long
    '初始化已分配数据集合
    If mEditType = EM_安排_增加 Then Exit Function
     On Error GoTo Hd
    If mEditType = EM_安排_查阅 Or mEditType = EM_安排_修改 Or mEditType = EM_计划_增加 Then
        strSQL = "Select 序号, 星期 As 限制项目, To_Char(开始时间, 'hh24:mi:ss') As 开始时间, To_Char(结束时间, 'hh24:mi:ss') As 结束时间,"
        strSQL = strSQL & vbCrLf & "         是否预约 , 限制数量,To_Char(开始时间, 'hh24') || ':00:00' As 时点,To_Char(开始时间, 'hh24:mi') || '-' || To_Char(结束时间, 'hh24:mi') As 时间段"
        strSQL = strSQL & vbCrLf & " From 挂号安排时段 Where 安排ID=[1] "
        strSQL = strSQL & vbCrLf & " Order By 星期"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
    ElseIf mEditType = EM_计划_查阅 Or EM_计划_修改 Then
        strSQL = "Select 序号, 星期 As 限制项目, To_Char(开始时间, 'hh24:mi:ss') As 开始时间, To_Char(结束时间, 'hh24:mi:ss') As 结束时间,"
        strSQL = strSQL & vbCrLf & "         是否预约 , 限制数量, To_Char(开始时间, 'hh24') || ':00:00' As 时点,To_Char(开始时间, 'hh24:mi') || '-' || To_Char(结束时间, 'hh24:mi') As 时间段"
        strSQL = strSQL & vbCrLf & " From 挂号计划时段 Where 计划ID=[1] "
        strSQL = strSQL & vbCrLf & " Order By 星期"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng计划Id)
    End If
    Do While Not rsTmp.EOF
            With mrsAssign
                .AddNew
                !限制项目 = Nvl(rsTmp!限制项目)
                !开始时间 = Nvl(rsTmp!开始时间, "00:00:00")
                !结束时间 = Nvl(rsTmp!结束时间, "00:00:00")
                !时间段 = Nvl(rsTmp!时间段, "__:__-__:__")
                !时间间隔 = DateDiff("n", CDate(!开始时间), CDate(!结束时间))
                !限制数量 = Val(Nvl(rsTmp!限制数量))
                !是否预约 = Val(Nvl(rsTmp!是否预约))
                !时点 = Nvl(rsTmp!时点, "00:00:00")
                !序号 = Val(Nvl(rsTmp!序号))
                lng固定 = 0
                If Not mrsHistory Is Nothing Then
                mrsHistory.Filter = "限制项目='" & Nvl(rsTmp!限制项目) & "'"
                    If mrsHistory.RecordCount > 0 Then
                        If CStr(mrsHistory!发生时间) >= CStr(Nvl(rsTmp!开始时间, "00:00:00")) Then
                            lng固定 = 1
                        End If
                    End If
                End If
                !已使用 = lng固定
                .Update
                
            End With
        rsTmp.MoveNext
    Loop
    Call AssignManage
'    If mblnInit Then
'        For i = 0 To 6
'            If tbPage.Item(i).Visible And tbPage.Item(i).Enabled Then
'                tbPage.Item(i).Selected = True
'                Call tbPage_SelectedChanged(tbPage.Item(i))
'                Exit For
'            End If
'        Next
'    End If
    InitAssignRs = True
Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   Dim str限制项目 As String
   If Not mblnInit Then Exit Sub
   
   If Item.Index <> mlngSelIndex And mlngSelIndex <> -1 Then '
     If mlngSelIndex <> -1 And mblnChange Then
        If VsTimeValidate(mlngSelIndex) = False Then
            mblnOnChange = True
            tbPage.Item(mlngSelIndex).Selected = True
            mblnOnChange = False
            Exit Sub
        End If
     End If
     
     str限制项目 = GetVsGridCaption(mlngSelIndex)
     If MoveAssign(str限制项目) = False Then
        If mlngSelIndex <> -1 Then tbPage.Item(mlngSelIndex).Selected = True
        Exit Sub
     End If
   End If
   
   If mblnOnChange Then Exit Sub
   mlngSelIndex = Item.Index
   SetStyle mbln序号控制, Item.Index
   
   LoadTimePlan Item.Caption
   setVsGridSNStyle Item.Index
End Sub

Private Sub Init时间段()
  '--------------------------------
  '功能:获取上下班时间段
  '--------------------------------
    Dim strTmp      As String
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim strDat      As String
    On Error GoTo Hd
    strTmp = zlDatabase.GetPara("上午上下班时间", glngSys, , "07:00:00 AND 12:00:00")
    strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_时间.dat_上午上班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_上午上班 = CDate("08:00:00")
    End If
   
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_时间.dat_上午下班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_上午下班 = CDate("1900-01-01 12:00:00")
    End If
    strTmp = zlDatabase.GetPara("下午上下班时间", glngSys, , "14:00:00 AND 18:00:00")
    
     strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_时间.dat_下午上班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_下午上班 = CDate("1900-01-01 14:00:00")
    End If
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_时间.dat_下午下班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_下午下班 = CDate("1900-01-01 18:00:00")
    End If
    With t_时间
         If .dat_上午上班 > .dat_上午下班 Then
            .dat_上午下班 = DateAdd("d", 1, .dat_上午下班)
         End If
         If .dat_上午上班 > .dat_上午下班 Then
            .dat_上午下班 = DateAdd("d", 1, .dat_上午下班)
         End If
    End With
    strSQL = _
    "       Select 时间段, 上班, 下班 " & vbNewLine & _
    "       From (" & vbNewLine & _
    "           With Tb As (Select 时间段,To_Date('1900-01-01 ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As 开始时间," & vbNewLine & _
    "                               To_Date(Decode(Sign(开始时间 - 终止时间), -1, '1900-01-01 ', '1900-01-02 ') ||To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As 终止时间," & _
    "                               Sign(开始时间 - 终止时间) As 隔天, " & vbNewLine & _
    "                                To_Date('" & Format(t_时间.dat_上午上班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 上午上班时间, " & vbNewLine & _
    "                                To_Date('" & Format(t_时间.dat_上午下班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 上午下班时间, " & vbNewLine & _
    "                                 To_Date('" & Format(t_时间.dat_下午上班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 下午上班时间," & vbNewLine & _
    "                                 To_Date('" & Format(t_时间.dat_下午下班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 下午下班时间"
    strSQL = strSQL & vbNewLine & _
    "                       From 时间段 )" & vbNewLine & _
    "           Select 时间段, '无' As 标签, 0 As 标志, 开始时间 As 上班, 终止时间 As 下班, 开始时间, 终止时间," & _
    "                  上午上班时间 As 上班时间, 上午下班时间 As 下班时间" & vbNewLine & _
    "            From Tb  Where (开始时间 >= 上午下班时间 Or 终止时间 <= 上午上班时间) And " & _
    "                      (开始时间 >= 下午下班时间 Or 终止时间 <= 下午上班时间) " & vbNewLine & _
    "           Union All" & vbNewLine & _
    "           Select 时间段, '有-上午' As 标签, 1 As 标志, Decode(Sign(上午上班时间 - 开始时间), 1, 上午上班时间, 开始时间) As 上班, " & vbNewLine & _
    "                        Decode(Sign(终止时间 - 上午下班时间), 1, 上午下班时间, 终止时间) As 下班, 开始时间, 终止时间, " & _
    "                        上午上班时间 As 上班时间, 上午下班时间 As 下班时间 " & vbNewLine & _
    "           From Tb a Where 时间段 Not In (Select 时间段 From Tb Where 开始时间 >= 上午下班时间 Or 终止时间 <= 上午上班时间) " & vbNewLine & _
    "           Union All " & vbNewLine & _
    "            Select 时间段, '有-下午' As 标签, 1 As 标志, Decode(Sign(下午上班时间 - 开始时间), 1, 下午上班时间, 开始时间) As 上班, " & _
    "                   Decode(Sign(终止时间 - 下午下班时间), 1, 下午下班时间, 终止时间) As 下班, 开始时间, 终止时间, 下午上班时间 As 上班时间, 下午下班时间 As 下班时间 " & vbNewLine & _
    "         From Tb a   Where 时间段 Not In (Select 时间段 From Tb Where 开始时间 >= 下午下班时间 Or 终止时间 <= 下午上班时间)" & vbNewLine & _
    "            ) b" & vbNewLine & _
    "         Order By 时间段,上班"
     Set mrs上班时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : AssignReapportion
'Description : 序号重新分配
'Author      : 李光福
'Date        : 05-November-2012 14:53:16
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'lng间隔时间           Long              ByVal                .时间间隔
'str限制项目           String            ByVal                .星期
'Output      :  分配是否成功
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function AssignReapportion(ByVal lng间隔时间 As Long, ByVal str限制项目 As String) As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng限号 As Long
    Dim lng限约 As Long
    Dim dat开始时间 As Date
    Dim dat结束时间 As Date
    Dim lng序号 As Long
    Dim strTmp As String
    Dim str时段 As String
    Dim str限制时间 As String
    Dim lng默认间隔 As Long
    Dim lng分配个数 As Long
    Dim lng固定数量 As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim dat时点 As Date
    If mrs上班时间段 Is Nothing Then
        Call Init时间段
    End If

    If mrs上班时间段 Is Nothing Then Exit Function
    mrsRegPlan.Filter = "限制项目='" & str限制项目 & "'"
    If mrsRegPlan.RecordCount = 0 Then mrsRegPlan.Filter = 0: Exit Function
    lng限号 = Nvl(mrsRegPlan!限号数, 0): lng限约 = Nvl(mrsRegPlan!限约数, 0)
    If lng限约 = 0 Then lng限约 = lng限号
    If lng限号 = 0 Then
        MsgBox "当前号别在" & str限制项目 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
        Exit Function
    End If


    str时段 = mrsRegPlan!排班
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    If mrs上班时间段.RecordCount = 0 Then
        MsgBox "不存在时段为[" & str时段 & "]的上下班时段,请检查!", vbOKOnly, Me.Caption
        Exit Function

    End If

    mrsAssign.Filter = "限制项目='" & str限制项目 & "' And 已使用=0"
    Do While Not mrsAssign.EOF
        mrsAssign.Delete adAffectCurrent
        mrsAssign.MoveNext
    Loop
    mrsAssign.Filter = "限制项目='" & str限制项目 & "'"
    If mrsAssign.RecordCount <> 0 Then
        lng固定数量 = mrsAssign.RecordCount
        lng默认间隔 = Val(Nvl(mrsAssign!时间间隔, lng间隔时间))
        Do While Not mrsAssign.EOF
            lng分配个数 = lng分配个数 + Val(Nvl(mrsAssign!限制数量))
            mrsAssign.MoveNext
        Loop
    End If
    mrsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs上班时间段.EOF
        dat开始时间 = CDate("1900-01-01 " & Format(mrs上班时间段!上班, "hh:mm:ss"))
        If Format(mrs上班时间段!上班, "hh:mm:ss") > Format(mrs上班时间段!下班, "hh:mm:ss") Then
            dat结束时间 = CDate("1900-01-02 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        Else
            dat结束时间 = CDate("1900-01-01 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        End If

        If blnExit Then Exit Do
        dat时点 = dat开始时间
        mrs上班时间段.MoveNext

        If mbln序号控制 Then

            For i = j To lng限号
                ' If lngStart > lng限约 Then blnExit = True: Exit For
                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                End If
                If i > lng固定数量 Then
                    With mrsAssign
                        .AddNew
                        !限制项目 = str限制项目
                        !开始时间 = Format(dat时点, "hh:mm:00")
                        !时点 = Format(dat时点, "hh:00:00")
                        !结束时间 = Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm:00")
                        !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm")
                        !时间间隔 = lng间隔时间
                        !限制数量 = IIf(lng分配个数 >= lng限号, 0, 1)
                        !是否预约 = 0
                        !序号 = i
                        !已使用 = 0
                        .Update
                        lng分配个数 = lng分配个数 + IIf(lng分配个数 >= lng限号, 0, 1)
                    End With
                Else
                    mrsAssign.Filter = "序号=" & i
                    If mrsAssign.RecordCount > 0 Then
                        lng默认间隔 = Nvl(mrsAssign!时间间隔, lng默认间隔)
                    Else
                        lng默认间隔 = lng间隔时间
                    End If
                End If
                dat时点 = DateAdd("n", IIf(i > lng固定数量, lng间隔时间, lng默认间隔), dat时点)
            Next

        Else    '非序号控制

            Do While Not Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss")

                ' If lngStart > lng限约 Then blnExit = True: Exit For
                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then Exit Do

                If i > lng固定数量 Then
                    With mrsAssign
                        .AddNew
                        !限制项目 = str限制项目
                        !开始时间 = Format(dat时点, "hh:mm:00")
                        !时点 = Format(dat时点, "hh:00:00")
                        !结束时间 = Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm:00")
                        !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm")
                        !时间间隔 = lng间隔时间
                        !限制数量 = IIf(lng分配个数 >= lng限约, 0, 1)
                        !是否预约 = 1
                        !序号 = i
                        !已使用 = 0
                        .Update
                        lng分配个数 = lng分配个数 + IIf(lng分配个数 >= lng限约, 0, 1)
                    End With
                Else
                    mrsAssign.Filter = "序号=" & i
                    If mrsAssign.RecordCount > 0 Then
                        lng默认间隔 = Nvl(mrsAssign!时间间隔, lng默认间隔)
                    Else
                        lng默认间隔 = lng间隔时间
                    End If
                End If
                dat时点 = DateAdd("n", IIf(i > lng固定数量, lng间隔时间, lng默认间隔), dat时点)
                i = i + 1
            Loop


        End If
        If i > lng限号 And mbln序号控制 Then
            blnExit = True
        End If
    Loop
    AssignReapportion = True
End Function


'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : AssignReapportion
'Description : 序号重新分配
'Author      : 李光福
'Date        : 05-November-2012 14:53:16
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'lng间隔时间           Long              ByVal                .时间间隔
'str限制项目           String            ByVal                .星期
'Output      :  分配是否成功
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function AssignReapportion1(ByVal cllTime As Collection, ByVal str限制项目 As String) As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng限号 As Long
    Dim lng限约 As Long
    Dim dat开始时间 As Date
    Dim dat结束时间 As Date
    Dim lng序号 As Long
    Dim strTmp As String
    Dim str时段 As String
    Dim str限制时间 As String
    Dim lng默认间隔 As Long
    Dim lng间隔时间 As Long
    Dim lng间隔 As Long
    Dim lng分配个数 As Long
    Dim lng固定数量 As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim dat时点  As Date
    Dim strPre时点 As String
    
    If Not mbln序号控制 Then Exit Function
    
    If mrs上班时间段 Is Nothing Then
       Call Init时间段
    End If
    
    If mrs上班时间段 Is Nothing Then Exit Function
    
    mrsRegPlan.Filter = "限制项目='" & str限制项目 & "'"
    If mrsRegPlan.RecordCount = 0 Then mrsRegPlan.Filter = 0: Exit Function
    lng限号 = Nvl(mrsRegPlan!限号数, 0): lng限约 = Nvl(mrsRegPlan!限约数, 0)
    
    If lng限约 = 0 Then lng限约 = lng限号
    
    If lng限号 = 0 Then
        MsgBox "当前号别在" & str限制项目 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    
    str时段 = mrsRegPlan!排班
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    If mrs上班时间段.RecordCount = 0 Then
        MsgBox "不存在时段为[" & str时段 & "]的上下班时段,请检查!", vbOKOnly, Me.Caption
        Exit Function
    
    End If
    
    mrsAssign.Filter = "限制项目='" & str限制项目 & "' And 已使用=0"
    Do While Not mrsAssign.EOF
        mrsAssign.Delete adAffectCurrent
        mrsAssign.MoveNext
    Loop
    mrsAssign.Filter = "限制项目='" & str限制项目 & "'"
    If mrsAssign.RecordCount <> 0 Then
            lng固定数量 = mrsAssign.RecordCount
            lng默认间隔 = Val(Nvl(mrsAssign!时间间隔, lng间隔时间))
            Do While Not mrsAssign.EOF
                lng分配个数 = lng分配个数 + Val(Nvl(mrsAssign!限制数量))
                mrsAssign.MoveNext
            Loop
    End If
    mrsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs上班时间段.EOF
        dat开始时间 = CDate("1900-01-01 " & Format(mrs上班时间段!上班, "hh:mm:ss"))
        If Format(mrs上班时间段!上班, "hh:mm:ss") > Format(mrs上班时间段!下班, "hh:mm:ss") Then
            dat结束时间 = CDate("1900-01-02 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        Else
            dat结束时间 = CDate("1900-01-01 " & Format(mrs上班时间段!下班, "hh:mm:ss"))
        End If
        
        If blnExit Then Exit Do
        dat时点 = dat开始时间
        mrs上班时间段.MoveNext
        
       
        
            For i = j To lng限号
               ' If lngStart > lng限约 Then blnExit = True: Exit For
                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(dat结束时间, "yyyy-MM-dd hh:mm:ss") Then
                   j = i
                   Exit For
                End If
                If i > lng固定数量 Then
                    With mrsAssign
                        .AddNew
                        !限制项目 = str限制项目
                        !开始时间 = Format(dat时点, "hh:mm:00")
                        !时点 = Format(dat时点, "hh:00:00")
                        !结束时间 = Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm:00")
                        !时间段 = Format(dat时点, "hh:mm") & "-" & Format(DateAdd("n", lng间隔时间, dat时点), "hh:mm")
                        !时间间隔 = lng间隔时间
                        !限制数量 = IIf(lng分配个数 >= lng限号, 0, 1)
                        !是否预约 = 0
                        !序号 = i
                        !已使用 = 0
                        .Update
                         lng分配个数 = lng分配个数 + IIf(lng分配个数 >= lng限号, 0, 1)
                    End With
                Else
                    mrsAssign.Filter = "限制项目='" & str限制项目 & "' And 序号=" & i
                    If mrsAssign.RecordCount = 0 Then
                        lng间隔 = lng默认间隔
                    Else
                        lng间隔 = Val(Nvl(mrsAssign!时间间隔, lng默认间隔))
                    End If
                End If
                dat时点 = DateAdd("n", IIf(i > lng固定数量, lng间隔时间, lng间隔), dat时点)
            Next
           
       
        If i > lng限号 And mbln序号控制 Then
                blnExit = True
        End If
    Loop
    AssignReapportion1 = True
End Function


'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : AssignManage
'Description : 对已经分配的号码进行限号数限约数的规则进行处理
'Author      : 李光福
'Date        : 05-November-2012 14:48:05
'Input       :
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function AssignManage() As Boolean
    Dim varData As Variant, varTemp As Variant, i As Long
    Dim j As Long, lngIndex As Long, p As Long, strTemp As String
    Dim lng限号数 As Long, lng限约数 As Long, lng分配数量 As Long
    Dim lng分配预约 As Long, lngTmp  As Long, lngTemp As Long
    Dim str最大时间 As String, blnChange As Boolean
     
    varData = Split(mstr安排, "|")
    lngIndex = -1
    For i = 0 To 6
        strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
        '如果当天应诊时间改变
        If InStr("|" & mstr安排, "|" & strTemp & ",") = 0 Or InStr("|" & mstr应诊时段 & "|", "|" & strTemp & "|") = 0 Then
            mrsAssign.Filter = "限制项目='" & strTemp & "'"
            Do While Not mrsAssign.EOF
                mrsAssign.Delete adAffectCurrent
                mrsAssign.Update
                mrsAssign.MoveNext
            Loop
        End If
    Next
    For i = 0 To UBound(varData)
        ''周一,限号数,限约数|周二,限号数,限约数|....
        varTemp = Split(varData(i) & ",,,,", ",")
        If varTemp(0) <> "" Then
            lng限号数 = Val(varTemp(1)): lng限约数 = Val(varTemp(2))
            If lng限约数 = 0 Then lng限约数 = lng限号数
            str最大时间 = ""
            If Not mrsHistory Is Nothing Then
                mrsHistory.Filter = "限制项目='" & varTemp(0) & "'"
                If mrsHistory.RecordCount = 0 Then
                   str最大时间 = ""
                Else
                   str最大时间 = Nvl(mrsHistory!发生时间)
                End If
            End If
            mrsAssign.Filter = "限制项目='" & varTemp(0) & "'"
            mrsAssign.Sort = "序号"
             
         
'             If str最大时间 <> "" Then
'                Do While Not mrsAssign.EOF
'                   If str最大时间 > Nvl(mrsAssign!开始时间) Then mrsAssign.Delete adAffectCurrent
'                   mrsAssign.MoveNext
'                Loop
'             End If
             
              lng分配数量 = 0
              blnChange = False
             Do While Not mrsAssign.EOF
                If lng分配数量 + Val(Nvl(mrsAssign!限制数量)) > IIf(mbln序号控制, lng限号数, lng限约数) Then
                    blnChange = True
                    If Val(Nvl(mrsAssign!已使用)) = 0 Then
                        lngTmp = Val(mrsAssign!限制数量)
                        lngTemp = lng分配数量 + lngTmp - IIf(mbln序号控制, lng限号数, lng限约数)
                        If lngTmp <= lngTemp Then
                            lngTmp = 0
                        Else
                            lngTmp = lngTmp - lngTemp
                            lng分配数量 = lng限号数
                        End If
                        mrsAssign!限制数量 = lngTmp
                        mrsAssign.Update
                        If mbln序号控制 Then
                            mrsAssign.Delete adAffectCurrent
                        End If
                    End If
                Else
                    lng分配数量 = lng分配数量 + Val(Nvl(mrsAssign!限制数量))
                End If
                mrsAssign.MoveNext
             Loop
             If blnChange Then
                mrsAssign.Filter = "限制项目='" & varTemp(0) & "' And 限制数量>0"
                lng分配数量 = 0
                If mrsAssign.RecordCount = 0 Then mrsAssign.Filter = 0: AssignManage = True: Exit Function
                mrsAssign.Sort = "序号 desc"
                mrsAssign.MoveFirst
                'lng分配数量
                Do While Not mrsAssign.EOF
                   lng分配数量 = lng分配数量 + Val(Nvl(mrsAssign!限制数量))
                   mrsAssign.MoveNext
                Loop
                mrsAssign.MoveFirst
                If lng分配数量 > IIf(mbln序号控制, lng限号数, lng限约数) Then
                   Do While Not mrsAssign.EOF
                      If Val(Nvl(mrsAssign!已使用)) = 0 Then
                           lngTmp = Val(Nvl(mrsAssign!限制数量))
                           lngTemp = lng分配数量 - lng限号数
                           If lngTemp >= lngTmp Then
                               mrsAssign!限制数量 = 0
                               mrsAssign.Update
                               lng分配数量 = lng分配数量 - lngTmp
                           Else
                               lngTmp = lngTmp - lngTemp
                               mrsAssign!限制数量 = lngTmp
                               mrsAssign.Update
                               lng分配数量 = lng分配数量 - lngTemp
                           End If
                      End If
                      If lng分配数量 <= lng限号数 Then Exit Do
                      mrsAssign.MoveNext
                   Loop
                End If
             End If
        End If
    Next
    mrsAssign.Filter = 0
    If Not mrsHistory Is Nothing Then mrsHistory.Filter = 0
    AssignManage = True
End Function

Private Function VsTimeValidate(ByVal lngIndex As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证设置的限约数是否符合要求
    '入参:lngIndex-指定的页面(星期对应的索引):-1时,表示按所有的页面进行检查
    '出参:
    '返回:校对成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-15 10:17:37
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngStep As Long, i As Long, j  As Long
    Dim lng预约数   As Long, lng限号数 As Long, lng限约数 As Long, lng号数 As Long
    Dim str星期   As String, str限制项目 As String
    Dim lngPage As Long, lngPages As Long, lngStartPage As Long
    Dim blnNotSetTime As Boolean '允许不设置时间段
    Dim blnAllowNums As Boolean '允许限号数不一致
    Dim blnAllowYYNums As Boolean '允许预约数与设置的预约数不一致
    Dim strCommand As String, bln时段 As Boolean '判断设置了时段的,需要检查其他时段页是否设置
    On Error GoTo errHandle
        
    lngStartPage = 0: lngPages = tbPage.ItemCount - 1
    If lngIndex <> -1 Then lngStartPage = lngIndex: lngPages = lngIndex
    bln时段 = False
    For lngPage = lngStartPage To lngPages
        If mbln序号控制 Then
            With vsTime(lngPage)
                For i = 0 To .Rows - 1 Step 2
                    For j = 1 To .Cols - 1
                       If .TextMatrix(i, j) <> "" Then
                           bln时段 = True: Exit For
                       End If
                    Next
                Next
            End With
        Else
                With vsTime(lngPage)
                    For i = 1 To .Rows - 1
                        For j = 1 To .Cols - 1 Step 2
                            If .TextMatrix(i, j) <> "" Then
                               bln时段 = True: Exit For
                            End If
                        Next
                    Next
                End With
        End If
    Next
    '未启用时段
    If bln时段 = False Then VsTimeValidate = True: Exit Function
    
    For lngPage = lngStartPage To lngPages
        str限制项目 = GetVsGridCaption(lngPage)
        mrsRegPlan.Filter = "限制项目='" & str限制项目 & "'"
        If mrsRegPlan.RecordCount = 0 Then
            mrsRegPlan.Filter = 0
        Else
                lng限号数 = Val(Nvl(mrsRegPlan!限号数)): lng限约数 = Val(Nvl(mrsRegPlan!限约数))
                If lng限约数 = 0 Then lng限约数 = lng限号数
                lng号数 = 0: lng预约数 = 0
                
                If mbln序号控制 Then
                    '专家号检查限约数是否大于限号数
                    With vsTime(lngPage)
                        For i = 0 To .Rows - 1 Step 2
                            For j = 1 To .Cols - 1
                               If .TextMatrix(i, j) <> "" Then
                                     If .Cell(flexcpForeColor, i, j, i, j) = vbBlue Then
                                         lng预约数 = lng预约数 + 1
                                     End If
                                     lng号数 = lng号数 + 1
                               End If
                            Next
                        Next
                    End With
                    If lng号数 < lng限号数 Then
                        If lng号数 = 0 Then
                           If lngIndex = -1 Then
                                If blnNotSetTime = False And bln时段 Then
                                        strCommand = zlCommFun.ShowMsgbox("提醒", "    在分时段页面中未设置『" & str限制项目 & "』的时段,你确定不设置时间段?" & vbCrLf & vbCrLf & _
                                         "『是』:表示允许不设置时间段进行保存" & vbCrLf & vbCrLf & _
                                         "『忽略』:表示遇到类似的未设置时间段的问题允许保存,但不再提示。" & vbCrLf & vbCrLf & _
                                         "『否』:表示不允许不设置时间段,返回重新设置" & vbCrLf, "是(&O),忽略(&I),否(&C)", Me, vbQuestion)
                                        Select Case strCommand
                                        Case "是"
                                        Case "忽略"
                                             blnNotSetTime = True
                                         Case Else
                                            RaiseEvent zlSaveTimePageSelected(str限制项目)
                                            mblnNotBrush = True
                                            tbPage.Item(lngPage).Selected = True
                                            If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                            mblnNotBrush = False
                                            Exit Function
                                         End Select
                                End If
                           Else
                                If MsgBox("在分时段页面中未设置『" & str限制项目 & "』的时段,你确定不设置时间段?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    If lngIndex = -1 Then
                                        RaiseEvent zlSaveTimePageSelected(str限制项目)
                                        mblnNotBrush = True
                                        tbPage.Item(lngPage).Selected = True
                                        If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                        mblnNotBrush = False
                                    End If
                                    Exit Function
                                End If
                            End If
                        Else
                                If lngIndex = -1 Then
                                        If blnAllowNums = False Then
                                        
                                                strCommand = zlCommFun.ShowMsgbox("提醒", "    在分时段页面中的『" & str限制项目 & "』所设置时间段的号数(" & lng号数 & ")与限号数(" & lng限号数 & ") 不等,你确定按当前设置的时段保存?" & vbCrLf & vbCrLf & _
                                                 "『是』:表示允许限号数与号数不一致" & vbCrLf & vbCrLf & _
                                                 "『忽略』:表示允许限号数与号数不一致，遇到类似的问题,不再提示。" & vbCrLf & vbCrLf & _
                                                 "『否』:表示不允许限号数与号数不一致,返回重新设置" & vbCrLf, "是(&O),忽略(&I),否(&C)", Me, vbQuestion)
                                                Select Case strCommand
                                                 Case "是"
                                                 Case "忽略"
                                                     blnAllowNums = True
                                                 Case Else
                                                    RaiseEvent zlSaveTimePageSelected(str限制项目)
                                                    mblnNotBrush = True
                                                    tbPage.Item(lngPage).Selected = True
                                                    If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                                    mblnNotBrush = False
                                                     Exit Function
                                                 End Select
                                        End If
                                   Else
                                        If MsgBox("在分时段页面中的『" & str限制项目 & "』所设置时间段的号数(" & lng号数 & ")与限号数(" & lng限约数 & ") 不等,你确定按当前设置的时段保存?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                End If
                        End If
                    ElseIf lng号数 > lng限号数 Then
                        Call MsgBox("在分时段页面中的『" & str限制项目 & "』所设置时间段的号数(" & lng号数 & ")大于了限号数(" & lng限约数 & ") 你不能按当前设置的时段保存!", vbQuestion + vbOKOnly + vbDefaultButton2, gstrSysName)
                        If lngIndex = -1 Then
                            RaiseEvent zlSaveTimePageSelected(str限制项目)
                            mblnNotBrush = True
                            tbPage.Item(lngPage).Selected = True
                            If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                            mblnNotBrush = False
                        End If
                        Exit Function
                    End If
                Else
                     '普通号检查限约数是否大于限号数
                    With vsTime(lngPage)
                        For i = 1 To .Rows - 1
                            For j = 1 To .Cols - 1 Step 2
                                If .TextMatrix(i, j) <> "" Then
                                    lng预约数 = lng预约数 + Val(.TextMatrix(i, j))
                                End If
                            Next
                        Next
                    End With
                End If
                If lng预约数 > lng限约数 Then
                   MsgBox "在分时段页面中的『" & str限制项目 & "』所设置的预约数(" & lng预约数 & ")大于了" & IIf(lng限号数 = lng限约数, "限号数(" & lng限约数 & ")", "限约数(" & lng限约数 & ")") & ",你不能按当前设置保存!", vbOKOnly, Me.Caption
                    If lngIndex = -1 Then
                        RaiseEvent zlSaveTimePageSelected(str限制项目)
                        mblnNotBrush = True
                        tbPage.Item(lngPage).Selected = True
                        If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                        mblnNotBrush = False
                    End If
                   Exit Function
                End If
                If lng预约数 < lng限约数 And lng预约数 <> 0 Then
                    If lngIndex = -1 Then
                           If blnAllowYYNums = False Then
                                   strCommand = zlCommFun.ShowMsgbox("提醒", "    在分时段页面中的『" & str限制项目 & "』所设置的实际预约数(" & lng预约数 & ") 与限约数(" & lng限约数 & ") 不等,你确定按当前设置的时段保存?" & vbCrLf & vbCrLf & _
                                    "『是』:表示允许限约数与预约数不一致" & vbCrLf & vbCrLf & _
                                    "『忽略』:表示允许限约数与预约数不一致，遇到类似的问题,不再提示。" & vbCrLf & vbCrLf & _
                                    "『否』:表示不允许限约数与预约数不一致,返回重新设置" & vbCrLf, "是(&O),忽略(&I),否(&C)", Me, vbQuestion)
                                    Select Case strCommand
                                    Case "是"
                                    Case "忽略"
                                        blnAllowYYNums = True
                                    Case Else
                                       RaiseEvent zlSaveTimePageSelected(str限制项目)
                                       mblnNotBrush = True
                                       tbPage.Item(lngPage).Selected = True
                                       If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                       mblnNotBrush = False
                                        Exit Function
                                    End Select
                           End If
                      Else
                            If MsgBox("在分时段页面中的『" & str限制项目 & "』所设置的实际预约数(" & lng预约数 & ") 与限约数(" & lng限约数 & ") 不等,你确定按当前设置的时段保存?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
        End If
    Next
    VsTimeValidate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : MoveAssign
'Description : 将分配的序号保存到本地数据集合中
'Author      : 李光福
'Date        : 05-November-2012 15:06:42
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'str限制项目           String            ByVal                .星期
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function MoveAssign(ByVal str限制项目 As String) As Boolean
    '分配调整的序号到数据集中
    Dim nIndex As Long
    Dim lng序号 As Long
    Dim i As Long, j As Long
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim lng限制 As Long
    Dim bln预约 As Boolean
    Dim str最大时间 As String
    If Not mblnChange Then MoveAssign = True: Exit Function
    
    nIndex = GetVsGridIndex(str限制项目)
    
    '删掉没有使用部分
    mrsAssign.Filter = "限制项目='" & str限制项目 & "' and 已使用=0"
    If mrsAssign.RecordCount > 0 Then
        Do While Not mrsAssign.EOF
            mrsAssign.Delete
            mrsAssign.MoveNext
        Loop
    End If
    
    If Not mbln序号控制 Then
        With vsTime(nIndex)
          lng序号 = 0
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                   If .TextMatrix(i, j) <> "" Then
                    
                    str开始时间 = Split(.TextMatrix(i, j), "-")(0)
                    str结束时间 = Split(.TextMatrix(i, j), "-")(1)
                    lng限制 = Val(.TextMatrix(i, j + 1))
                    lng序号 = lng序号 + 1
                    bln预约 = True
                    
                    str最大时间 = ""
                    If Not mrsHistory Is Nothing Then
                        mrsHistory.Filter = "限制项目='" & str限制项目 & "'"
                        If mrsHistory.RecordCount = 0 Then
                            str最大时间 = ""
                            mrsHistory.Filter = 0
                        Else
                            str最大时间 = Nvl(mrsHistory!发生时间)
                            mrsHistory.Filter = 0
                        End If
                    End If
                    
                    If (str最大时间 <> "" And str开始时间 > str最大时间) Or str最大时间 = "" Then
                        With mrsAssign
                            .AddNew
                            !限制项目 = str限制项目
                            !开始时间 = str开始时间
                            !结束时间 = str结束时间
                            !时间段 = str开始时间 & "-" & str结束时间
                            !限制数量 = lng限制
                            !序号 = lng序号
                            !已使用 = 0
                            !是否预约 = 1
                            .Update
                        End With
                    End If
                   End If
                Next
            Next
        End With
        mblnChange = False
        MoveAssign = True
        Exit Function
    End If
    
    
    '序号控制
    
    With vsTime(nIndex)
        For i = 0 To .Rows - 1 Step 2
            For j = 1 To .Cols - 1
                If Trim(.TextMatrix(i, j)) <> "" Then
                        str开始时间 = Split(.TextMatrix(i + 1, j) & "-", "-")(0)
                        str结束时间 = Split(.TextMatrix(i + 1, j) & "-", "-")(0)
                        lng序号 = Val(.TextMatrix(i, j))
                        lng限制 = 1
                        bln预约 = .Cell(flexcpForeColor, i, j) = vbBlue
                    If .Cell(flexcpFontUnderline, i, j) = False Then
                       
                        With mrsAssign
                            .AddNew
                            !限制项目 = str限制项目
                            !开始时间 = str开始时间
                            !结束时间 = str结束时间
                            !时点 = Format(str开始时间, "hh:00:00")
                            !时间段 = str开始时间 & "-" & str结束时间
                            !限制数量 = lng限制
                            !序号 = lng序号
                            !已使用 = 0
                            !是否预约 = IIf(bln预约, 1, 0)
                            .Update
                        End With
                    ElseIf .Cell(flexcpFontUnderline, i, j) Then
                        ' 固定的信息,可能改变是否预约,现在也只可改变是否预约
                        With mrsAssign
                            .Filter = "序号=" & lng序号 & " And 开始时间='" & Format(str开始时间, "hh:mm:00") & "'"
                            If .RecordCount > 0 Then
                                !是否预约 = IIf(bln预约, 1, 0)
                                .Update
                            End If
                        End With
                    End If
                End If
            Next
        Next
    End With
    mblnChange = False
    MoveAssign = True
    Exit Function
End Function
Private Function ConvertToDate(ByVal strDate As String, Optional ByVal haveYear = False) As String
    '**********************************************************
    '把字符串转换成oracle数据库能够识别的日期
    '**********************************************************
    Select Case haveYear
    Case True:
        ConvertToDate = "To_Date('" & strDate & "', 'YYYY-MM-DD HH24:MI:SS')"
    Case False:
        ConvertToDate = "To_Date('" & strDate & "', 'HH24:MI:SS')"
    End Select
End Function

Private Sub SetStyle(ByVal bln序号控制 As Boolean, ByVal lngIndex As Long)
    '设置
    Dim i As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    If lngIndex > vsTime.UBound Then Exit Sub
    If Not mblnInit Then Exit Sub
    With vsTime(lngIndex)
        If bln序号控制 Then
             
            If .Cols <= 1 Then Exit Sub
            .Rows = 0
            .FixedCols = 1
            .MergeCellsFixed = flexMergeFree
            .MergeCol(0) = True
            .FixedAlignment(0) = flexAlignRightTop
            .ColAlignment(0) = flexAlignRightTop
            lngWidth = 1275
'            lngHeight = 800
'            For i = 1 To .Cols - 1
'                .ColWidth(i) = lngWidth
'                .ColAlignment(i) = 4
'            Next
'            .ColAlignment(0) = 3
'            .ColWidth(0) = lngWidth
'            For i = 0 To .Rows - 1
'                 .RowHeight(i) = lngHeight
'            Next
'           If .Rows > 0 And .Cols > 0 Then
'                .Cell(flexcpFontBold, 0, 1, .Rows - 1, .Cols - 1) = True
'                .Cell(flexcpFontSize, 0, 1, .Rows - 1, .Cols - 1) = 9
'                .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 18
'           End If
           
        Else
             .Clear
             .Cols = 8: .Rows = 1
             .MergeCol(0) = False
            .FixedCols = 0
            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedRows = 1
            
            .RowHeightMax = 400: .RowHeightMin = 400
            For i = 0 To .Cols - 1 Step 2
              .TextMatrix(0, i) = "时间段"
            Next
            For i = 1 To .Cols - 1 Step 2
              .TextMatrix(0, i) = "预约人数"
            Next
            For i = 0 To .Cols - 1
               .ColAlignment(i) = flexAlignCenterCenter
               .ColWidth(i) = 1200
            Next
        End If
'        If bln时段 Then
'            .Clear
'            .FixedCols = 1
'            .FixedRows = 0
'            .Rows = 1
'        Else
'
'        End If
    End With
End Sub

Private Sub setVsGridSNStyle(ByVal lngIndex As Long)
 '如果分时段在vsFex表哥填充好数据后需要重新设置表哥样式
 '****************************************
'对表格样式进行设置
'****************************************
    Dim i           As Long
    Dim lngWidth    As Long
    Dim X           As Long
    Dim Y           As Long
    Dim j           As Long
    Dim lngHeight   As Long
   
    If vsTime(lngIndex).Cols <= 1 Then Exit Sub
    If mbln序号控制 Then
        With vsTime(lngIndex)
            For i = 1 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
             Next
             .ColWidth(0) = 1200
             .FixedAlignment(0) = flexAlignRightTop
             .ColAlignment(0) = flexAlignRightTop
             If .Rows > 0 Then
                .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
                .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
             End If
    '对时间段设置间隔背景
    
     
         End With
    Else
    
    End If
    With vsTime(lngIndex)
         If (mbln序号控制 And .Rows = 0) Or (mbln序号控制 = False And .Rows = 1) Then Exit Sub
         For i = IIf(mbln序号控制, 0, 1) To .Rows - 1 Step 2
             .Cell(flexcpBackColor, i, IIf(mbln序号控制, 1, 0), i, .Cols - 1) = &HE0E0D3
         Next
    End With

End Sub

 
 
 
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : LoadTimePlan
'Description :
'Author      : 李光福
'Date        : 05-November-2012 14:41:41
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'str限制项目           String            ByVal             星期           .
'Output      :  设置时间断是否成功
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function LoadTimePlan(ByVal str限制项目 As String) As Boolean
    Dim nIndex As Integer
    Dim i As Long, r As Long
    Dim strTime As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim str时点 As String
    Dim strData As String
    If mrsAssign Is Nothing Then Exit Function
    nIndex = GetVsGridIndex(str限制项目)
    cmd预约(nIndex).Visible = False
    cmd删除(nIndex).Visible = False
    If Not mbln序号控制 Then
        With vsTime(nIndex)
             
             mrsAssign.Filter = "限制项目='" & str限制项目 & "'"
            mrsAssign.Sort = "序号 asc "
               r = 1: i = -1
            Do While Not mrsAssign.EOF
                i = i + 1
                If i * 2 > .Cols - 2 Then r = r + 1: i = 0
                strData = Val(Nvl(mrsAssign!限制数量))
                strTime = mrsAssign!时间段
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                If Val(Nvl(mrsAssign!已使用)) = 1 Then
                    .Cell(flexcpFontUnderline, r, i * 2, r, i * 2 + 1) = True
                Else
                   '不做颜色处理
                End If
                mrsAssign.MoveNext
            Loop
             mrsAssign.Filter = 0
        End With
        LoadTimePlan = True
        Exit Function
    End If
    
    '-序号控制
    With vsTime(nIndex)
        .Cols = 1: .FixedCols = 1
        .Rows = 0: .FixedRows = 0
        .Cols = 2: .Clear
        lngRow = -1: lngCol = 0
        mrsAssign.Filter = "限制项目='" & str限制项目 & "'"
        If mrsAssign.RecordCount = 0 Then mrsAssign.Filter = 0: Exit Function
        i = 1
        mrsAssign.Sort = "序号 asc "
        Do While Not mrsAssign.EOF
             lngCol = lngCol + 1
             If str时点 <> Nvl(mrsAssign!时点) Then lngRow = lngRow + 2: lngCol = 1
             If lngCol = 1 Then
                str时点 = Nvl(mrsAssign!时点)
                If lngRow > .Rows - 1 Then .Rows = .Rows + 2
                 .TextMatrix(lngRow - 1, 0) = Format(str时点, "hh:mm")
                 .TextMatrix(lngRow, 0) = Format(str时点, "hh:mm")
             End If
             strData = mrsAssign!序号
             strTime = mrsAssign!时间段
            If lngCol > .Cols - 1 Then .Cols = .Cols + 1
            If lngRow > .Rows - 1 Then .Rows = .Rows + 2
             .TextMatrix(lngRow - 1, lngCol) = strData
             .TextMatrix(lngRow, lngCol) = strTime
            If Val(Nvl(mrsAssign!是否预约)) = 1 Then
                .Cell(flexcpForeColor, lngRow - 1, lngCol, lngRow, lngCol) = vbBlue
                .Cell(flexcpFontBold, lngRow - 1, lngCol, lngRow, lngCol) = True
            End If
            If Val(Nvl(mrsAssign!已使用)) = 1 Then
                    .Cell(flexcpFontUnderline, lngRow - 1, lngCol, lngRow, lngCol) = True
            Else
               '不做颜色处理
            End If
            mrsAssign.MoveNext
        Loop
        If .Rows = 0 Then .Rows = 1
    End With
End Function
 
 

Private Sub txtTimeOut_Change()
    If Val(txtTimeOut.Text) > 1440 Then txtTimeOut.Text = 1440
End Sub
 
Private Sub vsTime_BeforeRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Not mbln序号控制 Then
        vsTime(Index).Editable = IIf(NewCol Mod 2 = 1, flexEDKbd, flexEDNone)
          cmd预约(Index).Visible = False: Exit Sub
    End If
    If NewRow < 0 Or NewCol < 0 Then Exit Sub
    
    SetCtrlMove Index, NewRow - (NewRow) Mod 2, NewCol
    If mbln序号控制 Then vsTime(Index).Editable = flexEDNone: Exit Sub
    
    With vsTime(Index)
        .Editable = IIf(NewCol Mod 2 = 1, flexEDKbd, flexEDNone)
    End With
End Sub

Private Sub SetCtrlMove(ByVal Index As Integer, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnDel As Boolean
    With vsTime(Index)
        If mbln序号控制 Then
            If Trim(.TextMatrix(NewRow, NewCol)) = "" Then
                cmd删除(Index).Visible = False
                cmd预约(Index).Visible = False
                Exit Sub
            End If
            cmd删除(Index).Left = .Cell(flexcpLeft, NewRow, NewCol) + .Cell(flexcpWidth, NewRow, NewCol) - cmd删除(Index).Width
            If .Row Mod 2 <> 0 Then
                cmd删除(Index).Top = .Cell(flexcpTop, NewRow, NewCol)
            Else
                cmd删除(Index).Top = .Cell(flexcpTop, NewRow, NewCol)
            End If
            cmd预约(Index).Left = .Cell(flexcpLeft, NewRow, NewCol)
            cmd预约(Index).Top = cmd删除(Index).Top
            If NewCol < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(NewRow, NewCol + 1)) = ""
            Else
                blnDel = True
            End If
             
            blnDel = blnDel And Trim(.TextMatrix(NewRow, NewCol)) <> "" And Not .Cell(flexcpFontUnderline, NewRow, NewCol)
            cmd删除(Index).Visible = blnDel And mbln序号控制
            cmd预约(Index).Visible = True 'Val(txt限约.Text) <> 0
        Else
            cmd预约(Index).Left = .Cell(flexcpTop, NewRow, NewCol)
            cmd预约(Index).Top = .Cell(flexcpLeft, NewRow, NewCol)
            cmd预约(Index).Visible = False
'            cmd预约.Visible = Val(txt限约.Text) <> 0
        End If
    End With
End Sub

 

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : vsTime_KeyDown
'Description : 网格 按键事件,主要用于,序号控制 分时段,安
'Author      : 李光福
'Date        : 09-11-2012 05:58:34
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'Index             Integer           ByRef                .
'KeyCode           Integer           ByRef                .
'Shift             Integer           ByRef                .
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Sub vsTime_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'
    If Not mbln序号控制 Then Exit Sub
     
    With vsTime(Index)
           
        If (.Row < 0 Or .Col < 1) Or (.Row > .Rows - 1 Or .Col > .Cols - 1) Then Exit Sub '没在有效单元格内
        If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Sub
        If KeyCode = 13 Then
            Call cmd预约_Click(Index)
            Exit Sub
        End If
        
        If KeyCode = 46 Then 'delete
            '问题号:51429
            If cmd删除(Index).Visible = False Then Exit Sub
            If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Sub
            Call cmd删除_Click(Index)
        End If
     End With
End Sub

Private Sub vsTime_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13) Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vsTime_LostFocus(Index As Integer)
 '
 If Trim(vsTime(Index).EditText) <> "" Then
    With vsTime(Index)
        .TextMatrix(.Row, .Col) = .EditText
        mblnChange = True
    End With
 End If
'If mblnChange Then Stop
End Sub

Private Sub vsTime_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mblnChange = True
End Sub


Public Property Get IsInit() As Boolean
Attribute IsInit.VB_Description = "是否经过了初始化"
    IsInit = mblnInit
End Property



'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : zl_CheckMoveAssign
'Description : 检查是否序号分配是否改变,如果已改变则,调用函数,重新分配序号
'Author      : 李光福
'Date        : 14-11-2012 10:53:40
'Input       :
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Public Function zl_CheckMoveAssign(Optional ByVal lngIndex As Long = -1) As Boolean
    Dim str限制项目 As String
    If lngIndex = -1 Then lngIndex = mlngSelIndex
    If lngIndex = -1 Then zl_CheckMoveAssign = True: Exit Function
    If Not mblnChange Then zl_CheckMoveAssign = True: Exit Function
    
    If lngIndex < 0 Or lngIndex > 6 Then Exit Function
    If Not VsTimeValidate(lngIndex) Then Exit Function
    
    str限制项目 = GetVsGridCaption(lngIndex)
    zl_CheckMoveAssign = MoveAssign(str限制项目)
End Function

Public Property Get 序号控制() As Boolean
        序号控制 = mbln序号控制
End Property

Public Property Let 序号控制(ByVal vNewValue As Boolean)
        mbln序号控制 = vNewValue
End Property
