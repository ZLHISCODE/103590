VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO0EA1~1.OCX"
Begin VB.Form frmEPRAuditPati 
   Caption         =   "病人病历书写审查"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10455
   Icon            =   "frmEPRAuditPati.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10455
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6555
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRAuditPati.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15558
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgPati 
      Height          =   5655
      Left            =   105
      TabIndex        =   1
      Top             =   720
      Width           =   3630
      _cx             =   6403
      _cy             =   9975
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgAudit 
      Height          =   2400
      Left            =   3945
      TabIndex        =   2
      Top             =   750
      Width           =   6285
      _cx             =   11086
      _cy             =   4233
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
      WordWrap        =   -1  'True
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1755
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRAuditPati.frx":70E4
      Left            =   525
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRAuditPati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'常量
'-----------------------------------------------------
Private Enum mCol
    标志 = 0: 事件缘由: 应写病历: 监测点: 基点时间: 要求时间: 完成时间: 完成记录id: 当前时间: 备注说明
End Enum

Const conPane_Pati = 1
Const conPane_Audit = 2
Const conPane_Word = 3

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mlngDeptId As Long      '科室id
Private mstrDeptName As String  '科室名
Private mintKind As Integer     '病历种类
Private mstrDateFrom As String  '开始日期
Private mstrDateTo As String    '结束日期
Private mstrEvent As String     '病人事件范围

Private WithEvents mfrmWord As frmDockEPRContent     '病历内容窗格
Attribute mfrmWord.VB_VarHelpID = -1

'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rsTemp As New ADODB.Recordset
Dim strSQL As String
Dim lngCount As Long, lngRow As Long, lngCol As Long

Public Sub ShowMe(frmParent As Object, lngDeptId As Long, strDeptName As String, _
    intKind As Integer, strDateFrom As String, strDateTo As String, _
    Optional strEvent As String)
    mlngDeptId = lngDeptId: mstrDeptName = strDeptName
    mintKind = intKind: mstrDateFrom = strDateFrom: mstrDateTo = strDateTo
    mstrEvent = strEvent
    Me.Caption = Me.Caption & " - " & mstrDeptName
    
    Call RefreshData
    Me.Show vbModal, frmParent
End Sub

Private Sub RefreshData()
    Select Case mintKind
    Case 1
        Select Case mstrEvent
        Case "门诊"
            strSQL = "Select 病人id, ID, 门诊号, 姓名, 性别, To_Char(执行时间, 'yyyy-mm-dd hh24:mi') As 就诊时间, 执行人 As 医生" & vbNewLine & _
                    "From 病人挂号记录" & vbNewLine & _
                    "Where 执行部门id + 0 = [1] And Nvl(执行状态, 0) <> 0 And Nvl(急诊, 0) <> 1 And" & vbNewLine & _
                    "      登记时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By 执行时间"
        Case "急诊"
            strSQL = "Select 病人id, ID, 门诊号, 姓名, 性别, To_Char(执行时间, 'yyyy-mm-dd hh24:mi') As 就诊时间, 执行人 As 医生" & vbNewLine & _
                    "From 病人挂号记录" & vbNewLine & _
                    "Where 执行部门id + 0 = [1] And Nvl(执行状态, 0) <> 0 And Nvl(急诊, 0) = 1 And" & vbNewLine & _
                    "      登记时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By 执行时间"
        Case Else
            strSQL = "Select 病人id, ID, 门诊号, 姓名, 性别, To_Char(执行时间, 'yyyy-mm-dd hh24:mi') As 就诊时间, 执行人 As 医生" & vbNewLine & _
                    "From 病人挂号记录" & vbNewLine & _
                    "Where 执行部门id + 0 = [1] And Nvl(执行状态, 0) <> 0 And" & vbNewLine & _
                    "      登记时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By 执行时间"
        End Select
    Case 2
        Select Case mstrEvent
        Case "入院"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.入院时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select 病人id, 主页id, To_Char(Max(开始时间), 'yyyy-mm-dd hh24:mi') As 入院时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 科室id + 0 = [1] And 开始原因 In (1, 2, 9) And 开始时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By 病人id, 主页id) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.入院时间"
        Case "转入"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.转入时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id, To_Char(开始时间, 'yyyy-mm-dd hh24:mi') As 转入时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 科室id + 0 = [1] And 开始原因 = 3 And 开始时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.转入时间"
        Case "出院"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, To_Char(P.出院日期, 'yyyy-mm-dd hh24:mi') As 出院日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.出院科室id + 0 = [1] And P.出院方式 <> '死亡' And" & vbNewLine & _
                    "      P.出院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.出院日期"
        Case "死亡"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, To_Char(P.出院日期, 'yyyy-mm-dd hh24:mi') As 死亡日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.出院科室id + 0 = [1] And P.出院方式 = '死亡' And" & vbNewLine & _
                    "      P.出院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.出院日期"
        Case "转出"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.转出时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id, To_Char(终止时间, 'yyyy-mm-dd hh24:mi') As 转出时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 科室id + 0 = [1] And 终止原因 = 3 And 终止时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.转出时间"
        Case "手术"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, 手术时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select R.病人id, R.主页id, To_Char(S.首次时间, 'yyyy-mm-dd hh24:mi') As 手术时间" & vbNewLine & _
                    "       From 病人医嘱记录 R, 病人医嘱发送 S" & vbNewLine & _
                    "       Where R.ID = S.医嘱id And R.诊疗类别 = 'F' And R.相关id Is Null And R.医嘱期效 = 1 And" & vbNewLine & _
                    "             (R.医嘱状态 = 8 Or R.医嘱状态 = 9) And R.病人科室id + 0 = [1] And" & vbNewLine & _
                    "             S.首次时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.手术时间"
        Case Else
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, P.入院日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 科室id = [1] And" & vbNewLine & _
                    "             (开始原因 In (1, 2, 3, 9) And 开始时间 <= To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or" & vbNewLine & _
                    "             终止原因 In (1, 3, 10) And (终止时间 >= To_Date([2], 'yyyy-mm-dd') Or 终止时间 Is Null))) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By P.入院日期"
        End Select
    Case 4
        Select Case mstrEvent
        Case "入院"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.入院时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select 病人id, 主页id, To_Char(Max(开始时间), 'yyyy-mm-dd hh24:mi') As 入院时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 病区id + 0 = [1] And 开始原因 In (1, 2, 9) And 开始时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By 病人id, 主页id) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.入院时间"
        Case "转入"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.转入时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id, To_Char(开始时间, 'yyyy-mm-dd hh24:mi') As 转入时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 病区id + 0 = [1] And 开始原因 = 3 And 开始时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.转入时间"
        Case "出院"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, To_Char(P.出院日期, 'yyyy-mm-dd hh24:mi') As 出院日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.当前病区id + 0 = [1] And P.出院方式 <> '死亡' And" & vbNewLine & _
                    "      P.出院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.出院日期"
        Case "死亡"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, To_Char(P.出院日期, 'yyyy-mm-dd hh24:mi') As 死亡日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.当前病区id + 0 = [1] And P.出院方式 = '死亡' And" & vbNewLine & _
                    "      P.出院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.出院日期"
        Case "转出"
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.转出时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id, To_Char(终止时间, 'yyyy-mm-dd hh24:mi') As 转出时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 病区id + 0 = [1] And 终止原因 = 3 And 终止时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.转出时间"
        Case Else
            strSQL = "Select P.病人id, P.主页id, P.住院号, I.姓名, I.性别, P.入院日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 病区id = [1] And" & vbNewLine & _
                    "             (开始原因 In (1, 2, 3, 9) And 开始时间 <= To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or" & vbNewLine & _
                    "             终止原因 In (1, 3, 10) And (终止时间 >= To_Date([2], 'yyyy-mm-dd') Or 终止时间 Is Null))) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By P.入院日期"
        End Select
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo)
    With Me.vfgPati
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(1) = 0: .ColHidden(1) = True
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
    End With
    Call vfgPati_RowColChange
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    If Me.ActiveControl.Name = Me.vfgPati.Name Then
        Set objPrint.Body = Me.vfgPati
        objPrint.Title.Text = mstrDeptName & mstrEvent & "病人清单"
    Else
        Set objPrint.Body = Me.vfgAudit
        objPrint.Title.Text = "病人病历时限报告"
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add(Me.vfgPati.TextMatrix(Me.vfgPati.FixedRows - 1, 2) & ":" & Me.vfgPati.TextMatrix(Me.vfgPati.Row, 2))
        Call objAppRow.Add("姓名:" & Me.vfgPati.TextMatrix(Me.vfgPati.Row, 3))
        Call objPrint.UnderAppRows.Add(objAppRow)
    End If
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_File_Open:
        Dim f As New frmEPRView
        f.ShowMe Me, CLng(Me.vfgAudit.TextMatrix(Me.vfgAudit.Row, mCol.完成记录id)), True
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview:  Call zlRptPrint(0)
    Case conMenu_File_Print:    Call zlRptPrint(1)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_Exit:     Unload Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh: Call RefreshData
    Case conMenu_View_Jump
    Case conMenu_Tool_Monitor
        Dim lngRecordId As Long
        With Me.vfgAudit
            lngRecordId = Val(.TextMatrix(.Row, mCol.完成记录id))
        End With
        Call frmEPRAuditMonitor.zlRefList(lngRecordId)
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)

    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Open, conMenu_Tool_Monitor
        With Me.vfgAudit
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.完成记录id)) > 0)
        End With
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Me.ActiveControl.Name = Me.vfgPati.Name Then
            Control.Enabled = (Me.vfgPati.Rows > Me.vfgPati.FixedRows)
        Else
            Control.Enabled = (Me.vfgAudit.Rows > Me.vfgAudit.FixedRows)
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Pati
        Item.Handle = Me.vfgPati.hWnd
    Case conPane_Audit
        Item.Handle = Me.vfgAudit.hWnd
    Case conPane_Word
        If mfrmWord Is Nothing Then Set mfrmWord = New frmDockEPRContent
        Item.Handle = mfrmWord.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagStretched
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开(&O)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "跳转(&J)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    cbrMenuBar.ID = conMenu_ToolPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "内容监测(&T)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    With Me.cbsThis.ActiveMenuBar.Controls
        Select Case mintKind
        Case 1: Set cbrControl = .Add(xtpControlLabel, 0, "门诊病历(" & mstrDateFrom & "至" & mstrDateTo & ")")
        Case 2: Set cbrControl = .Add(xtpControlLabel, 0, "住院病历(" & mstrDateFrom & "至" & mstrDateTo & ")")
        Case 4: Set cbrControl = .Add(xtpControlLabel, 0, "护理病历(" & mstrDateFrom & "至" & mstrDateTo & ")")
        End Select
        cbrControl.flags = xtpFlagRightAlign
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Jump
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "内容监测"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置显示停靠窗格
    If mfrmWord Is Nothing Then Set mfrmWord = New frmDockEPRContent
    
    Dim panThis As Pane, panChild As Pane
    Set panThis = dkpMan.CreatePane(conPane_Pati, 300, 400, DockLeftOf, Nothing)
    panThis.Title = mstrEvent & "病人列表"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_Audit, 700, 100, DockRightOf, Nothing)
    panThis.Title = "病历时限审查"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panChild = dkpMan.CreatePane(conPane_Word, 700, 300, DockBottomOf, panThis)
    panChild.Title = "病历内容"
    panChild.Options = PaneNoCaption

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmWord: Set mfrmWord = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vfgAudit_GotFocus()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub vfgAudit_RowColChange()
    Dim lngRecordId As Long
    
    Err = 0: On Error Resume Next
    With Me.vfgAudit
        lngRecordId = Val(.TextMatrix(.Row, mCol.完成记录id))
    End With
    Err = 0: On Error GoTo 0
    If Me.Tag <> "" Then Exit Sub
    Call mfrmWord.zlRefresh(lngRecordId, "", True)
End Sub

Private Sub vfgPati_GotFocus()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub vfgPati_RowColChange()
    Dim lngPatiID As Long, lngPageId As Long
    Dim lngBalance As Long
    
    If Me.Tag <> "" Then Exit Sub
    lngPatiID = Me.vfgPati.TextMatrix(Me.vfgPati.Row, 0)
    lngPageId = Me.vfgPati.TextMatrix(Me.vfgPati.Row, 1)
    
    gstrSQL = "Zl_病历时限监测_Neaten(" & lngPatiID & "," & lngPageId & "," & mintKind & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '提取时限监测数据
    gstrSQL = "Select To_Char(事件时间, 'yyyy-mm-dd hh24:mi ') || 变动事件 As 事件缘由, 病历编号 || '-' || 病历名称 As 应写病历," & _
            "        Decode(唯一, 1, '书写', '第' || 周期号 || '次书写') As 监测点, 基点时间, 要求时间, 完成时间, 完成记录id, Sysdate As 当前时间, Null As 备注说明" & _
            " From 病历时限监测" & _
            " Where 病人id = [1] And 主页id = [2] And (病历种类 = [3] Or 病历种类 in (5,6) And [3]<>4) And 要求时间 - Sysdate < 2" & _
            " Order By 病历种类,事件时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId, mintKind)
    With Me.vfgAudit
        .Clear
        Set .DataSource = rsTemp
        
        .MergeCells = flexMergeFree: .MergeCol(mCol.事件缘由) = True: .MergeCol(mCol.应写病历) = True
        .ColWidth(mCol.标志) = 250: .ColWidth(mCol.基点时间) = 1100: .ColWidth(mCol.要求时间) = 1100: .ColWidth(mCol.完成时间) = 1100
        .ColWidth(mCol.完成记录id) = 0: .ColWidth(mCol.当前时间) = 0: .ColWidth(mCol.备注说明) = 2200
        
        .FixedAlignment(mCol.标志) = flexAlignCenterCenter
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftTop
        Next
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mCol.完成时间) = "" Then
                If .TextMatrix(lngCount, mCol.完成记录id) = "" Then
                    .TextMatrix(lngCount, mCol.备注说明) = "未书写"
                Else
                    .TextMatrix(lngCount, mCol.备注说明) = "正在书写"
                End If
                lngBalance = Int((CDate(.TextMatrix(lngCount, mCol.当前时间)) - CDate(.TextMatrix(lngCount, mCol.要求时间))) * 24)
                .TextMatrix(lngCount, mCol.标志) = "！"
                If lngBalance >= 0 Then
                    .Cell(flexcpForeColor, lngCount, mCol.标志, lngCount, mCol.标志) = RGB(255, 0, 0)
                    .TextMatrix(lngCount, mCol.备注说明) = .TextMatrix(lngCount, mCol.备注说明) & IIf(lngBalance = 0, "", ",已超过" & lngBalance & "小时")
                    .Cell(flexcpForeColor, lngCount, mCol.备注说明, lngCount, mCol.备注说明) = RGB(255, 0, 0)
                Else
                    If Abs(lngBalance) < 4 Then
                        .Cell(flexcpForeColor, lngCount, mCol.标志, lngCount, mCol.标志) = RGB(128, 128, 0)
                        .TextMatrix(lngCount, mCol.备注说明) = .TextMatrix(lngCount, mCol.备注说明) & ",剩余" & Abs(lngBalance) & "小时,请尽快完成"
                    Else
                        .Cell(flexcpForeColor, lngCount, mCol.标志, lngCount, mCol.标志) = RGB(0, 0, 255)
                        .TextMatrix(lngCount, mCol.备注说明) = .TextMatrix(lngCount, mCol.备注说明) & ",剩余" & Abs(lngBalance) & "小时,请按时完成"
                    End If
                End If
            Else
                lngBalance = Int((CDate(.TextMatrix(lngCount, mCol.完成时间)) - CDate(.TextMatrix(lngCount, mCol.要求时间))) * 24)
                If lngBalance > 0 Then
                    .TextMatrix(lngCount, mCol.标志) = "⊕"
                    .Cell(flexcpForeColor, lngCount, mCol.标志, lngCount, mCol.标志) = RGB(255, 0, 0)
                    .TextMatrix(lngCount, mCol.备注说明) = "完成,但超过" & lngBalance & "小时"
                    .Cell(flexcpForeColor, lngCount, mCol.备注说明, lngCount, mCol.备注说明) = RGB(255, 0, 0)
                Else
                    .TextMatrix(lngCount, mCol.备注说明) = "正常完成"
                End If
            End If
            .TextMatrix(lngCount, mCol.基点时间) = Format(.TextMatrix(lngCount, mCol.基点时间), "MM-dd hh:mm")
            .TextMatrix(lngCount, mCol.要求时间) = Format(.TextMatrix(lngCount, mCol.要求时间), "MM-dd hh:mm")
            .TextMatrix(lngCount, mCol.完成时间) = Format(.TextMatrix(lngCount, mCol.完成时间), "MM-dd hh:mm")
        Next
    End With
    Call vfgAudit_RowColChange
End Sub
