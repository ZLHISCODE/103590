VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO0EA1~1.OCX"
Begin VB.Form frmEPRAuditFile 
   Caption         =   "病历文件分类审查"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   Icon            =   "frmEPRAuditFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9630
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRAuditFile.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14102
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
   Begin VSFlex8Ctl.VSFlexGrid vfgFile 
      Height          =   1980
      Left            =   75
      TabIndex        =   1
      Top             =   750
      Width           =   4500
      _cx             =   7937
      _cy             =   3492
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
   Begin VSFlex8Ctl.VSFlexGrid vfgEPRs 
      Height          =   2925
      Left            =   60
      TabIndex        =   2
      Top             =   2835
      Width           =   4500
      _cx             =   7937
      _cy             =   5159
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   2235
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRAuditFile.frx":0E1C
      Left            =   525
      Top             =   150
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRAuditFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'常量
'-----------------------------------------------------
Const conPane_File = 1
Const conPane_EPRs = 2
Const conPane_Word = 3

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mlngDeptId As Long      '科室id
Private mstrDeptName As String  '科室名
Private mintKind As Integer     '病历种类
Private mstrDateFrom As String  '开始日期
Private mstrDateTo As String    '结束日期

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

Public Sub ShowMe(frmParent As Object, lngDeptId As Long, strDeptName As String, intKind As Integer, strDateFrom As String, strDateTo As String)
    mlngDeptId = lngDeptId: mstrDeptName = strDeptName
    mintKind = intKind: mstrDateFrom = strDateFrom: mstrDateTo = strDateTo
    Me.Caption = Me.Caption & " - " & mstrDeptName
    
    Call RefreshData
    Me.Show vbModal, frmParent
End Sub

Private Sub RefreshData()
    Dim intOut24h As Byte    '是否区分24小时出院或死亡：0-不区分,1-区分；根据是否定义24小时事件对应病历确定
    Select Case mintKind
    Case 1
        strSQL = "Select F.ID, F.编号, F.名称, F.事件 || '时书写' As 要求, P.人次 As 应写数, W.已完成, W.在书写" & vbNewLine & _
                "From (Select F.ID, F.编号, F.名称, F.事件, F.唯一, F.书写时限" & vbNewLine & _
                "       From (Select F.ID, F.编号, F.名称, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "              From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "              Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 1) F" & vbNewLine & _
                "       Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [1]) F," & vbNewLine & _
                "     (Select E.事件, Decode(E.事件, '门诊', 门诊, '初诊', 初诊, '复诊', 复诊, '急诊', 急诊) As 人次" & vbNewLine & _
                "       From (Select Sum(Decode(急诊, 1, 0, 1)) As 门诊, Sum(Decode(急诊, 1, 0, Decode(复诊, 1, 0, 1))) As 初诊," & vbNewLine & _
                "                     Sum(Decode(急诊, 1, 0, Decode(复诊, 1, 1, 0))) As 复诊, Sum(Decode(急诊, 1, 1, 0)) As 急诊" & vbNewLine & _
                "              From 病人挂号记录" & vbNewLine & _
                "              Where 执行部门id = [1] And Nvl(执行状态, 0) <> 0 And 登记时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "                    To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                "            (Select Decode(Rownum, 1, '门诊', 2, '初诊', 3, '复诊', 4, '急诊') As 事件 From 病历书写事件 Where Rownum < 5) E) P," & vbNewLine & _
                "     (Select 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成, Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "       From 电子病历记录" & vbNewLine & _
                "       Where 病历种类 = 1 And 科室id + 0 = [1] And 创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By 文件id) W" & vbNewLine & _
                "Where F.事件 = P.事件 And P.人次 > 0 And F.ID = W.文件id(+)" & vbNewLine & _
                "Order By F.编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo)
    Case 2
        strSQL = "Select Sign(Nvl(Count(*), 0))" & vbNewLine & _
                "From (Select F.ID, F.通用, A.科室id" & vbNewLine & _
                "       From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "       Where F.ID = A.文件id(+) And F.ID = Q.文件id And Q.事件 In ('24小时出院', '24小时死亡') And F.种类 = 2) F" & vbNewLine & _
                "Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId)
        If rsTemp.RecordCount <= 0 Then
            intOut24h = 0
        Else
            intOut24h = rsTemp.Fields(0).Value
        End If
        strSQL = "Select F.ID, F.编号, F.名称, F.事件 || Decode(Sign(F.书写时限), -1, '前', '后') || '书写' As 要求," & vbNewLine & _
                "       Decode(F.唯一, 1, To_Char(P.人次), '<循环>') As 应写数, W.已完成, W.在书写" & vbNewLine & _
                "From (Select F.ID, F.编号, F.名称, F.事件, F.唯一, F.书写时限" & vbNewLine & _
                "       From (Select F.ID, F.编号, F.名称, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "              From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "              Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 2) F" & vbNewLine & _
                "       Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [1]) F," & vbNewLine
        If intOut24h = 1 Then
            strSQL = strSQL & "     (Select E.事件, '后' As 时机, Decode(E.事件, '入院', 入院, '首次入院', 首次入院, '再次入院', 再次入院) As 人次" & vbNewLine & _
                    "       From (Select Count(*) As 入院, Sum(Decode(再入院, 1, 0, 1)) As 首次入院," & vbNewLine & _
                    "                     Sum(Decode(再入院, 1, '再次入院', 0)) As 再次入院" & vbNewLine & _
                    "              From 病案主页" & vbNewLine & _
                    "              Where 入院科室id + 0 = 36 And Nvl(出院日期, Sysdate + 1) - 入院日期 > 1 And" & vbNewLine & _
                    "                    入院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                    "            (Select Decode(Rownum, 1, '入院', 2, '首次入院', 3, '再次入院') As 事件 From 病历书写事件 Where Rownum < 4) E" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(Sign(出院日期 - 入院日期 - 1), -1, Decode(出院方式, '死亡', '24小时死亡', '24小时出院')," & vbNewLine & _
                    "                      Decode(出院方式, '死亡', '死亡', '出院')) As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                    "       From 病案主页" & vbNewLine & _
                    "       Where 出院科室id + 0 = [1] And 出院日期 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By Decode(Sign(出院日期 - 入院日期 - 1), -1, Decode(出院方式, '死亡', '24小时死亡', '24小时出院')," & vbNewLine & _
                    "                        Decode(出院方式, '死亡', '死亡', '出院'))" & vbNewLine
        Else
            strSQL = strSQL & "     (Select E.事件, '后' As 时机, Decode(E.事件, '入院', 入院, '首次入院', 首次入院, '再次入院', 再次入院) As 人次" & vbNewLine & _
                    "       From (Select Count(*) As 入院, Sum(Decode(再入院, 1, 0, 1)) As 首次入院," & vbNewLine & _
                    "                     Sum(Decode(再入院, 1, '再次入院', 0)) As 再次入院" & vbNewLine & _
                    "              From 病案主页" & vbNewLine & _
                    "              Where 入院科室id + 0 = 36 And" & vbNewLine & _
                    "                    入院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                    "            (Select Decode(Rownum, 1, '入院', 2, '首次入院', 3, '再次入院') As 事件 From 病历书写事件 Where Rownum < 4) E" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(出院方式, '死亡', '死亡', '出院') As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                    "       From 病案主页" & vbNewLine & _
                    "       Where 出院科室id + 0 = [1] And 出院日期 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By Decode(出院方式, '死亡', '死亡', '出院')" & vbNewLine
        End If
        strSQL = strSQL & "       Union All" & vbNewLine & _
                "       Select Decode(开始原因, 3, '转科', 7, '交班') As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病人变动记录" & vbNewLine & _
                "       Where 科室id + 0 = [1] And 开始原因 In (3, 7) And Nvl(附加床位, 0) = 0 And" & vbNewLine & _
                "             开始时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(开始原因, 3, '转科', 7, '交班')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(终止原因, 3, '转科', 7, '交班') As 事件, '前' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病人变动记录" & vbNewLine & _
                "       Where 科室id + 0 = [1] And 终止原因 In (3, 7) And Nvl(附加床位, 0) = 0 And" & vbNewLine & _
                "             终止时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(终止原因, 3, '转科', 7, '交班')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select R.事件, E.时机, Decode(E.时机, '前', R.前人次, '后', R.后人次) As 人次" & vbNewLine & _
                "       From (Select Decode(R.诊疗类别, 'F', '手术', Decode(I.操作类型, '7', '会诊', '抢救')) As 事件," & vbNewLine & _
                "                     Sum(Decode(R.病人科室id, [1], 1, 0)) As 前人次, Sum(Decode(R.执行科室id, [1], 1, 0)) As 后人次" & vbNewLine & _
                "              From 病人医嘱记录 R, 诊疗项目目录 I, 病人医嘱发送 S" & vbNewLine & _
                "              Where R.ID = S.医嘱id And R.诊疗项目id = I.ID And" & vbNewLine & _
                "                    (R.诊疗类别 = 'F' Or R.诊疗类别 = 'Z' And I.操作类型 In ('7', '8')) And R.相关id Is Null And" & vbNewLine & _
                "                    R.医嘱期效 = 1 And (R.医嘱状态 = 8 Or R.医嘱状态 = 9) And" & vbNewLine & _
                "                    (R.病人科室id + 0 = [1] Or R.执行科室id + 0 = [1]) And S.首次时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "                    To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "              Group By Decode(R.诊疗类别, 'F', '手术', Decode(I.操作类型, '7', '会诊', '抢救'))) R," & vbNewLine & _
                "            (Select Decode(Rownum, 1, '前', 2, '后') As 时机 From 病历书写事件 Where Rownum < 3) E) P," & vbNewLine
        
        strSQL = strSQL & "     (Select 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成, Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "       From 电子病历记录" & vbNewLine & _
                "       Where 病历种类 = 2 And 科室id + 0 = [1] And 创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By 文件id) W" & vbNewLine & _
                "Where F.事件 = P.事件 And Decode(Sign(F.书写时限), -1, '前', '后') = P.时机 And P.人次 > 0 And F.ID = W.文件id(+)" & vbNewLine & _
                "Order By F.编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo)
    Case 4
        strSQL = "Select F.ID, F.编号, F.名称, F.事件 || Decode(Sign(F.书写时限), -1, '前', '后') || '书写' As 要求," & vbNewLine & _
                "       Decode(F.唯一, 1, To_Char(P.人次), '<循环>') As 应写数, W.已完成, W.在书写" & vbNewLine & _
                "From (Select F.ID, F.编号, F.名称, F.事件, F.唯一, F.书写时限" & vbNewLine & _
                "       From (Select F.ID, F.编号, F.名称, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "              From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "              Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 4) F" & vbNewLine & _
                "       Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [1]) F," & vbNewLine
        strSQL = strSQL & "     (Select '入院' As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病案主页" & vbNewLine & _
                "       Where 入院病区id + 0 = [1] And 入院日期 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(出院方式, '死亡', '死亡', '出院') As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病案主页" & vbNewLine & _
                "       Where 当前病区id + 0 = [1] And 出院日期 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(出院方式, '死亡', '死亡', '出院')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(开始原因, 3, '转科', 8, '交班') As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病人变动记录" & vbNewLine & _
                "       Where 病区id + 0 = [1] And 开始原因 In (3, 8) And Nvl(附加床位, 0) = 0 And" & vbNewLine & _
                "             开始时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(开始原因, 3, '转科', 8, '交班')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(终止原因, 3, '转科', 8, '交班') As 事件, '前' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病人变动记录" & vbNewLine & _
                "       Where 病区id + 0 = [1] And 终止原因 In (3, 8) And Nvl(附加床位, 0) = 0 And" & vbNewLine & _
                "             终止时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(终止原因, 3, '转科', 8, '交班')) P," & vbNewLine
        strSQL = strSQL & "     (Select 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成, Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "       From 电子病历记录" & vbNewLine & _
                "       Where 病历种类 = 4 And 科室id + 0 = [1] And 创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By 文件id) W" & vbNewLine & _
                "Where F.事件 = P.事件 And Decode(Sign(F.书写时限), -1, '前', '后') = P.时机 And P.人次 > 0 And F.ID = W.文件id(+)" & vbNewLine & _
                "Order By F.编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo)
    End Select
    With Me.vfgFile
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 0: .ColHidden(0) = True
        .ColAlignment(4) = flexAlignRightCenter
        For lngCount = 1 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
    End With
    Call vfgFile_RowColChange
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    If Me.ActiveControl.Name = Me.vfgFile.Name Then
        Set objPrint.Body = Me.vfgFile
        objPrint.Title.Text = mstrDeptName & "病历文件列表"
    Else
        Set objPrint.Body = Me.vfgEPRs
        objPrint.Title.Text = mstrDeptName & Me.vfgFile.TextMatrix(Me.vfgFile.Row, 2) & "清单"
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
    Case conMenu_File_Open
        Dim f As New frmEPRView
        f.ShowMe Me, CLng(Me.vfgEPRs.TextMatrix(Me.vfgFile.Row, 0)), True
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
        Call frmEPRAuditMonitor.zlRefList(Val(Me.vfgEPRs.TextMatrix(Me.vfgEPRs.Row, 0)))
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
        Control.Enabled = (Val(Me.vfgEPRs.TextMatrix(Me.vfgEPRs.Row, 0)) > 0)
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Me.ActiveControl.Name = Me.vfgFile.Name Then
            Control.Enabled = (Me.vfgFile.Rows > Me.vfgFile.FixedRows)
        Else
            Control.Enabled = (Me.vfgEPRs.Rows > Me.vfgEPRs.FixedRows)
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_File
        Item.Handle = Me.vfgFile.hWnd
    Case conPane_EPRs
        Item.Handle = Me.vfgEPRs.hWnd
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
    '设置词句显示停靠窗格
    If mfrmWord Is Nothing Then Set mfrmWord = New frmDockEPRContent
    
    Dim panThis As Pane, panChild As Pane
    Set panThis = dkpMan.CreatePane(conPane_File, 400, 100, DockLeftOf, Nothing)
    panThis.Title = "病历文件清单"
    panThis.Options = PaneNoCaption
    
    Set panChild = dkpMan.CreatePane(conPane_EPRs, 400, 300, DockBottomOf, panThis)
    panChild.Title = "病历书写记录"
    panChild.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_Word, 600, 400, DockRightOf, Nothing)
    panThis.Title = "病历内容"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable

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

Private Sub vfgEPRs_GotFocus()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub vfgEPRs_RowColChange()
    Dim lngRecordId As Long
    
    Err = 0: On Error Resume Next
    With Me.vfgEPRs
        lngRecordId = Val(.TextMatrix(.Row, 0))
    End With
    Err = 0: On Error GoTo 0
    If Me.Tag <> "" Then Exit Sub
    Call mfrmWord.zlRefresh(lngRecordId, "", True)
End Sub

Private Sub vfgFile_GotFocus()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub vfgFile_RowColChange()
    Dim lngFileID As Long       '病历文件id
    
    If Me.Tag <> "" Then Exit Sub
    lngFileID = Val(Me.vfgFile.TextMatrix(Me.vfgFile.Row, 0))
    
    Select Case mintKind
    Case 1
        strSQL = "Select W.ID, P.病人id, P.门诊号, P.姓名, P.性别, To_Char(P.执行时间, 'mm-dd hh24:mi') As 就诊日期, W.创建人 As 书写人," & vbNewLine & _
                "       To_Char(W.完成时间, 'mm-dd hh24:mi') As 完成时间" & vbNewLine & _
                "From 电子病历记录 W, 病人挂号记录 P" & vbNewLine & _
                "Where W.主页id = P.ID And W.病历种类 = 1 And W.科室id + 0 = [1] And W.文件id + 0 = [4] And" & vbNewLine & _
                "      W.创建时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By P.执行时间"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo, lngFileID)
    Case 2
        strSQL = "Select W.ID, I.病人id, P.住院号, I.姓名, I.性别, To_Char(P.入院日期, 'mm-dd hh24:mi') As 入院日期, W.创建人 As 书写人," & vbNewLine & _
                "       To_Char(W.完成时间, 'mm-dd hh24:mi') As 完成时间" & vbNewLine & _
                "From 电子病历记录 W, 病案主页 P, 病人信息 I" & vbNewLine & _
                "Where I.病人id = P.病人id And P.病人id = W.病人id And P.主页id = W.主页id And W.病历种类 = 2 And W.科室id + 0 = [1] And" & vbNewLine & _
                "      W.文件id + 0 = [4] And W.创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "      To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By 入院日期"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo, lngFileID)
    Case 4
        strSQL = "Select W.ID, I.病人id, P.住院号, I.姓名, I.性别, To_Char(P.入院日期, 'mm-dd hh24:mi') As 入院日期, W.创建人 As 书写人," & vbNewLine & _
                "       To_Char(W.完成时间, 'mm-dd hh24:mi') As 完成时间" & vbNewLine & _
                "From 电子病历记录 W, 病案主页 P, 病人信息 I" & vbNewLine & _
                "Where I.病人id = P.病人id And P.病人id = W.病人id And P.主页id = W.主页id And W.病历种类 = 4 And W.科室id + 0 = [1] And" & vbNewLine & _
                "      W.文件id + 0 = [4] And W.创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "      To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By 入院日期"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo, lngFileID)
    End Select
    
    With Me.vfgEPRs
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 0: .ColHidden(0) = True
        For lngCount = 1 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
    End With
    Call vfgEPRs_RowColChange
End Sub
