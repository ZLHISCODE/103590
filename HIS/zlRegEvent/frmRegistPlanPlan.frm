VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmRegistPlanPlan 
   BorderStyle     =   0  'None
   Caption         =   "计划安排号别"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsPlan 
      Height          =   2145
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3405
      _cx             =   6006
      _cy             =   3784
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483641
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   26
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRegistPlanPlan.frx":0000
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
      ExplorerBar     =   7
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   30
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   1
         Top             =   195
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmRegistPlanPlan.frx":032F
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmRegistPlanPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mArrFilter As Variant
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2009-09-09 15:45:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer, i As Integer, objGrid As VSFlexGrid
   i = 0
    With vsPlan
        .Redraw = flexRDNone
        .Rows = 3: .FixedRows = 2
        .FixedCols = 1
        .Cols = 40:   .Clear
        .FrozenCols = 2
        .TextMatrix(0, i) = "  ": .ColWidth(i) = 285
        .TextMatrix(1, i) = "  ":  .ColKey(i) = "标志": i = i + 1
        
        .TextMatrix(0, i) = "ID": .ColHidden(i) = True: .ColWidth(i) = 0
        .TextMatrix(1, i) = "ID": .ColKey(i) = "ID": i = i + 1
         
        .TextMatrix(0, i) = "号类": .ColWidth(i) = 720
        .TextMatrix(1, i) = "号类": .ColKey(i) = "号类": i = i + 1

        .TextMatrix(0, i) = "号别": .ColWidth(i) = 480
        .TextMatrix(1, i) = "号别": .ColKey(i) = "号别": i = i + 1

        .TextMatrix(0, i) = "科室": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "科室": .ColKey(i) = "科室": i = i + 1
        .TextMatrix(0, i) = "项目": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "项目": .ColKey(i) = "项目": i = i + 1
        .TextMatrix(0, i) = "医生":: .ColWidth(i) = 1000
        .TextMatrix(1, i) = "医生": .ColKey(i) = "医生": i = i + 1
        .TextMatrix(0, i) = "建档": .ColWidth(i) = 495
        .TextMatrix(1, i) = "建档": .ColKey(i) = "建档": i = i + 1
        .TextMatrix(0, i) = "周日": .ColWidth(i) = 450
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周日-安排": i = i + 1
        .TextMatrix(0, i) = "周日": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周日-限号": i = i + 1
        .TextMatrix(0, i) = "周日": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周日-限约": i = i + 1

        .TextMatrix(0, i) = "周一": .ColWidth(i) = 450
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周一-安排": i = i + 1
        .TextMatrix(0, i) = "周一": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周一-限号": i = i + 1
        .TextMatrix(0, i) = "周一": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周一-限约": i = i + 1

        .TextMatrix(0, i) = "周二": .ColWidth(i) = 450
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周二-安排": i = i + 1
        .TextMatrix(0, i) = "周二": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周二-限号": i = i + 1
        .TextMatrix(0, i) = "周二": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周二-限约": i = i + 1

        .TextMatrix(0, i) = "周三": .ColWidth(i) = 450
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周三-安排": i = i + 1
        .TextMatrix(0, i) = "周三": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周三-限号": i = i + 1
        .TextMatrix(0, i) = "周三": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周三-限约": i = i + 1

        .TextMatrix(0, i) = "周四": .ColWidth(i) = 450
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周四-安排": i = i + 1
        .TextMatrix(0, i) = "周四": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周四-限号": i = i + 1
        .TextMatrix(0, i) = "周四": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周四-限约": i = i + 1

        .TextMatrix(0, i) = "周五": .ColWidth(i) = 450
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周五-安排": i = i + 1
        .TextMatrix(0, i) = "周五": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周五-限号": i = i + 1
        .TextMatrix(0, i) = "周五": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周五-限约": i = i + 1

        .TextMatrix(0, i) = "周六": .ColWidth(i) = 450
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周六-安排": i = i + 1
        .TextMatrix(0, i) = "周六": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周六-限号": i = i + 1
        .TextMatrix(0, i) = "周六": .ColWidth(i) = 450
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周六-限约": i = i + 1
        .TextMatrix(0, i) = "分诊方式": .ColWidth(i) = 855
        .TextMatrix(1, i) = "分诊方式": .ColKey(i) = "分诊方式": i = i + 1
        .TextMatrix(0, i) = "IDS": .ColWidth(i) = 0: .ColHidden(i) = True
        .TextMatrix(1, i) = "IDS": .ColKey(i) = "IDS": i = i + 1
        .TextMatrix(0, i) = "生效时间": .ColWidth(i) = 2000
        .TextMatrix(1, i) = "生效时间": .ColKey(i) = "生效时间": i = i + 1
        .TextMatrix(0, i) = "失效时间": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "失效时间": .ColKey(i) = "失效时间": i = i + 1
        .TextMatrix(0, i) = "序号" & vbCrLf & "控制": .ColWidth(i) = 765
        .TextMatrix(1, i) = "序号" & vbCrLf & "控制": .ColKey(i) = "序号控制": i = i + 1
        
        .TextMatrix(0, i) = "安排人": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "安排人": .ColKey(i) = "安排人": i = i + 1
        .TextMatrix(0, i) = "安排时间": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "安排时间": .ColKey(i) = "安排时间": i = i + 1
        
        .TextMatrix(0, i) = "审核人": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "审核人": .ColKey(i) = "审核人": i = i + 1
        .TextMatrix(0, i) = "审核时间": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "审核时间": .ColKey(i) = "审核时间": i = i + 1
        .TextMatrix(0, i) = "实际执行时间": .ColWidth(i) = 1500
        .TextMatrix(1, i) = "实际执行时间": .ColKey(i) = "实际执行时间": i = i + 1
        .TextMatrix(0, i) = "应诊诊室": .ColWidth(i) = 2000
        .TextMatrix(1, i) = "应诊诊室": .ColKey(i) = "应诊诊室": i = i + 1
        .Cell(flexcpText, 0, 0, .Rows - 1) = " "
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        For i = 0 To .Cols - 1
            .MergeCol(i) = True:
            .FixedAlignment(i) = flexAlignCenterCenter
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            Select Case .ColKey(i)
            Case "ID", "标志", "IDS"
                 .ColData(i) = "-1|1"
            Case "号类", "号别", "生效时间"
                .ColData(i) = "1|0"
            End Select
        Next
         .MergeRow(0) = True: .MergeRow(1) = True
        .Redraw = flexRDBuffered
    End With
End Sub
 
Public Sub zlRefreshData(ByVal ArrFilter As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-15 11:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mArrFilter = ArrFilter
    Call LoadDataToList
End Sub
Private Sub LoadDataToList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, lngRow As Long, strSQL As String
    Dim blnHistory As Boolean, strStartDate As String, lngPriID As Long
    Dim strTable As String
    Dim strWhere As String
    
    Err = 0: On Error GoTo Errhand:
    If CStr(mArrFilter("有效期")(0)) <> "1901-01-01" Then
        strFilter = "  And Nvl(A.开始时间,To_Date('3000-01-01','YYYY-MM-DD'))>=[4]   And Nvl(A.终止时间,To_Date('1900-01-01','YYYY-MM-DD'))<=[5]"
    End If
    If Val(mArrFilter("科室ID")) > 0 Then strFilter = strFilter & " And A.科室ID=[1]"
    If Val(mArrFilter("科室ID")) = -1 Then strFilter = strFilter & " And A.科室ID in (Select 部门ID From 部门人员 where 人员id=" & UserInfo.ID & ") "
    
    '105869，根据计划的医生过滤
    Select Case mArrFilter("医生ID")(1)
    Case "ID"
         strFilter = strFilter & "  And C.医生ID=[2]"
    Case "UPR"
         strFilter = strFilter & " And Upper(C.医生姓名)=[3]"
    Case "NONE"
         strFilter = strFilter & " And C.医生姓名=[3]"
    End Select

 
    If CStr(mArrFilter("生效时间")(0)) <> "1901-01-01" And CStr(mArrFilter("安排时间")(0)) <> "1901-01-01" And CStr(mArrFilter("审核时间")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.生效时间 between [6] and [7] or C.安排时间 between [8] and [9] or C.审核时间 between [10] and [11]) "
    ElseIf CStr(mArrFilter("生效时间")(0)) <> "1901-01-01" And CStr(mArrFilter("安排时间")(0)) <> "1901-01-01" And CStr(mArrFilter("审核时间")(0)) = "1901-01-01" Then
        strFilter = strFilter & "  And (C.生效时间 between [6] and [7] or C.安排时间 between [8] and [9] ) "
    ElseIf CStr(mArrFilter("生效时间")(0)) <> "1901-01-01" And CStr(mArrFilter("安排时间")(0)) = "1901-01-01" And CStr(mArrFilter("审核时间")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.生效时间 between [6] and [7]   or C.审核时间 between [10] and [11]) "
    ElseIf CStr(mArrFilter("生效时间")(0)) = "1901-01-01" And CStr(mArrFilter("安排时间")(0)) <> "1901-01-01" And CStr(mArrFilter("审核时间")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.安排时间 between [8] and [9]  or C.审核时间 between [10] and [11]) "
    ElseIf CStr(mArrFilter("生效时间")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.安排时间 between [6] and [7])  "
    ElseIf CStr(mArrFilter("安排时间")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.安排时间 between [8] and [9])  "
    ElseIf CStr(mArrFilter("审核时间")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.审核时间 between [10] and [11])  "
    End If
    
    If Val(mArrFilter("仅显未生效计划")) = 1 Then strFilter = strFilter & " and C.生效时间>nvl(A.开始时间,to_date('1901-01-01','yyyy-mm-dd'))"
    If Val(mArrFilter("仅显示未审计划")) = 1 Then strFilter = strFilter & " and  C.审核时间 IS NULL "


    strTable = "" & _
    "   Select C.ID, " & _
    "             Sum(Decode(B.限制项目,'周日',B.限号数,0)) as 周日限号, Sum(Decode(B.限制项目,'周日',B.限约数,0))  as 周日限约," & _
    "             Sum(Decode(B.限制项目,'周一',B.限号数,0)) as 周一限号, Sum(Decode(B.限制项目,'周一',B.限约数,0))  as 周一限约," & _
    "             Sum(Decode(B.限制项目,'周二',B.限号数,0)) as 周二限号, Sum(Decode(B.限制项目,'周二',B.限约数,0))  as 周二限约," & _
    "             Sum(Decode(B.限制项目,'周三',B.限号数,0)) as 周三限号, Sum(Decode(B.限制项目,'周三',B.限约数,0))  as 周三限约," & _
    "             Sum(Decode(B.限制项目,'周四',B.限号数,0)) as 周四限号, Sum(Decode(B.限制项目,'周四',B.限约数,0))  as 周四限约," & _
    "             Sum(Decode(B.限制项目,'周五',B.限号数,0)) as 周五限号, Sum(Decode(B.限制项目,'周五',B.限约数,0))  as 周五限约," & _
    "             Sum(Decode(B.限制项目,'周六',B.限号数,0)) as 周六限号, Sum(Decode(B.限制项目,'周六',B.限约数,0))  as 周六限约" & _
    "   From 挂号安排计划 C,挂号计划限制 B,挂号安排 A  " & _
    "   Where C.ID=B.计划ID(+)   and C.安排ID=A.ID  " & strFilter & _
    "   Group by C.ID"
    
    '105869，取计划的医生及收费项目
    strSQL = " " & _
        "   Select P.*,B.名称 As 项目,D.名称 As 科室 " & _
        "   From ( " & _
        "     Select  row_number()  over (Partition By 计划id Order By 计划id,级数 Desc) As 序号1,M.* " & _
        "     From ( " & _
        "       Select Level As 级数, Sys_Connect_By_Path(门诊诊室, ';') 门诊诊室集, Q.*  " & _
        "       From (  Select  C.Id as 计划ID,C.安排ID ,A.号类,  A.号码,  A.科室id,  C.项目id, C.医生姓名,  C.医生id,     " & _
        "                              C.周日,C1.周日限号,C1.周日限约,C.周一,C1.周一限号,C1.周一限约,C.周二,C1.周二限号,C1.周二限约, " & _
        "                              C.周三,C1.周三限号,C1.周三限约,C.周四,C1.周四限号,C1.周四限约,C.周五,C1.周五限号,C1.周五限约, " & _
        "                              C.周六,C1.周六限号,C1.周六限约, " & _
        "                              A.病案必须,   Decode(Nvl(C.分诊方式,0),0,'不分诊',1,'指定诊室',2,'动态分诊',3,'平均分诊') as 分诊方式 ,  C.序号控制," & _
        "                              to_char(A.开始时间,'yyyy-mm-dd hh24:mi:ss') 开始时间,  to_char(A.终止时间,'yyyy-mm-dd hh24:mi:ss') 终止时间," & _
        "                              to_char(C.生效时间,'yyyy-mm-dd hh24:mi:ss') as 生效时间,to_char(C.失效时间,'yyyy-mm-dd hh24:mi:ss') as 失效时间," & _
        "                              to_char(C.实际生效,'yyyy-mm-dd hh24:mi:ss') as 实际执行时间,            " & _
        "                              C.安排人,to_char(C.安排时间,'yyyy-mm-dd hh24:mi:ss') as 安排时间,            " & _
        "                              C.审核人,to_char(C.审核时间,'yyyy-mm-dd hh24:mi:ss') as 审核时间 , " & _
        "                              b.门诊诊室,row_number() over (Partition By 计划ID Order By 计划id,门诊诊室) As 序号 " & _
        "           From  (" & strTable & ") C1,挂号安排计划 C,挂号安排 A,挂号计划诊室 B " & _
        "           Where C.ID=C1.ID And C.安排ID =A.Id And C.Id=B.计划ID(+)   " & _
        "           Order By 计划ID,门诊诊室 ) Q " & _
        "        Connect By 计划id= Prior 计划id And 序号-1 =Prior 序号 " & _
        "        )  M ) P,收费项目目录 B,部门表 D " & _
        "    Where P.序号1=1 And P.项目id=b.Id And P.科室id =d.Id(+) And (B.站点='" & gstrNodeNo & "' Or b.站点 is Null)   " & _
        "    Order By 号码, 生效时间 Desc, 计划ID DESC"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        Val(mArrFilter("科室ID")), _
        Val(mArrFilter("医生ID")(0)), _
        CStr(mArrFilter("医生ID")(0)), _
        CDate(mArrFilter("有效期")(0)), CDate(mArrFilter("有效期")(1)), _
        CDate(mArrFilter("生效时间")(0)), CDate(mArrFilter("生效时间")(1)), _
        CDate(mArrFilter("安排时间")(0)), CDate(mArrFilter("安排时间")(1)), _
        CDate(mArrFilter("审核时间")(0)), CDate(mArrFilter("审核时间")(1)), _
        "")
      
    With Me.vsPlan
        If .Row > 0 And .Row <= .Rows - 1 Then lngPriID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        .Clear 1
        .Rows = 3: lngRow = 2
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 2
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!计划Id)
            .Cell(flexcpData, lngRow, .ColIndex("ID")) = Nvl(rsTemp!安排ID)
            .TextMatrix(lngRow, .ColIndex("号类")) = Nvl(rsTemp!号类)
            .TextMatrix(lngRow, .ColIndex("号别")) = Nvl(rsTemp!号码)
            .TextMatrix(lngRow, .ColIndex("科室")) = Nvl(rsTemp!科室)
            .TextMatrix(lngRow, .ColIndex("项目")) = Nvl(rsTemp!项目)
            .TextMatrix(lngRow, .ColIndex("医生")) = Nvl(rsTemp!医生姓名)
            .TextMatrix(lngRow, .ColIndex("周日-安排")) = Nvl(rsTemp!周日)
            .TextMatrix(lngRow, .ColIndex("周日-限号")) = Format(Val(Nvl(rsTemp!周日限号)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周日-限约")) = Format(Val(Nvl(rsTemp!周日限约)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周一-安排")) = Nvl(rsTemp!周一)
            .TextMatrix(lngRow, .ColIndex("周一-限号")) = Format(Val(Nvl(rsTemp!周一限号)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周一-限约")) = Format(Val(Nvl(rsTemp!周一限约)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周二-安排")) = Nvl(rsTemp!周二)
            .TextMatrix(lngRow, .ColIndex("周二-限号")) = Format(Val(Nvl(rsTemp!周二限号)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周二-限约")) = Format(Val(Nvl(rsTemp!周二限约)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周三-安排")) = Nvl(rsTemp!周三)
            .TextMatrix(lngRow, .ColIndex("周三-限号")) = Format(Val(Nvl(rsTemp!周三限号)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周三-限约")) = Format(Val(Nvl(rsTemp!周三限约)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周四-安排")) = Nvl(rsTemp!周四)
            .TextMatrix(lngRow, .ColIndex("周四-限号")) = Format(Val(Nvl(rsTemp!周四限号)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周四-限约")) = Format(Val(Nvl(rsTemp!周四限约)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周五-安排")) = Nvl(rsTemp!周五)
            .TextMatrix(lngRow, .ColIndex("周五-限号")) = Format(Val(Nvl(rsTemp!周五限号)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周五-限约")) = Format(Val(Nvl(rsTemp!周五限约)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周六-安排")) = Nvl(rsTemp!周六)
            .TextMatrix(lngRow, .ColIndex("周六-限号")) = Format(Val(Nvl(rsTemp!周六限号)), "###;;")
            .TextMatrix(lngRow, .ColIndex("周六-限约")) = Format(Val(Nvl(rsTemp!周六限约)), "###;;")
            .TextMatrix(lngRow, .ColIndex("建档")) = IIf(Val(Nvl(rsTemp!病案必须)) = 0, "", "√")
            .TextMatrix(lngRow, .ColIndex("分诊方式")) = Nvl(rsTemp!分诊方式)
            .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!科室ID) & "_" & Nvl(rsTemp!项目ID) & "_" & Nvl(rsTemp!医生ID)
            If Nvl(rsTemp!门诊诊室集) <> "" Then
                .TextMatrix(lngRow, .ColIndex("应诊诊室")) = Mid(Nvl(rsTemp!门诊诊室集), 2)  ' Read计划应诊诊室(lng安排ID, Val(Nvl(rsTemp!计划ID)), False) ' Nvl(rsTemp!门诊诊室)
            End If
            
            If Not IsNull(rsTemp!生效时间) Then
                .TextMatrix(lngRow, .ColIndex("生效时间")) = Format(rsTemp!生效时间, "yyyy-MM-dd HH:mm:ss")
                If Format(Nvl(rsTemp!生效时间), "yyyy-MM-dd HH:mm:ss") <= Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And Nvl(rsTemp!审核时间) <> "" Then
                    '已经生效,不能更改
                    .Cell(flexcpData, lngRow, .ColIndex("生效时间")) = 1
                Else
                    '未生效,能更改
                    .Cell(flexcpData, lngRow, .ColIndex("生效时间")) = 0
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("失效时间")) = Nvl(rsTemp!失效时间)
            .TextMatrix(lngRow, .ColIndex("序号控制")) = IIf(Val(Nvl(rsTemp!序号控制)) = 0, "", "√")
            
            .TextMatrix(lngRow, .ColIndex("安排人")) = Nvl(rsTemp!安排人)
            .TextMatrix(lngRow, .ColIndex("安排时间")) = Nvl(rsTemp!安排时间)
            .TextMatrix(lngRow, .ColIndex("审核人")) = Nvl(rsTemp!审核人)
            .TextMatrix(lngRow, .ColIndex("审核时间")) = Nvl(rsTemp!审核时间)
            If Nvl(rsTemp!实际执行时间) < "3000-01-01" Then
                .TextMatrix(lngRow, .ColIndex("实际执行时间")) = Nvl(rsTemp!实际执行时间)
            End If
            If Val(.Cell(flexcpData, lngRow, .ColIndex("生效时间"))) = 1 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000010
            Else
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
            End If
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
       ' .AutoSizeMode = flexAutoSizeColWidth
        '.AutoSize 0, .Cols - 1
        If lngPriID <> 0 Then
            lngRow = .FindRow(lngPriID, 0, .ColIndex("ID"), , True)
            If lngRow > 0 Then .Row = lngRow
        Else
            .Row = 1
        End If
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsPlan, Me.Caption, "计划信息列表", True
        .ColWidth(.ColIndex("标志")) = 285
        .Redraw = flexRDBuffered
    End With
   Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
     Me.vsPlan.Redraw = flexRDBuffered
End Sub

Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    Call InitVsGrid
    Call vsPlan_LostFocus
    vsPlan_GotFocus
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsPlan
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2009-09-09 11:24:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "挂号计划表"
    

    If CStr(mArrFilter("有效期")(0)) <> "1901-01-01" Then
        objRow.Add "效期范围：" & CStr(mArrFilter("有效期")(0)) & "至" & CStr(mArrFilter("有效期")(1))
    End If
    If Val(mArrFilter("科室ID")) > 0 Then
        strSQL = "Select 名称 From 部门表 where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mArrFilter("科室ID")))
        If rsTemp.EOF Then
            objRow.Add "科室：所有科室"
        Else
            objRow.Add "科室：" & Nvl(rsTemp!名称)
        End If
    ElseIf Val(mArrFilter("科室ID")) = -1 Then
        objRow.Add "科室：操作员所属科室"
    Else
        objRow.Add "科室：所有科室"
    End If
    Select Case mArrFilter("医生ID")(1)
    Case "ID"
        strSQL = "Select 姓名 From 人员表 where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mArrFilter("医生ID")(0)))
        If rsTemp.EOF Then
            objRow.Add "医生：所有"
        Else
            objRow.Add "医生：" & Nvl(rsTemp!姓名)
        End If
    Case "UPR", "NONE"
            objRow.Add "医生：" & CStr(mArrFilter("医生ID")(0))
    End Select
    objPrint.UnderAppRows.Add objRow
    
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo Errhand:
    With vsPlan
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("标志") Then .ColWidth(intCol) = 0
        Next
    End With
    
    Set objPrint.Body = vsPlan
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    
    With vsPlan
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
    With vsPlan
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
End Sub
Public Sub zlCallCustomReprot(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用相关的自定义报表
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-15 11:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, str号类 As String
    '科室ID_项目ID_医生ID
    With vsPlan
        varData = Split(.TextMatrix(.Row, .ColIndex("IDS")) & "___", "_")
        str号类 = Trim(.TextMatrix(.Row, .ColIndex("号类")))
        If str号类 <> "" Then
            Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain, _
                "号类=" & str号类, "号别=" & Trim(.TextMatrix(.Row, .ColIndex("号别"))), _
                "科室=" & Val(varData(0)), _
                "项目=" & Val(varData(1)), _
                "医生=" & Val(varData(2)))
        Else
            Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
        End If
    End With
End Sub
Public Property Get zlGet安排ID(Optional blnPlanID As Boolean = True) As Long
    With vsPlan
        If blnPlanID Then
            zlGet安排ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        Else
            zlGet安排ID = Val(.Cell(flexcpData, .Row, .ColIndex("ID")))
        End If
    End With
End Property

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "计划信息列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsPlan_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "计划信息列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsPlan_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub
Private Sub vsPlan_GotFocus()
    vsPlan.BackColorSel = &H8000000D
End Sub

Private Sub vsPlan_LostFocus()
    vsPlan.BackColorSel = GRD_LOSTFOCUS_COLORSEL
End Sub
Private Sub vsPlan_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "计划信息列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsPlan, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "计划信息列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
Public Property Get zlPlanStatus() As Long
    Dim lngID As Long
    '获取计划安排的当前状态
    '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效
    With vsPlan
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then zlPlanStatus = 0: Exit Property
        If .TextMatrix(.Row, .ColIndex("审核时间")) <> "" Then
            zlPlanStatus = 2
            If Val(.Cell(flexcpData, .Row, .ColIndex("生效时间"))) = 1 Then
                zlPlanStatus = 3
            End If
        Else
              zlPlanStatus = 1
        End If
    End With
End Property

Public Sub zlActtion()
    zlControl.ControlSetFocus vsPlan, True
End Sub

