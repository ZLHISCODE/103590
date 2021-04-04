VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBookingDefer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "预约延期"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   Icon            =   "frmBookingDefer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8085
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6840
      TabIndex        =   6
      Top             =   90
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6840
      TabIndex        =   5
      Top             =   570
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   3900
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   6879
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      MousePointer    =   1
      FormatString    =   "^ 号序|^     时间|^       NO|^    票据号|^     姓名|^ 性别|^ 年龄|^     门诊号"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBookingDefer.frx":038A
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComCtl2.DTPicker dtpDefer 
      Height          =   300
      Left            =   3720
      TabIndex        =   1
      Top             =   2490
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   104595459
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpBooking 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   2490
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   104595459
      CurrentDate     =   36588
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPlan 
      Height          =   2235
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   3942
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      MousePointer    =   1
      FormatString    =   "^  号类|  号别|^       科室|^    医生|限约|开始时间|终止时间|序号控制"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBookingDefer.frx":06A4
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblNewDate 
      Caption         =   "延期至"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblOldDate 
      Caption         =   "预约日期"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
End
Attribute VB_Name = "frmBookingDefer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsList As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strBegin As String, strEnd As String, strTMP As String
    Dim strSNS As String, strDay As String, str号别 As String
    Dim i As Long, intDay As Integer, intCol As Integer
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If mrsList Is Nothing Then Exit Sub
    If mrsList.RecordCount = 0 Then Exit Sub
    str号别 = mshPlan.TextMatrix(mshPlan.Row, GetPlanCol("号别"))
    If str号别 = "" Then Exit Sub
    If dtpDefer.Value <= dtpBooking.Value Then
        MsgBox "指定的延期时间必须大于旧的预约时间!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '检查该号别当天是否有安排
    mrsList.Sort = "时间 ASC"
    mrsList.MoveLast
    strBegin = mrsList!时间
    
    mrsList.Sort = "时间 DESC"
    mrsList.MoveFirst
    strEnd = mrsList!时间
    
    intDay = Weekday(dtpDefer.Value, vbSunday)
    strDay = Choose(intDay, "周日", "周一", "周二", "周三", "周四", "周五", "周六")
    
    On Error GoTo errH
    strSQL = "Select 开始时间,终止时间" & vbNewLine & _
            "From (Select B.开始时间, Decode(Sign(B.终止时间 - B.开始时间), 1, B.终止时间, B.终止时间 + 1) 终止时间" & vbNewLine & _
            "       From 挂号安排 A, 时间段 B" & vbNewLine & _
            "       Where A.号码 = [1] And A." & strDay & " = B.时间段 And ([2] Between A.开始时间 And A.终止时间 Or A.开始时间 IS Null))" & vbNewLine & _
            "Where To_Date(To_char(开始时间,'yyyy-mm-dd ')||'" & strBegin & "','yyyy-mm-dd hh24:mi:ss') Between 开始时间 And 终止时间 " & _
            " And To_Date(To_char(开始时间,'yyyy-mm-dd ')||'" & strEnd & "','yyyy-mm-dd hh24:mi:ss') Between 开始时间 And 终止时间"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号别, CDate(Format(dtpDefer.Value, "yyyy-MM-dd 00:00:00")))
    If rsTmp.RecordCount = 0 Then
        MsgBox "指定的延期日期没有该医生有效的挂号安排!" & vbCrLf & _
            "请检查延期日期和当前号别的挂号安排.", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '检查号序是否已用,仅提示,过程处理时跳过
    intCol = GetListCol("号序")
    For i = 1 To mshList.Rows - 1
        strTMP = Trim(mshList.TextMatrix(i, intCol))
        If strTMP <> "" Then strSNS = strSNS & ",'" & strTMP & "'"
    Next
    If strSNS <> "" Then
        strSQL = "Select 序号 From 挂号序号状态 Where Trunc(日期) = [1] And Instr([2], ','''||序号||'''') > 0 And Not (状态=3 And 操作员姓名=[3])"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpDefer.Value, "yyyy-MM-dd 00:00:00")), strSNS, CStr(UserInfo.姓名))
        strTMP = ""
        For i = 1 To rsTmp.RecordCount
            strTMP = strTMP & "," & rsTmp!序号
            rsTmp.MoveNext
        Next
        
        If strTMP <> "" Then
            MsgBox "注意:延期时间" & Format(dtpDefer.Value, "yyyy-MM-dd") & "的以下序号已被使用:" & vbCrLf & Mid(strTMP, 2) & vbCrLf & _
                "使用这些序号的预约挂号单将不会执行延期!", vbInformation, gstrSysName
        End If
    End If
    
    strSQL = "zl_病人预约挂号_Defer('" & str号别 & "',To_date('" & Format(dtpBooking.Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS'),To_date('" & _
            Format(dtpDefer.Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    Call dtpBooking_Change
    
    Exit Sub
errH:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpBooking_Change()
     If dtpDefer.Value <= dtpBooking.Value Then
        dtpDefer.Value = DateAdd("d", 1, dtpBooking.Value)
        dtpDefer.MinDate = dtpDefer.Value
    Else
        dtpDefer.MinDate = DateAdd("d", 1, dtpBooking.Value)
    End If
    
    Call SetPlanGrid
    Call ShowPlan(dtpBooking.Value)
    Call mshPlan_EnterCell
End Sub

Private Sub Form_Load()
    Dim Datsys As Date
    
    Datsys = zlDatabase.Currentdate
    dtpBooking.Value = DateAdd("d", 1, Datsys)
    dtpBooking.MinDate = dtpBooking.Value
    dtpDefer.Value = DateAdd("d", 1, dtpBooking.Value)
    dtpDefer.MinDate = dtpDefer.Value
    
    Call SetPlanGrid
    Call ShowPlan(dtpBooking.Value)
End Sub

Private Sub ShowPlan(datBooking As Date)
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "Select Distinct B.号类,B.号码 号别, C.名称 科室, B.医生姓名 医生, Nvl(F.限约数,0) 限约," & vbNewLine & _
            " To_Char(B.开始时间,'YYYY-MM-DD') 开始时间,To_Char(B.终止时间,'YYYY-MM-DD') 终止时间,Decode(Nvl(B.序号控制,0),1,'√',' ') as 序号控制" & vbNewLine & _
            "From 门诊费用记录 A, 挂号安排 B, 部门表 C,挂号安排限制 F " & vbNewLine & _
            "Where A.发生时间 Between [1] And [2] And A.记录性质 = 4 And A.记录状态 = 0 And A.序号=1" & vbNewLine & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
            "      And A.计算单位 = B.号码 And B.科室id = C.ID And B.id=F.安排ID(+) " & _
            "      And  Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =f.限制项目(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpBooking.Value, "yyyy-MM-dd 00:00:00")), _
                CDate(Format(dtpBooking.Value, "yyyy-MM-dd 23:59:59")))
    With mshPlan
        .ToolTipText = "共 " & rsTmp.RecordCount & " 条记录."
        .Rows = IIf(rsTmp.RecordCount = 0, 1, rsTmp.RecordCount) + 1
        For i = 1 To rsTmp.RecordCount
            For j = 0 To rsTmp.Fields.Count - 1
                .TextMatrix(i, j) = "" & rsTmp.Fields(j).Value
            Next
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetPlanGrid()
    Dim i As Integer, strHead As String
    
    strHead = "号类,1,600|号别,1,600|科室,1,1050|医生,4,800|限约,4,500|开始时间,4,1000|终止时间,4,1000|序号控制,4,850"
       
    With mshPlan
        .Redraw = False
        .Clear: .Rows = 2
        .FixedRows = 1
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = flexAlignCenterCenter
        Next
        
        If Not Visible Then Call RestoreFlexState(mshPlan, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 300
        
        .Redraw = True
    End With
End Sub


Private Sub SetListGrid()
    Dim i As Integer, strHead As String
    
    strHead = "号序,1,500|时间,4,1200|NO,4,1250|票据号,4,1250|病人,4,1250|性别,4,500|年龄,4,800|门诊号,1,1450"
       
    With mshList
        .Redraw = False
        .Clear: .Rows = 2
        .FixedRows = 1
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = flexAlignCenterCenter
        Next
        
        If Not Visible Then Call RestoreFlexState(mshPlan, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 300
        
        .Redraw = True
    End With
End Sub

Private Sub ShowList(datBooking As Date, str号别 As String)
    Dim strSQL As String, i As Long, j As Long
    On Error GoTo errH
    
    strSQL = "Select A.发药窗口 号序,To_Char(A.发生时间,'hh24:mi:ss') 时间,A.NO, A.实际票号 票据号, A.姓名 病人, A.性别, A.年龄, A.标识号 As 门诊号" & vbNewLine & _
        "From 门诊费用记录 A" & vbNewLine & _
        "Where A.发生时间 Between [1] And [2] And A.计算单位 = [3] And A.记录性质 = 4 And A.记录状态 = 0 And A.序号 = 1 Order by to_number(号序)"


    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpBooking.Value, "yyyy-MM-dd 00:00:00")), _
                CDate(Format(dtpBooking.Value, "yyyy-MM-dd 23:59:59")), str号别)
    With mshList
        .ToolTipText = "共 " & mrsList.RecordCount & " 条记录."
        .Rows = IIf(mrsList.RecordCount = 0, 1, mrsList.RecordCount) + 1
        For i = 1 To mrsList.RecordCount
            For j = 0 To mrsList.Fields.Count - 1
                .TextMatrix(i, j) = "" & mrsList.Fields(j).Value
            Next
            mrsList.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If errCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mrsList = Nothing
    
    Call SaveFlexState(mshPlan, App.ProductName & "\" & Me.Name)
    Call SaveFlexState(mshList, App.ProductName & "\" & Me.Name)
End Sub

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCol As Integer, intRow As Integer
    
    intCol = mshList.MouseCol
    intRow = mshList.MouseRow
    If intRow = 0 Then
        mshList.ColData(intCol) = (mshList.ColData(intCol) + 1) Mod 2
        mshList.ColSel = mshList.Col
        mshList.Sort = Val(mshList.ColData(intCol)) + 1 '1-升,2-降
    End If
End Sub

Private Sub mshPlan_EnterCell()
    Dim i As Integer, blnPre As Boolean
    Dim intRow As Integer, intCol As Integer, str号别 As String
    
    blnPre = mshPlan.Redraw
    intRow = mshPlan.Row: intCol = mshPlan.Col
    mshPlan.Redraw = False
    
    For i = 0 To mshPlan.Cols - 1
        mshPlan.Col = i
        mshPlan.CellBackColor = mshPlan.BackColorSel
        mshPlan.CellForeColor = mshPlan.ForeColorSel
    Next
    
    mshPlan.Row = intRow:  mshPlan.Col = intCol
    mshPlan.Redraw = blnPre
    
    str号别 = mshPlan.TextMatrix(mshPlan.Row, GetPlanCol("号别"))
    Call SetListGrid
    Call ShowList(dtpBooking.Value, str号别)

    cmdOK.Enabled = (str号别 <> "")
End Sub

Private Function GetPlanCol(strName As String) As Integer
    Dim i As Integer
    For i = 0 To mshPlan.Cols - 1
        If mshPlan.TextMatrix(0, i) = strName Then
            GetPlanCol = i: Exit For
        End If
    Next
End Function

Private Function GetListCol(strName As String) As Integer
    Dim i As Integer
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strName Then
            GetListCol = i: Exit For
        End If
    Next
End Function

Private Sub mshPlan_LeaveCell()
    Dim i As Integer, blnPre As Boolean
    
    blnPre = mshPlan.Redraw
    mshPlan.Redraw = False
    
    For i = 0 To mshPlan.Cols - 1
        mshPlan.Col = i
        mshPlan.CellBackColor = mshPlan.BackColor
        mshPlan.CellForeColor = mshPlan.ForeColor
    Next
    
    mshPlan.Redraw = blnPre
End Sub

Private Sub mshPlan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshPlan.MouseRow = 0 Then
        mshPlan.MousePointer = 99
    Else
        mshPlan.MousePointer = 0
    End If
End Sub

Private Sub mshPlan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCol As Integer, intRow As Integer
    
    intCol = mshPlan.MouseCol
    intRow = mshPlan.MouseRow
    If intRow = 0 Then
        mshPlan.ColData(intCol) = (mshPlan.ColData(intCol) + 1) Mod 2
        mshPlan.ColSel = mshPlan.Col
        mshPlan.Sort = Val(mshPlan.ColData(intCol)) + 1 '1-升,2-降
    End If
End Sub

Private Sub mshPlan_SelChange()
    If mshPlan.Rows = 2 Then Exit Sub
    mshPlan.RowSel = mshPlan.Row
End Sub




Private Sub mshList_EnterCell()
    Dim i As Integer, blnPre As Boolean
    Dim intRow As Integer, intCol As Integer
    
    blnPre = mshList.Redraw
    intRow = mshList.Row: intCol = mshList.Col
    mshList.Redraw = False
    
    For i = 0 To mshList.Cols - 1
        mshList.Col = i
        mshList.CellBackColor = mshList.BackColorSel
        mshList.CellForeColor = mshList.ForeColorSel
    Next
    
    mshList.Row = intRow:  mshList.Col = intCol
    mshList.Redraw = blnPre
End Sub

Private Sub mshList_LeaveCell()
    Dim i As Integer, blnPre As Boolean
    
    blnPre = mshList.Redraw
    mshList.Redraw = False
    
    For i = 0 To mshList.Cols - 1
        mshList.Col = i
        mshList.CellBackColor = mshList.BackColor
        mshList.CellForeColor = mshList.ForeColor
    Next
    
    mshList.Redraw = blnPre
End Sub

Private Sub mshList_SelChange()
    If mshList.Rows = 2 Then Exit Sub
    mshList.RowSel = mshList.Row
End Sub
