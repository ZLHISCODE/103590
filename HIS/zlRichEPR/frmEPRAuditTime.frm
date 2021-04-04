VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRAuditTime 
   BorderStyle     =   0  'None
   Caption         =   "病历内容监测"
   ClientHeight    =   3840
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   6975
   Icon            =   "frmEPRAuditTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   0
      Left            =   465
      ScaleHeight     =   2535
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   270
      Width           =   5145
      Begin VSFlex8Ctl.VSFlexGrid vfgAudit 
         Height          =   1560
         Left            =   375
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   3960
         _cx             =   6985
         _cy             =   2752
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
         GridLinesFixed  =   2
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
   End
End
Attribute VB_Name = "frmEPRAuditTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    标志 = 0: 姓名: 病人ID: 主页ID: 事件缘由: 应写病历: 应写科室: 监测点: 基点时间: 要求时间: 完成时间: 完成记录id: 当前时间: 备注说明
End Enum

Private mintKind As Integer     '病历种类
Private mstrDateFrom As String  '开始日期
Private mstrDateTo As String    '结束日期
Private mlngMoual As Long
Private mclsAudit As clsVsf
Private mintType As Integer
Public Event AfterDocumentChanged(ByVal lngEPRKey As Long)
Public Event SelectVfgRow(ByVal strPatiInfo As String)
Public Event GotFocus()

'######################################################################################################################

Public Function zlInitData(ByVal frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
'    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始控件") = False Or ExecuteCommand("初始数据") = False Then Exit Function
    
End Function

Public Sub zlClearData()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mclsAudit.ClearGrid
End Sub

Public Sub zlPrintData(ByVal bytMode As Byte, Optional ByVal strPatiInfo As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow


    Set objPrint.Body = vfgAudit
    objPrint.Title.Text = "病历时限监测记录"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add(strPatiInfo)
    Call objPrint.UnderAppRows.Add(objAppRow)

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

Public Function GetCurrentEPRKey() As Long
    With vfgAudit
        GetCurrentEPRKey = Val(.TextMatrix(.Row, mCol.完成记录id))
    End With
End Function

Public Function zlRefreshData(ByVal lngPatientKey As Long, ByVal lngPatientPageKey As Long, ByVal intKind As Integer, _
    Optional ByVal lngDeptId As Long, Optional ByVal intType As Integer = 1, Optional ByVal intState As Integer = 0, _
    Optional ByVal dtEndTime As Date) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：lngDeptId外部科室ID，intType：1-当前病人；2-我的病人；3-本科病人
    '返回：
    '******************************************************************************************************************
    
    Dim lngPatiID As Long, lngPageId As Long
    Dim lngBalance As Long
    Dim rs As New ADODB.Recordset
    Dim lngCount As Long
    
    If lngPatientKey = 0 Then Exit Function
    
    mintKind = intKind
    mintType = intType
    '提取时限监测数据
    Call ExecuteCommand("时限监测", lngPatientKey, lngPatientPageKey)
     Select Case intType
            Case 1 '当前病人
            gstrSQL = "Select 0 As 标记,'' as 姓名,a.病人id ,a.主页id ,To_Char(事件时间, 'yyyy-mm-dd hh24:mi ') || 事件 As 事件缘由, 病历编号 || '-' || 病历名称 As 应写病历,b.名称 As 应写科室," & _
            "        Decode(唯一, 1, '书写', '第' || 周期号 || '次书写') As 监测点, 开始时间 基点时间, 到期时间 要求时间, 完成时间, 完成记录id, Sysdate As 当前时间, Null As 备注说明" & _
            " From 电子病历时机 a,部门表 b" & _
            " Where a.病人id = [1] And a.主页id = [2] And (a.病历种类 = [3] Or a.病历种类 in (5,6) And [3]<>4) And a.到期时间 - Sysdate < 2 And a.科室id=b.ID" & _
            " Order By a.事件时间,a.病历编号,A.开始时间"
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatientKey, lngPatientPageKey, intKind)
            Case 2 '我的病人
            gstrSQL = "Select 0 As 标记,d.姓名,a.病人id ,a.主页id ,To_Char(事件时间, 'yyyy-mm-dd hh24:mi ') || 事件 As 事件缘由, 病历编号 || '-' || 病历名称 As 应写病历,b.名称 As 应写科室," & _
            "        Decode(唯一, 1, '书写', '第' || 周期号 || '次书写') As 监测点, 开始时间 基点时间, 到期时间 要求时间, 完成时间, 完成记录id, Sysdate As 当前时间, Null As 备注说明" & _
            " From 电子病历时机 a,部门表 b,病案主页 c,病人信息 D" & _
            " Where  a.主页id = [1] And (a.病历种类 = [2] Or a.病历种类 in (5,6) And [2]<>4) And a.到期时间 - Sysdate < 2 And a.科室id=b.ID and C.病人ID=a.病人id  " & _
            " and c.主页id=a.主页id And d.病人id=c.病人id and"
             If intState = 2 Then
                gstrSQL = gstrSQL & " c.出院日期>=[4]"
              Else
                gstrSQL = gstrSQL & " D.在院=1"
             End If
             If intKind = 1 Then
             gstrSQL = gstrSQL & " and c.门诊医师=[3] " & " Order By a.事件时间,a.病历编号,A.开始时间"
             ElseIf intKind = 2 Then
             gstrSQL = gstrSQL & " and c.住院医师=[3] " & " Order By a.事件时间,a.病历编号,A.开始时间"
             Else
             gstrSQL = gstrSQL & " and c.责任护士=[3] " & " Order By a.事件时间,a.病历编号,A.开始时间"
             End If
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatientPageKey, intKind, gstrUserName, dtEndTime)
            Case 3 '本科病人
             gstrSQL = "Select 0 As 标记,d.姓名,a.病人id ,a.主页id ,To_Char(事件时间, 'yyyy-mm-dd hh24:mi ') || 事件 As 事件缘由, 病历编号 || '-' || 病历名称 As 应写病历,b.名称 As 应写科室," & _
            "        Decode(唯一, 1, '书写', '第' || 周期号 || '次书写') As 监测点, 开始时间 基点时间, 到期时间 要求时间, 完成时间, 完成记录id, Sysdate As 当前时间, Null As 备注说明" & _
            " From 电子病历时机 a,部门表 b,病案主页 c,病人信息 D" & _
            " Where a.主页id = [1] And (a.病历种类 = [2] Or a.病历种类 in (5,6) And [2]<>4) And a.到期时间 - Sysdate < 2 And a.科室id=b.ID and"
             If intState = 2 Then
                gstrSQL = gstrSQL & " c.出院日期>=[4]"
             Else
                gstrSQL = gstrSQL & " D.在院=1"
             End If
             gstrSQL = gstrSQL & " and C.病人ID=a.病人id  and c.主页id=a.主页id And d.病人id=c.病人id and c.出院科室ID=[3] Order By a.事件时间,a.病历编号,A.开始时间"
             Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatientPageKey, intKind, lngDeptId, dtEndTime)
     End Select
    
    With Me.vfgAudit
        .Clear
        .FixedCols = 0
        Set .DataSource = rs
       
        .MergeCells = flexMergeFree: .MergeCol(mCol.事件缘由) = True: .MergeCol(mCol.应写病历) = True: .MergeCol(mCol.姓名) = True
        .ColWidth(mCol.标志) = 250: .ColWidth(mCol.基点时间) = 1800: .ColWidth(mCol.要求时间) = 1800: .ColWidth(mCol.完成时间) = 1800
        .ColWidth(mCol.完成记录id) = 0: .ColWidth(mCol.当前时间) = 0: .ColWidth(mCol.备注说明) = 2200: .ColWidth(mCol.病人ID) = 0
        .ColWidth(mCol.主页ID) = 0
         If mintType = 1 Then
         .ColWidth(mCol.姓名) = 0
         Else
         .ColWidth(mCol.姓名) = 1000
         End If
        .FixedCols = 1
        .TextMatrix(0, mCol.标志) = ""
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
            .TextMatrix(lngCount, mCol.基点时间) = Format(.TextMatrix(lngCount, mCol.基点时间), "yyyy-MM-dd HH:mm")
            .TextMatrix(lngCount, mCol.要求时间) = Format(.TextMatrix(lngCount, mCol.要求时间), "yyyy-MM-dd HH:mm")
            .TextMatrix(lngCount, mCol.完成时间) = Format(.TextMatrix(lngCount, mCol.完成时间), "yyyy-MM-dd HH:mm")
        Next
        .Row = 0
    End With
    
    zlRefreshData = True
    
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
                
        '------------------------------------------------------------------------------------------------------------------
        Set mclsAudit = New clsVsf
        With mclsAudit

            Call .Initialize(Me.Controls, vfgAudit, True, False)
            Call .ClearColumn

            Call .AppendColumn("", 250, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("事件缘由", 1800, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("姓名", 1800, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("病人ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("主页ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("应写病历", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("应写科室", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("监测点", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("基点时间", 1800, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
            Call .AppendColumn("要求时间", 1800, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
            Call .AppendColumn("完成时间", 1800, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
            Call .AppendColumn("完成记录id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("当前时间", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
            Call .AppendColumn("备注说明", 900, flexAlignLeftCenter, flexDTString, "", "", True)

        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"


    '------------------------------------------------------------------------------------------------------------------
    Case "时限监测"
        
        strSQL = "zl_电子病历时机_makeup(" & Val(varParam(0)) & "," & Val(varParam(1)) & "," & IIf(mintKind = 1, 1, 2) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "病历时限监测")
        
    End Select

    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:

End Function

Private Sub Form_Resize()

    On Error Resume Next
    
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsAudit Is Nothing) Then Set mclsAudit = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    vfgAudit.Move 0, 0, picPane(Index).Width, picPane(Index).Height
End Sub

Private Sub vfgAudit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strPatiInfo As String, lngPatiID As Long, lngPageId As Long
    Dim rsTemp As New ADODB.Recordset
    With vfgAudit
        If OldRow <> NewRow And NewRow > 0 Then
            
            RaiseEvent AfterDocumentChanged(Val(.TextMatrix(NewRow, mCol.完成记录id)))
            lngPatiID = Val(.TextMatrix(NewRow, mCol.病人ID))
            lngPageId = Val(.TextMatrix(NewRow, mCol.主页ID))
            If mintType <> 1 Then
                If mintKind = 1 Then
                    gstrSQL = "Select r.门诊号, r.No, r.姓名, r.性别, r.年龄, r.登记时间 From 病人挂号记录 r Where r.Id =[1] And r.记录性质=1  and r.记录状态=1"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPageId)
                    With rsTemp
                        If .RecordCount <= 0 Then strPatiInfo = "该病人不存在，可能存在数据错误！"
                        strPatiInfo = "门诊号:" & !门诊号 & "(No:" & !NO & ")    姓名:" & !姓名 & "(" & !性别 & ")" & _
                                    "  日期:" & Format(!登记时间, "yyyy-MM-dd hh:mm")
                    End With
                Else
                    gstrSQL = "Select b.住院号, a.姓名, a.性别, a.年龄, b.出院病床 As 床号, b.入院日期" & _
                            " From 病人信息 a, 病案主页 b" & _
                            " Where a.病人id = b.病人id And b.病人id = [1] And Nvl(b.主页id, 0) = [2]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId)
                    With rsTemp
                        If .RecordCount <= 0 Then strPatiInfo = "该病人不存在，可能刚好被身份合并等！"
                        strPatiInfo = "住院号:" & !住院号 & "(第" & lngPageId & "次住院)    姓名:" & !姓名 & "(" & !性别 & ")" & _
                                    "  日期:" & Format(!入院日期, "yyyy-MM-dd hh:mm")
                    End With
                End If
                RaiseEvent SelectVfgRow(strPatiInfo)
            End If
            
            
        End If
    End With
End Sub

Private Sub vfgAudit_GotFocus()
    RaiseEvent GotFocus
End Sub

