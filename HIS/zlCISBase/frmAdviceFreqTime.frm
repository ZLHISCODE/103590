VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmAdviceFreqTime 
   Caption         =   "执行时间方案"
   ClientHeight    =   4875
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   7740
   Icon            =   "frmAdviceFreqTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7740
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   6480
      TabIndex        =   7
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "插入(&I)"
      Height          =   350
      Left            =   6495
      TabIndex        =   3
      Top             =   2430
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6495
      TabIndex        =   2
      Top             =   1200
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6495
      TabIndex        =   1
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdVaild 
      Caption         =   "合法性(&V)"
      Height          =   350
      Left            =   6495
      TabIndex        =   4
      Top             =   1980
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6495
      TabIndex        =   5
      Top             =   4350
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit mshTime 
      Height          =   3990
      Left            =   105
      TabIndex        =   0
      Top             =   765
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7038
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAdviceFreqTime.frx":030A
      Height          =   560
      Left            =   780
      TabIndex        =   6
      Top             =   90
      Width           =   5670
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   165
      Picture         =   "frmAdviceFreqTime.frx":03A3
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAdviceFreqTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrCode As String
Private mint频率次数 As Integer
Private mint频率间隔 As Integer
Private mstr间隔单位 As String
Private mint适用范围 As Integer '1-西医,2-中医
Private mblnChange As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim i As Long
    
    If mshTime.Row >= mshTime.MsfObj.FixedRows Then
        If mshTime.Rows = mshTime.MsfObj.FixedRows + 1 Then
            For i = 1 To mshTime.Cols - 1
                mshTime.TextMatrix(mshTime.MsfObj.FixedRows, i) = ""
            Next
        Else
            Call mshTime.MsfObj.RemoveItem(mshTime.Row)
        End If
    End If
    Call AdjustOrder
    mshTime.SetFocus
End Sub

Private Sub cmdInsert_Click()
    mshTime.MsfObj.AddItem "", mshTime.Row
    mshTime.Col = 1
    Call AdjustOrder
    mshTime.SetFocus
End Sub

Private Function CheckValid(lngRow As Long, lngCol As Long, strErr As String, Optional arrSql As Variant) As Boolean
    Dim strTime As String, i As Long, j As Long
    
    '先检查合法性
    arrSql = Array()
    For i = mshTime.MsfObj.FixedRows To mshTime.Rows - 1
        If mshTime.TextMatrix(i, 1) <> "" Then
            For j = 2 To mshTime.Cols - 1
                If mshTime.TextMatrix(i, j) = "" Then
                    strErr = "序号为 " & mshTime.TextMatrix(i, 0) & " 的时间方案中输入的数据不完整。"
                    lngRow = i: lngCol = j: Exit For
                End If
            Next
            If j <= mshTime.Cols - 1 Then Exit For
            
            strTime = ""
            If mstr间隔单位 = "周" Or mstr间隔单位 = "天" And mint频率间隔 > 1 Then
                For j = 2 To mshTime.Cols - 1 Step 2
                    strTime = strTime & "-" & mshTime.TextMatrix(i, j) & "/" & mshTime.TextMatrix(i, j + 1)
                Next
            ElseIf mstr间隔单位 = "小时" Or mstr间隔单位 = "天" And mint频率间隔 = 1 Then
                For j = 2 To mshTime.Cols - 1
                    strTime = strTime & "-" & mshTime.TextMatrix(i, j)
                Next
            End If
            strTime = Mid(strTime, 2)
            
            If Not ExeTimeValid(strTime, lngCol, strErr) Then
                If lngCol = 0 Then
                    lngCol = 2
                    strErr = "序号为 " & mshTime.TextMatrix(i, 0) & " 的时间方案中" & strErr
                Else
                    strErr = "序号为 " & mshTime.TextMatrix(i, 0) & " 的时间方案中" & mshTime.TextMatrix(0, lngCol) & strErr
                End If
                lngRow = i: Exit For
            End If
            
            If zlCommFun.ActualLen(strTime) > 50 Then
                strErr = "时间方案内容太长，这可能是由于频率次数太多所致，请先对该频率项目作适当调整。"
                lngRow = i: Exit For
            End If
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "ZL_诊疗频率时间_Insert('" & mstrCode & "'," & mshTime.TextMatrix(i, 0) & "," & _
                "'" & strTime & "'," & IIf(mshTime.RowData(i) = 0, "NULL", mshTime.RowData(i)) & ")"
        Else
            If i <> mshTime.Rows - 1 Then
                strErr = "序号为 " & mshTime.TextMatrix(i, 0) & " 的时间方案中输入的数据不完整。"
                lngRow = i: lngCol = 1: Exit For
            Else
                For j = 2 To mshTime.Cols - 1
                    If mshTime.TextMatrix(i, j) <> "" Then
                        strErr = "序号为 " & mshTime.TextMatrix(i, 0) & " 的时间方案中输入的数据不完整。"
                        lngRow = i: lngCol = 1: Exit For
                    End If
                Next
                If j <= mshTime.Cols - 1 Then Exit For
            End If
        End If
    Next
    CheckValid = Not (i <= mshTime.Rows - 1)
End Function

Private Sub cmdOK_Click()
    Dim arrSql As Variant, strErr As String
    Dim lngRow As Long, lngCol As Long, i As Long
    
    If Not CheckValid(lngRow, lngCol, strErr, arrSql) Then
        mshTime.Row = lngRow: mshTime.Col = lngCol
        Call mshTime_EnterCell(lngRow, lngCol)
        If lngRow - mshTime.Height \ mshTime.RowHeight(0) \ 2 < mshTime.MsfObj.FixedRows Then
            mshTime.MsfObj.TopRow = mshTime.MsfObj.FixedRows
        Else
            mshTime.MsfObj.TopRow = lngRow - mshTime.Height \ mshTime.RowHeight(0) \ 2
        End If
        If strErr <> "" Then MsgBox strErr, vbInformation, gstrSysName
        mshTime.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure("ZL_诊疗频率时间_Delete('" & mstrCode & "')", Me.Caption)
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    mblnChange = False
    gblnOK = True
    Unload Me
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdVaild_Click()
    Dim strErr As String, lngRow As Long, lngCol As Long
    
    If Not CheckValid(lngRow, lngCol, strErr) Then
        mshTime.Row = lngRow: mshTime.Col = lngCol
        Call mshTime_EnterCell(lngRow, lngCol)
        If lngRow - mshTime.Height \ mshTime.RowHeight(0) \ 2 < mshTime.MsfObj.FixedRows Then
            mshTime.MsfObj.TopRow = mshTime.MsfObj.FixedRows
        Else
            mshTime.MsfObj.TopRow = lngRow - mshTime.Height \ mshTime.RowHeight(0) \ 2
        End If
        If strErr <> "" Then MsgBox strErr, vbInformation, gstrSysName
        mshTime.SetFocus
        Exit Sub
    Else
        MsgBox "所有输入数据正确。", vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If mshTime.TxtVisible Then
            mshTime.Text = "": mshTime.TxtVisible = False: mshTime.SetFocus
        Else
            Call cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    RestoreWinState Me, App.ProductName
    gblnOK = False
    mblnChange = False
        
    '频率项目信息
    strSql = "Select * From 诊疗频率项目 Where 编码=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrCode)
    
    mint频率次数 = Nvl(rsTmp!频率次数, 0)
    mint频率间隔 = Nvl(rsTmp!频率间隔, 0)
    mstr间隔单位 = Nvl(rsTmp!间隔单位)
    mint适用范围 = IIf(IsNull(rsTmp!适用范围), 1, rsTmp!适用范围)
    
    lblCaption.Caption = Replace(lblCaption.Caption, "XXXXX", rsTmp!名称)
    lblCaption.Caption = Replace(lblCaption.Caption, "YYYYY", IIf(mint适用范围 = 1, "给药途径", "中药用法"))
                    
    '显示表格格式及数据
    Call ShowTimeScheme(mstrCode)
    
    '列编辑属性
    mshTime.ColData(0) = 5
    mshTime.ColData(1) = 1
    For i = 2 To mshTime.Cols - 1
        mshTime.ColData(i) = 4
    Next
    mshTime.LocateCol = 1
    mshTime.PrimaryCol = 1

    mshTime.Col = 1
    mshTime.Row = mshTime.MsfObj.FixedRows
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With mshTime
        .Width = Me.ScaleWidth - .Left - cmdOK.Width - 350
        .Height = Me.ScaleHeight - .Top - 60
    End With
    
    cmdOK.Left = mshTime.Left + mshTime.Width + 200
    cmdCancel.Left = cmdOK.Left
    cmdVaild.Left = cmdOK.Left
    cmdInsert.Left = cmdOK.Left
    cmdHelp.Left = cmdOK.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("你已修改了相关内容，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    SaveWinState Me, App.ProductName
    mstrCode = ""
End Sub

Private Function ShowTimeScheme(ByVal str编码 As String) As Boolean
'功能：根据当前频率项目显示它的时间方案表
'参数：str编码=频率项目编码
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim arrTime As Variant
    
    On Error GoTo errH
    
    With mshTime.MsfObj
        .Clear
        .ClearStructure
        .FixedCols = 0: .FixedRows = 0
        .Rows = 0: .Cols = 0
        
        '频率时间方案
        strSql = _
            "Select A.方案序号,A.时间方案,A.给药途径ID,B.名称,B.编码" & _
            " From 诊疗频率时间 A,诊疗项目目录 B" & _
            " Where A.给药途径ID=B.ID(+) And A.执行频率=[1]" & _
            " Order by 方案序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str编码)
        
        '时间方案表头
        If mstr间隔单位 = "周" Or mstr间隔单位 = "天" And mint频率间隔 > 1 Then
            .Cols = 2 + mint频率次数 * 2
            .Rows = IIf(rsTmp.EOF, 1, rsTmp.RecordCount) + 2
            .FixedRows = 2
            .FixedCols = 1
            
            .TextMatrix(0, 0) = "序号": .TextMatrix(1, 0) = .TextMatrix(0, 0)
            .TextMatrix(0, 1) = IIf(mint适用范围 = 1, "给药途径", "中药用法"): .TextMatrix(1, 1) = .TextMatrix(0, 1)
            For i = 2 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "第" & ((i - 2) \ 2) + 1 & "次"
                .TextMatrix(0, i + 1) = .TextMatrix(0, i)
                If mstr间隔单位 = "周" Then
                    .TextMatrix(1, i) = "星期"
                    .TextMatrix(1, i + 1) = "时间"
                    .ColWidth(i) = 450
                    .ColWidth(i + 1) = 1000
                Else
                    .TextMatrix(1, i) = "天"
                    .TextMatrix(1, i + 1) = "时间"
                    .ColWidth(i) = 300
                    .ColWidth(i + 1) = 1000
                End If
                .ColAlignment(i) = 4
                .ColAlignment(i + 1) = 1
            Next
        ElseIf mstr间隔单位 = "小时" Or mstr间隔单位 = "天" And mint频率间隔 = 1 Then
            .Cols = 2 + mint频率次数
            .Rows = IIf(rsTmp.EOF, 1, rsTmp.RecordCount) + 1
            .FixedRows = 1
            .FixedCols = 1
            
            .TextMatrix(0, 0) = "序号"
            .TextMatrix(0, 1) = IIf(mint适用范围 = 1, "给药途径", "中药用法")
            For i = 2 To .Cols - 1
                .TextMatrix(0, i) = "第" & i - 1 & "次"
                .ColWidth(i) = 1000
                .ColAlignment(i) = 1
            Next
        End If
        .ColWidth(0) = 450
        .ColWidth(1) = 1800
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        For i = 0 To .Cols - 1
            .ColAlignmentFixed(i) = 4
        Next
        .MergeCells = flexMergeRestrictAll
        .MergeCol(0) = True: .MergeCol(1) = True
        .MergeRow(0) = True: .MergeRow(1) = True
        
        '时间数据
        Call AdjustOrder
        For i = 1 To rsTmp.RecordCount
            .RowData(i + .FixedRows - 1) = IIf(IsNull(rsTmp!给药途径ID), 0, rsTmp!给药途径ID)
            .TextMatrix(i + .FixedRows - 1, 0) = rsTmp!方案序号
            .TextMatrix(i + .FixedRows - 1, 1) = IIf(IsNull(rsTmp!名称), "<不确定>", rsTmp!编码 & "-" & rsTmp!名称)
            
            arrTime = Split(rsTmp!时间方案, "-")
            If mstr间隔单位 = "周" Or mstr间隔单位 = "天" And mint频率间隔 > 1 Then
                For j = 0 To mint频率次数 - 1
                    .TextMatrix(i + .FixedRows - 1, j * 2 + 2) = Split(arrTime(j), "/")(0)
                    .TextMatrix(i + .FixedRows - 1, j * 2 + 3) = Split(arrTime(j), "/")(1)
                Next
            ElseIf mstr间隔单位 = "小时" Or mstr间隔单位 = "天" And mint频率间隔 = 1 Then
                For j = 0 To mint频率次数 - 1
                    .TextMatrix(i + .FixedRows - 1, j + 2) = arrTime(j)
                Next
            End If
            rsTmp.MoveNext
        Next
    End With
    
    ShowTimeScheme = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mshTime_AfterAddRow(Row As Long)
    Call AdjustOrder(Row)
End Sub

Private Sub mshTime_AfterDeleteRow()
    Call AdjustOrder(mshTime.Row)
    mblnChange = True
End Sub

Private Sub AdjustOrder(Optional ByVal lngRow As Long)
    Dim i As Long
    
    If lngRow = 0 Then lngRow = mshTime.MsfObj.FixedRows
    
    For i = lngRow To mshTime.Rows - 1
        mshTime.TextMatrix(i, 0) = i - mshTime.MsfObj.FixedRows + 1
    Next
End Sub

Public Function ExeTimeValid(ByVal strTime As String, lngCol As Long, strErr As String) As Boolean
'功能：检查指定的执行时间是否合法
'返回：lngCol=具体出错的列
    Dim arrTime() As String, strTmp As String, i As Integer
    Dim strPreTime As String, intPreDay As Long, intCurDay As Long
    
    If strTime = "" Then Exit Function
    
    If mstr间隔单位 = "周" Then
        '1/8:00-3/15:00-5/9:00；1/8:00-3/15-5/9:00
        If Not StringMask(strTime, "0123456789:-/") Then
            strErr = "输入了非法字符。": Exit Function
        End If
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> mint频率次数 Then
            strErr = "输入了非法字符。": Exit Function
        End If
        
        For i = 0 To UBound(arrTime)
            If UBound(Split(arrTime(i), "/")) <> 1 Then
                strErr = "输入了非法字符。"
                lngCol = i * 2 + 2: Exit Function
            End If
            
            '星期部份
            strTmp = Split(arrTime(i), "/")(0)
            If InStr(strTmp, ":") > 0 Or strTmp = "" Then
                strErr = "的星期部份输入不正确。"
                lngCol = i * 2 + 2: Exit Function
            End If
            intCurDay = Val(strTmp)
            If intCurDay < 1 Or intCurDay > 7 Then
                strErr = "的星期必须在 1-7 之间。"
                lngCol = i * 2 + 2: Exit Function
            End If
            If intPreDay <> 0 Then
                If intCurDay < intPreDay Then
                    strErr = "的星期数不能小于上一次的星期数。"
                    lngCol = i * 2 + 2: Exit Function
                End If
            End If
            
            '绝对时间部分
            strTmp = Split(arrTime(i), "/")(1)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then
                strErr = "输入了多个时间分隔符"":""。"
                lngCol = i * 2 + 3: Exit Function
            End If
            If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then
                strErr = "的小时数没有输入或输入的小时数大于或等于了24小时。"
                lngCol = i * 2 + 3: Exit Function
            End If
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then
                strErr = "的分钟数没有输入或输入的分钟数大于或等于了60分钟。"
                lngCol = i * 2 + 3: Exit Function
            End If
            If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then
                    strErr = "在同一天的时间必须大于前一次的时间。"
                    lngCol = i * 2 + 3: Exit Function
                End If
            End If
            
            strPreTime = Format(strTmp, "HH:mm")
            intPreDay = intCurDay
        Next
    ElseIf mstr间隔单位 = "天" Then
        If mint频率间隔 = 1 Then
            '8:00-12:00-14:00；8:00-12-14:00
            If Not StringMask(strTime, "0123456789:-") Then
                strErr = "输入了非法字符。": Exit Function
            End If
            
            arrTime = Split(strTime, "-")
            If UBound(arrTime) + 1 <> mint频率次数 Then
                strErr = "输入了非法字符。": Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                strTmp = arrTime(i)
                
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then
                    strErr = "输入了多个时间分隔符"":""。"
                    lngCol = i + 2: Exit Function
                End If
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then
                    strErr = "的小时数没有输入或输入的小时数大于或等于了24小时。"
                    lngCol = i + 2: Exit Function
                End If
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then
                    strErr = "的分钟数没有输入或输入的分钟数大于或等于了60分钟。"
                    lngCol = i + 2: Exit Function
                End If
                If strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then
                        strErr = "的时间必须大于前一次的时间。"
                        lngCol = i + 2: Exit Function
                    End If
                End If
                strPreTime = Format(strTmp, "HH:mm")
            Next
        Else
            '1/8:00-1/15:00-2/9:00；1/8:00-1/15-2/9:00
            If Not StringMask(strTime, "0123456789:-/") Then
                strErr = "输入了非法字符。": Exit Function
            End If
            
            arrTime = Split(strTime, "-")
            If UBound(arrTime) + 1 <> mint频率次数 Then
                strErr = "输入了非法字符。": Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                If UBound(Split(arrTime(i), "/")) <> 1 Then
                    strErr = "输入了非法字符。"
                    lngCol = i * 2 + 2: Exit Function
                End If
                
                '相对天数部分
                strTmp = Split(arrTime(i), "/")(0)
                If InStr(strTmp, ":") > 0 Or strTmp = "" Then
                    strErr = "的天数部份输入不正确。"
                    lngCol = i * 2 + 2: Exit Function
                End If
                intCurDay = Val(strTmp)
                If intCurDay < 1 Or intCurDay > mint频率间隔 Then
                    strErr = "的天数必须在 1-" & mint频率间隔 & " 之间。"
                    lngCol = i * 2 + 2: Exit Function
                End If
                If intPreDay <> 0 Then
                    If intCurDay < intPreDay Then
                        strErr = "的天数不能小于上一次的天数。"
                        lngCol = i * 2 + 2: Exit Function
                    End If
                End If
                
                '绝对时间部分
                strTmp = Split(arrTime(i), "/")(1)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then
                    strErr = "输入了多个时间分隔符"":""。"
                    lngCol = i * 2 + 3: Exit Function
                End If
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then
                    strErr = "的小时数没有输入或输入的小时数大于或等于了24小时。"
                    lngCol = i * 2 + 3: Exit Function
                End If
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then
                    strErr = "的分钟数没有输入或输入的分钟数大于或等于了60分钟。"
                    lngCol = i * 2 + 3: Exit Function
                End If
                If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then
                        strErr = "在同一天的时间必须大于前一次的时间。"
                        lngCol = i * 2 + 3: Exit Function
                    End If
                End If
                
                strPreTime = Format(strTmp, "HH:mm")
                intPreDay = intCurDay
            Next
        End If
    ElseIf mstr间隔单位 = "小时" Then
        '1:30-2-3:30
        If Not StringMask(strTime, "0123456789:-") Then
            strErr = "输入了非法字符。": Exit Function
        End If
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> mint频率次数 Then
            strErr = "输入了非法字符。": Exit Function
        End If
        
        For i = 0 To UBound(arrTime)
            strTmp = arrTime(i)
            
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then
                strErr = "输入了多个时间分隔符"":""。"
                lngCol = i + 2: Exit Function
            End If
            If Val(Split(strTmp, ":")(0)) < 1 Or Val(Split(strTmp, ":")(0)) > mint频率间隔 Or Split(strTmp, ":")(0) = "" Then
                strErr = "的小时数没有输入或输入的小时数不在 1-" & mint频率间隔 & "小时 之间。"
                lngCol = i + 2: Exit Function
            End If
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then
                strErr = "的分钟数没有输入或输入的分钟数大于或等于了60分钟。"
                lngCol = i + 2: Exit Function
            End If
            If strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then
                    strErr = "的时间必须大于前一次的时间。"
                    lngCol = i + 2: Exit Function
                End If
            End If
            strPreTime = Format(strTmp, "HH:mm")
        Next
    End If
    
    ExeTimeValid = True
End Function

Public Function StringMask(ByVal strText As String, ByVal strMask As String) As Boolean
'功能：检查字符串是否只包含指定的字符
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strMask, Mid(strText, i, 1)) = 0 Then Exit Function
    Next
    StringMask = True
End Function

Private Sub mshTime_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long
    
    If Row = mshTime.MsfObj.FixedRows And mshTime.Rows = mshTime.MsfObj.FixedRows + 1 Then
        Cancel = True
        For i = 1 To mshTime.Cols - 1
            mshTime.TextMatrix(Row, i) = ""
        Next
        Call AdjustOrder(Row)
        mblnChange = True
    End If
End Sub

Private Sub mshTime_EditChange(curText As String)
    If Visible Then mblnChange = True
End Sub

Private Sub mshTime_EditKeyPress(KeyAscii As Integer)
    If mshTime.ColData(mshTime.Col) = 4 Then
        If InStr("01234567890:" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub mshTime_EnterCell(Row As Long, Col As Long)
    If mshTime.ColData(Col) = 4 Then
        mshTime.MaxLength = 5
    Else
        mshTime.MaxLength = 0
    End If
End Sub

Private Sub mshTime_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strText As String, strLike As String
    Dim vPoint As POINTAPI
    
    If KeyCode = 13 And mshTime.Col = 1 And mshTime.TxtVisible Then
        '用法信息
        strLike = gstrMatch
        strText = Replace(UCase(mshTime.Text), "'", "''")
        strSql = _
            " Select Distinct ID,排序ID,编码,名称 From (" & _
            " Select 0 as 排序ID,0 as ID,'-' as 编码,'<不确定>' as 名称,NULL as 名称ID,NULL as 简码ID From Dual Union ALL" & _
            " Select 1 as 排序ID,A.ID,A.编码,A.名称,B.名称 as 名称ID,B.简码 as 简码ID" & _
            " From 诊疗项目目录 A,诊疗项目别名 B" & _
            " Where A.类别='E' And A.操作类型='" & IIf(mint适用范围 = 1, 2, 4) & "'" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And A.ID=B.诊疗项目ID)" & _
            " Where (编码 Like '" & strText & "%'" & _
            " Or Upper(名称) Like '" & strLike & strText & "%'" & _
            " Or Upper(名称ID) Like '" & strLike & strText & "%'" & _
            " Or Upper(简码ID) Like '" & strLike & strText & "%')" & _
            " Order by 排序ID,编码"
        With mshTime.MsfObj
            vPoint = zlControl.GetCoordPos(.hWnd, .CellLeft - 30, .CellTop - 45)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "用法", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
        End With
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有找到匹配的" & IIf(mint适用范围 = 1, "给药途径", "中药用法") & "。", vbInformation, gstrSysName
            End If
            mshTime.TxtVisible = False
            Cancel = True
        Else
            mshTime.TxtVisible = False
            mshTime.RowData(mshTime.Row) = rsTmp!ID
            mshTime.TextMatrix(mshTime.Row, mshTime.Col) = IIf(rsTmp!编码 = "-", rsTmp!名称, rsTmp!编码 & "-" & rsTmp!名称)
            mblnChange = True
        End If
    End If
End Sub

Private Sub mshTime_CommandClick()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    '用法信息
    strSql = _
        " Select 0 as 排序ID,0 as ID,'-' as 编码,'<不确定>' as 名称 From Dual Union ALL" & _
        " Select 1 as 排序ID,A.ID,A.编码,A.名称 From 诊疗项目目录 A" & _
        " Where A.类别='E' And A.操作类型='" & IIf(mint适用范围 = 1, 2, 4) & "'" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
        " Order by 排序ID,编码"
    With mshTime.MsfObj
        vPoint = zlControl.GetCoordPos(.hWnd, .CellLeft - 30, .CellTop - 45)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "用法", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
    End With
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有" & IIf(mint适用范围 = 1, "给药途径", "中药用法") & "数据,请先到诊疗项目管理中设置。", vbInformation, gstrSysName
        End If
    Else
        mshTime.RowData(mshTime.Row) = rsTmp!ID
        mshTime.TextMatrix(mshTime.Row, mshTime.Col) = IIf(rsTmp!编码 = "-", rsTmp!名称, rsTmp!编码 & "-" & rsTmp!名称)
    End If
End Sub

Private Sub mshTime_KeyPress(KeyAscii As Integer)
    If mshTime.ColData(mshTime.Col) = 4 Then
        If InStr("01234567890:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub
