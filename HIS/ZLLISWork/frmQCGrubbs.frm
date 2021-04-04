VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmQCGrubbs 
   BorderStyle     =   0  'None
   Caption         =   "Grubbs质控记录表"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cboQCitem 
      Height          =   300
      Left            =   4380
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   6750
      Width           =   2595
   End
   Begin VB.OptionButton opt质控品 
      Caption         =   "473843A低值质控品"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   6855
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgData 
      Height          =   5010
      Left            =   450
      TabIndex        =   0
      Top             =   1620
      Width           =   9105
      _cx             =   16060
      _cy             =   8837
      Appearance      =   2
      BorderStyle     =   0
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
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
      Cols            =   12
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2475
      TabIndex        =   2
      Top             =   435
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmQCGrubbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    日期 = 1:  次数: 测定值: 均值: SD: SI上限: SI下限: N: n3s: n2s: 结果: 检验者
End Enum

Private mstrResList As String
Private mlngItemID As Long
Private mstrFromDate As String
Private mstrToDate As String
Private mstr质控品期限 As String

Public Function zlRefresh(strResList As String, lngItemID As Long, strFromDate As String, strToDate As String, str质控品期限 As String) As Boolean
    '功能：刷新本窗体的数据显示内容
    '参数： strResList  当前选择的质控品id串，以逗号分隔
    '       lngItemId   当前项目id
    '       strFromDate 开始日期
    '       strToDate   结束日期
    Dim rsTemp As New adodb.Recordset
    Dim intCounts As Integer
    Dim lngResId As Long
    Dim lngCount As Long
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr质控品期限 = str质控品期限
    Me.Tag = "不刷新"
    
    lngResId = 0
    intCounts = Me.cboQCitem.ListCount
    For lngCount = intCounts - 1 To 1 Step -1
        If Me.cboQCitem.ListIndex = lngCount Then lngResId = Val(Me.cboQCitem.ItemData(lngCount))
'        Unload Me.opt质控品(Me.opt质控品.UBound)
    Next
    cboQCitem.Clear
    
    Me.opt质控品(0).Enabled = False
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "Select A.ID, A.批号 || '-' || A.名称 As 质控品, B.对数质控图 From 检验质控品 A,检验仪器 B Where A.仪器ID=B.ID(+) And Instr(',' || [1] || ',', ',' || A.ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList)
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.cboQCitem.ListCount Then cboQCitem.AddItem "" & !质控品
            cboQCitem.ItemData(cboQCitem.NewIndex) = !ID
'            If .AbsolutePosition > Me.opt质控品.Count Then Load Me.opt质控品(.AbsolutePosition - 1)
'            Me.opt质控品(.AbsolutePosition - 1).Caption = "" & !质控品
'            Me.opt质控品(.AbsolutePosition - 1).Tag = !ID
'            Me.opt质控品(.AbsolutePosition - 1).Width = Me.TextWidth(Me.opt质控品(.AbsolutePosition - 1).Caption) + 360
'            Me.opt质控品(.AbsolutePosition - 1).Value = (lngResId = !ID)
'            Me.opt质控品(.AbsolutePosition - 1).Visible = True
'            Me.opt质控品(.AbsolutePosition - 1).Enabled = True
            
            .MoveNext
        Loop
    End With
    If rsTemp.RecordCount > 0 Then Me.cboQCitem.ListIndex = 0
    Me.Tag = ""
    Call RefGrid
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog

End Function

Private Sub RefGrid()
'
    Dim rsTemp As New adodb.Recordset
    Dim lngResId As Long, strLable As String, strUnit As String
    Dim intFormatNum As Integer, curTotal As Currency, strData As String
    Dim lng次数 As Long, strLast日期 As String, cur均值, curSD As Currency, curMax As Currency, curMin As Currency
    Dim curSI上 As Currency, curSI下 As Currency, curn3s As Currency, curn2s As Currency, curCV As Currency
    Dim lngCount As Long, lngRow As Long, iCol As Integer
    On Error GoTo ErrHandle
    lngResId = 0
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then lngResId = Val(Me.cboQCitem.ItemData(lngCount))
    Next
    If lngResId = 0 Then
        Me.opt质控品(0).Enabled = False
        Me.opt质控品(0).Value = True
        lngResId = Val(Me.opt质控品(0).Tag)
        Me.opt质控品(0).Enabled = True
    End If
    
    '获取小数位数
'    gstrSql = "Select nvl(小数位数,2) as 小数位数 from 检验仪器项目 where 项目ID = [1] "
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngItemId)
'    If rsTemp.EOF = False Then intFormatNum = Val("" & rsTemp("小数位数"))
    intFormatNum = 3
    Call initVfgData
    
    '获得基本的文字信息

    gstrSql = "Select Distinct RPad('单位：' || '" & gstrUnitName & "', 56, ' ') || '日期：' As 行0," & vbNewLine & _
            "                RPad('仪器：' || D.名称, 56, ' ') || '试剂来源：' || M.试剂 As 行1," & vbNewLine & _
            "                RPad('项目：' || I.项目, 56, ' ') || '校准物来源：' || M.校准物 As 行2" & vbNewLine & _
            "From 检验仪器 D, 检验质控品 M, (Select 中文名 || ',' || 英文名 As 项目 From 诊治所见项目 Where ID = [2]) I" & vbNewLine & _
            "Where D.ID = M.仪器id And Instr(',' || [1] || ',', ',' || M.ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrResList, mlngItemID)
    Me.vfgData.Visible = True
    Me.lblInfo.Visible = False
    If rsTemp.RecordCount <= 0 Then
       Me.lblInfo.Caption = "该质控品信息不全面！"
       Me.lblInfo.Visible = True
       Me.vfgData.Visible = False
       Exit Sub
    End If
    '表头附项
    With vfgData
        For iCol = .FixedCols To .Cols - 1
            
            .TextMatrix(1, iCol) = "  " & rsTemp!行0 & Format(mstrFromDate, "yyyy年MM月dd日") & "～" & Format(mstrToDate, "yyyy年MM月dd日") & vbCrLf & _
                                   "  " & rsTemp!行1 & vbCrLf & "  " & rsTemp!行2
        Next
        .Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = flexAlignLeftCenter
    End With
    
    Call Form_Resize
    
    gstrSql = "Select Q.检验时间 as 日期,Q.测试次数, Zl_Lis_Tonumber(Q.质控品id, R.检验项目id, R.检验结果,R.ID) As 结果,R.弃用结果,Q.检验人 " & vbNewLine & _
            "From 检验质控记录 Q, 检验普通结果 R,检验质控品 M, 检验质控均值 X " & vbNewLine & _
            "Where Q.标本id = R.检验标本id And Q.标本id = R.检验标本id And" & vbNewLine & _
            "      Nvl(R.弃用结果,0)=0 And Q.质控品id =[1] And R.检验项目id + 0 = [2] And" & vbNewLine & _
            "      Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd') And " & vbNewLine & _
            "      (Q.检验时间 Between X.开始日期 And NVL(X.结束日期,M.结束日期)) And " & vbNewLine & _
            "       Q.质控品id=M.id And M.id=X.质控品id  And  X.项目ID = [2] And " & vbNewLine & _
            "      Instr(';'||[5]||';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "Order By Q.检验时间,Q.测试次数"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstrFromDate, mstrToDate, mstr质控品期限)
    lng次数 = 0
    With vfgData
        Do Until rsTemp.EOF
            If strLast日期 <> Format("" & rsTemp!日期, "yyyy-MM-dd") & "," & rsTemp!测试次数 Then
                lng次数 = lng次数 + 1
                .TextMatrix(.Rows - 1, mCol.日期) = Format("" & rsTemp!日期, "yyyy-MM-dd")
                .TextMatrix(.Rows - 1, mCol.次数) = lng次数
                .TextMatrix(.Rows - 1, mCol.测定值) = Format(Val("" & rsTemp!结果), "0." & String(intFormatNum, "0"))
                .TextMatrix(.Rows - 1, mCol.检验者) = "" & rsTemp!检验人
                    
                .Rows = .Rows + 1
                If lng次数 >= 20 Then Exit Do
            ElseIf strLast日期 <> "" Then
                .TextMatrix(.Rows - 2, mCol.测定值) = Format(Val("" & rsTemp!结果), "0." & String(intFormatNum, "0"))
                .TextMatrix(.Rows - 2, mCol.检验者) = "" & rsTemp!检验人
            End If
            strLast日期 = Format("" & rsTemp!日期, "yyyy-MM-dd") & "," & rsTemp!测试次数
            rsTemp.MoveNext
        Loop
        curTotal = 0
        
        If .Rows > 4 Then .Rows = .Rows - 1
        
        If .Rows > 3 Then
            gstrSql = "Select n,n3s,n2s From 质控即刻法 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
        
            For lngRow = 3 To .Rows - 1
                strData = strData & "," & Val(.TextMatrix(lngRow, mCol.测定值))
                If Val(.TextMatrix(lngRow, mCol.次数)) > 2 Then
                
                    curn3s = Val("" & rsTemp!n3s)
                    curn2s = Val("" & rsTemp!n2s)
                    
                    .TextMatrix(lngRow, mCol.N) = rsTemp!N
                    .TextMatrix(lngRow, mCol.n3s) = curn3s
                    .TextMatrix(lngRow, mCol.n2s) = curn2s
                    

                    cur均值 = s(strData): curSD = stdev(strData)
                    curMax = Max(strData): curMin = Min(strData)
                    If curSD <> 0 Then
                        curSI上 = curMax / curSD - cur均值 / curSD: curSI下 = cur均值 / curSD - curMin / curSD
                    End If
                    .TextMatrix(lngRow, mCol.均值) = Format(cur均值, "0." & String(intFormatNum, "0"))
                    .TextMatrix(lngRow, mCol.SD) = Format(curSD, "0." & String(intFormatNum, "0"))
                    .TextMatrix(lngRow, mCol.SI上限) = Format(curSI上, "0.00")
                    .TextMatrix(lngRow, mCol.SI下限) = Format(curSI下, "0.00")
                    
                    If curSI上 > curn3s Or curSI下 > curn3s Then              '090504 有一个大于3s 则失控
                        .TextMatrix(lngRow, mCol.结果) = "#" '失控
                        .Cell(flexcpForeColor, lngRow, mCol.结果) = &H40C0&   '090504 都小于2s 则在控
                    ElseIf curSI上 < curn2s And curSI下 < curn2s Then
                        .TextMatrix(lngRow, mCol.结果) = "*" '在控
                    Else
                        .TextMatrix(lngRow, mCol.结果) = "！" '报警
                        .Cell(flexcpForeColor, lngRow, mCol.结果) = &H80C0FF
                    End If
                    rsTemp.MoveNext
                End If
            Next
        End If
        '最后一行
        
        .Rows = .Rows + 1
        lngRow = .Rows - 1
        
        .TextMatrix(.Rows - 1, mCol.日期) = lng次数 & "次在控数据测定："
        .TextMatrix(.Rows - 1, mCol.次数) = lng次数 & "次在控数据测定："
        
        If Val(.TextMatrix(.Rows - 2, mCol.测定值)) <> 0 And _
           Val(.TextMatrix(.Rows - 2, mCol.次数)) > 2 And _
           Val(.TextMatrix(.Rows - 2, mCol.次数)) < 21 Then
            curCV = curSD / Val(.TextMatrix(.Rows - 2, mCol.测定值)) * 100
        End If
        For iCol = mCol.测定值 To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = "均值=" & Format(cur均值, "0.000") & Space(10) & "SD=" & Format(curSD, "0.000") & Space(10) & "CV%=" & Format(curCV, "0.000")
        Next
        .MergeRow(.Rows - 1) = True
        
        .Select 2, .FixedCols, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 2, .FixedCols
        '表尾 说明
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.日期) = "说明："
        For iCol = mCol.次数 To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = "1、SI上限值和下限值  < n2s  时为在控，n2s — n3s 之间为告警状态，> n3s  时"
        Next
        .MergeRow(.Rows - 1) = True
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.日期) = ""
        For iCol = mCol.次数 To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = " 　结果以“*”表示在控，“！”表示告警，“#”表示失控。"
        Next
        .MergeRow(.Rows - 1) = True
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.日期) = ""
        For iCol = mCol.次数 To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = "2、失控数据及原因要填写“室内质控失控报告”。"
        Next
        .MergeRow(.Rows - 1) = True
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.日期) = ""
        For iCol = mCol.次数 To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = "3、遇告警和失控状态，纠错重测。舍弃告警或失控的数值，其它测定数值继续使用。"
        Next
        .MergeRow(.Rows - 1) = True
        
        .Cell(flexcpAlignment, lngRow, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .Cell(flexcpAlignment, lngRow + 1, .FixedCols) = flexAlignRightCenter
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Function s(ByVal strVal As String) As Currency
'   均值
    Dim varInData As Variant, curX As Currency, i As Integer
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    For i = LBound(varInData) To UBound(varInData)
        curX = curX + Val(varInData(i))
    Next
    If i > 0 Then
        s = curX / i
    End If
End Function
Private Function stdev(ByVal strVal As String) As Currency
    '标准差
    Dim varInData As Variant, curX As Currency, i As Integer, cur均值 As Currency
    
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    cur均值 = s(strVal)
    For i = LBound(varInData) To UBound(varInData)
        curX = curX + (Val(varInData(i)) - cur均值) ^ 2
    Next
    If i - 1 > 0 Then
        stdev = Sqr(curX / (i - 1))
    End If
    'Sqr (∑(xn - x拨) ^ 2 / (N - 1))
End Function

Private Function Max(ByVal strVal As String) As Currency
    Dim varInData As Variant, curX As Currency, i As Integer
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    For i = LBound(varInData) To UBound(varInData)
        If i = LBound(varInData) Then
            curX = Val(varInData(i))
        Else
            If curX < Val(varInData(i)) Then curX = Val(varInData(i))
        End If
    Next
    Max = curX
End Function

Private Function Min(ByVal strVal As String) As Currency
    Dim varInData As Variant, curX As Currency, i As Integer
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    For i = LBound(varInData) To UBound(varInData)
        If i = LBound(varInData) Then
            curX = Val(varInData(i))
        Else
            If curX > Val(varInData(i)) Then curX = Val(varInData(i))
        End If
    Next
    Min = curX
End Function

Private Sub initVfgData()
    Dim iCol As Integer
    With vfgData
        .Editable = flexEDNone
        .GridLines = flexGridNone
        .Rows = 4: .Cols = 13
        .FixedCols = 1: .FixedRows = 3
        .MergeCells = flexMergeRestrictRows
        .BackColorFixed = .BackColor
        .ForeColorFixed = .ForeColor
        .GridColorFixed = .GridColor
        .GridLinesFixed = flexGridNone
        
        '-- 表头
        For iCol = 0 To 1
            .MergeRow(iCol) = True
        Next
        
        For iCol = .FixedCols To .Cols - 1
            .TextMatrix(0, iCol) = "即刻法-室内质控记录表"
        Next
        .Cell(flexcpFontSize, 0, .FixedCols, 0, .Cols - 1) = 18
        .Cell(flexcpFontBold, 0, .FixedCols, 0, .Cols - 1) = True
        .RowHeight(0) = 500
        .RowHeight(1) = 700
        .TextMatrix(2, mCol.日期) = "日期": .ColWidth(mCol.日期) = 1100: .ColAlignment(mCol.日期) = flexAlignCenterCenter
        .TextMatrix(2, mCol.次数) = "次数": .ColWidth(mCol.次数) = 600: .ColAlignment(mCol.次数) = flexAlignCenterCenter
        .TextMatrix(2, mCol.测定值) = "测定值": .ColWidth(mCol.测定值) = 900: .ColAlignment(mCol.测定值) = flexAlignCenterCenter
        .TextMatrix(2, mCol.均值) = "均值": .ColWidth(mCol.均值) = 900: .ColAlignment(mCol.均值) = flexAlignCenterCenter
        .TextMatrix(2, mCol.SD) = "SD": .ColWidth(mCol.SD) = 800: .ColAlignment(mCol.SD) = flexAlignCenterCenter
        .TextMatrix(2, mCol.SI上限) = "SI上限": .ColWidth(mCol.SI上限) = 800: .ColAlignment(mCol.SI上限) = flexAlignCenterCenter
        .TextMatrix(2, mCol.SI下限) = "SI下限": .ColWidth(mCol.SI下限) = 800: .ColAlignment(mCol.SI下限) = flexAlignCenterCenter
        .TextMatrix(2, mCol.N) = "n": .ColWidth(mCol.N) = 600: .ColAlignment(mCol.N) = flexAlignCenterCenter
        .TextMatrix(2, mCol.n3s) = "n3s": .ColWidth(mCol.n3s) = 600: .ColAlignment(mCol.n3s) = flexAlignCenterCenter
        .TextMatrix(2, mCol.n2s) = "n2s": .ColWidth(mCol.n2s) = 600: .ColAlignment(mCol.n2s) = flexAlignCenterCenter
        .TextMatrix(2, mCol.结果) = "结果": .ColWidth(mCol.结果) = 500: .ColAlignment(mCol.结果) = flexAlignCenterCenter
        .TextMatrix(2, mCol.检验者) = "检验者": .ColWidth(mCol.检验者) = 1100: .ColAlignment(mCol.检验者) = flexAlignCenterCenter
        
        .ColWidth(0) = 2000
    End With
End Sub

Public Sub ReportPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览
    
    '写数据到临时表
    Dim lngRow As Long, strSQL As String, lngResId As Long, lngCount As Long
    
    With Me.vfgData
        If .Rows <= 6 Then Exit Sub
        strSQL = "ZL_即刻法打印_Clear"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        For lngRow = 3 To Me.vfgData.Rows - 1
            If InStr(.TextMatrix(lngRow, mCol.日期), "-") > 0 And _
               Val(.TextMatrix(lngRow, mCol.次数)) > 0 And _
               Val(.TextMatrix(lngRow, mCol.次数)) < 21 Then
                strSQL = "ZL_即刻法打印_Insert('" & .TextMatrix(lngRow, mCol.日期) & "','" & _
                                                .TextMatrix(lngRow, mCol.次数) & "','" & _
                                                .TextMatrix(lngRow, mCol.测定值) & "','" & _
                                                .TextMatrix(lngRow, mCol.均值) & "','" & _
                                                .TextMatrix(lngRow, mCol.SD) & "','" & _
                                                .TextMatrix(lngRow, mCol.SI上限) & "','" & _
                                                .TextMatrix(lngRow, mCol.SI下限) & "','" & _
                                                .TextMatrix(lngRow, mCol.结果) & "','" & _
                                                .TextMatrix(lngRow, mCol.检验者) & "')"
               zlDatabase.ExecuteProcedure strSQL, Me.Caption
           End If
        Next
    End With
    '-------------------------------------------------
    '调用打印部件处理
    lngResId = 0
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then lngResId = Val(Me.cboQCitem.ItemData(lngCount)): Exit For
    Next
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1209_7", Me, _
                    "质控品ID=" & lngResId, _
                    "项目ID=" & mlngItemID, _
                    "开始日期=" & mstrFromDate, _
                    "结束日期=" & mstrToDate, _
                    IIf(bytMode, 1, 2))

End Sub

Private Sub Form_Load()
    initVfgData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngCount As Long

    With Me.vfgData
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - Me.cboQCitem.Height - Screen.TwipsPerPixelY * 4
    End With
    
    With Me.cboQCitem
        .Left = Me.ScaleLeft + Screen.TwipsPerPixelX * 2
        .Top = Me.ScaleHeight - .Height - Screen.TwipsPerPixelY * 2
    End With
    
    With Me.opt质控品(0)
        .Left = Me.ScaleLeft + Screen.TwipsPerPixelX * 2
        .Top = Me.ScaleHeight - .Height - Screen.TwipsPerPixelY * 2
    End With
    For lngCount = 1 To Me.opt质控品.Count
        With Me.opt质控品(lngCount)
            .Left = Me.opt质控品(lngCount - 1).Left + Me.opt质控品(lngCount - 1).Width + Screen.TwipsPerPixelX * 10
            .Top = Me.opt质控品(lngCount - 1).Top
        End With
    Next
    
End Sub

Private Sub opt质控品_Click(Index As Integer)
    If Me.Visible = False Then Exit Sub
    If Me.opt质控品(Index).Enabled = False Then Exit Sub
    If Me.Tag = "不刷新" Then Exit Sub
    Call RefGrid
End Sub

Private Sub cboQCitem_Click()
    If Me.Visible = False Then Exit Sub
    If Me.Tag = "不刷新" Then Exit Sub
    Call RefGrid
End Sub
