VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCChartGS 
   BorderStyle     =   0  'None
   Caption         =   "Grubbs质控图"
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
      Left            =   3450
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   6810
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
      Height          =   2000
      Left            =   390
      TabIndex        =   0
      Top             =   4740
      Width           =   9180
      _cx             =   16192
      _cy             =   3528
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
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   4020
      Left            =   1005
      TabIndex        =   2
      Top             =   405
      Width           =   7020
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   12382
      _ExtentY        =   7091
      _StockProps     =   0
      ControlProperties=   "frmQCChartGS.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartGS"
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
'            If .AbsolutePosition <> 1 Then Load Me.chtThis(.AbsolutePosition - 1)
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
                .TextMatrix(5, lng次数 + 1) = Format("" & rsTemp!日期, "MM-dd")
                .TextMatrix(6, lng次数 + 1) = Format(Val("" & rsTemp!结果), "0." & String(intFormatNum, "0"))
                '.TextMatrix(.Rows - 1, mCol.检验者) = "" & rsTemp!检验人
                If lng次数 >= 20 Then Exit Do
            ElseIf strLast日期 <> "" Then
                .TextMatrix(6, lng次数 + 1) = Format(Val("" & rsTemp!结果), "0." & String(intFormatNum, "0"))
            End If
            strLast日期 = Format("" & rsTemp!日期, "yyyy-MM-dd") & "," & rsTemp!测试次数
            rsTemp.MoveNext
        Loop
        curTotal = 0
        
        gstrSql = "Select n,n3s,n2s From 质控即刻法 "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
        For iCol = 2 To .Cols - 1
            
            strData = strData & "," & Val(.TextMatrix(6, iCol))
            If Val(.TextMatrix(0, iCol)) > 2 Then
            
                curn3s = Val("" & rsTemp!n3s)
                curn2s = Val("" & rsTemp!n2s)
                
                .TextMatrix(1, iCol) = Format(curn3s, "0.00")
                .TextMatrix(2, iCol) = Format(curn2s, "0.00")
                
                If Not (.TextMatrix(6, iCol) = "" And .TextMatrix(5, iCol) = "") Then
                    cur均值 = s(strData): curSD = stdev(strData)
                    curMax = Max(strData): curMin = Min(strData)
                    If curSD <> 0 Then
                        curSI上 = (curMax - cur均值) / curSD
                        curSI下 = (cur均值 - curMin) / curSD
                    End If
                    .TextMatrix(3, iCol) = IIf(Format(curSI上, "0.00") = "0.00", "", Format(curSI上, "0.00"))
                    .TextMatrix(4, iCol) = IIf(Format(curSI下, "0.00") = "0.00", "", Format(curSI下, "0.00"))
                End If
                rsTemp.MoveNext
            End If
        Next

        .Cell(flexcpAlignment, 0, 2, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 1, .Rows - 1, 1) = flexAlignLeftCenter
    End With
    
    '绘图
    '获得基本的文字信息


    '表头附项
    With chtThis
        .IsBatched = True
        .Reset
        .AllowUserChanges = False
        'Setup the Header
        With .Header
            .Text = "检验科Grubbs质量控制图" & vbCrLf & " " & vbCrLf & " "
            .Adjust = oc2dAdjustCenter
            .Font.Bold = True
            .Font.Size = 16
        End With
        
        .IsBatched = False
        .ChartLabels.RemoveAll
            '行0
        gstrSql = "Select Distinct RPad('单位：' || '" & gstrUnitName & "', 56, ' ') || '日期：' As 行0," & vbNewLine & _
                "                RPad('仪器：' || D.名称, 56, ' ') || '试剂来源：' || M.试剂 As 行1," & vbNewLine & _
                "                RPad('项目：' || I.项目, 56, ' ') || '校准物来源：' || M.校准物 As 行2" & vbNewLine & _
                "From 检验仪器 D, 检验质控品 M, (Select 中文名 || ',' || 英文名 As 项目 From 诊治所见项目 Where ID = [2]) I" & vbNewLine & _
                "Where D.ID = M.仪器id And Instr(',' || [1] || ',', ',' || M.ID || ',') > 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrResList, mlngItemID)
        If rsTemp.RecordCount <= 0 Then Exit Sub
        
        .ChartLabels.Add
        .ChartLabels(1).AttachMethod = oc2dAttachCoord
        .ChartLabels(1).Anchor = oc2dAnchorNorth
        .ChartLabels(1).Text = rsTemp!行0 & Format(mstrFromDate, "yyyy年MM月dd日") & "～" & Format(mstrToDate, "yyyy年MM月dd日")
        .ChartLabels(1).AttachCoord.x = (.ChartLabels(1).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 20
        '行1
        .ChartLabels.Add
        .ChartLabels(2).AttachMethod = oc2dAttachCoord
        .ChartLabels(2).Adjust = oc2dAdjustRight
        .ChartLabels(2).Text = rsTemp!行1
'        .ChartLabels(2).AttachCoord.X = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 180
        .ChartLabels(2).AttachCoord.x = (.ChartLabels(2).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        
        
        '行2
        .ChartLabels.Add
        .ChartLabels(3).AttachMethod = oc2dAttachCoord
        .ChartLabels(3).Adjust = oc2dAdjustRight
        .ChartLabels(3).Text = rsTemp!行2
'        .ChartLabels(2).AttachCoord.X = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 180
        .ChartLabels(3).AttachCoord.x = (.ChartLabels(3).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
        .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(1).Location.Height + 10
        .IsBatched = True
        Set rsTemp = Nothing
        'Make some changes to the X-Axis
        With .ChartArea.Axes("X")
            .AnnotationMethod = oc2dAnnotateValueLabels
            .MajorGrid.Style.Pattern = oc2dLineSolid
            .MajorGrid.Spacing.Value = 1
            .MajorGrid.Style.Width = 1
            .MajorGrid.Style.COLOR = &HC0C0C0
            .Min = 0
            .Max = 20
            With .ValueLabels
                .RemoveAll
                For iCol = 1 To 20
                    .Add iCol, iCol
                Next
            End With
        End With
       
        'Make some changes to the Y-Axis
        With .ChartArea.Axes("Y")
            .AnnotationMethod = oc2dAnnotateValueLabels
            .MajorGrid.Style.Pattern = oc2dLineSolid
            .MajorGrid.Spacing.Value = 5
            .MajorGrid.Style.Width = 1
            .MajorGrid.Style.COLOR = &HC0C0C0 '&HE0E0E0
            .Min = 0
            .Max = 35
            
            With .ValueLabels
                .RemoveAll
                .Add 0, "0.00"
                .Add 5, "0.50"
                .Add 10, "1.00"
                .Add 15, "1.50"
                .Add 20, "2.00"
                .Add 25, "2.50"
                .Add 30, "3.00"
                .Add 35, "3.50"
            End With
        End With
    
    
        With .ChartGroups(1)
            .ChartType = oc2dTypePlot
            With .Data
                .NumSeries = 0
                .NumSeries = 4
                
                .NumPoints(1) = 20
                .NumPoints(2) = 20
                .NumPoints(3) = 20
                .NumPoints(4) = 20
                
                 's3n
                For iCol = 4 To 21
                    .Y(1, iCol - 1) = Val(vfgData.TextMatrix(1, iCol)) * 10
                Next
                's2n
                For iCol = 4 To 21
                    .Y(2, iCol - 1) = Val(vfgData.TextMatrix(2, iCol)) * 10
                Next
                'si上限
                For iCol = 4 To 21
                    .Y(3, iCol - 1) = Val(vfgData.TextMatrix(3, iCol)) * 10
                Next
                'si上限
                For iCol = 4 To 21
                    .Y(4, iCol - 1) = Val(vfgData.TextMatrix(4, iCol)) * 10
                Next
            End With
            
            .Styles(1).Symbol.Shape = oc2dShapeNone: .Styles(1).Line.COLOR = vbRed
            .Styles(2).Symbol.Shape = oc2dShapeNone: .Styles(2).Line.COLOR = vbBlue
            .Styles(3).Symbol.Shape = oc2dShapeNone: .Styles(3).Line.COLOR = vbBlack
            .Styles(4).Symbol.Shape = oc2dShapeNone: .Styles(4).Line.COLOR = vbGreen
        End With
        .IsBatched = False
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
        .BackColor = chtThis.Interior.BackgroundColor
        .Clear
        .Editable = flexEDNone
        .GridLines = flexGridNone
        .Rows = 7: .Cols = 22
        .FixedCols = 2: .FixedRows = 1
        .MergeCells = flexMergeRestrictRows
        
        .BackColorFixed = .BackColor
        .ForeColorFixed = .ForeColor
        .GridColorFixed = .GridColor
        .GridLinesFixed = flexGridNone
        
        '-- 表头
        For iCol = .FixedCols To .Cols - 1
            .TextMatrix(0, iCol) = iCol - 1
            .ColAlignment(iCol) = flexAlignCenterCenter
            .ColWidth(iCol) = 600
        Next
        
        .TextMatrix(1, 1) = "n3s": .ColWidth(1) = 1000: .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(2, 1) = "n2s"
        .TextMatrix(3, 1) = "SI上限"
        .TextMatrix(4, 1) = "SI下限"
        .TextMatrix(5, 1) = "日期"
        .TextMatrix(6, 1) = "测定结果"
        .RowHeight(6) = 0:
        .RowHidden(6) = True
        
        .Select 0, 1, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, .FixedCols
        
        .ColWidth(0) = 1000
    End With
    
End Sub

Public Sub ChartPrint()
    '功能:将数据复制到可打印的对象，将当前图形保存
    
    
    '写数据到临时表
    Dim lngCol As Long, strsql As String, lngResId As Long, lngCount As Long
    Dim strN3S As String, strN2s As String, strSI上 As String, strSI下 As String, str日期 As String
    Dim rsTmp As adodb.Recordset
    With Me.vfgData

        For lngCol = 2 To Me.vfgData.Cols - 1
            If Val(.TextMatrix(0, lngCol)) > 0 And _
               Val(.TextMatrix(0, lngCol)) < 21 Then
                 'n3s ,n2s ,SI上限,SI下限,日期
                strN3S = strN3S & "," & Trim(.TextMatrix(1, lngCol))
                strN2s = strN2s & "," & Trim(.TextMatrix(2, lngCol))
                strSI上 = strSI上 & "," & Trim(.TextMatrix(3, lngCol))
                strSI下 = strSI下 & "," & Trim(.TextMatrix(4, lngCol))
                str日期 = str日期 & "," & Trim(.TextMatrix(5, lngCol))
            End If
        Next
        If strN3S <> "" Then strN3S = Mid(strN3S, 2)
        If strN2s <> "" Then strN2s = Mid(strN2s, 2)
        If strSI上 <> "" Then strSI上 = Mid(strSI上, 2)
        If strSI下 <> "" Then strSI下 = Mid(strSI下, 2)
        If str日期 <> "" Then str日期 = Mid(str日期, 2)
        
        strsql = "ZL_即刻图打印_Insert('" & strN3S & "','" & strN2s & "','" & strSI上 & "','" & strSI下 & "','" & str日期 & "')"
        zlDatabase.ExecuteProcedure strsql, Me.Caption
    End With
    If Dir(App.path & "\QC_Tmp0") <> "" Then Kill App.path & "\QC_Tmp0"
    chtThis.Save App.path & "\QC_Tmp0"

End Sub

Public Function ZLGetGS_QCID() As Long
    '功能       得到当前使用的质控品的ID
    Dim lngCount As Long
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then ZLGetGS_QCID = Val(Me.cboQCitem.ItemData(lngCount)): Exit For
    Next
End Function

Public Sub ChartSaveAs()
    With Me.comDlg
        .CancelError = True
        .DialogTitle = "另存为"
        .filter = "(图形文件)|*.jpg"
        .FileName = Me.chtThis.Header.Text & Format(mstrToDate, "yyyyMMdd") & ".jpg"
        Err = 0: On Error Resume Next
        .ShowSave
        If Err <> 0 Then Exit Sub
        If .FileName = "" Then Exit Sub
        Me.chtThis.SaveImageAsJpeg .FileName, 100, False, False, False
    End With
End Sub

Public Sub ChartCopy()
    Me.chtThis.CopyToClipboard (oc2dFormatBitmap)
End Sub

Private Sub Form_Load()
    initVfgData
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngCount As Long
    With Me.chtThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop
        .Height = Me.ScaleHeight - Me.cboQCitem.Height - Me.vfgData.Height - Screen.TwipsPerPixelY * 4
    End With
    With Me.vfgData
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.chtThis.Top + Me.chtThis.Height
    End With
    
    With Me.opt质控品(0)
        .Left = Me.ScaleLeft + Screen.TwipsPerPixelX * 2
        .Top = Me.ScaleHeight - .Height - Screen.TwipsPerPixelY * 2
    End With
    
    With Me.cboQCitem
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
    Dim intLoop As Integer
    If Me.Visible = False Then Exit Sub

'    If Me.opt质控品(Index).Enabled = False Then Exit Sub
    If Me.Tag = "不刷新" Then Exit Sub
    Call RefGrid
End Sub

