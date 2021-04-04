VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCChartLJ 
   BorderStyle     =   0  'None
   Caption         =   "Levey_Jennings图"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cboQCitem 
      Height          =   300
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4980
      Width           =   2595
   End
   Begin VB.ComboBox cbo显示 
      Height          =   300
      Left            =   5790
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4950
      Width           =   1785
   End
   Begin VB.OptionButton opt质控品 
      Caption         =   "473843A低值质控品"
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   4920
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2475
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   4410
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   7365
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   12991
      _ExtentY        =   7779
      _StockProps     =   0
      ControlProperties=   "frmQCChartLJ.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartLJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrResList As String
Private mlngItemID As Long
Private mstrFromDate As String
Private mstrToDate As String
Private mstr质控品期限 As String

Dim lngCount As Long
Private mArr() As String
Private mbln对数质控图 As Boolean
Private mint补位显示 As Integer
Private mLastXY As String           '避免鼠标显示检验结果时重复刷新

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function ChartPrint() As Integer
    '返回有几个图片
    Dim intLoop As Integer
    Dim intIndex As Integer
    Dim intCount As Integer
    For intLoop = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = intLoop Then
            intIndex = intLoop
        End If
    Next
    For intLoop = 0 To chtThis.Count - 1
        With Me.chtThis(intLoop)
            If .Visible = True Then
    '        .PrintChart oc2dFormatBitmap, oc2dScaleToFit, 0, 0, 0, 0
                .Save App.path & "\QC_Tmp" & intCount
                intCount = intCount + 1
            End If
        End With
    Next
    ChartPrint = intCount
End Function


Public Sub ChartSaveAs()
    Dim strBatCode As String
    Dim intLoop As Integer
    Dim intIndex As Integer
    For intLoop = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = intLoop Then
            intIndex = intLoop
        End If
    Next
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then strBatCode = Me.cboQCitem.Text: Exit For
    Next
    With Me.comDlg
        .CancelError = True
        .DialogTitle = "另存为"
        .filter = "(图形文件)|*.jpg"
        .FileName = strBatCode & Me.Caption & Format(mstrToDate, "yyyyMMdd") & ".jpg"
        Err = 0: On Error Resume Next
        .ShowSave
        If Err <> 0 Then Exit Sub
        If .FileName = "" Then Exit Sub
        Me.chtThis(intIndex).SaveImageAsJpeg .FileName, 100, False, False, False
    End With
End Sub

Public Sub ChartCopy()
    Dim intLoop As Integer
    Dim intIndex As Integer
    For intLoop = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = intLoop Then
            intIndex = intLoop
        End If
    Next
    Me.chtThis(intIndex).CopyToClipboard (oc2dFormatBitmap)
End Sub

Public Function zlRefresh(strResList As String, lngItemID As Long, strFromDate As String, strToDate As String, str质控品期限 As String, Optional ByVal int补位显示 = 1) As Boolean
    '功能：刷新本窗体的数据显示内容
    '参数： strResList  当前选择的质控品id串，以逗号分隔
    '       lngItemId   当前项目id
    '       strFromDate 开始日期
    '       strToDate   结束日期
    '       strDateSpace 以;分隔的质控品的期间
    '
    Dim rsTemp As New adodb.Recordset
    Dim intCounts As Integer
    Dim lngResId As Long
    Dim int图像Index As Integer
    Dim int当前质控Index As Integer
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr质控品期限 = str质控品期限
    mint补位显示 = int补位显示

    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then
            int当前质控Index = lngCount
        End If
    Next
    
    int图像Index = Me.cbo显示.ListIndex
    
    
    lngResId = 0
   
    intCounts = Me.cboQCitem.ListCount
    For lngCount = intCounts - 1 To 1 Step -1
'        If Me.opt质控品(lngCount).Value Then lngResId = Val(Me.opt质控品(Me.opt质控品.UBound).Tag)
'        Unload Me.opt质控品(Me.opt质控品.UBound)
        Unload chtThis(Me.chtThis.UBound)
    Next
    cboQCitem.Clear
    Me.opt质控品(0).Enabled = False
    Err = 0: On Error GoTo ErrHand
    mbln对数质控图 = False
    
    gstrSql = "Select A.ID, A.批号 || '-' || A.名称 As 质控品, B.对数质控图 From 检验质控品 A,检验仪器 B Where A.仪器ID=B.ID(+) And Instr(',' || [1] || ',', ',' || A.ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList)
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.cboQCitem.ListCount Then cboQCitem.AddItem "" & !质控品
            If .AbsolutePosition <> 1 Then Load Me.chtThis(.AbsolutePosition - 1)
            cboQCitem.ItemData(cboQCitem.NewIndex) = !ID
'            If .AbsolutePosition > Me.opt质控品.Count Then Load Me.opt质控品(.AbsolutePosition - 1): Load Me.chtThis(.AbsolutePosition - 1)
'            Me.opt质控品(.AbsolutePosition - 1).Caption = "" & !质控品
'            Me.opt质控品(.AbsolutePosition - 1).Tag = !ID
'            Me.opt质控品(.AbsolutePosition - 1).Width = Me.TextWidth(Me.opt质控品(.AbsolutePosition - 1).Caption) + 360
'            Me.opt质控品(.AbsolutePosition - 1).Value = (lngResId = !ID)
'            Me.opt质控品(.AbsolutePosition - 1).Visible = True
'            Me.opt质控品(.AbsolutePosition - 1).Enabled = True
            mbln对数质控图 = Val("" & !对数质控图) = 1
            .MoveNext
        Loop
    End With
    If rsTemp.RecordCount > 0 Then Me.cboQCitem.ListIndex = 0
    Call Form_Resize
    
    
    Me.cbo显示.Clear
    Me.cbo显示.Tag = "不刷新"
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If lngCount < 3 Then
            Me.cbo显示.AddItem lngCount + 1 & "幅图"
        Else
            Exit For
        End If
        If lngCount = int当前质控Index Then
            Me.cboQCitem.ListIndex = lngCount
        End If
        
    Next
    If Me.cbo显示.ListCount > 0 Then
        Me.cbo显示.ListIndex = IIf(int图像Index = -1, 0, int图像Index)
    End If
    Me.cbo显示.Tag = ""
    
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        Call RefChart(CInt(lngCount))
    Next
    
    Call Form_Resize
    DoEvents
    If Me.cboQCitem.ListIndex >= 0 Then
        Call chtThis_Resize(Me.cboQCitem.ListIndex, chtThis(Me.cboQCitem.ListIndex).Width, chtThis(Me.cboQCitem.ListIndex).Height)
    End If
    zlRefresh = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RefChart(intIndex As Integer)
    '功能：刷新图形显示
    Dim rsTemp As New adodb.Recordset
    Dim lngResId As Long, strLable As String, strUnit As String
    Dim dblAvg As Double, dblSD As Double, dblMax As Double
    Dim aryX() As Variant, aryY() As Variant
    Dim strCalc As String           '计算结果
    Dim strStartDate As String, strEndDate As String
    Dim str超限数据 As String '保存超过上下限的数据，用于显示
    Dim intLoop As Integer, dateLoop As Date '用于补足30天的数据
    Dim lngX As Long '记录X的序号
    Dim bln合并行 As Boolean, str小数 As String, lngTmp As Long, strTmp As String
    Dim strAllCount As String, strCurCount '所有次数,当前次数
    lngResId = 0
'    For lngCount = 0 To Me.opt质控品.UBound
'        If Me.opt质控品(lngCount).Value Then lngResId = Val(Me.opt质控品(lngCount).Tag): Exit For
'    Next
    lngResId = Val(Me.cboQCitem.ItemData(intIndex))
    If lngResId = 0 Then
        Me.opt质控品(0).Enabled = False
        Me.opt质控品(0).Value = True
        lngResId = Val(Me.opt质控品(0).Tag)
        Me.opt质控品(0).Enabled = True
    End If
    
    '设置图形的基本形态
    With Me.chtThis(intIndex)
        .IsBatched = True
        .Reset
        .AllowUserChanges = False
        .ChartArea.Axes("Y").Min = 0: .ChartArea.Axes("Y").Max = 1
        .ChartArea.Axes("X").Min = 0: .ChartArea.Axes("X").Max = 1
        With .ChartGroups(1)
            .ChartType = oc2dTypePlot
            With .Data
                .NumSeries = 0
                .LayOut = oc2dDataArray
                .NumSeries = 15
                .NumPoints(1) = 0
            End With
            .Styles(1).Symbol.Shape = oc2dShapeNone: .Styles(1).Line.COLOR = RGB(0, 0, 0)
            .Styles(2).Symbol.Shape = oc2dShapeNone: .Styles(2).Line.COLOR = RGB(0, 128, 0)
            .Styles(3).Symbol.Shape = oc2dShapeNone: .Styles(3).Line.COLOR = RGB(0, 128, 0)
            .Styles(4).Symbol.Shape = oc2dShapeNone: .Styles(4).Line.COLOR = RGB(200, 200, 0)
            .Styles(5).Symbol.Shape = oc2dShapeNone: .Styles(5).Line.COLOR = RGB(200, 200, 0)
            .Styles(6).Symbol.Shape = oc2dShapeNone: .Styles(6).Line.COLOR = RGB(255, 0, 0)
            .Styles(7).Symbol.Shape = oc2dShapeNone: .Styles(7).Line.COLOR = RGB(255, 0, 0)
            .Styles(8).Symbol.Shape = oc2dShapeNone: .Styles(8).Line.COLOR = RGB(0, 0, 0)
            .Styles(9).Symbol.Shape = oc2dShapeNone: .Styles(9).Line.COLOR = RGB(0, 0, 0)
            .Styles(10).Symbol.Shape = oc2dShapeDot: .Styles(10).Line.COLOR = RGB(0, 0, 160): .Styles(10).Symbol.COLOR = RGB(0, 0, 160)
            .Styles(11).Symbol.Shape = oc2dShapeDot: .Styles(11).Line.Pattern = oc2dLineNone: .Styles(11).Symbol.COLOR = RGB(255, 0, 0)
            .Styles(12).Symbol.Shape = oc2dShapeDot: .Styles(12).Line.Pattern = oc2dLineNone: .Styles(12).Symbol.COLOR = RGB(255, 0, 0)
            .Styles(13).Symbol.Shape = oc2dShapeDot: .Styles(13).Line.Pattern = oc2dLineNone: .Styles(13).Symbol.COLOR = RGB(255, 0, 0)
            .Styles(14).Symbol.Shape = oc2dShapeDot: .Styles(14).Line.Pattern = oc2dLineNone: .Styles(14).Symbol.COLOR = RGB(255, 0, 0)
            .Styles(15).Symbol.Shape = oc2dShapeDot: .Styles(15).Line.Pattern = oc2dLineNone: .Styles(15).Symbol.COLOR = RGB(255, 0, 0)
        End With
        .IsBatched = False
    End With
    
    '获得基本的文字信息
    Err = 0: On Error GoTo ErrHand
'    gstrSql = "Select RPad('单位：' || '" & gstrUnitName & "', 46, ' ') || '日期：' As 行0," & vbNewLine & _
            "       RPad('仪器：' || D.名称, 46, ' ') ||" & vbNewLine & _
            "        RPad('均值：' || Replace(Replace(' 0' || X.均值, ' 0.', '0.'), ' 0', ''), 26, ' ') || '检测方法：' || L.方法 As 行1," & vbNewLine & _
            "       RPad('项目：' || I.中文名 || ',' || I.英文名, 46, ' ') ||" & vbNewLine & _
            "        RPad('SD值：' || Replace(Replace(' 0' || X.Sd, ' 0.', '0.'), ' 0', ''), 26, ' ') || '试剂来源：' || D.试剂来源 As 行2," & vbNewLine & _
            "       RPad('质控品：' || M.批号 || ',' || M.名称, 46, ' ') ||" & vbNewLine & _
            "        RPad('CV% ：' || Replace(Replace(' 0' || X.Cv * 100, ' 0.', '0.'), ' 0', ''), 26, ' ') || '校准物来源：' ||" & vbNewLine & _
            "        D.校准物来源 As 行3, X.均值, X.Sd, I.单位" & vbNewLine & _
            "From 检验仪器 D, 检验质控品 M, 检验质控均值 X, 诊治所见项目 I,检验质控品项目 L" & vbNewLine & _
            "Where D.ID = M.仪器id And M.ID = X.质控品id And X.项目id = I.ID And M.ID = [1] And X.项目id = [2] And" & vbNewLine & _
            "      M.id = L.质控品ID and L.项目ID = [2] And " & vbNewLine & _
            "      (To_Date([3], 'yyyy-MM-dd') Between X.开始日期 And Nvl(X.结束日期, M.结束日期)) And" & vbNewLine & _
            "      (To_Date([4], 'yyyy-MM-dd') Between X.开始日期 And Nvl(X.结束日期, M.结束日期))"
                
'    gstrSql = " Select A.开始日期, Nvl(A.结束日期, B.结束日期) As 结束日期" & vbNewLine & _
'                " From 检验质控均值 A, 检验质控品 B" & vbNewLine & _
'                " Where A.质控品id = B.ID And 质控品id = [1] And 项目id = [2] And 期间 = [3] "
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemId, mstrDateSpace)
'
'    If rsTemp.EOF = True Then
'        MsgBox "没有找到对应的期间<" & mstrDateSpace & ">!", vbInformation, gstrSysName: Exit Sub
'    End If
    Dim varTmp As Variant, intCount As Integer
    
    If InStr(mstr质控品期限, ";") > 0 Then
        varTmp = Split(mstr质控品期限, ";")
        For intCount = LBound(varTmp) To UBound(varTmp)
            If lngResId = Val(Split(varTmp(intCount), "=")(0)) Then
                strStartDate = Split(Split(varTmp(intCount), "=")(1), ",")(0)
                strEndDate = Split(Split(varTmp(intCount), "=")(1), ",")(1)
                Exit For
            End If
        Next
    End If
    
        '兼容以前的调用方式
    If strStartDate = "" Then
        strStartDate = mstrFromDate: strEndDate = mstrToDate
    Else
        If CDate(strStartDate) < CDate(mstrFromDate) Then strStartDate = mstrFromDate
        If CDate(strEndDate) > CDate(mstrToDate) Then strEndDate = mstrToDate
    End If
    
'    If CDate(mstrFromDate) < CDate(rsTemp("开始日期")) Then
'        strStartDate = Nvl(rsTemp("开始日期"))
'    End If
'    If CDate(mstrToDate) > CDate(rsTemp("结束日期")) Then
'        strEndDate = Nvl(rsTemp("结束日期"))
'    End If

    str小数 = "0000"
'    gstrSql = "Select Nvl(小数位数,2) As 小数 From 检验仪器项目  A,检验质控品 M Where m.仪器id=A.仪器Id And m.Id=[1]  And a.项目id=[2]"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID)
'    Do Until rsTemp.EOF
'        str小数 = String(Val("" & rsTemp!小数), "0")
'        rsTemp.MoveNext
'    Loop
                
    gstrSql = "Select RPad('单位：' || '" & gstrUnitName & "', 59, ' ') || ' 日期：' As 行0," & vbNewLine & _
                "       RPad('项目：' || I.中文名 || '/' || I.英文名, 30, ' ')|| RPad(' 方法：' || L.方法, 29, ' ')  ||RPad(' 仪器：' || D.名称, 25, ' ')  as 行1 ," & vbNewLine & _
                "        rpad('输入均值：' || Replace(Replace(' 0' || trim(to_char(X.均值,'999990." & str小数 & "')), ' 0.', '0.'), ' 0', '') || '(' || I.单位 || ')' || '   SD: ' ||" & vbNewLine & _
                "        Replace(Replace(' 0' || trim(to_char(X.Sd,'999990." & str小数 & "')), ' 0.', '0.'), ' 0', '') || '(' || I.单位 || ')' || '   CV: ' ||" & vbNewLine & _
                "        Replace(Replace(' 0' || trim(to_char(X.Cv * 100,'999990." & str小数 & "')), ' 0.', '0.'), ' 0', '') || '%',60,' ') || RPad('质控品：' || M.名称, 20, ' ') ||RPad('批号：' || M.批号, 20, ' ')   As 行2," & vbNewLine & _
                "        RPad('试剂：' || M.试剂, 20, ' ') || RPad('校准物：' || M.校准物, 20, ' ') as 行3, X.均值, X.Sd, I.单位" & vbNewLine & _
                "From 检验仪器 D, 检验质控品 M, 检验质控均值 X, 诊治所见项目 I, 检验质控品项目 L" & vbNewLine & _
                "Where D.ID = M.仪器id And M.ID = X.质控品id And X.项目id = I.ID And M.ID = [1] And X.项目id = [2] And M.ID = L.质控品id And" & vbNewLine & _
                "      L.项目id = [2] And " & vbNewLine & _
                "      Instr(';' || [3] || ';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0 "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstr质控品期限)
    If rsTemp.RecordCount <= 0 Then Me.chtThis(intIndex).Header.Text = "该质控品信息不全面！": Exit Sub
    'If rsTemp.RecordCount <= 0 Then MsgBox "该质控品信息不全面！", vbInformation, gstrSysName: Exit Sub
   
    '组织表头
    
'    strLable = rsTemp!行0 & Format(strStartDate, "yyyy年MM月dd日") & "～" & Format(strEndDate, "yyyy年MM月dd日")
'    strLable = strLable & vbCrLf & " " & rsTemp!行1 & vbCrLf & " " & rsTemp!行2 & vbCrLf & " " & rsTemp!行3
    
    
    
    dblAvg = Val("" & rsTemp!均值): dblSD = Val("" & rsTemp!SD): strUnit = "" & rsTemp!单位
    If dblAvg = 0 Or dblSD = 0 Then
'        MsgBox "尚未定值或SD为0，无法绘制" & Me.Caption & "！", vbInformation, gstrSysName: Exit Sub
        Me.chtThis(intIndex).Header.Text = "尚未定值或SD为0，无法绘制" & Me.Caption & "！": Exit Sub
    End If
    If Me.cbo显示.ListIndex > 0 Then
        '标题、XY轴设置
        With Me.chtThis(intIndex).Header
            .Text = cboQCitem.Text
            .Adjust = oc2dAdjustCenter
            .Font.Bold = True
            .Font.Size = 8
        End With
    Else
        '标题、XY轴设置
        With Me.chtThis(intIndex).Header
            .Text = "检验科Levey-Jennings" & IIf(mbln对数质控图, "对数", "") & "质量控制图" & vbCrLf & " " & vbCrLf & " "
            .Adjust = oc2dAdjustCenter
            .Font.Bold = True
            .Font.Size = 16
        End With
        With Me.chtThis(intIndex)
            strUnit = Nvl(rsTemp("单位"))
            .ChartLabels.RemoveAll
            '行0
            .ChartLabels.Add
            .ChartLabels(1).AttachMethod = oc2dAttachCoord
            .ChartLabels(1).Anchor = oc2dAnchorNorth
            .ChartLabels(1).Text = rsTemp!行0 & Format(strStartDate, "yyyy年MM月dd日") & "～" & Format(strEndDate, "yyyy年MM月dd日")
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
            
            strCalc = ""
            strLable = rsTemp!行3
            '处理计算均值，SD
            gstrSql = "Select Round(Avg(结果), 2) As 均值, Round(Stddev(结果), 2) As Sd, Count(*) As 次数" & vbNewLine & _
                "From (Select Trunc(Q.检验时间) As 日期," & vbNewLine & _
                "              Avg(zl_Lis_toNumber(Q.质控品ID,R.检验项目ID,R.检验结果,R.ID)) As 结果" & vbNewLine & _
                "       From 检验质控记录 Q, 检验普通结果 R,检验质控报告 T" & vbNewLine & _
                "       Where Q.标本id = R.检验标本id And Q.质控品id = [1] And R.检验项目id + 0 = [2] And" & vbNewLine & _
                "             Nvl(R.弃用结果,0)=0 And R.ID=T.结果ID(+) And Q.检验时间 Between   [3] and [4]  And Nvl(T.标记, 0) <> 2" & vbNewLine & _
                "       Group By Trunc(Q.检验时间))"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, CDate(strStartDate), CDate(strEndDate))
            
            If rsTemp.EOF = False Then
                If rsTemp("均值") = 0 Then
                    strCalc = "计算均值：" & Format(rsTemp("均值"), "0." & str小数) & "(" & strUnit & _
                                ")   SD: " & Format(rsTemp("SD"), "0." & str小数) & _
                                "(" & strUnit & ")   CV: " & Format(0, "0." & str小数) & "%"
                Else
                    strCalc = "计算均值：" & Format(rsTemp("均值"), "0." & str小数) & "(" & strUnit & _
                                ")   SD: " & Format(rsTemp("SD"), "0." & str小数) & _
                                "(" & strUnit & ")   CV: " & Format(rsTemp("SD") / rsTemp("均值") * 100, "0." & str小数) & "%"
                End If
            End If
            If LenB(StrConv(strCalc, vbFromUnicode)) < 60 Then
                strCalc = strCalc & Space(60 - LenB(StrConv(strCalc, vbFromUnicode))) & strLable
            Else
                strCalc = strCalc & strLable
            End If
            '行3
            .ChartLabels.Add
            .ChartLabels(4).AttachMethod = oc2dAttachCoord
            .ChartLabels(4).Adjust = oc2dAdjustRight
            .ChartLabels(4).Text = strCalc
    '        .ChartLabels(3).AttachCoord.X = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 180
            .ChartLabels(4).AttachCoord.x = (.ChartLabels(4).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
            .ChartLabels(4).AttachCoord.Y = .ChartLabels(3).Location.Top + .ChartLabels(2).Location.Height + 10
            
            
                
        End With
    End If
    
    With Me.chtThis(intIndex).ChartArea.Axes("Y")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValues
        .Title.Text = "测定值" & IIf(strUnit = "", "", "(" & strUnit & ")")
    End With
    With Me.chtThis(intIndex).ChartArea.Axes("Y2")
        .AnnotationMethod = oc2dAnnotateValueLabels   '纵坐标2显示值提示
        .Title.Text = "控制线"
        .Multiplier = 1
        With .ValueLabels
            .RemoveAll
'            .Add Val(dblAvg), "CL=    " & Format(Val(dblAvg), "0.00")
'            .Add Val(dblAvg) + 1 * Val(dblSD), "CL+1SD=" & Format(Val(dblAvg) + 1 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) - 1 * Val(dblSD), "CL-1SD=" & Format(Val(dblAvg) - 1 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) + 2 * Val(dblSD), "CL+2SD=" & Format(Val(dblAvg) + 2 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) - 2 * Val(dblSD), "CL-2SD=" & Format(Val(dblAvg) - 2 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) + 3 * Val(dblSD), "CL+3SD=" & Format(Val(dblAvg) + 3 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) - 3 * Val(dblSD), "CL-3SD=" & Format(Val(dblAvg) - 3 * Val(dblSD), "0.00")
            .Add Val(dblAvg), Format(Val(dblAvg), "##0.00##") & " CL"
            .Add Val(dblAvg) + 1 * Val(dblSD), Format(Round(Val(dblAvg) + 1 * Val(dblSD), 4), "##0.00##") & " CL+1SD"
            .Add Val(dblAvg) - 1 * Val(dblSD), Format(Round(Val(dblAvg) - 1 * Val(dblSD), 4), "##0.00##") & " CL-1SD"
            .Add Val(dblAvg) + 2 * Val(dblSD), Format(Round(Val(dblAvg) + 2 * Val(dblSD), 4), "##0.00##") & " CL+2SD"
            .Add Val(dblAvg) - 2 * Val(dblSD), Format(Round(Val(dblAvg) - 2 * Val(dblSD), 4), "##0.00##") & " CL-2SD"
            .Add Val(dblAvg) + 3 * Val(dblSD), Format(Round(Val(dblAvg) + 3 * Val(dblSD), 4), "##0.00##") & " CL+3SD"
            .Add Val(dblAvg) - 3 * Val(dblSD), Format(Round(Val(dblAvg) - 3 * Val(dblSD), 4), "##0.00##") & " CL-3SD"
        End With
    End With
    With Me.chtThis(intIndex).ChartArea.Axes("X")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels   '横坐标显示值提示
        .Title.Text = "日期"
    End With
    
    '数据组织
'    gstrSql = "Select 检验时间, 测试次数, Max(Decode(标记, 2, 0, 结果)) As 在控, Max(Decode(标记, 2, 结果, 0)) As 失控" & vbNewLine & _
'            "From (Select Q.检验时间, Q.测试次数, T.标记," & vbNewLine & _
'            "              Decode(I.值序列, Null, Zl_To_Number(R.检验结果)," & vbNewLine & _
'            "                      Length(Substr(I.值序列, 1, Instr(I.值序列, ';' || RTrim(R.检验结果) || ';'))) -" & vbNewLine & _
'            "                       Nvl(Length(Replace(Substr(I.值序列, 1, Instr(I.值序列, ';' || RTrim(R.检验结果) || ';')), ';')), 0)) As 结果" & vbNewLine & _
'            "       From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T," & vbNewLine & _
'            "            (Select Decode(结果类型, 3, Decode(RTrim(取值序列), '', '', ';' || RTrim(取值序列) || ';'), '') As 值序列" & vbNewLine & _
'            "              From 检验项目" & vbNewLine & _
'            "              Where 诊治项目id = [2]) I" & vbNewLine & _
'            "       Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And /*Nvl(R.是否检验, 0) = 1 And*/ Q.质控品id + 0 = [1] And" & vbNewLine & _
'            "             R.检验项目id + 0 = [2] And" & vbNewLine & _
'            "             (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')))" & vbNewLine & _
'            "Group By 检验时间, 测试次数" & vbNewLine & _
'            "Order By 检验时间, 测试次数"
                
'    gstrSql = "Select 检验时间,  Max(Decode(标记, 2, 0, 结果)) As 在控, Max(Decode(标记, 2, 结果, 0)) As 失控" & vbNewLine & _

    
    Set rsTemp = GetQCChartData(lngResId, mlngItemID, strStartDate, strEndDate)
    
    
    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.RemoveAll
    str超限数据 = ""
    With rsTemp
        If .RecordCount < 30 And mint补位显示 = 1 Then
            intLoop = .RecordCount
            ReDim Preserve aryX(31)
            ReDim Preserve aryY(31, 14)
        Else
            intLoop = 0
            ReDim aryX(.RecordCount)
            ReDim aryY(.RecordCount, 14)
        End If

        For lngTmp = LBound(aryY) To UBound(aryY)
            aryX(lngTmp) = lngTmp
            aryY(lngTmp, 0) = Val(dblAvg)
            aryY(lngTmp, 1) = Val(dblAvg) + 1 * Val(dblSD)
            aryY(lngTmp, 2) = Val(dblAvg) - 1 * Val(dblSD)
            aryY(lngTmp, 3) = Val(dblAvg) + 2 * Val(dblSD)
            aryY(lngTmp, 4) = Val(dblAvg) - 2 * Val(dblSD)
            aryY(lngTmp, 5) = Val(dblAvg) + 3 * Val(dblSD)
            aryY(lngTmp, 6) = Val(dblAvg) - 3 * Val(dblSD)
            aryY(lngTmp, 7) = Val(dblAvg) + 4 * Val(dblSD)
            aryY(lngTmp, 8) = Val(dblAvg) - 4 * Val(dblSD)
            aryY(lngTmp, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 10) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 11) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 12) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 13) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 14) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
        Next
        

        dblMax = 4 * Val(dblSD)
        
        Do While Not .EOF
'            Me.ChtThis.ChartArea.Axes("X").ValueLabels.Add .AbsolutePosition, .AbsolutePosition
            bln合并行 = False
            If lngX > 0 Then
                If Not (aryY(lngX, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue And dateLoop = Format(Nvl(!检验时间), "yyyy-MM-dd")) Then
                    lngX = lngX + 1
                    If Format(Nvl(!检验时间), "dd") <> "01" Then
                        Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!检验时间), "dd")
                    Else
                        Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!检验时间), "mm" & "月")
                    End If
                Else
                    bln合并行 = True
                    intLoop = intLoop - 1
                End If
            Else
                lngX = lngX + 1
                If Format(Nvl(!检验时间), "dd") <> "01" Then
                    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!检验时间), "dd")
                Else
                    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!检验时间), "mm" & "月")
                End If
            End If

            dateLoop = Format(Nvl(!检验时间), "yyyy-MM-dd")
            strAllCount = Trim$("" & !测试次数)
            aryX(lngX) = lngX
            aryY(lngX, 0) = Val(dblAvg)
            aryY(lngX, 1) = Val(dblAvg) + 1 * Val(dblSD)
            aryY(lngX, 2) = Val(dblAvg) - 1 * Val(dblSD)
            aryY(lngX, 3) = Val(dblAvg) + 2 * Val(dblSD)
            aryY(lngX, 4) = Val(dblAvg) - 2 * Val(dblSD)
            aryY(lngX, 5) = Val(dblAvg) + 3 * Val(dblSD)
            aryY(lngX, 6) = Val(dblAvg) - 3 * Val(dblSD)
            aryY(lngX, 7) = Val(dblAvg) + 4 * Val(dblSD)
            aryY(lngX, 8) = Val(dblAvg) - 4 * Val(dblSD)
            
            
            If "" & !在控 <> "" Then
'                aryY(lngX, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue 'Val(dblAvg)
'            Else
                If Abs(Val("" & !在控) - Val(dblAvg)) > dblMax Then
                    aryY(lngX, 9) = IIf((Val("" & !在控) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !在控)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                    aryY(lngX, 9) = Val("" & !在控) + 0.03 * dblSD
                Else
                    aryY(lngX, 9) = Val("" & !在控)
                End If
                strTmp = Val("" & !在控)
                If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                
                
                If InStr(strAllCount, ",") > 0 Then
                    strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                    strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                Else
                    strCurCount = strAllCount
                End If
                strTmp = "检验结果:" & strTmp & " 日期:" & Format(Nvl(!检验时间), "yyyy-MM-dd") & " " & Trim("" & !时间) & " 第" & strCurCount & "次"
                str超限数据 = str超限数据 & "|" & lngX & ",9," & strTmp
                
'                If dblMax < Abs(Val("" & !在控) - Val(dblAvg)) Then dblMax = Abs(Val("" & !在控) - Val(dblAvg))
            End If
            
            If Not bln合并行 Then
                If "" & !失控1 <> "" Then
'                    aryY(lngX, 10) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !失控1) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 10) = IIf((Val("" & !失控1) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !失控1)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 10) = Val("" & !失控1) + 0.03 * dblSD
                    Else
                        aryY(lngX, 10) = Val("" & !失控1)
                    End If
                    strTmp = Val("" & !失控1)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ","))
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "检验结果:" & strTmp & " 日期:" & Format(Nvl(!检验时间), "yyyy-MM-dd") & " " & Trim("" & !时间) & " 第" & strCurCount & "次"
                    str超限数据 = str超限数据 & "|" & lngX & ",10," & strTmp
    '                If dblMax < Abs(Val("" & !失控) - Val(dblAvg)) Then dblMax = Abs(Val("" & !失控) - Val(dblAvg))
                End If
                
                If "" & !失控2 <> "" Then
'                    aryY(lngX, 11) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !失控2) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 11) = IIf((Val("" & !失控2) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !失控2)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 11) = Val("" & !失控2) + 0.03 * dblSD
                    Else
                        aryY(lngX, 11) = Val("" & !失控2)
                    End If
                    strTmp = Val("" & !失控2)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "检验结果:" & strTmp & " 日期:" & Format(Nvl(!检验时间), "yyyy-MM-dd") & " " & Trim("" & !时间) & " 第" & strCurCount & "次"
                    str超限数据 = str超限数据 & "|" & lngX & ",11," & strTmp
                End If
                
                If "" & !失控3 <> "" Then
'                    aryY(lngX, 12) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !失控3) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 12) = IIf((Val("" & !失控3) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !失控3)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 12) = Val("" & !失控3) + 0.03 * dblSD
                    Else
                        aryY(lngX, 12) = Val("" & !失控3)
                    End If
                    strTmp = Val("" & !失控3)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "检验结果:" & strTmp & " 日期:" & Format(Nvl(!检验时间), "yyyy-MM-dd") & " " & Trim("" & !时间) & " 第" & strCurCount & "次"
                    str超限数据 = str超限数据 & "|" & lngX & ",12," & strTmp
    '                If dblMax < Abs(Val("" & !失控) - Val(dblAvg)) Then dblMax = Abs(Val("" & !失控) - Val(dblAvg))
                End If
                
                If "" & !失控4 <> "" Then
'                    aryY(lngX, 13) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !失控4) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 13) = IIf((Val("" & !失控4) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !失控4)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 13) = Val("" & !失控4) + 0.03 * dblSD
                    Else
                        aryY(lngX, 13) = Val("" & !失控4)
                    End If
                    strTmp = Val("" & !失控4)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "检验结果:" & strTmp & " 日期:" & Format(Nvl(!检验时间), "yyyy-MM-dd") & " " & Trim("" & !时间) & " 第" & strCurCount & "次"
                    str超限数据 = str超限数据 & "|" & lngX & ",13," & strTmp
    '                If dblMax < Abs(Val("" & !失控) - Val(dblAvg)) Then dblMax = Abs(Val("" & !失控) - Val(dblAvg))
                End If
                
                If "" & !失控5 <> "" Then
'                    aryY(lngX, 14) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !失控5) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 14) = IIf((Val("" & !失控5) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !失控4)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 14) = Val("" & !失控5) + 0.03 * dblSD
                    Else
                        aryY(lngX, 14) = Val("" & !失控5)
                    End If
                    strTmp = Val("" & !失控5)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "检验结果:" & strTmp & " 日期:" & Format(Nvl(!检验时间), "yyyy-MM-dd") & " " & Trim("" & !时间) & " 第" & strCurCount & "次"
                    str超限数据 = str超限数据 & "|" & lngX & ",14," & strTmp
                    
    '                If dblMax < Abs(Val("" & !失控) - Val(dblAvg)) Then dblMax = Abs(Val("" & !失控) - Val(dblAvg))
                End If
            End If
            .MoveNext
        Loop
        
        Do While lngX < UBound(aryX)
            '中间可能有合并的点，这里补齐数据
            lngX = lngX + 1
            aryX(lngX) = lngX
            aryY(lngX, 0) = Val(dblAvg)
            aryY(lngX, 1) = Val(dblAvg) + 1 * Val(dblSD)
            aryY(lngX, 2) = Val(dblAvg) - 1 * Val(dblSD)
            aryY(lngX, 3) = Val(dblAvg) + 2 * Val(dblSD)
            aryY(lngX, 4) = Val(dblAvg) - 2 * Val(dblSD)
            aryY(lngX, 5) = Val(dblAvg) + 3 * Val(dblSD)
            aryY(lngX, 6) = Val(dblAvg) - 3 * Val(dblSD)
            aryY(lngX, 7) = Val(dblAvg) + 4 * Val(dblSD)
            aryY(lngX, 8) = Val(dblAvg) - 4 * Val(dblSD)
            
            aryY(lngX, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 10) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 11) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 12) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 13) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 14) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
        Loop
    End With
    
    '如果不足30天的数据,补齐30天的数据
    If intLoop > 0 And mint补位显示 = 1 Then
        For intLoop = intLoop + 1 To 31
            
            dateLoop = DateAdd("d", 1, dateLoop)
            If dateLoop <= CDate(strEndDate) Then
                If Format(Nvl(dateLoop), "dd") <> "01" Then
                    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "dd")
                Else
                    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "mm" & "月")
                End If
            End If
            aryX(intLoop) = intLoop
            aryY(intLoop, 0) = Val(dblAvg)
            aryY(intLoop, 1) = Val(dblAvg) + 1 * Val(dblSD)
            aryY(intLoop, 2) = Val(dblAvg) - 1 * Val(dblSD)
            aryY(intLoop, 3) = Val(dblAvg) + 2 * Val(dblSD)
            aryY(intLoop, 4) = Val(dblAvg) - 2 * Val(dblSD)
            aryY(intLoop, 5) = Val(dblAvg) + 3 * Val(dblSD)
            aryY(intLoop, 6) = Val(dblAvg) - 3 * Val(dblSD)
            aryY(intLoop, 7) = Val(dblAvg) + 4 * Val(dblSD)
            aryY(intLoop, 8) = Val(dblAvg) - 4 * Val(dblSD)
            
            aryY(intLoop, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 10) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 11) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 12) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 13) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 14) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
        Next
    End If

    If str超限数据 <> "" Then
        str超限数据 = Mid(str超限数据, 2)
        ReDim mArr(intIndex + 1)
        mArr(intIndex) = str超限数据
    End If
    '变更刷新内部数据
    With Me.chtThis(intIndex)
        .IsBatched = True
        With .ChartGroups(1).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(aryY)
        End With
        With .ChartArea.Axes("Y")
            .Min = Val(dblAvg) - Val(dblMax)
            .Origin = .Min
            .Max = Val(dblAvg) + Val(dblMax)
            .MajorGrid.Spacing.IsDefault = False
            .AnnotationMethod = oc2dAnnotateValues
        End With
        With .ChartArea.Axes("X")
            .Min = 0: .Max = aryX(UBound(aryX))
        End With
        .IsBatched = False
    End With
    Call chtThis_Resize(intIndex, chtThis(intIndex).Width, chtThis(intIndex).Height)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetQCChartData(ByVal lngResId As Long, ByVal lngItemID As Long, ByVal strDateS As String, ByVal strDateE As String) As adodb.Recordset
    '取绘图用的数据集
    Dim rsTemp As adodb.Recordset, rsQcData As adodb.Recordset
    Dim strLastDate As String, i As Integer
    On Error GoTo errH
    Set rsQcData = New adodb.Recordset
    rsQcData.Fields.Append "检验时间", adVarChar, 50
    rsQcData.Fields.Append "时间", adVarChar, 8
    rsQcData.Fields.Append "测试次数", adVarChar, 10
    rsQcData.Fields.Append "在控", adVarChar, 30
    rsQcData.Fields.Append "失控1", adVarChar, 30
    rsQcData.Fields.Append "失控2", adVarChar, 30
    rsQcData.Fields.Append "失控3", adVarChar, 30
    rsQcData.Fields.Append "失控4", adVarChar, 30
    rsQcData.Fields.Append "失控5", adVarChar, 30


    rsQcData.CursorLocation = adUseClient
    rsQcData.LockType = adLockOptimistic
    rsQcData.CursorType = adOpenStatic
    rsQcData.Open

                
    gstrSql = "Select Q.检验时间,Q.时间, Q.测试次数, Nvl(T.标记,0) as 标记," & vbNewLine & _
                "                     Zl_Lis_Tonumber(Q.质控品ID, R.检验项目id, R.检验结果,R.ID) As 结果" & vbNewLine & _
                "              From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T" & vbNewLine & _
                "              Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And Q.质控品id = [1] And" & vbNewLine & _
                "                    Nvl(R.弃用结果,0)=0 And Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd') And " & vbNewLine & _
                "                    R.检验项目id + 0 = [2] order by  Q.检验时间, Nvl(T.标记,0), Q.测试次数 "
                
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, lngItemID, strDateS, strDateE)
    
    Do Until rsTemp.EOF
        
        If Val("" & rsTemp.Fields("标记").Value) = 2 Then
            '失控点
            If strLastDate <> Format(rsTemp.Fields("检验时间").Value, "yyyy-MM-dd") Then
                rsQcData.AddNew
                rsQcData("检验时间") = Format(rsTemp.Fields("检验时间").Value, "yyyy-MM-dd")
                rsQcData("时间") = Trim("" & rsTemp.Fields("时间").Value)
                rsQcData("测试次数") = rsTemp.Fields("测试次数").Value
                rsQcData("在控") = ""
                rsQcData("失控1") = Trim("" & rsTemp.Fields("结果").Value)
                rsQcData("失控2") = ""
                rsQcData("失控3") = ""
                rsQcData("失控4") = ""
                rsQcData("失控5") = ""
            Else
               For i = 1 To 5
                    If Trim("" & rsQcData.Fields(3 + i).Value) = "" Then
                         rsQcData.Fields(2).Value = rsQcData.Fields(2).Value & "," & Trim("" & rsTemp.Fields("测试次数").Value)
                         rsQcData.Fields(3 + i).Value = Trim("" & rsTemp.Fields("结果").Value)
                         Exit For
                    End If
               Next
            End If
        Else
            '在控与警告
            rsQcData.AddNew
            rsQcData("检验时间") = Format(rsTemp.Fields("检验时间").Value, "yyyy-MM-dd")
            rsQcData("时间") = Trim("" & rsTemp.Fields("时间").Value)
            rsQcData("测试次数") = rsTemp.Fields("测试次数").Value
            rsQcData("在控") = Trim("" & rsTemp.Fields("结果").Value)
            rsQcData("失控1") = ""
            rsQcData("失控2") = ""
            rsQcData("失控3") = ""
            rsQcData("失控4") = ""
            rsQcData("失控5") = ""
        End If
        strLastDate = Format(rsTemp.Fields("检验时间").Value, "yyyy-MM-dd")
        
        rsTemp.MoveNext
    Loop
    If rsQcData.RecordCount > 0 Then rsQcData.MoveFirst
    
    Set GetQCChartData = rsQcData
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo显示_Click()
    Dim intLoop As Integer
    ReDim mArr(Me.cboQCitem.ListCount)
    If Me.cbo显示.Tag = "不刷新" Then Exit Sub
    For intLoop = 0 To Me.cbo显示.ListCount - 1
        Call RefChart(intLoop)
    Next
    Call Form_Resize
End Sub

Private Sub ChtThis_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim px As Long
    Dim py As Long
    Dim Series As Long
    Dim Point As Long
    Dim Distance As Long
    Dim Region As Long
    Dim i As Integer, strTmp As String
    Dim varTmp As Variant
    
    On Error Resume Next
    
    px = x / Screen.TwipsPerPixelX
    py = Y / Screen.TwipsPerPixelY
    If mLastXY = px & "," & py Then Exit Sub
    mLastXY = px & "," & py
    
    If (Button = 0) Then
        With chtThis(Index)
            Region = .ChartGroups(1).CoordToDataIndex(px, py, oc2dFocusXY, Series, Point, Distance)
            If (Series > 9 And Point > 0) And (Distance <= 5) Then
                If (Region = oc2dRegionInChartArea) Then
                    .ToolTipText = .ChartGroups(1).Data(Series, Point)
                    
                        If mArr(Index) <> "" Then
                            varTmp = Split(mArr(Index), "|")
                            For i = LBound(varTmp) To UBound(varTmp)
                                strTmp = varTmp(i)
                                If strTmp <> "" Then
                                    If Split(strTmp, ",")(0) = Point - 1 And Split(strTmp, ",")(1) = Series - 1 Then
                                        .ToolTipText = Split(strTmp, ",")(2)
                                    End If
                                End If
                            Next
                        End If
                    
                    If Left(.ToolTipText, 1) = "." Then .ToolTipText = "0" & .ToolTipText
                End If
            Else
                .ToolTipText = ""
                .Footer.Text = ""
            End If
            .Refresh
        End With
    End If
End Sub

Private Sub chtThis_Resize(Index As Integer, ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    With Me.chtThis(Index)
        '行1
        .ChartLabels(1).AttachCoord.x = .Header.Location.Left + (.ChartLabels(1).Location.Width / 2) - 80
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 30
        '行2
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 80
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '行3
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 80
        .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
        '行3
        .ChartLabels(4).AttachCoord.x = .Header.Location.Left + (.ChartLabels(4).Location.Width / 2) - 80
        .ChartLabels(4).AttachCoord.Y = .ChartLabels(3).Location.Top + .ChartLabels(3).Location.Height + 10
    End With
End Sub

Private Sub Form_Load()
        
'    ReDim mArr(ChtThis.Count)
End Sub

'--------------------------------------------
'以下为控件事件处理
'--------------------------------------------
Private Sub opt质控品_Click(Index As Integer)
    Dim intLoop As Integer
    If Me.Visible = False Then Exit Sub
    If Me.opt质控品(Index).Enabled = False Then Exit Sub
    If Me.cbo显示.Tag = "不刷新" Then Exit Sub
    
    Call Form_Resize
    For intLoop = 0 To Me.chtThis.Count - 1
'        If intLoop = Index Then
'            Me.ChtThis(intLoop).Visible = True
'        Else
'            Me.ChtThis(intLoop).Visible = False
'        End If
        Call RefChart(intLoop)
    Next
    
End Sub

Private Sub cboQCitem_Click()
    Dim intLoop As Integer
    If Me.Visible = False Then Exit Sub
'    If Me.opt质控品(Index).Enabled = False Then Exit Sub
    If Me.cbo显示.Tag = "不刷新" Then Exit Sub
    
    Call Form_Resize
    DoEvents

'        If intLoop = Index Then
'            Me.ChtThis(intLoop).Visible = True
'        Else
'            Me.ChtThis(intLoop).Visible = False
'        End If
        Call RefChart(cboQCitem.ListIndex)
'         Me.chtThis(cboQCitem.ListIndex).Visible = False
End Sub

Private Sub Form_Resize()
    Dim intLoop As Integer
    Dim intIndex As Integer
    Err = 0: On Error Resume Next
    
    For intLoop = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = intLoop Then
            intIndex = intLoop
        End If
        Me.chtThis(intLoop).Visible = False
    Next
    Select Case Me.cbo显示.ListIndex + 1
        Case 1
            With Me.chtThis(intIndex)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.ScaleTop: .Height = Me.ScaleHeight - Me.opt质控品(0).Height - Screen.TwipsPerPixelY * 4
            End With
        Case 2
            With Me.chtThis(intIndex)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.ScaleTop
                .Height = (Me.ScaleHeight - Me.opt质控品(0).Height - Screen.TwipsPerPixelY * 4) / 2
            End With
            With Me.chtThis(intIndex + 1)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.chtThis(intIndex).Top + Me.chtThis(intIndex).Height
                .Height = Me.chtThis(intIndex).Height
            End With
        Case 3
            With Me.chtThis(intIndex)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.ScaleTop
                .Height = (Me.ScaleHeight - Me.opt质控品(0).Height - Screen.TwipsPerPixelY * 4) / 3
            End With
            With Me.chtThis(intIndex + 1)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.chtThis(intIndex).Top + Me.chtThis(intIndex).Height
                .Height = Me.chtThis(intIndex).Height
            End With
            With Me.chtThis(intIndex + 2)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.chtThis(intIndex + 1).Top + Me.chtThis(intIndex + 1).Height
                .Height = Me.chtThis(intIndex).Height
            End With
        Case 4
        Case 5
        Case 6
    End Select
'    With Me.ChtThis(0)
'        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
'        .Top = Me.ScaleTop: .Height = Me.ScaleHeight - Me.opt质控品(0).Height - Screen.TwipsPerPixelY * 4
'    End With
    
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
    
    With Me.cbo显示
        .Top = Me.opt质控品(0).Top
        .Left = Me.ScaleWidth - .Width - Screen.TwipsPerPixelX * 2
    End With
End Sub


Public Function ZLGetLJ_QCID() As Long
    '功能       得到当前使用的质控品的ID
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then ZLGetLJ_QCID = Val(Me.cboQCitem.ItemData(lngCount)): Exit For
    Next
End Function

Public Function ZLGetLJ_QCIDStr() As String
    '功能       得到当前使用的质控品ID串
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
'        If Me.opt质控品(lngCount).Enabled = True Then
            ZLGetLJ_QCIDStr = ZLGetLJ_QCIDStr & "," & Val(Me.cboQCitem.ItemData(lngCount))
'        End If
    Next
    ZLGetLJ_QCIDStr = Mid(ZLGetLJ_QCIDStr, 2)
End Function



