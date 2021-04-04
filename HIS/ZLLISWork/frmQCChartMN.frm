VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCChartMN 
   BorderStyle     =   0  'None
   Caption         =   "Monica图"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cboQCitem 
      Height          =   300
      Left            =   2970
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4950
      Width           =   2595
   End
   Begin VB.OptionButton opt质控品 
      Caption         =   "473843A低值质控品"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   4905
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2475
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   4410
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   7005
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   12356
      _ExtentY        =   7779
      _StockProps     =   0
      ControlProperties=   "frmQCChartMN.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartMN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrResList As String
Private mlngItemID As Long
Private mstrFromDate As String
Private mstrToDate As String

Dim lngCount As Long
Private mstr质控品期限 As String
'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Sub ChartPrint()
    With Me.chtThis
'        .PrintChart oc2dFormatBitmap, oc2dScaleToFit, 0, 0, 0, 0
        .Save App.path & "\QC_Tmp0"
    End With
End Sub

Public Sub ChartSaveAs()
    Dim strBatCode As String
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
        Me.chtThis.SaveImageAsJpeg .FileName, 100, False, False, False
    End With
End Sub

Public Sub ChartCopy()
    Me.chtThis.CopyToClipboard (oc2dFormatBitmap)
End Sub

Public Function zlRefresh(strResList As String, lngItemID As Long, strFromDate As String, strToDate As String, str质控品期限 As String) As Boolean
    '功能：刷新本窗体的数据显示内容
    '参数： strResList  当前选择的质控品id串，以逗号分隔
    '       lngItemId   当前项目id
    '       strFromDate 开始日期
    '       strToDate   结束日期
    Dim rsTemp As New adodb.Recordset
    Dim intCounts As Integer
    Dim lngResId As Long
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr质控品期限 = str质控品期限
    lngResId = 0
    Me.Tag = "不刷新"
    intCounts = Me.cboQCitem.ListCount
    For lngCount = intCounts - 1 To 1 Step -1
        If Me.cboQCitem.ListIndex = lngCount Then lngResId = Val(Me.cboQCitem.ItemData(lngCount))
'        Unload Me.opt质控品(Me.opt质控品.UBound)
    Next
    cboQCitem.Clear
    
    Me.opt质控品(0).Enabled = False
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID, 批号 || '-' || 名称 As 质控品 From 检验质控品 Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
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
    Call Form_Resize
    Call RefChart
    
    zlRefresh = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RefChart()
    '功能：刷新图形显示
    Dim rsTemp As New adodb.Recordset
    Dim lngResId As Long, strLable As String, strUnit As String
    Dim dblAvg As Double, dblSD As Double, dblMax As Double
    Dim aryX() As Variant, aryY() As Variant, ary2() As Variant, lngHoles As Long
    
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
    
    '设置图形的基本形态
    Me.chtThis.Reset
    Me.chtThis.AllowUserChanges = False
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    With Me.chtThis.ChartArea
        .Axes("Y").Min = 0: .Axes("Y").Max = 1
        .Axes("X").Min = 0: .Axes("X").Max = 1
    End With
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 8
            .NumPoints(1) = 0
        End With
        .Styles(1).Symbol.Shape = oc2dShapeNone: .Styles(1).Line.COLOR = RGB(0, 0, 0)
        .Styles(2).Symbol.Shape = oc2dShapeNone: .Styles(2).Line.COLOR = RGB(200, 200, 0)
        .Styles(3).Symbol.Shape = oc2dShapeNone: .Styles(3).Line.COLOR = RGB(200, 200, 0)
        .Styles(4).Symbol.Shape = oc2dShapeNone: .Styles(4).Line.COLOR = RGB(255, 0, 0)
        .Styles(5).Symbol.Shape = oc2dShapeNone: .Styles(5).Line.COLOR = RGB(255, 0, 0)
        .Styles(6).Symbol.Shape = oc2dShapeOpenDiamond: .Styles(6).Line.Pattern = oc2dLineNone: .Styles(6).Symbol.COLOR = RGB(0, 64, 64)
        .Styles(7).Symbol.Shape = oc2dShapeDiamond: .Styles(7).Line.Pattern = oc2dLineNone: .Styles(7).Symbol.COLOR = RGB(0, 64, 64)
        .Styles(8).Symbol.Shape = oc2dShapeNone: .Styles(8).Line.COLOR = RGB(0, 0, 160): .Styles(8).Symbol.COLOR = RGB(0, 0, 160)
    End With
    With Me.chtThis.ChartGroups(2)
        .ChartType = oc2dTypeHiLo
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 2
            .NumPoints(1) = 0
        End With
        .Styles(1).Line.COLOR = RGB(0, 64, 64)
    End With
    
    '获得基本的文字信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select RPad('单位：' || '" & gstrUnitName & "', 46, ' ') || '日期范围：' As 行0," & vbNewLine & _
            "       RPad('仪器：' || D.名称, 46, ' ') ||" & vbNewLine & _
            "        RPad('参考靶值：' || Replace(Replace(' 0' || X.靶值, ' 0.', '0.'), ' 0', ''), 26, ' ') || '检测方法：' || L.方法 As 行1," & vbNewLine & _
            "       RPad('项目：' || I.中文名 || ',' || I.英文名, 46, ' ') ||" & vbNewLine & _
            "        RPad('参考SD值：' || Replace(Replace(' 0' || X.Sd, ' 0.', '0.'), ' 0', ''), 26, ' ') || '试剂来源：' || M.试剂 As 行2," & vbNewLine & _
            "       RPad('质控品：' || M.批号 || ',' || M.名称, 46, ' ') ||" & vbNewLine & _
            "        RPad('参考CCV%：' || Replace(Replace(' 0' || X.Cv, ' 0.', '0.'), ' 0', ''), 26, ' ') || '校准物来源：' || M.校准物 As 行3," & vbNewLine & _
            "       X.靶值, X.Sd, I.单位" & vbNewLine & _
            "From 检验仪器 D, 检验质控品 M, 检验质控品项目 X, 诊治所见项目 I,检验质控品项目 L " & vbNewLine & _
            "Where D.ID = M.仪器id And M.ID = X.质控品id And X.项目id = I.ID And M.ID = [1] And X.项目id = [2] " & vbNewLine & _
            " And M.ID = L.质控品ID And L.项目ID = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID)
    If rsTemp.RecordCount <= 0 Then Me.chtThis.Header.Text = "该质控品信息不全面！": Exit Sub
    strLable = rsTemp!行0 & Format(mstrFromDate, "yyyy年MM月dd日") & "～" & Format(mstrToDate, "yyyy年MM月dd日")
    strLable = strLable & vbCrLf & rsTemp!行1 & vbCrLf & rsTemp!行2 & vbCrLf & rsTemp!行3
    dblAvg = Val("" & rsTemp!靶值): dblSD = Val("" & rsTemp!SD): strUnit = "" & rsTemp!单位
    If dblAvg = 0 Or dblSD = 0 Then
        
        'MsgBox "没有设置参考靶值和CCV，无法绘制" & Me.Caption & "！", vbInformation, gstrSysName: Exit Sub
        Me.chtThis.Header.Text = "没有设置参考靶值和CCV，无法绘制" & Me.Caption & "！": Exit Sub
    End If
    
    '标题和XY轴设置
    With Me.chtThis.Header
        .Text = strLable
        .Adjust = oc2dAdjustLeft
    End With
    With Me.chtThis.ChartArea.Axes("Y")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValues
        .Title.Text = "测定值" & IIf(strUnit = "", "", "(" & strUnit & ")")
    End With
    With Me.chtThis.ChartArea.Axes("Y2")
        .AnnotationMethod = oc2dAnnotateValueLabels   '纵坐标2显示值提示
        .Title.Text = "控制线"
        .Multiplier = 1
        With .ValueLabels
            .RemoveAll
            .Add Val(dblAvg), "T         =" & Format(Val(dblAvg), "0.00")
            .Add Val(dblAvg) + 0.8 * Val(dblSD), "T+0.8CCV*T=" & Format(Val(dblAvg) + 0.8 * Val(dblSD), "0.00")
            .Add Val(dblAvg) - 0.8 * Val(dblSD), "T-0.8CCV*T=" & Format(Val(dblAvg) - 0.8 * Val(dblSD), "0.00")
            .Add Val(dblAvg) + 1.5 * Val(dblSD), "T+1.5CCV*T=" & Format(Val(dblAvg) + 1.5 * Val(dblSD), "0.00")
            .Add Val(dblAvg) - 1.5 * Val(dblSD), "T-1.5CCV*T=" & Format(Val(dblAvg) - 1.5 * Val(dblSD), "0.00")
        End With
    End With
    With Me.chtThis.ChartArea.Axes("X")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels   '横坐标显示值提示
        .Title.Text = "日期"
        .AnnotationRotationAngle = 30
    End With
    
    '数据组织
    gstrSql = "Select 检验时间, Max(Decode(次数, '1-', 结果, 0)) As 结果1, Max(Decode(次数, '1-', 0, 结果)) As 结果2" & vbNewLine & _
            "From (Select Q.检验时间, Q.测试次数 || '-' || Decode(Nvl(T.标记,0),2,2,Null) As 次数," & vbNewLine & _
            "              zl_Lis_toNumber(Q.质控品id,R.检验项目id, R.检验结果,R.id) As 结果" & vbNewLine & _
            "       From 检验质控记录 Q, 检验普通结果 R,检验质控报告 T, 检验质控品 M, 检验质控均值 X " & vbNewLine & _
            "       Where Q.标本id = R.检验标本id And /*Nvl(R.是否检验, 0) = 1 And*/ Q.质控品id + 0 = [1] And R.检验项目id + 0 = [2] And" & vbNewLine & _
            "             Nvl(R.弃用结果,0)=0 And R.ID=T.结果ID(+) And (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
            "             And (Q.检验时间 Between X.开始日期 And NVL(X.结束日期,M.结束日期)) And " & vbNewLine & _
            "              Q.质控品id=M.id And M.id=X.质控品id  And  X.项目ID = [2] And " & vbNewLine & _
            "             Instr(';'||[5]||';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "       )" & vbNewLine & _
            "Group By 检验时间" & vbNewLine & _
            "Order By 检验时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstrFromDate, mstrToDate, mstr质控品期限)
    
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    lngHoles = 0
    With rsTemp
        ReDim aryX(.RecordCount)
        ReDim aryY(.RecordCount, 7)
        ReDim ary2(.RecordCount, 1)
        aryY(0, 0) = Val(dblAvg)
        aryY(0, 1) = Val(dblAvg) + 0.8 * Val(dblSD)
        aryY(0, 2) = Val(dblAvg) - 0.8 * Val(dblSD)
        aryY(0, 3) = Val(dblAvg) + 1.5 * Val(dblSD)
        aryY(0, 4) = Val(dblAvg) - 1.5 * Val(dblSD)
        aryY(0, 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 7) = Me.chtThis.ChartGroups(1).Data.HoleValue
        ary2(0, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
        ary2(0, 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
        dblMax = 3 * Val(dblSD)
        Do While Not .EOF
            Me.chtThis.ChartArea.Axes("X").ValueLabels.Add .AbsolutePosition, Format(!检验时间, "M月d日")
            aryX(.AbsolutePosition) = .AbsolutePosition
            aryY(.AbsolutePosition, 0) = Val(dblAvg)
            aryY(.AbsolutePosition, 1) = Val(dblAvg) + 0.8 * Val(dblSD)
            aryY(.AbsolutePosition, 2) = Val(dblAvg) - 0.8 * Val(dblSD)
            aryY(.AbsolutePosition, 3) = Val(dblAvg) + 1.5 * Val(dblSD)
            aryY(.AbsolutePosition, 4) = Val(dblAvg) - 1.5 * Val(dblSD)
            If Val("" & !结果1) = 0 Then
                aryY(.AbsolutePosition, 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Else
                aryY(.AbsolutePosition, 5) = Val("" & !结果1)
                If dblMax < Abs(Val(aryY(.AbsolutePosition, 5)) - Val(dblAvg)) Then dblMax = Abs(Val(aryY(.AbsolutePosition, 5)) - Val(dblAvg))
            End If
            If Val("" & !结果2) = 0 Then
                aryY(.AbsolutePosition, 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Else
                aryY(.AbsolutePosition, 6) = Val("" & !结果2)
                If dblMax < Abs(Val(aryY(.AbsolutePosition, 6)) - Val(dblAvg)) Then dblMax = Abs(Val(aryY(.AbsolutePosition, 6)) - Val(dblAvg))
            End If
            If Val("" & !结果1) = 0 Or Val("" & !结果2) = 0 Then
                aryY(.AbsolutePosition, 7) = Me.chtThis.ChartGroups(1).Data.HoleValue: lngHoles = lngHoles + 1
            Else
                aryY(.AbsolutePosition, 7) = (aryY(.AbsolutePosition, 5) + aryY(.AbsolutePosition, 6)) / 2
            End If
            ary2(.AbsolutePosition, 0) = aryY(.AbsolutePosition, 5)
            ary2(.AbsolutePosition, 1) = aryY(.AbsolutePosition, 6)
            .MoveNext
        Loop
    End With
    If lngHoles > 3 Then
        Me.chtThis.Footer.Text = "注：由于该质控品有" & lngHoles & "天没有同时进行两次测试，影响了该控制图的表现。"
        Me.chtThis.Footer.Adjust = oc2dAdjustLeft
    Else
        Me.chtThis.Footer.Text = ""
    End If

    '变更刷新内部数据
    With Me.chtThis
        .IsBatched = True
        With .ChartGroups(1).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(aryY)
        End With
        With .ChartArea.Axes("Y")
            .Min = Val(dblAvg) - Val(dblMax)
            .Max = Val(dblAvg) + Val(dblMax)
        End With
        With .ChartArea.Axes("X")
            .Min = 0
            .Max = aryX(UBound(aryX))
        End With
        With .ChartGroups(2).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(ary2)
        End With
        .IsBatched = False
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ChtThis_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim px As Long
    Dim py As Long
    Dim Series As Long
    Dim Point As Long
    Dim Distance As Long
    Dim Region As Long
    
    On Error Resume Next
    
    px = x / Screen.TwipsPerPixelX
    py = Y / Screen.TwipsPerPixelY
    
    If (Button = 0) Then
        With chtThis
            Region = .ChartGroups(1).CoordToDataIndex(px, py, oc2dFocusXY, Series, Point, Distance)
            If (Series > 0 And Point > 0) And (Distance <= 5) Then
                If (Region = oc2dRegionInChartArea) Then
                    .ToolTipText = .ChartGroups(1).Data(Series, Point)
                End If
            Else
                .ToolTipText = ""
                .Footer.Text = ""
            End If
            .Refresh
        End With
    End If
End Sub

'--------------------------------------------
'以下为控件事件处理
'--------------------------------------------
Private Sub opt质控品_Click(Index As Integer)
    If Me.Visible = False Then Exit Sub
    If Me.opt质控品(Index).Enabled = False Then Exit Sub
    If Me.Tag = "不刷新" Then Exit Sub
    Call RefChart
End Sub

Private Sub cboQCitem_Click()
    If Me.Visible = False Then Exit Sub
    If Me.Tag = "不刷新" Then Exit Sub
    Call RefChart
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.chtThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight - Me.cboQCitem.Height - Screen.TwipsPerPixelY * 4
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

Public Function ZLGetMN_QCID() As Long
    '功能       得到当前使用的质控品的ID
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then ZLGetMN_QCID = Val(Me.cboQCitem.ItemData(lngCount)): Exit For
    Next
End Function

