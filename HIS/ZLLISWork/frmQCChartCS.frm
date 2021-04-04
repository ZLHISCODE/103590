VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCChartCS 
   BorderStyle     =   0  'None
   Caption         =   "累积和图"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cboQCitem 
      Height          =   300
      Left            =   2730
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4950
      Width           =   2595
   End
   Begin VB.CheckBox chkUnion 
      Caption         =   "联合Levey_Jennings图"
      Height          =   180
      Left            =   5475
      TabIndex        =   2
      Top             =   5040
      Width           =   2115
   End
   Begin VB.OptionButton opt质控品 
      Caption         =   "473843A低值质控品"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   255
      TabIndex        =   0
      Top             =   4980
      Width           =   2475
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   4020
      Left            =   180
      TabIndex        =   1
      Top             =   165
      Width           =   7020
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   12382
      _ExtentY        =   7091
      _StockProps     =   0
      ControlProperties=   "frmQCChartCS.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartCS"
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
Dim mstr质控品期限 As String

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
    gstrSql = "Select ID, 批号 || '-' || 名称 As 质控品 From 检验质控品 Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
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
    Call Form_Resize
    Me.Tag = ""
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
    Dim lngAllTimes As Long, lngSelTimes As Long
    Dim dblAvg As Double, dblSD As Double, dblMax As Double, dblK As Double, dblH As Double '相关的控制值参数
    Dim aryX() As Variant, aryY() As Variant, arySum() As Variant
    
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
    With Me.chtThis
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
                .NumSeries = 7
                .NumPoints(1) = 0
            End With
        .Styles(1).Symbol.Shape = oc2dShapeNone: .Styles(1).Line.COLOR = RGB(0, 0, 0)
        .Styles(2).Symbol.Shape = oc2dShapeNone: .Styles(2).Line.COLOR = RGB(0, 128, 0)
        .Styles(3).Symbol.Shape = oc2dShapeNone: .Styles(3).Line.COLOR = RGB(0, 128, 0)
        .Styles(4).Symbol.Shape = oc2dShapeNone: .Styles(4).Line.COLOR = RGB(255, 0, 0)
        .Styles(5).Symbol.Shape = oc2dShapeNone: .Styles(5).Line.COLOR = RGB(255, 0, 0)
        .Styles(6).Symbol.Shape = oc2dShapeDot: .Styles(6).Line.COLOR = RGB(0, 0, 160): .Styles(6).Symbol.COLOR = RGB(0, 0, 160)
        .Styles(7).Symbol.Shape = oc2dShapeDiamond: .Styles(7).Line.COLOR = RGB(0, 128, 255): .Styles(7).Symbol.COLOR = RGB(0, 128, 255)
        End With
        .IsBatched = False
    End With
    Call chkUnion_Click
    
    '获得基本的文字信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select RPad('单位：' || '" & gstrUnitName & "', 46, ' ') || '日期：' As 行0," & vbNewLine & _
            "       RPad('仪器：' || D.名称, 46, ' ') ||" & vbNewLine & _
            "        RPad('均值：' || Replace(Replace(' 0' || X.均值, ' 0.', '0.'), ' 0', ''), 26, ' ') || '检测方法：' || L.方法 As 行1," & vbNewLine & _
            "       RPad('项目：' || I.中文名 || ',' || I.英文名, 46, ' ') ||" & vbNewLine & _
            "        RPad('SD值：' || Replace(Replace(' 0' || X.Sd, ' 0.', '0.'), ' 0', ''), 26, ' ') || '试剂来源：' || M.试剂 As 行2," & vbNewLine & _
            "       RPad('质控品：' || M.批号 || ',' || M.名称, 46, ' ') || RPad('规则：' || R.名称, 26, ' ') || '校准物来源：' ||" & vbNewLine & _
            "        M.校准物 As 行3, X.均值, X.Sd, I.单位, R.K, R.H" & vbNewLine & _
            "From 检验仪器 D, 检验质控品 M, 检验质控均值 X, 诊治所见项目 I, 检验仪器规则 A, 检验质控规则 R,检验质控品项目 L " & vbNewLine & _
            "Where D.ID = M.仪器id And M.ID = X.质控品id And X.项目id = I.ID And D.ID = A.仪器id And A.规则id = R.ID And R.种类 = 3 And" & vbNewLine & _
            "      M.ID = [1] And X.项目id = [2] And M.ID = L.质控品ID and L.项目ID = [2] And " & vbNewLine & _
            "      Instr(';' || [3] || ';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstr质控品期限)
    If rsTemp.RecordCount <= 0 Then Me.chtThis.Header.Text = "该质控品信息不全面！": Exit Sub
    strLable = rsTemp!行0 & Format(mstrFromDate, "yyyy年MM月dd日") & "～" & Format(mstrToDate, "yyyy年MM月dd日")
    strLable = strLable & vbCrLf & rsTemp!行1 & vbCrLf & rsTemp!行2 & vbCrLf & rsTemp!行3
    dblAvg = Val("" & rsTemp!均值): dblSD = Val("" & rsTemp!SD): strUnit = "" & rsTemp!单位
    dblK = rsTemp!k: dblH = rsTemp!H
    If dblAvg = 0 Or dblSD = 0 Then
         Me.chtThis.Header.Text = "尚未定值，无法绘制" & Me.Caption & "！": Exit Sub
    End If
    
    '标题、XY轴设置
    With Me.chtThis.Header
        .Text = strLable
        .Adjust = oc2dAdjustLeft
    End With
    With Me.chtThis.ChartArea.Axes("Y")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValues
        .Title.Text = "累积和" & IIf(strUnit = "", "", "(" & strUnit & ")")
    End With
    With Me.chtThis.ChartArea.Axes("Y2")
        .AnnotationMethod = oc2dAnnotateValueLabels
        .Title.Text = "控制线"
        .Multiplier = 1
        With .ValueLabels
            .RemoveAll
            .Add 0, "0"
            .Add 0 + Val(dblK) * Val(dblSD), "Ku= " & Format(Val(dblK) * Val(dblSD), "0.00")
            .Add 0 - Val(dblK) * Val(dblSD), "Kl=" & Format(-Val(dblK) * Val(dblSD), "0.00")
            .Add 0 + Val(dblH) * Val(dblSD), "Hu= " & Format(Val(dblH) * Val(dblSD), "0.00")
            .Add 0 - Val(dblH) * Val(dblSD), "Hl=" & Format(-Val(dblH) * Val(dblSD), "0.00")
        End With
    End With
    With Me.chtThis.ChartArea.Axes("X")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels   '横坐标显示值提示
        .AnnotationPlacement = oc2dAnnotateMinimum
        .Title.Text = "测试次数"
    End With
    
    '数据组织
    gstrSql = "Select 检验时间, 次数, Nvl(结果, 0) As 结果" & vbNewLine & _
            "From (Select Q.检验时间, To_Char(Q.测试次数, '000') || '-' || Decode(Nvl(T.标记, 0), 2, Q.测试次数, 999) As 次数," & vbNewLine & _
            "              zl_Lis_ToNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果" & vbNewLine & _
            "       From 检验质控记录 Q, 检验普通结果 R,检验质控报告 T,检验质控品 M,检验质控均值 X " & vbNewLine & _
            "       Where Q.标本id = R.检验标本id And Nvl(R.弃用结果,0)=0 And /*Nvl(R.是否检验, 0) = 1 And*/ Q.质控品id = [1] And R.检验项目id + 0 = [2] And" & vbNewLine & _
            "             R.ID=T.结果ID(+) And Q.检验时间 + 0 <= To_Date([3], 'yyyy-MM-dd')" & vbNewLine & _
            "       And Instr(';' || [4] || ';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "       And Q.质控品id=M.ID And M.id=X.质控品ID and  X.项目id = [2] And Q.检验时间 between X.开始日期 and Nvl(X.结束日期, M.结束日期) " & vbNewLine & _
            "      )Order By 检验时间, 次数"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstrToDate, mstr质控品期限)
    '首先计算累积和
    With rsTemp
        ReDim arySum(.RecordCount, 1)
        arySum(0, 0) = 0: arySum(0, 1) = 0: lngSelTimes = 0: lngAllTimes = .RecordCount
        Do While Not .EOF
            If Val("" & !结果) = 0 Then
                arySum(.AbsolutePosition, 0) = 0
            ElseIf Abs(Val("" & !结果) - Val(dblAvg)) <= Val(dblK) * Val(dblSD) Then
                arySum(.AbsolutePosition, 0) = 0
            ElseIf Sgn(Val("" & !结果) - Val(dblAvg)) <> Sgn(arySum(.AbsolutePosition - 1, 0)) Then
                arySum(.AbsolutePosition, 0) = Sgn(Val("" & !结果) - Val(dblAvg)) * (Abs(Val("" & !结果) - Val(dblAvg)) - Val(dblK) * Val(dblSD))
            Else
                arySum(.AbsolutePosition, 0) = arySum(.AbsolutePosition - 1, 0) + Sgn(Val("" & !结果) - Val(dblAvg)) * (Abs(Val("" & !结果) - Val(dblAvg)) - Val(dblK) * Val(dblSD))
            End If
            If Format(!检验时间, "yyyy-MM-dd") >= mstrFromDate Then lngSelTimes = lngSelTimes + 1
            arySum(.AbsolutePosition, 1) = Val("" & !结果) - Val(dblAvg)
            .MoveNext
        Loop
    End With
    
    '将区间范围的数据给予绘图数组
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    ReDim aryX(lngSelTimes)
    ReDim aryY(lngSelTimes, 6)
    aryY(0, 0) = 0
    aryY(0, 1) = 0 + Val(dblK) * Val(dblSD)
    aryY(0, 2) = 0 - Val(dblK) * Val(dblSD)
    aryY(0, 3) = 0 + Val(dblH) * Val(dblSD)
    aryY(0, 4) = 0 - Val(dblH) * Val(dblSD)
    aryY(0, 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
    aryY(0, 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
    dblMax = Val(dblH) * Val(dblSD)
    For lngCount = 1 To lngSelTimes
        Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, lngCount
        aryX(lngCount) = lngCount
        aryY(lngCount, 0) = 0
        aryY(lngCount, 1) = 0 + Val(dblK) * Val(dblSD)
        aryY(lngCount, 2) = 0 - Val(dblK) * Val(dblSD)
        aryY(lngCount, 3) = 0 + Val(dblH) * Val(dblSD)
        aryY(lngCount, 4) = 0 - Val(dblH) * Val(dblSD)
        aryY(lngCount, 5) = arySum(lngAllTimes - lngSelTimes + lngCount, 0)
        aryY(lngCount, 6) = arySum(lngAllTimes - lngSelTimes + lngCount, 1)
        If dblMax < Abs(arySum(lngAllTimes - lngSelTimes + lngCount, 0)) Then dblMax = Abs(arySum(lngAllTimes - lngSelTimes + lngCount, 0))
        If dblMax < Abs(arySum(lngAllTimes - lngSelTimes + lngCount, 1)) Then dblMax = Abs(arySum(lngAllTimes - lngSelTimes + lngCount, 1))
    Next

    '变更刷新内部数据
    With Me.chtThis
        .IsBatched = True
        With .ChartGroups(1).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(aryY)
        End With
        With .ChartArea.Axes("Y")
            .Min = 0 - Val(dblMax) - 0.01
            .Max = 0 + Val(dblMax) + 0.01
        End With
        With .ChartArea.Axes("X")
            .Min = 0: .Max = aryX(UBound(aryX))
        End With
        .IsBatched = False
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'--------------------------------------------
'以下为控件事件处理
'--------------------------------------------
Private Sub chkUnion_Click()
    With Me.chtThis.ChartGroups(1)
        If Me.chkUnion.Value = vbChecked Then
            .Styles(7).Line.Pattern = oc2dLineDotted: .Styles(7).Symbol.Shape = oc2dShapeDiamond
        Else
            .Styles(7).Line.Pattern = oc2dLineNone: .Styles(7).Symbol.Shape = oc2dShapeNone
        End If
    End With
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
    
    
    With Me.chkUnion
        .Left = Me.ScaleWidth - Me.chkUnion.Width
        .Top = Me.opt质控品(0).Top
    End With
End Sub

Public Function ZLGetCS_QCID() As Long
    '功能       得到当前使用的质控品的ID
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then ZLGetCS_QCID = Val(Me.cboQCitem.ItemData(lngCount)): Exit For
    Next
End Function

