VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmQCChartLJAverage 
   BorderStyle     =   0  'None
   Caption         =   "均值Levey_Jennings图"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   4410
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   7365
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   12991
      _ExtentY        =   7779
      _StockProps     =   0
      ControlProperties=   "frmQCChartLJAverage.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartLJAverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long
Private mstr仪器名称 As String
Private mstrFromDate As String
Private mstrToDate As String
Private mrsData As New adodb.Recordset
Private mrsAverage As New adodb.Recordset

Dim lngCount As Long
Private mArr() As String

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function ChartPrint() As Integer
    '返回有几个图片
    With Me.chtThis
        If .Visible = True Then
            .Save App.path & "\QCLJAverage_Tmp"
        End If
    End With
End Function


Public Sub ChartSaveAs()
    Dim strBatCode As String
    Dim intLoop As Integer
    Dim intIndex As Integer

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
    Call Me.chtThis.CopyToClipboard(oc2dFormatBitmap)
End Sub

Public Function zlRefresh(lngItemID As Long, str仪器名称 As String, strFromDate As String, strToDate As String, _
                        rsDate As adodb.Recordset, rsAverage As adodb.Recordset) As Boolean
    '功能：刷新本窗体的数据显示内容
    '参数：
    '       lngItemId   当前项目id
    '       str仪器名称 当前选中仪器
    '       strFromDate 开始日期
    '       strToDate   结束日期

    Dim rsTemp As New adodb.Recordset
    
    mlngItemID = lngItemID
    mstr仪器名称 = str仪器名称
    mstrFromDate = strFromDate
    mstrToDate = Format(CDate(strToDate), "yyyy-MM-dd 23:59:59")
    Set mrsData = rsDate
    Set mrsAverage = rsAverage
    
    Err = 0: On Error GoTo ErrHand
    Call Form_Resize
    
    Call RefChart
    zlRefresh = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RefChart()
    '功能：刷新图形显示
    Dim rsTemp As New adodb.Recordset
    Dim strLable As String, strUnit As String
    Dim dblAvg As Double, dblSD As Double, dblMax As Double
    Dim aryX() As Variant, aryY() As Variant
    Dim strCalc As String           '计算结果
    Dim strStartDate As String, strEndDate As String
    Dim str超限数据 As String '保存超过上下限的数据，用于显示
    Dim intLoop As Integer, dateLoop As Date '用于补足30天的数据
    Dim lngX As Long '记录X的序号
    Dim bln合并行 As Boolean, str小数 As String
    
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

    Dim varTmp As Variant, intCount As Integer
    
    
    strStartDate = mstrFromDate
    strEndDate = mstrToDate

    str小数 = "00"
    
    gstrSql = "Select RPad('项目：' || 中文名 || '/' || 英文名, 30, ' ') || RPad(' 单位：' || 单位, 29, ' ') || RPad(' 仪器：' || '" & mstr仪器名称 & "', 25, ' ') As 行1,单位 From 诊治所见项目 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
    
    
    '标题、XY轴设置
    With Me.chtThis.Header
        .Text = "检验科均值Levey-Jennings图" & vbCrLf & " " & vbCrLf & " "
        .Adjust = oc2dAdjustCenter
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    With Me.chtThis
        strUnit = Nvl(rsTemp("单位"))
        .ChartLabels.RemoveAll
        '行0
        .ChartLabels.Add
        .ChartLabels(1).AttachMethod = oc2dAttachCoord
        .ChartLabels(1).Anchor = oc2dAnchorNorth
        
        If LenB(StrConv(gstrUnitName, vbFromUnicode)) + LenB(StrConv("单位： ", vbFromUnicode)) < 60 Then
            .ChartLabels(1).Text = "单位：" & gstrUnitName & Space(60 - LenB(StrConv(gstrUnitName, vbFromUnicode)) - LenB(StrConv("单位： ", vbFromUnicode))) & " 日期：" & Format(strStartDate, "yyyy年MM月dd日") & "～" & Format(strEndDate, "yyyy年MM月dd日")
        Else
            .ChartLabels(1).Text = "单位：" & gstrUnitName & " 日期：" & Format(strStartDate, "yyyy年MM月dd日") & "～" & Format(strEndDate, "yyyy年MM月dd日")
        End If
        .ChartLabels(1).AttachCoord.x = (.ChartLabels(1).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 20
        '行1
        .ChartLabels.Add
        .ChartLabels(2).AttachMethod = oc2dAttachCoord
        .ChartLabels(2).Adjust = oc2dAdjustRight
        .ChartLabels(2).Text = rsTemp!行1
        .ChartLabels(2).AttachCoord.x = (.ChartLabels(2).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        
        strCalc = ""
        
'        '处理计算均值，SD
'        gstrSql = "Select Round(Avg(结果), 2) As 均值, Round(Stddev(结果), 2) As Sd, Count(*) As 次数 " & _
'                  "From (Select Trunc(a.检验时间) As 日期,Avg(b.检验结果) As 结果 " & _
'                        "From 检验标本记录 A, 检验普通结果 b Where a.审核人 Is Not Null And a.id=b.检验标本ID " & _
'                        "And b.检验项目id + 0 = [1] And a.检验时间 Between [2] And [3] " & _
'                  "Group By Trunc(a.检验时间))"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, CDate(strStartDate), CDate(strEndDate))
'       mrs
        If Not mrsAverage.EOF Then
            If mrsAverage("均值") = 0 Then
                strCalc = "计算均值：" & Format(mrsAverage("均值"), "0." & str小数) & "(" & strUnit & _
                            ")   SD: " & Format(mrsAverage("SD"), "0." & str小数) & _
                            "(" & strUnit & ")   CV: " & Format(0, "0." & str小数) & "%"
            Else
                strCalc = "计算均值：" & Format(mrsAverage("均值"), "0." & str小数) & "(" & strUnit & _
                            ")   SD: " & Format(mrsAverage("SD"), "0." & str小数) & _
                            "(" & strUnit & ")   CV: " & Format(mrsAverage("SD") / mrsAverage("均值") * 100, "0." & str小数) & "%"
            End If
        End If
        
        dblAvg = Val("" & mrsAverage!均值): dblSD = Val("" & mrsAverage!SD)
        
        If dblAvg = 0 Or dblSD = 0 Then Exit Sub
        
        If LenB(StrConv(strCalc, vbFromUnicode)) < 60 Then
            strCalc = strCalc & Space(60 - LenB(StrConv(strCalc, vbFromUnicode))) & strLable
        Else
            strCalc = strCalc & strLable
        End If
        '行2
        .ChartLabels.Add
        .ChartLabels(3).AttachMethod = oc2dAttachCoord
        .ChartLabels(3).Adjust = oc2dAdjustRight
        .ChartLabels(3).Text = strCalc
        .ChartLabels(3).AttachCoord.x = (.ChartLabels(3).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
        .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(1).Location.Height + 10

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
            
            .Add Val(dblAvg), Format(Val(dblAvg), "0.00") & " CL"
            .Add Val(dblAvg) + 1 * Val(dblSD), Format(Val(dblAvg) + 1 * Val(dblSD), "0." & str小数) & " CL+1SD"
            .Add Val(dblAvg) - 1 * Val(dblSD), Format(Val(dblAvg) - 1 * Val(dblSD), "0." & str小数) & " CL-1SD"
            .Add Val(dblAvg) + 2 * Val(dblSD), Format(Val(dblAvg) + 2 * Val(dblSD), "0." & str小数) & " CL+2SD"
            .Add Val(dblAvg) - 2 * Val(dblSD), Format(Val(dblAvg) - 2 * Val(dblSD), "0." & str小数) & " CL-2SD"
            .Add Val(dblAvg) + 3 * Val(dblSD), Format(Val(dblAvg) + 3 * Val(dblSD), "0." & str小数) & " CL+3SD"
            .Add Val(dblAvg) - 3 * Val(dblSD), Format(Val(dblAvg) - 3 * Val(dblSD), "0." & str小数) & " CL-3SD"
        End With
    End With
    
    With Me.chtThis.ChartArea.Axes("X")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels   '横坐标显示值提示
        .Title.Text = "日期"
    End With
    
    '数据组织
'    gstrSql = "Select Trunc(a.检验时间) As 检验时间, Avg(b.检验结果) As 结果 " & _
'                "From 检验标本记录 A, 检验普通结果 b Where a.审核人 Is Not Null And a.id=b.检验标本ID " & _
'                "And b.检验项目id + 0 = [1] And a.检验时间 Between   To_Date([2],'YYYY-MM-DD') " & _
'                "And To_Date([3],'YYYY-MM-DD') Group By Trunc(a.检验时间)  "
'
'
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, strStartDate, strEndDate)
    
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    str超限数据 = ""
    
    With mrsData
        If .RecordCount < 30 Then
            intLoop = .RecordCount
            ReDim Preserve aryX(31)
            ReDim Preserve aryY(31, 14)
        Else
            intLoop = 0
            ReDim aryX(.RecordCount)
            ReDim aryY(.RecordCount, 14)
        End If

        aryY(0, 0) = Val(dblAvg)
        aryY(0, 1) = Val(dblAvg) + 1 * Val(dblSD)
        aryY(0, 2) = Val(dblAvg) - 1 * Val(dblSD)
        aryY(0, 3) = Val(dblAvg) + 2 * Val(dblSD)
        aryY(0, 4) = Val(dblAvg) - 2 * Val(dblSD)
        aryY(0, 5) = Val(dblAvg) + 3 * Val(dblSD)
        aryY(0, 6) = Val(dblAvg) - 3 * Val(dblSD)
        aryY(0, 7) = Val(dblAvg) + 4 * Val(dblSD)
        aryY(0, 8) = Val(dblAvg) - 4 * Val(dblSD)
        aryY(0, 9) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 10) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 11) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 12) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 13) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 14) = Me.chtThis.ChartGroups(1).Data.HoleValue
        dblMax = 4 * Val(dblSD)
        
        .MoveFirst
        Do While Not .EOF

            bln合并行 = False
            If lngX > 0 Then
                If Not (aryY(lngX, 9) = Me.chtThis.ChartGroups(1).Data.HoleValue And dateLoop = Format(Nvl(!日期), "yyyy-MM-dd")) Then
                    lngX = lngX + 1
                    If Format(Nvl(!日期), "dd") <> "01" Then
                        Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!日期), "dd")
                    Else
                        Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!日期), "mm" & "月")
                    End If
                Else
                    bln合并行 = True
                    intLoop = intLoop - 1
                End If
            Else
                lngX = lngX + 1
                If Format(Nvl(!日期), "dd") <> "01" Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!日期), "dd")
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!日期), "mm" & "月")
                End If
            End If

            dateLoop = Format(Nvl(!日期), "yyyy-MM-dd")
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
            
            
            If Val("" & !结果) = 0 Then
                aryY(lngX, 9) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Else
                If Abs(Val("" & !结果) - Val(dblAvg)) > dblMax Then
                    aryY(lngX, 9) = IIf((Val("" & !结果) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax, Val(dblAvg) + dblMax)
                    str超限数据 = str超限数据 & "|" & lngX & ",9," & Round(Val("" & !结果), 2)
                Else
                    aryY(lngX, 9) = Round(Val("" & !结果), 2)
                End If
            End If
            
            aryY(lngX, 10) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(lngX, 11) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(lngX, 12) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(lngX, 13) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(lngX, 14) = Me.chtThis.ChartGroups(1).Data.HoleValue
            
            .MoveNext
        Loop
        
    End With
    
    '如果不足30天的数据,补齐30天的数据
    If intLoop <> 0 Then
        For intLoop = intLoop + 1 To 31
            
            dateLoop = DateAdd("d", 1, dateLoop)
            If dateLoop <= CDate(strEndDate) Then
                If Format(Nvl(dateLoop), "dd") <> "01" Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "dd")
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "mm" & "月")
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
            
            aryY(intLoop, 9) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(intLoop, 10) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(intLoop, 11) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(intLoop, 12) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(intLoop, 13) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(intLoop, 14) = Me.chtThis.ChartGroups(1).Data.HoleValue
        Next
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
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'--------------------------------------------
'以下为控件事件处理
'--------------------------------------------

Private Sub ChtThis_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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

    If (Button = 0) Then
        With chtThis
            Region = .ChartGroups(1).CoordToDataIndex(px, py, oc2dFocusXY, Series, Point, Distance)
            If (Series > 0 And Point > 0) And (Distance <= 5) Then
                If (Region = oc2dRegionInChartArea) Then
                    .ToolTipText = .ChartGroups(1).Data(Series, Point)
                    If Series >= 7 And Series <= 9 Then
                        If mArr(0) <> "" Then
                            varTmp = Split(mArr(0), "|")
                            For i = LBound(varTmp) To UBound(varTmp)
                                strTmp = varTmp(i)
                                If strTmp <> "" Then
                                    If Split(strTmp, ",")(0) = Point - 1 Then
                                        .ToolTipText = Split(strTmp, ",")(2)
                                    End If
                                End If
                            Next
                        End If
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

Private Sub chtThis_Resize(ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    With Me.chtThis
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


Private Sub Form_Resize()
    With Me.chtThis
        .Visible = True
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight
    End With
End Sub



