VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmQCChartZS 
   BorderStyle     =   0  'None
   Caption         =   "Z-分数图"
   ClientHeight    =   5352
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7752
   LinkTopic       =   "Form1"
   ScaleHeight     =   5352
   ScaleWidth      =   7752
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chk质控品 
      Caption         =   "473843A低值质控品"
      Enabled         =   0   'False
      Height          =   240
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   5055
      Value           =   1  'Checked
      Width           =   1830
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   3690
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6630
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   11695
      _ExtentY        =   6509
      _StockProps     =   0
      ControlProperties=   "frmQCChartZS.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartZS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrResList As String
Private mlngItemID As Long
Private mstrFromDate As String
Private mstrToDate As String
Private mdblAVG As Double                           '均值
Private mdblSD As Double                            'SD
Private mintFormatNum As Integer                    '格式化小数位置
Private mstr质控品期限 As String
Dim lngCount As Long

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
    With Me.comDlg
        .CancelError = True
        .DialogTitle = "另存为"
        .filter = "(图形文件)|*.jpg"
        .FileName = Me.Caption & Format(mstrToDate, "yyyyMMdd") & ".jpg"
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
    '       str质控品期限 格式是: 质控品id=开始日期，结束日期， 用 ;分隔多个质控品，
    Dim rsTemp As New ADODB.Recordset
    Dim intCounts As Integer
    Dim lngResId As Long
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr质控品期限 = str质控品期限
    
    intCounts = Me.chk质控品.Count
    For lngCount = intCounts - 1 To 1 Step -1
        Unload Me.chk质控品(Me.chk质控品.UBound)
    Next
    Me.chk质控品(0).Enabled = False
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID, 批号 || '-' || 名称 As 质控品 From 检验质控品 Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList)
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.chk质控品.Count Then Load Me.chk质控品(.AbsolutePosition - 1)
            Me.chk质控品(.AbsolutePosition - 1).Caption = !质控品 & " (─" & AskLevelNote(.AbsolutePosition) & "─)"
            Me.chk质控品(.AbsolutePosition - 1).Tag = !ID
            Me.chk质控品(.AbsolutePosition - 1).Width = Me.TextWidth(Me.chk质控品(.AbsolutePosition - 1).Caption) + 360
            Me.chk质控品(.AbsolutePosition - 1).Value = vbChecked
            Me.chk质控品(.AbsolutePosition - 1).Visible = True
            Me.chk质控品(.AbsolutePosition - 1).Enabled = True
            .MoveNext
        Loop
    End With
    
    Call RefChart
    Call Form_Resize
    zlRefresh = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AskLevelShape(lngLevel As Long) As Long
    '功能：确定不同序号质控品的线条点形状
    Select Case lngLevel
    Case 1: AskLevelShape = oc2dShapeDot
    Case 2: AskLevelShape = oc2dShapeBox
    Case 3: AskLevelShape = oc2dShapeTriangle
    Case 4: AskLevelShape = oc2dShapeDiamond
    Case 5: AskLevelShape = oc2dShapeStar
    Case 6: AskLevelShape = oc2dShapeCircle
    Case 7: AskLevelShape = oc2dShapeSquare
    Case 8: AskLevelShape = oc2dShapeOpenTriangle
    Case 9: AskLevelShape = oc2dShapeOpenDiamond
    Case Else: AskLevelShape = oc2dShapeCross
    End Select
End Function

Private Function AskLevelNote(lngLevel As Long) As String
    '功能：确定不同序号质控品的线条点形状说明
    Select Case lngLevel
    Case 1: AskLevelNote = "●"
    Case 2: AskLevelNote = "■"
    Case 3: AskLevelNote = "▲"
    Case 4: AskLevelNote = "◆"
    Case 5: AskLevelNote = "＊"
    Case 6: AskLevelNote = "○"
    Case 7: AskLevelNote = "□"
    Case 8: AskLevelNote = "△"
    Case 9: AskLevelNote = "◇"
    Case Else: AskLevelNote = "×"
    End Select
End Function

Private Function AskLevelColor(lngLevel As Long) As Long
    '功能：确定不同序号质控品的线条颜色
    Select Case lngLevel
    Case 1: AskLevelColor = RGB(0, 0, 160)
    Case 2: AskLevelColor = RGB(0, 128, 255)
    Case 3: AskLevelColor = RGB(0, 128, 64)
    Case 4: AskLevelColor = RGB(0, 64, 128)
    Case 5: AskLevelColor = RGB(64, 128, 128)
    Case 6: AskLevelColor = RGB(128, 128, 192)
    Case 7: AskLevelColor = RGB(128, 128, 64)
    Case 8: AskLevelColor = RGB(128, 128, 128)
    Case 9: AskLevelColor = RGB(0, 255, 64)
    Case Else: AskLevelColor = RGB(0, 0, 0)
    End Select
    
End Function

Private Sub RefChart()
    '功能：刷新图形显示
    Dim rsTemp As New ADODB.Recordset
    Dim strLable As String, intRow As Integer, intCol As Integer
    Dim dblMax As Double
    Dim aryX() As Variant, aryY() As Variant
    Dim intLoop As Integer, dateLoop As Date
    Dim bln合并行 As Boolean
    Dim strLastData As String, strLastBadData As String
    Dim dblAvg1 As Double, dblSD1 As Double, dblAvg2 As Double, dblSD2 As Double, lngAVGcount As Long, intFormatNum As Integer
    '获取小数位数
    gstrSql = "Select 小数位数 from 检验仪器项目 where 项目ID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngItemID)
    If rsTemp.EOF = False Then mintFormatNum = Val(Nvl(rsTemp("小数位数"), 2))
    intFormatNum = mintFormatNum
    
    '--- 取均值与SD
    gstrSql = "Select x.质控品id, x.均值, Decode(x.Sd, Null, 1, 0, 1, x.Sd) As Sd, x.开始日期, Nvl(x.结束日期, m.结束日期) As 结束日期" & vbNewLine & _
            "From 检验质控品 M, 检验质控均值 X" & vbNewLine & _
            "Where m.Id = x.质控品id And x.项目id = [1] And" & vbNewLine & _
            "      Instr(';'|| [2] ||';',';' || x.质控品id || '=' || To_Char(x.开始日期, 'yyyy-MM-dd') || ',' || To_Char(Nvl(x.结束日期, m.结束日期), 'yyyy-mm-dd') || ';') > 0" & _
            " order by  质控品id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, mstr质控品期限)
    lngAVGcount = 0
    Do Until rsTemp.EOF
    
        For lngCount = 0 To Me.chk质控品.Count - 1
            If Me.chk质控品(lngCount).Value = 1 Then
                If Val("" & rsTemp!质控品id) = Val("" & Me.chk质控品(lngCount).Tag) Then
                    If lngAVGcount = 0 Then
                        dblAvg1 = Val("" & rsTemp!均值)
                        dblSD1 = Val("" & rsTemp!SD)
                        '均值，SD，默认小数，哪个精度高用哪个
                        If InStr("" & rsTemp!均值, ".") > 0 Then
                            If intFormatNum < Len(Mid("" & rsTemp!均值, InStr("" & rsTemp!均值, ".") + 1)) Then
                                intFormatNum = Len(Mid("" & rsTemp!均值, InStr("" & rsTemp!均值, ".") + 1))
                            End If
                        Else
                            If intFormatNum < 0 Then intFormatNum = 0
                        End If
                        
                        If InStr("" & rsTemp!SD, ".") > 0 Then
                            If intFormatNum < Len(Mid("" & rsTemp!SD, InStr("" & rsTemp!SD, ".") + 1)) Then
                                intFormatNum = Len(Mid("" & rsTemp!SD, InStr("" & rsTemp!SD, ".") + 1))
                            End If
                        Else
                            If intFormatNum < 0 Then intFormatNum = 0
                        End If
                        
                        lngAVGcount = lngAVGcount + 1
                    Else
                        dblAvg2 = Val("" & rsTemp!均值)
                        dblSD2 = Val("" & rsTemp!SD)
                        lngAVGcount = lngAVGcount + 1
                        Exit Do
                    End If
                End If
            End If
        Next
        rsTemp.MoveNext
    Loop
        
        
    '获得基本的文字信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select Distinct RPad('单位：' || '" & gstrUnitName & "', 56, ' ') || '日期：' As 行0," & vbNewLine & _
            "                RPad('仪器：' || D.名称, 56, ' ') || '试剂来源：' || M.试剂 As 行1," & vbNewLine & _
            "                RPad('项目：' || I.项目, 56, ' ') || '校准物来源：' || M.校准物 As 行2" & vbNewLine & _
            "From 检验仪器 D, 检验质控品 M, (Select 中文名 || ',' || 英文名 As 项目 From 诊治所见项目 Where ID = [2]) I" & vbNewLine & _
            "Where D.ID = M.仪器id And Instr(',' || [1] || ',', ',' || M.ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrResList, mlngItemID)
    If rsTemp.RecordCount <= 0 Then Me.chtThis.Header.Text = "该质控品信息不全面！": Exit Sub
    strLable = rsTemp!行0 & Format(mstrFromDate, "yyyy年MM月dd日") & "～" & Format(mstrToDate, "yyyy年MM月dd日")
    strLable = strLable & vbCrLf & rsTemp!行1 & vbCrLf & rsTemp!行2
    
    '将序列数字设置为0，清除图形显示
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    With Me.chtThis.Header
        .Text = "检验科Z-分数图" & vbCrLf & " " & vbCrLf & " "
        .Adjust = oc2dAdjustCenter
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    With Me.chtThis
        .ChartLabels.RemoveAll
        '行1
        .ChartLabels.Add
        .ChartLabels(1).AttachMethod = oc2dAttachCoord
        .ChartLabels(1).Text = rsTemp!行0 & Format(mstrFromDate, "yyyy年MM月dd日") & "～" & Format(mstrToDate, "yyyy年MM月dd日")
        .ChartLabels(1).AttachCoord.x = .Header.Location.Left + (.ChartLabels(1).Location.Width / 2) - 150
        .ChartLabels(1).AttachCoord.y = .Header.Location.Top + .Header.Location.Height - 30
        '行2
        .ChartLabels.Add
        .ChartLabels(2).AttachMethod = oc2dAttachCoord
        .ChartLabels(2).Adjust = oc2dAdjustRight
        .ChartLabels(2).Text = rsTemp!行1
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 150
        .ChartLabels(2).AttachCoord.y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '行3
        .ChartLabels.Add
        .ChartLabels(3).AttachMethod = oc2dAttachCoord
        .ChartLabels(3).Adjust = oc2dAdjustRight
        .ChartLabels(3).Text = rsTemp!行2
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 150
        .ChartLabels(3).AttachCoord.y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
        
    End With
    
    With Me.chtThis.Footer
        .Text = ""
        For lngCount = 0 To Me.chk质控品.Count - 1
            If chk质控品(lngCount).Value = 1 Then
                .Text = .Text & Space(6) & Me.chk质控品(lngCount).Caption
            End If
        Next
        .Text = Trim(.Text)
    End With
    
    '设置图形的基本形态
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 9 + Me.chk质控品.Count * 6
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
        For lngCount = 1 To Me.chk质控品.Count
            Me.chk质控品(lngCount - 1).ForeColor = AskLevelColor(lngCount)
            .Styles(9 + lngCount * 6 - 5).Symbol.Shape = AskLevelShape(lngCount)
            .Styles(9 + lngCount * 6 - 5).Line.COLOR = Me.chk质控品(lngCount - 1).ForeColor
            .Styles(9 + lngCount * 6 - 5).Symbol.COLOR = Me.chk质控品(lngCount - 1).ForeColor
            
            .Styles(9 + lngCount * 6 - 4).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6 - 4).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6 - 4).Symbol.COLOR = RGB(255, 0, 0)
            
            .Styles(9 + lngCount * 6 - 3).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6 - 3).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6 - 3).Symbol.COLOR = RGB(255, 0, 0)
            
            .Styles(9 + lngCount * 6 - 2).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6 - 2).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6 - 2).Symbol.COLOR = RGB(255, 0, 0)
            
            .Styles(9 + lngCount * 6 - 1).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6 - 1).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6 - 1).Symbol.COLOR = RGB(255, 0, 0)
            
            .Styles(9 + lngCount * 6).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6).Symbol.COLOR = RGB(255, 0, 0)
            Call chk质控品_Click(CInt(lngCount - 1))
        Next
    End With
    With Me.chtThis.ChartArea.Axes("Y")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels
        .Title.Text = "测定偏移(SD)"
        With .ValueLabels
            .RemoveAll
            If lngAVGcount = 1 Then
                .Add 4, Format(Val(dblAvg1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 1, Format(Val(dblAvg1) + 1 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 1, Format(Val(dblAvg1) - 1 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 2, Format(Val(dblAvg1) + 2 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 2, Format(Val(dblAvg1) - 2 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 3, Format(Val(dblAvg1) + 3 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 3, Format(Val(dblAvg1) - 3 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 4, Format(Val(dblAvg1) + 4 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 4, Format(Val(dblAvg1) - 4 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
            
            ElseIf lngAVGcount = 2 Then
                .Add 4, Format(Val(dblAvg1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 1, Format(Val(dblAvg1) + 1 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) + 1 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 1, Format(Val(dblAvg1) - 1 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) - 1 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 2, Format(Val(dblAvg1) + 2 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) + 2 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 2, Format(Val(dblAvg1) - 2 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) - 2 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 3, Format(Val(dblAvg1) + 3 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) + 3 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 3, Format(Val(dblAvg1) - 3 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) - 3 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 4, Format(Val(dblAvg1) + 4 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) + 4 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 4, Format(Val(dblAvg1) - 4 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) - 4 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
            Else
                .Add 4, 0
                .Add 4 + 1, 1
                .Add 4 - 1, -1
                .Add 4 + 2, 2
                .Add 4 - 2, -2
                .Add 4 + 3, 3
                .Add 4 - 3, -3
                .Add 4 + 4, 4
                .Add 4 - 4, -4
                
            End If
        End With
    End With
    With Me.chtThis.ChartArea.Axes("Y2")
        .AnnotationMethod = oc2dAnnotateValueLabels
        .Title.Text = "控制线"
        .Multiplier = 1
        With .ValueLabels
            .RemoveAll
            .Add 4, "CL"
            .Add 4 + 1, "CL+1SD": .Add 4 - 1, "CL-1SD"
            .Add 4 + 2, "CL+2SD": .Add 4 - 2, "CL-2SD"
            .Add 4 + 3, "CL+3SD": .Add 4 - 3, "CL-3SD"
        End With
    End With
    With Me.chtThis.ChartArea.Axes("X")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels
        .Title.Text = "日期"
    End With
    
    '数据组织
'    gstrSql = "Select Q.检验时间, Q.测试次数, Q.质控品id,x.均值,x.SD, Round(Max(Decode(Q.标记, 2, 0, (Q.结果 - X.均值) / X.Sd)), 4) As 在控," & vbNewLine & _
            "       Round(Max(Decode(标记, 2, (Q.结果 - X.均值) / X.Sd, 0)), 4) As 失控" & vbNewLine & _
            "From (Select Q.检验时间, Q.测试次数, Q.质控品id, T.标记," & vbNewLine & _
            "              Decode(I.值序列, Null, Zl_To_Number(R.检验结果)," & vbNewLine & _
            "                      Length(Substr(I.值序列, 1, Instr(I.值序列, ';' || RTrim(R.检验结果) || ';'))) -" & vbNewLine & _
            "                       Nvl(Length(Replace(Substr(I.值序列, 1, Instr(I.值序列, ';' || RTrim(R.检验结果) || ';')), ';')), 0)) As 结果" & vbNewLine & _
            "       From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T," & vbNewLine & _
            "            (Select Decode(结果类型, 3, Decode(RTrim(取值序列), '', '', ';' || RTrim(取值序列) || ';'), '') As 值序列" & vbNewLine & _
            "              From 检验项目" & vbNewLine & _
            "              Where 诊治项目id = [2]) I" & vbNewLine & _
            "       Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And /*Nvl(R.是否检验, 0) = 1 And*/ " & vbNewLine & _
            "             Instr(',' || [1] || ',', ',' || Q.质控品id || ',') > 0 And R.检验项目id + 0 = [2] And" & vbNewLine & _
            "             (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))) Q," & vbNewLine & _
            "     (Select X.质控品id, X.均值, Decode(X.Sd, Null, 1, 0, 1, X.Sd) As Sd" & vbNewLine & _
            "       From 检验质控品 M, 检验质控均值 X" & vbNewLine & _
            "       Where M.ID = X.质控品id And Instr(',' || [1] || ',', ',' || X.质控品id || ',') > 0 And X.项目id = [2] And" & vbNewLine & _
            "             (To_Date([3], 'yyyy-MM-dd') Between X.开始日期 And Nvl(X.结束日期, M.结束日期)) And" & vbNewLine & _
            "             (To_Date([4], 'yyyy-MM-dd') Between X.开始日期 And Nvl(X.结束日期, M.结束日期))) X" & vbNewLine & _
            "Where Q.质控品id = X.质控品id" & vbNewLine & _
            "Group By Q.检验时间, Q.测试次数, Q.质控品id,x.均值,x.SD " & vbNewLine & _
            "Order By Q.检验时间, Q.测试次数"
            
    gstrSql = "Select 检验时间, 测试次数, 质控品id, 均值, Sd," & vbNewLine & _
                "       Max(在控) As 在控,max(失控1) As 失控1,max(失控2) As 失控2,max(失控3) As 失控3,max(失控4) As 失控4 ,max(失控5) As 失控5" & vbNewLine & _
                "From (Select Q.检验时间, Q.测试次数, Q.质控品id, X.均值, X.Sd, to_char(Round(Max(Decode(Q.标记, 2, 0, decode(Q.结果,null,'',(Q.结果 - X.均值) / X.Sd))), 4)) As 在控, 0 As 失控1," & vbNewLine & _
                "              0 As 失控2, 0 As 失控3, 0 As 失控4, 0 As 失控5" & vbNewLine & _
                "       From (Select Q.检验时间, Q.测试次数, Q.质控品id, T.标记," & vbNewLine & _
                "                     zl_Lis_ToNumber(Q.质控品ID,R.检验项目id,R.检验结果,R.id ) As 结果" & vbNewLine & _
                "              From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T" & vbNewLine & _
                "              Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And Nvl(R.弃用结果,0)=0 And /*Nvl(R.是否检验, 0) = 1 And*/" & vbNewLine & _
                "                    Instr(',' || [1] || ',', ',' || Q.质控品ID || ',') > 0 And R.检验项目id + 0 = [2] And" & vbNewLine & _
                "                    (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')) " & vbNewLine & _
                "                    " & vbNewLine & _
                "      ) Q, "
    gstrSql = gstrSql & "" & vbNewLine & _
                "     (Select X.质控品id, X.均值, Decode(X.Sd, Null, 1, 0, 1, X.Sd) As Sd, X.开始日期, Nvl(X.结束日期, M.结束日期) As 结束日期" & vbNewLine & _
                "       From 检验质控品 M, 检验质控均值 X" & vbNewLine & _
                "       Where M.ID = X.质控品id And X.项目id = [2] And" & vbNewLine & _
                "         Instr(';' || [5] || ';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0  " & vbNewLine & _
                "      ) X" & vbNewLine & _
                "Where Nvl(Q.标记, 0) <> 2 And Q.质控品id = X.质控品id And  Q.检验时间 Between X.开始日期 And X.结束日期" & vbNewLine & _
                "Group By Q.检验时间, Q.测试次数, Q.质控品id, X.均值, X.Sd" & vbNewLine & _
                "" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Q.检验时间, Q.测试次数, Q.质控品id, X.均值, X.Sd, '' As 在控, Round(Max(Decode(标记, 2, (Q.结果1 - X.均值) / X.Sd, 0)), 4) As 失控1," & vbNewLine & _
                "       Round(Max(Decode(标记, 2, (Q.结果2 - X.均值) / X.Sd, 0)), 4) As 失控2," & vbNewLine & _
                "       Round(Max(Decode(标记, 2, (Q.结果3 - X.均值) / X.Sd, 0)), 4) As 失控3," & vbNewLine & _
                "       Round(Max(Decode(标记, 2, (Q.结果4 - X.均值) / X.Sd, 0)), 4) As 失控4," & vbNewLine & _
                "       Round(Max(Decode(标记, 2, (Q.结果5 - X.均值) / X.Sd, 0)), 4) As 失控5" & vbNewLine & _
                "From (Select 检验时间, 测试次数, 质控品id, 标记, Max(Decode(Mod(行号, 5), 1, 结果, '')) As 结果1," & vbNewLine & _
                "              Max(Decode(Mod(行号, 5), 2, 结果, '')) As 结果2, Max(Decode(Mod(行号, 5), 3, 结果, '')) As 结果3," & vbNewLine & _
                "              Max(Decode(Mod(行号, 5), 4, 结果, '')) As 结果4, Max(Decode(Mod(行号, 5), 0, 结果, '')) As 结果5" & vbNewLine & _
                "       From (Select Q.检验时间, Q.测试次数, Q.质控品id, T.标记," & vbNewLine & _
                "                     zl_Lis_ToNumber(Q.质控品ID,R.检验项目id,R.检验结果,R.id) As 结果," & vbNewLine & _
                "                     Rownum As 行号" & vbNewLine & _
                "              From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T "
    gstrSql = gstrSql & "" & vbNewLine & _
                "                     Where Q.标本id = R.检验标本id And R.ID = T.结果id And Nvl(R.弃用结果,0)=0 And /*Nvl(R.是否检验, 0) = 1 And*/" & vbNewLine & _
                "                           Instr(',' || [1] || ',', ',' || Q.质控品id || ',') > 0 And R.检验项目id + 0 = [2] And" & vbNewLine & _
                "                           (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')) And" & vbNewLine & _
                "                           Nvl(T.标记, 0) = 2  " & vbNewLine & _
                "                           )" & vbNewLine & _
                "              Group By 检验时间, 测试次数, 质控品id, 标记) Q," & vbNewLine & _
                "     (Select X.质控品id, X.均值, Decode(X.Sd, Null, 1, 0, 1, X.Sd) As Sd, X.开始日期, Nvl(X.结束日期, M.结束日期) As 结束日期" & vbNewLine & _
                "       From 检验质控品 M, 检验质控均值 X" & vbNewLine & _
                "       Where M.ID = X.质控品id And X.项目id = [2] And" & vbNewLine & _
                "         Instr(';' || [5] || ';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0  " & vbNewLine & _
                "      ) X" & vbNewLine & _
                "       Where Q.质控品id = X.质控品id And  Q.检验时间 Between X.开始日期 And X.结束日期 " & vbNewLine & _
                "       Group By Q.检验时间, Q.测试次数, Q.质控品id, X.均值, X.Sd)" & vbNewLine & _
                "group by 检验时间, 测试次数, 质控品id, 均值, Sd order by 检验时间,测试次数"



    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrResList, mlngItemID, mstrFromDate, mstrToDate, mstr质控品期限)
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    With rsTemp
        strLable = "": intRow = 0: strLastData = ""
        Do While Not .EOF
            '-1.新日期，应加标签
            '-2.日期相同，测试次数不同,上次和本次均在控，则加,失控不加
            
            If strLable <> Format(!检验时间, "yyyy-MM-dd") Then
                intRow = intRow + 1
            ElseIf strLable = Format(!检验时间, "yyyy-MM-dd") And strLastData <> Trim("" & !测试次数) And _
                (Trim("" & !在控) <> "" And strLastBadData <> "") Then
                intRow = intRow + 1

            End If
            strLable = Format(!检验时间, "yyyy-MM-dd")
            strLastData = Trim("" & !测试次数)
            strLastBadData = Trim("" & !在控)
            .MoveNext
        Loop
        If intRow < 30 Then
            intLoop = intRow
            ReDim aryX(31)
            ReDim aryY(31, 8 + Me.chk质控品.Count * 6)
        Else
            intLoop = 0
            ReDim aryX(intRow)
            ReDim aryY(intRow, 8 + Me.chk质控品.Count * 6)
        End If
        aryY(0, 0) = 4
        aryY(0, 1) = 4 + 1: aryY(0, 2) = 4 - 1
        aryY(0, 3) = 4 + 2: aryY(0, 4) = 4 - 2
        aryY(0, 5) = 4 + 3: aryY(0, 6) = 4 - 3
        aryY(0, 7) = 4 + 4: aryY(0, 8) = 4 - 4
        For lngCount = 1 To Me.chk质控品.Count
            aryY(0, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
        Next
        dblMax = 4
        strLable = "": intRow = 0
        If .RecordCount > 0 Then .MoveFirst
        strLastData = ""
        Do While Not .EOF
            mdblAVG = Val(Nvl(!均值))
            mdblSD = Val(Nvl(!SD))
            If mdblAVG = 0 Or mdblSD = 0 Then
                Me.chtThis.Header.Text = "尚未定值或SD为0，无法绘制" & Me.Caption & "！": Exit Sub
            End If
            If strLable <> Format(!检验时间, "yyyy-MM-dd") Then
                
                intRow = intRow + 1
'                Me.ChtThis.ChartArea.Axes("X").ValueLabels.Add intRow, intRow
                
                bln合并行 = False
                
                If Format(Nvl(!检验时间), "dd") <> "01" Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intRow, Format(Nvl(!检验时间), "dd")
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intRow, Format(Nvl(!检验时间), "mm" & "月")
                End If
                dateLoop = Format(Nvl(!检验时间), "yyyy-MM-dd")
                aryX(intRow) = intRow
                aryY(intRow, 0) = 4
                aryY(intRow, 1) = 4 + 1: aryY(intRow, 2) = 4 - 1
                aryY(intRow, 3) = 4 + 2: aryY(intRow, 4) = 4 - 2
                aryY(intRow, 5) = 4 + 3: aryY(intRow, 6) = 4 - 3
                aryY(intRow, 7) = 4 + 4: aryY(intRow, 8) = 4 - 4
                For lngCount = 1 To Me.chk质控品.Count
                    aryY(intRow, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
                Next
            ElseIf strLable = Format(!检验时间, "yyyy-MM-dd") And strLastData <> Trim("" & !测试次数) And _
                (Trim("" & !在控) <> "" And strLastBadData <> "") Then
                bln合并行 = False
                intRow = intRow + 1

                If Format(Nvl(!检验时间), "dd") <> "01" Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intRow, Format(Nvl(!检验时间), "dd")
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intRow, Format(Nvl(!检验时间), "mm" & "月")
                End If
                dateLoop = Format(Nvl(!检验时间), "yyyy-MM-dd")
                aryX(intRow) = intRow
                aryY(intRow, 0) = 4
                aryY(intRow, 1) = 4 + 1: aryY(intRow, 2) = 4 - 1
                aryY(intRow, 3) = 4 + 2: aryY(intRow, 4) = 4 - 2
                aryY(intRow, 5) = 4 + 3: aryY(intRow, 6) = 4 - 3
                aryY(intRow, 7) = 4 + 4: aryY(intRow, 8) = 4 - 4
                For lngCount = 1 To Me.chk质控品.Count
                    aryY(intRow, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
                Next
            Else
                bln合并行 = True
            End If
            
            strLable = Format(!检验时间, "yyyy-MM-dd")
            strLastData = Trim("" & !测试次数)
            strLastBadData = Trim("" & !在控)
                
            For lngCount = 1 To Me.chk质控品.Count
                If Val(Me.chk质控品(lngCount - 1).Tag) = Val("" & !质控品id) Then
                    
                    '当超出最大或最少值时定为最大或最少值
                    If Abs(Val("" & !在控)) > 4 Then
                        aryY(intRow, 8 + lngCount * 6 - 5) = 4 + IIf(Val("" & !在控) < -4, -4, 4)
                    Else
                        If Trim("" & !在控) = "" And bln合并行 = False Then
                            aryY(intRow, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
                        ElseIf Trim("" & !在控) <> "" Then
                            aryY(intRow, 8 + lngCount * 6 - 5) = 4 + Val("" & !在控)
                        End If
                    End If
                    
                    
                    If Val("" & !失控1) = 0 And bln合并行 = False Then
                        aryY(intRow, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '当超出最大或最少值时定为最大或最少值
                        If Abs(Val("" & !失控1)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6 - 4) = 4 + IIf(Val("" & !失控1) < -4, -4, 4)
                        ElseIf Val("" & !失控1) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6 - 4) = 4 + Val("" & !失控1)
                        End If
                    End If
                    
                    If Val("" & !失控2) = 0 And bln合并行 = False Then
                        aryY(intRow, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '当超出最大或最少值时定为最大或最少值
                        If Abs(Val("" & !失控2)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6 - 3) = 4 + IIf(Val("" & !失控2) < -4, -4, 4)
                        ElseIf Val("" & !失控2) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6 - 3) = 4 + Val("" & !失控2)
                        End If
                    End If
                    
                    If Val("" & !失控3) = 0 And bln合并行 = False Then
                        aryY(intRow, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '当超出最大或最少值时定为最大或最少值
                        If Abs(Val("" & !失控3)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6 - 2) = 4 + IIf(Val("" & !失控3) < -4, -4, 4)
                        ElseIf Val("" & !失控3) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6 - 2) = 4 + Val("" & !失控3)
                        End If
                    End If
                    
                    If Val("" & !失控4) = 0 And bln合并行 = False Then
                        aryY(intRow, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '当超出最大或最少值时定为最大或最少值
                        If Abs(Val("" & !失控4)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6 - 1) = 4 + IIf(Val("" & !失控4) < -4, -4, 4)
                        ElseIf Val("" & !失控4) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6 - 1) = 4 + Val("" & !失控4)
                        End If
                    End If
                    
                    If Val("" & !失控5) = 0 And bln合并行 = False Then
                        aryY(intRow, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '当超出最大或最少值时定为最大或最少值
                        If Abs(Val("" & !失控5)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6) = 4 + IIf(Val("" & !失控5) < -4, -4, 4)
                        ElseIf Val("" & !失控5) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6) = 4 + Val("" & !失控5)
                        End If
                    End If

                    Exit For
                End If
            Next
            .MoveNext
        Loop
    End With
    '如果不足30天的数据,补齐30天的数据
    'intLoop = 11
    If intLoop <> 0 Then
        For intLoop = intLoop + 1 To 31
            dateLoop = DateAdd("d", 1, dateLoop)
            If dateLoop <= CDate(mstrToDate) Then
                If Format(Nvl(dateLoop), "dd") <> "01" Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "dd")
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "mm" & "月")
                End If
            End If
            aryX(intLoop) = intLoop
            aryY(intLoop, 0) = 4
            aryY(intLoop, 1) = 4 + 1: aryY(intLoop, 2) = 4 - 1
            aryY(intLoop, 3) = 4 + 2: aryY(intLoop, 4) = 4 - 2
            aryY(intLoop, 5) = 4 + 3: aryY(intLoop, 6) = 4 - 3
            aryY(intLoop, 7) = 4 + 4: aryY(intLoop, 8) = 4 - 4
            
            For lngCount = 1 To Me.chk质控品.Count
                aryY(intLoop, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Next
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
            .Min = 4 - Val(dblMax)
            .Max = 4 + Val(dblMax)
        End With
        With .ChartArea.Axes("X")
            .Max = aryX(UBound(aryX))
        End With
        .IsBatched = False
        .AllowUserChanges = False
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
Private Sub chk质控品_Click(Index As Integer)
    If Me.Visible = False Then Exit Sub
    If Me.chk质控品(Index).Enabled = False Then Exit Sub
    With Me.chtThis.ChartGroups(1)
        If .Data.NumSeries < 9 + (Index + 1) * 2 - 1 Then Exit Sub
        If Me.chk质控品(Index).Value = vbChecked Then
            .Styles(9 + (Index + 1) * 6 - 5).Line.Pattern = oc2dLineSolid
            .Styles(9 + (Index + 1) * 6 - 5).Symbol.Size = 7
        Else
            .Styles(9 + (Index + 1) * 6 - 5).Line.Pattern = oc2dLineNone
            .Styles(9 + (Index + 1) * 6 - 5).Symbol.Size = 0
        End If
    End With
End Sub

Private Sub ChtThis_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim px As Long
    Dim py As Long
    Dim Series As Long
    Dim Point As Long
    Dim Distance As Long
    Dim Region As Long
    
    On Error Resume Next
    
    px = x / Screen.TwipsPerPixelX
    py = y / Screen.TwipsPerPixelY
    
    If (Button = 0) Then
        With chtThis
            Region = .ChartGroups(1).CoordToDataIndex(px, py, oc2dFocusXY, Series, Point, Distance)
            If (Series > 0 And Point > 0) And (Distance <= 5) Then
                If (Region = oc2dRegionInChartArea) Then
                    .ToolTipText = (Val(.ChartGroups(1).Data(Series, Point)) - 4)  '* mdblSD + mdblAVG
                    If mintFormatNum > 0 Then
                        .ToolTipText = Format(.ToolTipText, "###0." & Replace(Space(mintFormatNum), " ", "#"))
                    End If
                End If
            Else
'                .ToolTipText = ""
'                .Footer.Text = ""
            End If
            .Refresh
        End With
    End If
End Sub

Private Sub chtThis_Resize(ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    With Me.chtThis
        '行1
        .ChartLabels(1).AttachCoord.x = .Header.Location.Left + (.ChartLabels(1).Location.Width / 2) - 150
        .ChartLabels(1).AttachCoord.y = .Header.Location.Top + .Header.Location.Height - 30
        '行2
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 150
        .ChartLabels(2).AttachCoord.y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '行3
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 150
        .ChartLabels(3).AttachCoord.y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
        
    End With
    
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.chtThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight - Me.chk质控品(0).Height - Screen.TwipsPerPixelY * 4
    End With
    
    With Me.chk质控品(0)
        .Left = Me.ScaleLeft + Screen.TwipsPerPixelX * 2
        .Top = Me.ScaleHeight - .Height - Screen.TwipsPerPixelY * 2
    End With
    For lngCount = 1 To Me.chk质控品.Count
        With Me.chk质控品(lngCount)
            .Left = Me.chk质控品(lngCount - 1).Left + Me.chk质控品(lngCount - 1).Width + Screen.TwipsPerPixelX * 10
            .Top = Me.chk质控品(lngCount - 1).Top
        End With
    Next
End Sub






