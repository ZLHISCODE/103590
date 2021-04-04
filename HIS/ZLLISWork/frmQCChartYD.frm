VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCChartYD 
   BorderStyle     =   0  'None
   Caption         =   "Youden图"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cbo质控品 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   4530
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4935
      Width           =   3180
   End
   Begin VB.ComboBox cbo质控品 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   570
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4920
      Width           =   3180
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   2955
      Left            =   510
      TabIndex        =   1
      Top             =   645
      Width           =   6570
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   11589
      _ExtentY        =   5212
      _StockProps     =   0
      ControlProperties=   "frmQCChartYD.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl质控品 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "横向"
      Height          =   180
      Index           =   1
      Left            =   4050
      TabIndex        =   4
      Top             =   5010
      Width           =   360
   End
   Begin VB.Label lbl质控品 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "纵向"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   4995
      Width           =   360
   End
End
Attribute VB_Name = "frmQCChartYD"
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
    Dim rsTemp As New adodb.Recordset
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr质控品期限 = str质控品期限
    
    Me.Tag = "不刷新"
    
    Me.cbo质控品(0).Enabled = False: Me.cbo质控品(1).Enabled = False
    Me.cbo质控品(0).Clear: Me.cbo质控品(1).Clear
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID, 批号 || '-' || 名称 As 质控品 From 检验质控品 Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList)
    With rsTemp
        Do While Not .EOF
            For lngCount = 0 To 1
                Me.cbo质控品(lngCount).AddItem "" & !质控品
                Me.cbo质控品(lngCount).ItemData(Me.cbo质控品(lngCount).NewIndex) = !ID
            Next
            .MoveNext
        Loop
    End With
    If Me.cbo质控品(0).ListCount < 2 Then Me.chtThis.Header.Text = "至少需要两个质控品才能绘制Youden图！": Exit Function
    Me.cbo质控品(0).ListIndex = 0: Me.cbo质控品(1).ListIndex = 1
    Me.cbo质控品(0).Enabled = True: Me.cbo质控品(1).Enabled = True
    
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
    Dim strLable As String
    Dim lngResIdY As Long, lngResIdX As Long
    Dim aryX() As Variant, aryY() As Variant
    
    lngResIdY = Me.cbo质控品(0).ItemData(Me.cbo质控品(0).ListIndex)
    lngResIdX = Me.cbo质控品(1).ItemData(Me.cbo质控品(1).ListIndex)
    
    '获得基本的文字信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select RPad('单位：' || '" & gstrUnitName & "', 56, ' ') || '日期：' As 行0," & vbNewLine & _
            "         RPad('仪器：' || D.名称, 56, ' ') || '试剂来源：' || M.试剂 As 行1," & vbNewLine & _
            "         RPad('项目：' || I.项目, 56, ' ') || '校准物来源：' || M.校准物 As 行2" & vbNewLine & _
            "From 检验仪器 D, 检验质控品 M, (Select 中文名 || ',' || 英文名 As 项目 From 诊治所见项目 Where ID = [2]) I" & vbNewLine & _
            "Where D.ID = M.仪器id And M.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResIdY, mlngItemID)
    If rsTemp.RecordCount <= 0 Then Me.chtThis.Header.Text = "该质控品信息不全面！": Exit Sub
    strLable = rsTemp!行0 & Format(mstrFromDate, "yyyy年MM月dd日") & "～" & Format(mstrToDate, "yyyy年MM月dd日")
    strLable = strLable & vbCrLf & rsTemp!行1 & vbCrLf & rsTemp!行2
    
    '将序列数字设置为0，清除图形显示
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    With Me.chtThis.Header
        .Text = "检验科Youden图" & vbCrLf & " " & vbCrLf & " "
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
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 30
        '行2
        .ChartLabels.Add
        .ChartLabels(2).AttachMethod = oc2dAttachCoord
        .ChartLabels(2).Adjust = oc2dAdjustRight
        .ChartLabels(2).Text = rsTemp!行1
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 150
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '行3
        .ChartLabels.Add
        .ChartLabels(3).AttachMethod = oc2dAttachCoord
        .ChartLabels(3).Adjust = oc2dAdjustRight
        .ChartLabels(3).Text = rsTemp!行2
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 150
        .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
                
    End With
    
    '设置图形的基本形态
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 1
            .NumPoints(1) = 4
        End With
        .Styles(1).Symbol.Shape = oc2dShapeDot: .Styles(1).Symbol.COLOR = RGB(0, 0, 160)
        .Styles(1).Line.Pattern = oc2dLineNone
    End With
    With Me.chtThis.ChartArea
        With .Axes("Y")
            .MajorGrid.Spacing.IsDefault = True
            .MajorGrid.Style.Pattern = oc2dLineSolid
            .AnnotationMethod = oc2dAnnotateValueLabels
            .Title.Text = Me.cbo质控品(0).Text
            .TitleRotation = oc2dRotate90Degrees
        End With
        With .Axes("Y2")
            .AnnotationMethod = oc2dAnnotateValueLabels
            .Multiplier = 1
        End With
        With .Axes("X")
            .MajorGrid.Spacing.IsDefault = True
            .MajorGrid.Style.Pattern = oc2dLineSolid
            .AnnotationMethod = oc2dAnnotateValueLabels
            .Title.Text = Me.cbo质控品(1).Text
        End With
    End With
    
    '坐标标记
    Dim dblAvgY As Double, dblSdY As Double, dblMaxY As Double
    Dim dblAvgX As Double, dblSdX As Double, dblMaxX As Double
    gstrSql = "Select X.质控品id, X.均值, Decode(X.Sd, Null, 1, 0, 1, X.Sd) As Sd" & vbNewLine & _
            "From 检验质控品 M, 检验质控均值 X" & vbNewLine & _
            "Where M.ID = X.质控品id And (M.ID = [1] Or M.ID = [2]) And X.项目id = [3] And" & vbNewLine & _
            "   Instr(';' || [4] || ';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0 "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResIdY, lngResIdX, mlngItemID, mstr质控品期限)
    With rsTemp
        Do While Not .EOF
            If lngResIdY = !质控品id Then
                dblAvgY = Val("" & !均值): dblSdY = Val("" & !SD):
            ElseIf lngResIdX = !质控品id Then
                dblAvgX = Val("" & !均值): dblSdX = Val("" & !SD):
            End If
            .MoveNext
        Loop
    End With
    If dblAvgY = 0 Or dblAvgX = 0 Or dblSdY = 0 Or dblSdX = 0 Then
        Me.chtThis.Header.Text = "尚未定值或SD为0，无法绘制" & Me.Caption & "！": Exit Sub
    End If
    With Me.chtThis.ChartArea.Axes("Y").ValueLabels
        .RemoveAll
        .Add Val(dblAvgY), Format(Val(dblAvgY), "0.00")
        .Add Val(dblAvgY) + 1 * Val(dblSdY), Format(Val(dblAvgY) + 1 * Val(dblSdY), "0.00")
        .Add Val(dblAvgY) - 1 * Val(dblSdY), Format(Val(dblAvgY) - 1 * Val(dblSdY), "0.00")
        .Add Val(dblAvgY) + 2 * Val(dblSdY), Format(Val(dblAvgY) + 2 * Val(dblSdY), "0.00")
        .Add Val(dblAvgY) - 2 * Val(dblSdY), Format(Val(dblAvgY) - 2 * Val(dblSdY), "0.00")
        .Add Val(dblAvgY) + 3 * Val(dblSdY), " ": .Add Val(dblAvgY) - 3 * Val(dblSdY), " "
    End With
    With Me.chtThis.ChartArea.Axes("Y2").ValueLabels
        .RemoveAll
        .Add Val(dblAvgY), "CL"
        .Add Val(dblAvgY) + 1 * Val(dblSdY), "CL+1SD"
        .Add Val(dblAvgY) - 1 * Val(dblSdY), "CL-1SD"
        .Add Val(dblAvgY) + 2 * Val(dblSdY), "CL+2SD"
        .Add Val(dblAvgY) - 2 * Val(dblSdY), "CL-2SD"
        .Add Val(dblAvgY) + 3 * Val(dblSdY), " "
        .Add Val(dblAvgY) - 3 * Val(dblSdY), " "
    End With
    With Me.chtThis.ChartArea.Axes("X").ValueLabels
        .RemoveAll
        .Add Val(dblAvgX), "CL=" & Format(Val(dblAvgX), "0.00")
        .Add Val(dblAvgX) + 1 * Val(dblSdX), "CL+1SD=" & Format(Val(dblAvgX) + 1 * Val(dblSdX), "0.00")
        .Add Val(dblAvgX) - 1 * Val(dblSdX), "CL-1SD=" & Format(Val(dblAvgX) - 1 * Val(dblSdX), "0.00")
        .Add Val(dblAvgX) + 2 * Val(dblSdX), "CL+2SD=" & Format(Val(dblAvgX) + 2 * Val(dblSdX), "0.00")
        .Add Val(dblAvgX) - 2 * Val(dblSdX), "CL-2SD=" & Format(Val(dblAvgX) - 2 * Val(dblSdX), "0.00")
        .Add Val(dblAvgX) + 3 * Val(dblSdX), " ": .Add Val(dblAvgX) - 3 * Val(dblSdX), " "
    End With
    
    '数据组织
    gstrSql = "Select 检验时间, 次数, Nvl(Max(Decode(质控品id, [1], 结果)), 0) As Y, Nvl(Max(Decode(质控品id, [2], 结果)), 0) As X" & vbNewLine & _
            "From (Select Q.检验时间, To_Char(Q.测试次数, '000') || '-' || Decode(Nvl(T.标记, 0), 0, 999, Q.测试次数) As 次数," & vbNewLine & _
            "              Q.质控品id," & vbNewLine & _
            "              zl_Lis_ToNumber(Q.质控品ID,R.检验项目id,R.检验结果,R.id) As 结果" & vbNewLine & _
            "       From 检验质控记录 Q, 检验普通结果 R,检验质控报告 T,检验质控品 M, 检验质控均值 X " & vbNewLine & _
            "       Where Q.标本id = R.检验标本id And /*Nvl(R.是否检验, 0) = 1 And*/ R.检验项目id + 0 = [3] And" & vbNewLine & _
            "             Nvl(R.弃用结果,0)=0 And R.ID=T.结果ID(+) And (Q.质控品id = [1] Or Q.质控品id = [2]) And" & vbNewLine & _
            "             (Q.检验时间 Between To_Date([4], 'yyyy-MM-dd') And To_Date([5], 'yyyy-MM-dd')) And " & vbNewLine & _
            "             (Q.检验时间 Between X.开始日期 And NVL(X.结束日期,M.结束日期)) And " & vbNewLine & _
            "              Q.质控品id=M.id And M.id=X.质控品id  And  X.项目ID = [3] And " & vbNewLine & _
            "             Instr(';'||[6]||';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "      )" & vbNewLine & _
            "Group By 检验时间, 次数" & vbNewLine & _
            "Order By 检验时间, 次数"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResIdY, lngResIdX, mlngItemID, mstrFromDate, mstrToDate, mstr质控品期限)
    With rsTemp
        If .RecordCount > 0 Then
            ReDim aryX(.RecordCount + 1)
            ReDim aryY(.RecordCount + 1, 0)
        Else
            ReDim aryX(1)
            ReDim aryY(1, 0)
        End If
        aryX(0) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
        Do While Not .EOF
            If !x = 0 Then
                aryX(.AbsolutePosition) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Else
                aryX(.AbsolutePosition) = !x
            End If
            If !Y = 0 Then
                aryY(.AbsolutePosition, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Else
                aryY(.AbsolutePosition, 0) = !Y
            End If
            .MoveNext
        Loop
    End With

    '变更刷新内部数据
    With Me.chtThis
        .IsBatched = True
        '设置为正方式
        If .Width > .Height Then
            .ChartArea.Location.Height = .Height / Screen.TwipsPerPixelY - .ChartArea.Location.Top
            .ChartArea.Location.Width = .ChartArea.Location.Height + 100
        Else
            .ChartArea.Location.Width = .Width / Screen.TwipsPerPixelX - .ChartArea.Location.Left
            .ChartArea.Location.Height = .ChartArea.Location.Width - 100
        End If
        .ChartArea.Location.Left = .Width / Screen.TwipsPerPixelX / 2 - .ChartArea.Location.Width / 2
        
        With .ChartGroups(1).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(aryY)
        End With
        With .ChartArea.Axes("Y")
            .Min = dblAvgY - 3 * dblSdY
            .Max = dblAvgY + 3 * dblSdY
        End With
        With .ChartArea.Axes("X")
            .Min = dblAvgX - 3 * dblSdX
            .Max = dblAvgX + 3 * dblSdX
        End With
        .IsBatched = False
        .AllowUserChanges = False
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'--------------------------------------------
'以下为控件事件处理
'--------------------------------------------
Private Sub cbo质控品_Click(Index As Integer)
    Dim intBrother As Integer
    
    If Me.Visible = False Then Exit Sub
    If Me.cbo质控品(Index).Enabled = False Then Exit Sub
    
    If Index = 0 Then
        intBrother = 1
    Else
        intBrother = 0
    End If
    If Me.cbo质控品(Index).ListIndex = Me.cbo质控品(intBrother).ListIndex Then
        Me.cbo质控品(intBrother).Enabled = False
        For lngCount = 0 To Me.cbo质控品(intBrother).ListCount - 1
            If Me.cbo质控品(Index).ListIndex <> lngCount Then
                Me.cbo质控品(intBrother).ListIndex = lngCount
                Exit For
            End If
        Next
        Me.cbo质控品(intBrother).Enabled = True
    End If
    If Me.Tag = "不刷新" Then Exit Sub
    Call RefChart
    Me.chtThis.SetFocus
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

Private Sub chtThis_Resize(ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    With Me.chtThis
        '行1
        .ChartLabels(1).AttachCoord.x = .Header.Location.Left + (.ChartLabels(1).Location.Width / 2) - 150
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 30
        '行2
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 150
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '行3
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 150
        .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
        
        If .Width > .Height Then
            .ChartArea.Location.Height = .Height / Screen.TwipsPerPixelY - .ChartArea.Location.Top
            .ChartArea.Location.Width = .ChartArea.Location.Height + 100
        Else
            .ChartArea.Location.Width = .Width / Screen.TwipsPerPixelX - .ChartArea.Location.Left
            .ChartArea.Location.Height = .ChartArea.Location.Width - 100
        End If
        
        .ChartArea.Location.Left = .Width / Screen.TwipsPerPixelX / 2 - .ChartArea.Location.Width / 2
    End With
    
    
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.cbo质控品(0)
        .Left = Me.lbl质控品(0).Width + Screen.TwipsPerPixelX * 4
        .Width = Me.ScaleWidth / 2 - .Left
        .Top = Me.ScaleHeight - .Height
    End With
    With Me.lbl质控品(0)
        .Left = Screen.TwipsPerPixelX * 2
        .Top = Me.cbo质控品(0).Top + (Me.cbo质控品(0).Height - .Height) / 2
    End With
    
    With Me.cbo质控品(1)
        .Left = Me.ScaleWidth / 2 + Me.lbl质控品(1).Width + Screen.TwipsPerPixelX * 4
        .Width = Me.ScaleWidth - .Left
        .Top = Me.ScaleHeight - .Height
    End With
    With Me.lbl质控品(1)
        .Left = Me.ScaleWidth / 2 + Screen.TwipsPerPixelX * 2
        .Top = Me.cbo质控品(1).Top + (Me.cbo质控品(1).Height - .Height) / 2
    End With
    
    With Me.chtThis
        .Left = 0: .Width = Me.ScaleWidth
        .Top = 0: .Height = Me.ScaleHeight - .Top - Me.cbo质控品(0).Height
    End With
End Sub
