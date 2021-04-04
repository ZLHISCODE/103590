VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Begin VB.Form frmLabTrack 
   BorderStyle     =   0  'None
   Caption         =   "历史跟踪"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picChart 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   165
      ScaleHeight     =   2565
      ScaleWidth      =   8550
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2880
      Width           =   8550
      Begin VB.OptionButton opt内容 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "糖耐量(&3)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3660
         TabIndex        =   14
         Top             =   45
         Width           =   1260
      End
      Begin VB.OptionButton opt内容 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "变异率(&1)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   915
         TabIndex        =   12
         Top             =   45
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton opt内容 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "结果值(&2)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2250
         TabIndex        =   11
         Top             =   45
         Width           =   1260
      End
      Begin C1Chart2D8.Chart2D chtThis 
         Height          =   2085
         Left            =   30
         TabIndex        =   10
         Top             =   285
         Width           =   8520
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   15028
         _ExtentY        =   3678
         _StockProps     =   0
         ControlProperties=   "frmLabTrack.frx":0000
      End
      Begin VB.Label lbl项目 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目:RBC"
         Height          =   180
         Left            =   7665
         TabIndex        =   13
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lbl图形种类 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "图形内容:"
         Height          =   180
         Left            =   90
         TabIndex        =   9
         Top             =   45
         Width           =   810
      End
   End
   Begin VB.PictureBox picData 
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   120
      ScaleHeight     =   2445
      ScaleWidth      =   8550
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   8550
      Begin VSFlex8Ctl.VSFlexGrid vfgData 
         Height          =   2085
         Left            =   0
         TabIndex        =   7
         Top             =   315
         Width           =   8565
         _cx             =   15108
         _cy             =   3678
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
         BackColorFixed  =   15790320
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin VB.TextBox txt天数 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5085
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "10"
         Top             =   60
         Width           =   525
      End
      Begin VB.TextBox txt次数 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6945
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "3"
         Top             =   60
         Width           =   330
      End
      Begin VB.CheckBox chkHide 
         Appearance      =   0  'Flat
         Caption         =   "隐藏中文名"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   15
         TabIndex        =   1
         Top             =   75
         Width           =   1275
      End
      Begin VB.CommandButton cmdRefersh 
         Caption         =   "刷新"
         Height          =   350
         Left            =   7500
         TabIndex        =   6
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label lbl天数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最大跟踪天数:"
         Height          =   180
         Left            =   3945
         TabIndex        =   5
         Top             =   75
         Width           =   1170
      End
      Begin VB.Label lbl次数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最大跟踪次数:"
         Height          =   180
         Left            =   5790
         TabIndex        =   4
         Top             =   75
         Width           =   1170
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    结果类型 = 0: 中文名: 英文名: 报警率: 单位
End Enum

Private mlngRcdId As Long           '当前显示的样本记录的id
Private mstrEndTime As String       '本次检验时间
Private mintIdentMode As Integer    '历史比较病人识别方式

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim lngCount As Long, lngRow As Long, lngCol As Long

Private Function zlGetCV(ParamArray dbInput() As Variant) As Double
    '功能：返回多个数值的CV值的统计函数(变异系数)
    '参数：必须为数值数组
    Dim lngSubs As Long
    Dim dblSumAll As Double, dblSquSum As Double, dblSumSqu As Double
    Dim dblAV As Double, dblSD As Double
    
    If UBound(dbInput) < 1 Then zlGetCV = 0: Exit Function
    
    Err = 0: On Error GoTo 0
    dblSumAll = 0: dblSquSum = 0
    For lngSubs = LBound(dbInput) To UBound(dbInput)
        dblSumAll = dblSumAll + dbInput(lngSubs)
        dblSquSum = dblSquSum + dbInput(lngSubs) ^ 2
    Next
    If dblSumAll = 0 Then zlGetCV = 0: Exit Function
    dblSumSqu = dblSumAll ^ 2
    dblAV = dblSumAll / lngSubs
    dblSD = Sqr((dblSquSum - (dblSumSqu / lngSubs)) / (lngSubs - 1))
    zlGetCV = dblSD / dblAV * 100
    
End Function

Private Sub RefChart(Optional blnMust As Boolean)
    '功能：根据当前对比表显示指定内容的变化曲线
    '参数：是否强制重新获取数据进行刷新，否则当行未变化时，不进行刷新处理
    
    Dim aryX() As Variant, aryY() As Variant
    Dim intLoop As Integer, dblAvg As Double
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    If Val(Me.chtThis.Tag) <> Me.vfgData.Row Or blnMust Then
        Me.chtThis.Tag = Me.vfgData.Row
    Else
        Exit Sub
    End If
    
    '将序列数字设置为0，清除图形显示
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    
    If Me.vfgData.Row < Me.vfgData.FixedRows Then Me.lbl项目.Caption = "": Exit Sub
    
    '定性和定量项目不画图
    If Me.vfgData.TextMatrix(Me.vfgData.Row, mCol.结果类型) = "2" Or _
       Me.vfgData.TextMatrix(Me.vfgData.Row, mCol.结果类型) = "3" Then
       Me.chtThis.IsBatched = False
       Exit Sub
    End If
    
    
    '设置图形的基本形态
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot  '折线
        .Styles(oc2dTypePlot).Symbol.Shape = oc2dShapeBox
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 1
            .NumPoints(1) = 4
        End With
    End With
    With Me.chtThis.ChartArea
        .Axes("X").MajorGrid.Spacing.IsDefault = True
        .Axes("Y").MajorGrid.Spacing.IsDefault = True
        .Axes("X").AnnotationMethod = oc2dAnnotateValueLabels   '横坐标显示值提示
'        .Axes("X").AnnotationRotationAngle = 10
    End With
    
    If Me.opt内容(0).Value = True Then
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "变异率"
    ElseIf Me.opt内容(1).Value = True Then
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "结果值"
    Else
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "糖耐量"
    End If
    
    '数据组织
    Dim strMaxValue As String, strMinValue As String
    
    strMaxValue = 0
    If Me.opt内容(0).Value = True Or Me.opt内容(1).Value = True Then
        For intLoop = 0 To (Me.vfgData.Cols - mCol.单位 - 1) / 2 - 1
            If Val(vfgData.TextMatrix(vfgData.Row, mCol.单位 + 1 + intLoop * 2)) <> 0 Then
                j = j + 1
            End If
        Next
        If j = 0 Then j = 1
        ReDim aryX(j - 1)
        ReDim aryY(j - 1, 0)
    Else
        gstrSql = "Select 编码, 中文名, 英文名, 检验项目id, 检验结果, decode(别名,null,中文名,别名) as 名称 " & vbNewLine & _
                    "From (Select Decode(E.排列序号, Null, D.编码, E.排列序号) As 编码, D.中文名, D.英文名, B.检验项目id, B.检验结果, H.别名" & vbNewLine & _
                    "       From 检验标本记录 A, 检验普通结果 B, 检验仪器项目 C, 诊治所见项目 D, 检验项目 E, 检验报告项目 F, 诊疗项目目录 G," & vbNewLine & _
                    "            (Select 诊疗项目id, 名称 As 别名 From 诊疗项目别名 Where 性质 = 9 And 码类 = 1) H" & vbNewLine & _
                    "       Where A.ID = B.检验标本id And B.检验项目id = C.项目id And B.检验项目id = D.ID And Nvl(C.糖耐量项目, 0) = -1 And A.ID = [1] And" & vbNewLine & _
                    "             B.检验项目id = E.诊治项目id And B.检验项目id = F.报告项目id And F.诊疗项目id = G.ID And Nvl(G.组合项目, 0) = 0 And" & vbNewLine & _
                    "             G.ID = H.诊疗项目id(+)" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(E.排列序号, Null, D.编码, E.排列序号) As 编码, D.中文名, D.英文名, B.检验项目id, B.检验结果, H.别名" & vbNewLine & _
                    "       From 检验标本记录 A, 检验普通结果 B, 检验仪器项目 C, 诊治所见项目 D, 检验项目 E, 检验报告项目 F, 诊疗项目目录 G," & vbNewLine & _
                    "            (Select 诊疗项目id, 名称 As 别名 From 诊疗项目别名 Where 性质 = 9 And 码类 = 1) H" & vbNewLine & _
                    "       Where A.ID = B.检验标本id And B.检验项目id = C.项目id And B.检验项目id = D.ID And Nvl(C.糖耐量项目, 0) = -1 And A.合并id = [1] And" & vbNewLine & _
                    "             B.检验项目id = E.诊治项目id And B.检验项目id = F.报告项目id And F.诊疗项目id = G.ID And Nvl(G.组合项目, 0) = 0 And" & vbNewLine & _
                    "             G.ID = H.诊疗项目id(+))" & vbNewLine & _
                    "Order By 编码"

        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngRcdId)
        If rsTmp.RecordCount = 0 Then Exit Sub
        ReDim aryX(rsTmp.RecordCount - 1)
        ReDim aryY(rsTmp.RecordCount - 1, 0)
    End If
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    With Me.vfgData
        Me.lbl项目.Caption = "项目:" & .TextMatrix(.Row, mCol.中文名) & " (" & .TextMatrix(.Row, mCol.英文名) & ")"
        For lngCount = 0 To (Me.vfgData.Cols - mCol.单位 - 1) / 2 - 1
            If Val(.TextMatrix(.Row, mCol.单位 + 1 + lngCount * 2)) <> 0 Then
                aryX(i) = i

                If Me.opt内容(0).Value = True Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, .TextMatrix(0, mCol.单位 + 1 + lngCount * 2)
                    If Val(.TextMatrix(.Row, mCol.单位 + 2 + lngCount * 2)) = 0 And Val(.TextMatrix(.Row, mCol.单位 + 1 + lngCount * 2)) = 0 Then
    '                    aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        aryY(i, 0) = Val(.TextMatrix(.Row, mCol.单位 + 2 + lngCount * 2))
                    End If
                ElseIf Me.opt内容(1).Value = True Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, .TextMatrix(0, mCol.单位 + 1 + lngCount * 2)
                    If Val(.TextMatrix(.Row, mCol.单位 + 1 + lngCount * 2)) = 0 Then
    '                    aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        aryY(i, 0) = Val(.TextMatrix(.Row, mCol.单位 + 1 + lngCount * 2))
                    End If
                End If
                If Val(strMaxValue) < Abs(Val(aryY(i, 0))) Then
                    strMaxValue = Abs(Val(aryY(i, 0)))
                End If
                If Val(strMinValue) > Abs(Val(aryY(i, 0))) Then
                    strMinValue = Abs(Val(aryY(i, 0)))
                End If
                i = i + 1
            End If
        Next
    End With
    
    With Me.vfgData
        If Me.opt内容(2).Value = True Then
            For lngCount = LBound(aryX) To UBound(aryX)
                aryX(lngCount) = lngCount
                Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, Nvl(rsTmp("名称"))
                aryY(lngCount, 0) = Val(Nvl(rsTmp("检验结果")))
                rsTmp.MoveNext
                If Val(strMaxValue) < Abs(Val(aryY(lngCount, 0))) Then
                    strMaxValue = Abs(Val(aryY(lngCount, 0)))
                End If
                If Val(strMinValue) > Abs(Val(aryY(lngCount, 0))) Then
                    strMinValue = Abs(Val(aryY(lngCount, 0)))
                End If
            Next
        End If
    End With
    
    '变更刷新内部数据
    Me.chtThis.IsBatched = True
    Me.chtThis.ChartGroups(1).Data.NumPoints(1) = UBound(aryX) + 1
    Call Me.chtThis.ChartGroups(1).Data.CopyXVectorIn(1, aryX)
    Call Me.chtThis.ChartGroups(1).Data.CopyYArrayIn(aryY)
    
    If opt内容(0).Value = True Then
        Me.chtThis.ChartArea.Axes("Y").Origin = 0
        Me.chtThis.ChartArea.Axes("Y").Min = -1 * Val(strMaxValue)
        Me.chtThis.ChartArea.Axes("Y").Max = Val(strMaxValue)
    ElseIf opt内容(1).Value = True Then
        On Error Resume Next
        For intLoop = 0 To UBound(aryY, 1) - 1
            dblAvg = dblAvg + Val(aryY(intLoop, 0))
        Next
        If dblAvg <> 0 Then
            dblAvg = dblAvg / UBound(aryY, 1)
            Me.chtThis.ChartArea.Axes("Y").Origin = dblAvg
            If (dblAvg - Val(strMinValue)) < (Val(strMaxValue) - dblAvg) Then
                Me.chtThis.ChartArea.Axes("Y").Min = Val(dblAvg - (Val(strMaxValue) - dblAvg))
                Me.chtThis.ChartArea.Axes("Y").Max = Val(dblAvg + (Val(strMaxValue) - dblAvg))
            Else
                Me.chtThis.ChartArea.Axes("Y").Min = Val(dblAvg - (dblAvg - Val(strMinValue)))
                Me.chtThis.ChartArea.Axes("Y").Max = Val(dblAvg + (dblAvg - Val(strMinValue)))
            End If
        End If
    Else
        Me.chtThis.ChartArea.Axes("Y").Origin = 0
        Me.chtThis.ChartArea.Axes("Y").Min = 0
        Me.chtThis.ChartArea.Axes("Y").Max = Val(strMaxValue)
    End If
    Me.chtThis.IsBatched = False

End Sub

Private Sub setListFormat(Optional blnKeepData As Boolean)
    '功能：初始化设置参考值列表
    '参数： blnKeepData-是否保留数据，即只是重新设置格式
    With Me.vfgData
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 2: .FixedRows = 2: .Cols = mCol.单位 + 1: .FixedCols = .Cols
            For lngCol = 0 To mCol.单位: .TextMatrix(0, lngCol) = "项目": Next
            .TextMatrix(1, mCol.结果类型) = "结果类型"
            .TextMatrix(1, mCol.中文名) = "中文名"
            .TextMatrix(1, mCol.英文名) = "英文名"
            .TextMatrix(1, mCol.报警率) = "报警率"
            .TextMatrix(1, mCol.单位) = "单位"
            .MergeCells = flexMergeFixedOnly
            .MergeRow(0) = True
            .ColWidth(mCol.结果类型) = 0
            .ColWidth(mCol.中文名) = 1500
            .ColWidth(mCol.英文名) = 900
            .ColWidth(mCol.报警率) = 0
            .ColWidth(mCol.单位) = 500
        End If
        If .Cols > mCol.单位 + 1 Then
            .TextMatrix(0, mCol.单位 + 1) = "本次结果"
            .TextMatrix(1, mCol.单位 + 1) = "本次结果"
            .MergeCol(mCol.单位 + 1) = True
            .ColWidth(mCol.单位 + 2) = 0
        End If
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .AutoSize mCol.中文名, .Cols - 1
        If .Cols > .FixedCols Then .Col = .FixedCols
        If .Rows > .FixedRows Then .Row = .FixedRows
        Call RefChart(True)
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngRcdId As Long) As Boolean
    '功能：根据仪器id刷新当前显示内容
    '参数：当前项目id
    Dim lngDates As Long, lngTimes As Long
    Dim strRows As String, aryRows() As String
    Dim strCols As String, aryCols() As String
    Dim dblCurCV As Double     '计算的CV
    Dim strPatientName As String                    '病人姓名
    Dim strPatinetSex As String                     '病人姓别
    Dim lngPatientID As Long
    
    If lngRcdId = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    Err = 0: On Error GoTo ErrHand
    
    '获得当前检验的时间、项目要求的跟踪天数（取项目中最大的）
    If mlngRcdId <> lngRcdId Then
        mlngRcdId = lngRcdId
        gstrSql = "Select Nvl(L.核收时间, Sysdate) As 核收时间, Nvl(Max(跟踪天数), 0) As 天数" & vbNewLine & _
                "From 检验项目选项 O, 检验报告项目 X, 检验普通结果 R, 检验标本记录 L" & vbNewLine & _
                "Where O.诊疗项目id(+) = X.诊疗项目id And X.报告项目id = R.检验项目id And R.检验标本id = L.ID And L.ID = [1]" & vbNewLine & _
                "Group By Nvl(L.核收时间, Sysdate)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngRcdId)
        If rsTemp.RecordCount > 0 Then
            Me.txt天数.Text = rsTemp!天数
            mstrEndTime = Format(rsTemp!核收时间, "yyyy-MM-dd hh:mm:ss")
        Else
            Me.txt天数.Text = 30
            mstrEndTime = Format(Now(), "yyyy-MM-dd hh:mm:ss")
        End If
    End If
    If Val(Me.txt天数.Text) <= 0 Then Me.txt天数.Text = 30
    If Val(Me.txt次数.Text) <= 0 Then Me.txt次数 = 3
    
    lngDates = Val(Me.txt天数.Text)
    lngTimes = Val(Me.txt次数.Text)
    
'    If mintIdentMode <> 0 Then
        gstrSql = "select 姓名,病人ID,性别 from 检验标本记录 where id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngRcdId)
        If rsTemp.EOF = False Then strPatientName = Nvl(rsTemp("姓名")): lngPatientID = Nvl(rsTemp("病人ID"), 0): strPatinetSex = Nvl(rsTemp("性别"))
'    End If
    
    '查询历次数据装入：
    gstrSql = "Select /*+ RULE */ I.ID, I.名称 As 中文名, V.缩写 As 英文名, I.计算单位 As 单位, L.次数, L.核收时间, L.检验结果, V.变异报警率,V.结果类型 " & vbNewLine & _
            "From (Select L.检验项目id, L.次数, L.核收时间, L.检验结果 " & vbNewLine & _
            "       From (Select M.病人id As 病人id, M.姓名, M.性别, L.ID As 次数, L.核收时间, R.检验项目id, R.检验结果,L.标本类型 " & vbNewLine & _
            "              From 检验标本记录 L, 检验普通结果 R, 病人医嘱记录 M, " & _
            "                   (select 病人id,姓名,性别 from 病人信息 where " & IIf(mintIdentMode = 0, " 病人ID = [4] ", " 姓名 = [5] and 性别 = [6] ") & " ) N " & vbNewLine & _
            "              Where M.ID = L.医嘱id And L.ID = R.检验标本id And  " & vbNewLine & _
            "                    L.核收时间 Between [2]  And" & vbNewLine & _
            "                    [3] and L.病人id = N.病人id ) L," & vbNewLine & _
            "            (Select M.病人id As 病人id, M.姓名, M.性别, L.核收时间, R.检验项目id,L.标本类型 " & vbNewLine & _
            "              From 病人医嘱记录 M, 检验标本记录 L, 检验普通结果 R" & vbNewLine & _
            "              Where M.ID = L.医嘱id And L.ID = R.检验标本id And L.ID = [1]) C" & vbNewLine & _
            "        " & IIf(mintIdentMode = 0, "Where L.病人id = C.病人id   ", " Where  l.姓名 = c.姓名 And l.性别 = c.性别  ") & _
            "        And L.检验项目id+0 = C.检验项目id And L.标本类型 = C.标本类型  ) L, 检验项目 V, 检验报告项目 R, 诊疗项目目录 I" & vbNewLine & _
            "Where L.检验项目id = V.诊治项目id And L.检验项目id = R.报告项目id And R.诊疗项目id = I.ID And I.组合项目 <> 1" & vbNewLine & _
            "Order By I.编码, L.核收时间 desc"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngRcdId, CDate(Format(mstrEndTime, "yyyy-MM-dd 00:00:00")) - lngDates, _
                                       CDate(Format(mstrEndTime, "yyyy-MM-dd HH:MM:SS")), lngPatientID, strPatientName, strPatinetSex)
    
    Err = 0: On Error GoTo 0
    strRows = "": strCols = ""
    With Me.vfgData
        .Redraw = flexRDNone
        .Rows = .FixedRows: .Cols = .FixedCols
        lngRow = 0: lngCol = 0
        Do While Not rsTemp.EOF
            If InStr(1, strRows & ",", "," & rsTemp!ID & ",") = 0 Then
                strRows = strRows & "," & rsTemp!ID
                .Rows = .Rows + 1: lngRow = .Rows - 1
                .RowData(lngRow) = CLng(rsTemp!ID)
            Else
                aryRows = Split(strRows, ",")
                For lngCount = LBound(aryRows) To UBound(aryRows)
                    If Val(aryRows(lngCount)) = rsTemp!ID Then lngRow = .FixedRows - 1 + lngCount: Exit For
                Next
            End If
            .TextMatrix(lngRow, mCol.结果类型) = "" & rsTemp!结果类型
            .TextMatrix(lngRow, mCol.中文名) = "" & rsTemp!中文名
            .TextMatrix(lngRow, mCol.英文名) = "" & rsTemp!英文名
            .TextMatrix(lngRow, mCol.报警率) = Val("" & rsTemp!变异报警率)
            .TextMatrix(lngRow, mCol.单位) = "" & rsTemp!单位
            
            If InStr(1, strCols & ",", "," & rsTemp!次数 & ",") = 0 Then
                If UBound(Split(strCols, ",")) < lngTimes + 1 Then
                    strCols = strCols & "," & rsTemp!次数
                    .Cols = .Cols + 2: lngCol = .Cols - 1
                    .ColData(lngCol - 1) = CLng(rsTemp!次数): .ColData(lngCol) = CLng(rsTemp!次数)
                    .TextMatrix(0, lngCol - 1) = Format(rsTemp!核收时间, "yy-MM-dd HH:mm")
                    .TextMatrix(0, lngCol) = .TextMatrix(0, lngCol - 1)
                    .TextMatrix(1, lngCol - 1) = "结果值": .TextMatrix(1, lngCol) = "变异率"
                    .TextMatrix(lngRow, lngCol - 1) = "" & rsTemp!检验结果
                End If
            Else
                aryCols = Split(strCols, ",")
                For lngCount = LBound(aryCols) To UBound(aryCols)
                    If Val(aryCols(lngCount)) = rsTemp!次数 Then lngCol = .FixedCols - 1 + lngCount * 2: Exit For
                Next
                .TextMatrix(lngRow, lngCol - 1) = "" & rsTemp!检验结果
            End If
        
            rsTemp.MoveNext
        Loop
        
        '变异率计算填写和报警色处理
        For lngRow = .FixedRows To .Rows - 1
            .TextMatrix(lngRow, mCol.单位 + 1) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.单位 + 1)), " .", "0."), " ", "")
            For lngCol = mCol.单位 + 4 To .Cols - 1 Step 2
                .TextMatrix(lngRow, lngCol - 1) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, lngCol - 1)), " .", "0."), " ", "")
                If Val(.TextMatrix(lngRow, lngCol - 1)) = 0 Or Val(.TextMatrix(lngRow, mCol.单位 + 1)) = 0 Then
                    dblCurCV = 0
                Else
                    dblCurCV = (Val(.TextMatrix(lngRow, lngCol - 1)) - Val(.TextMatrix(lngRow, mCol.单位 + 1))) / Val(.TextMatrix(lngRow, mCol.单位 + 1)) * 100
                End If
                .TextMatrix(lngRow, lngCol) = Format(dblCurCV, "0.00;-0.00; ; ")
                If Val(.TextMatrix(lngRow, mCol.报警率)) <> 0 And Abs(dblCurCV) > Val(.TextMatrix(lngRow, mCol.报警率)) Then
                    .Cell(flexcpBackColor, lngRow, lngCol) = RGB(248, 194, 169)
                End If
            Next
        Next
        .Redraw = flexRDDirect
    End With
    Call setListFormat(True)
    
    zlRefresh = True: Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefresh = False
End Function

Private Sub chkHide_Click()
    Me.vfgData.ColHidden(mCol.中文名) = (Me.chkHide.Value = vbChecked)
    If Me.Visible Then Me.vfgData.SetFocus
End Sub

Private Sub chtThis_GotFocus()
    Me.dkpMan.RecalcLayout
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

Private Sub cmdRefersh_Click()
    Call Me.zlRefresh(mlngRcdId)
    Me.vfgData.SetFocus
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1: Item.Handle = Me.picData.hWnd
    Case 2: Item.Handle = Me.picChart.hWnd
    End Select
End Sub

Private Sub Form_Load()

    '获得本地参数设置
    mintIdentMode = zlDatabase.GetPara("历史病人识别", 100, 1208, 1)
    '隐藏中文名
    If Val(zlDatabase.GetPara("隐藏中文名", 100, 1208, 0)) = 0 Then
        Me.chkHide.Value = vbUnchecked
    Else
        Me.chkHide.Value = vbChecked
    End If
    Me.txt次数.Text = 3

    '基本格式设置
    '------------------------------------------------------
    mlngRcdId = 0
    Me.chkHide.BackColor = Me.picData.BackColor
    Me.opt内容(0).BackColor = Me.picChart.BackColor
    Me.opt内容(1).BackColor = Me.picChart.BackColor
    Me.opt内容(2).BackColor = Me.picChart.BackColor
    Call setListFormat

    '窗格划分
    '-----------------------------------------------------
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(1, 200, 400, DockTopOf, Nothing)
    panThis.Title = "历史对比表"
    panThis.Options = PaneNoCaption
    Set panThis = dkpMan.CreatePane(2, 200, 300, DockBottomOf, Nothing)
    panThis.Title = "历史对比图"
    panThis.Options = PaneNoCaption
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.chkHide.Value = vbUnchecked Then
        zlDatabase.SetPara "隐藏中文名", 0, 100, 1208
    Else
        zlDatabase.SetPara "隐藏中文名", 1, 100, 1208
    End If
    Me.dkpMan.DestroyAll
End Sub

Private Sub opt内容_Click(Index As Integer)
    Call RefChart(True)
    Me.vfgData.SetFocus
End Sub

Private Sub picChart_Resize()
    Err = 0: On Error Resume Next
    Me.lbl项目.Left = Me.ScaleWidth - Me.lbl项目.Width - 90
    With Me.chtThis
        .Left = 0: .Width = Me.picChart.ScaleWidth
        .Height = Me.picChart.ScaleHeight - .Top
    End With
End Sub

Private Sub picData_Resize()
    Err = 0: On Error Resume Next
    With Me.cmdRefersh
        .Left = Me.picData.ScaleWidth - .Width + 15
    End With
    Me.txt次数.Left = Me.cmdRefersh.Left - 900
    Me.lbl次数.Left = Me.txt次数.Left - Me.lbl次数.Width
    Me.txt天数.Left = Me.lbl次数.Left - 900
    Me.lbl天数.Left = Me.txt天数.Left - Me.lbl天数.Width
    Me.chkHide.Left = 45
    
    With Me.vfgData
        .Left = -15: .Width = Me.picData.ScaleWidth - .Left * 2
        .Height = Me.picData.ScaleHeight - .Top
    End With
End Sub

Private Sub txt次数_GotFocus()
    Me.txt次数.SelStart = 0: Me.txt次数.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt次数_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt天数_GotFocus()
    Me.txt天数.SelStart = 0: Me.txt天数.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt天数_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgData_RowColChange()
    Call RefChart
End Sub
