VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmQCChartFQ 
   BorderStyle     =   0  'None
   Caption         =   "质控统计数据"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   60
      Top             =   2865
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgData 
      Height          =   2745
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8640
      _cx             =   15240
      _cy             =   4842
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
      Rows            =   9
      Cols            =   5
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
      Height          =   3135
      Left            =   30
      TabIndex        =   1
      Top             =   2865
      Width           =   8625
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   15214
      _ExtentY        =   5530
      _StockProps     =   0
      ControlProperties=   "frmQCChartFQ.frx":0000
   End
End
Attribute VB_Name = "frmQCChartFQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ID = 0: 统计项: 质控品: 个数: 平均值: 中位数: 标准差: CV: 最小值: 最大值
End Enum

Private mstrResList As String
Private mlngItemID As Long
Private mstrFromDate As String
Private mstrToDate As String
Private mstr质控品期限 As String
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Sub ChartPrint()
    With Me.chtThis
        '.PrintChart oc2dFormatBitmap, oc2dScaleToFit, 0, 0, 0, 0
        .Save App.path & "\QC_Tmp0"
    End With
End Sub

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

Public Function zlRefresh(strResList As String, lngItemID As Long, strFromDate As String, strToDate As String, str质控品期限 As String) As Boolean
    '功能：刷新本窗体的数据显示内容
    '参数： strResList  当前选择的质控品id串，以逗号分隔
    '       lngItemId   当前项目id
    '       strFromDate 开始日期
    '       strToDate   结束日期
    Dim rsTemp As New adodb.Recordset
    Dim lngRow As Long, lngCol As Long
    Dim lngResId As Long, intFact As Integer, intCounts As Integer
    Dim intFormatNum As Integer                 '取小数的位数
    
    gstrSql = "select 小数位数 from 检验仪器项目 where 项目id = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngItemID)
    If rsTemp.EOF = False Then intFormatNum = Val(Nvl(rsTemp("小数位数")))
    
    Me.Tag = "不刷新"
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr质控品期限 = str质控品期限
    Err = 0: On Error GoTo ErrHand
    '获取失控报告
    gstrSql = "Select Q.累积, Q.日期, Q.失控, Q.质控品id," & vbNewLine & _
            "       zl_To_Number(Q.结果) As 结果" & vbNewLine & _
            "From (Select 0 As 累积, Q.检验时间 As 日期, Decode(T.标记, 2, 1, 0) As 失控, Q.质控品id, zl_Lis_ToNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果" & vbNewLine & _
            "       From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T" & vbNewLine & _
            "       Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And /*Nvl(R.是否检验, 0) = 1 And*/ " & vbNewLine & _
            "             Instr(',' || [1] || ',', ',' || Q.质控品id || ',') > 0 And R.检验项目id + 0 = [2] And" & vbNewLine & _
            "             Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 1 As 累积, Q.检验时间 As 日期, Decode(T.标记, 2, 1, 0) As 失控, Q.质控品id, zl_Lis_ToNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果" & vbNewLine & _
            "       From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T,检验质控品 M " & vbNewLine & _
            "       Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And Nvl(R.弃用结果,0)=0 And /*Nvl(R.是否检验, 0) = 1 And */ " & vbNewLine & _
            "             Instr(',' || [1] || ',', ',' || Q.质控品id || ',') > 0 And R.检验项目id + 0 = [2] And" & vbNewLine & _
            "             Q.检验时间 between trunc(M.开始日期) and  To_Date([4], 'yyyy-MM-dd') And Q.质控品ID = M.id) Q"
            
    gstrSql = "Select M.ID, Decode(Q.累积, 1, '累积') || Decode(D.原始, 1, '原始数据', '在控数据') As 统计项," & vbNewLine & _
            "       M.批号 || '-' || M.名称 As 质控品, Count(*) As 个数, Round(Avg(结果), 3) As 平均值, 0 As 中位数," & vbNewLine & _
            "       Round(Stddev(结果), 3) As 标准差, Round(Stddev(结果) / Avg(结果) * 100, 2) As ""CV%"", Min(结果) As 最小值," & vbNewLine & _
            "       Max(结果) As 最大值" & vbNewLine & _
            "From 检验质控品 M, (Select Rownum As 原始 From 检验质控规则 Where Rownum <= 2) D," & vbNewLine & _
            "     (" & gstrSql & ") Q, 检验质控均值 X" & vbNewLine & _
            "Where Q.质控品id = M.ID And (D.原始 = 1 Or D.原始 = 2 And Q.失控 = 0)  And M.id=X.质控品id And X.项目id=[2] And Q.日期 Between X.开始日期 And nvl(X.结束日期,M.结束日期)" & vbNewLine & _
            "       And Instr(';' || [5] || ';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0  " & vbNewLine & _
            "Group By M.ID, M.批号, M.名称, Q.累积, D.原始" & vbNewLine & _
            "Order By Q.累积, D.原始, M.批号"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList, lngItemID, strFromDate, strToDate, mstr质控品期限)
    
    With Me.vfgData
        .Redraw = flexRDNone
        .Clear
        .FixedCols = 0
        Set .DataSource = rsTemp
        .FixedCols = mCol.统计项 + 1
        .MergeCells = flexMergeFixedOnly
        .MergeCol(mCol.ID) = True
        .MergeCol(mCol.统计项) = True
        '取小数位数
        If intFormatNum > 0 Then
            .ColFormat(mCol.标准差) = Replace("###0." & Space(intFormatNum), " ", "#")
            .ColFormat(mCol.平均值) = Replace("###0." & Space(intFormatNum), " ", "#")
            .ColFormat(mCol.中位数) = Replace("###0." & Space(intFormatNum), " ", "#")
        End If
        .ColFormat(mCol.CV) = "###0.#"
        .ColWidth(mCol.ID) = 0: .ColHidden(mCol.ID) = True
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        For lngCount = .FixedRows To .Rows - 1
            If Left(.TextMatrix(lngCount, mCol.平均值), 1) = "." Then .TextMatrix(lngCount, mCol.平均值) = "0" & .TextMatrix(lngCount, mCol.平均值)
            If Left(.TextMatrix(lngCount, mCol.标准差), 1) = "." Then .TextMatrix(lngCount, mCol.标准差) = "0" & .TextMatrix(lngCount, mCol.标准差)
            If Left(.TextMatrix(lngCount, mCol.CV), 1) = "." Then .TextMatrix(lngCount, mCol.CV) = "0" & .TextMatrix(lngCount, mCol.CV)
            If Left(.TextMatrix(lngCount, mCol.最小值), 1) = "." Then .TextMatrix(lngCount, mCol.最小值) = "0" & .TextMatrix(lngCount, mCol.最小值)
            If Left(.TextMatrix(lngCount, mCol.最大值), 1) = "." Then .TextMatrix(lngCount, mCol.最大值) = "0" & .TextMatrix(lngCount, mCol.最大值)
            '求中位数
            lngResId = CLng(.TextMatrix(lngCount, mCol.ID))
            intFact = IIf(InStr(1, .TextMatrix(lngCount, mCol.统计项), "原始") > 0, 1, 0)
            intCounts = CLng(.TextMatrix(lngCount, mCol.个数))
            If Left(.TextMatrix(lngCount, mCol.统计项), 2) <> "累积" Then
                gstrSql = "Select Avg(结果) As 中位数" & vbNewLine & _
                        "From (Select Rownum As 序号, 结果" & vbNewLine & _
                        "       From (Select 结果" & vbNewLine & _
                        "              From (Select zl_Lis_ToNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果" & vbNewLine & _
                        "                     From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T,检验质控品 M,检验质控均值 X" & vbNewLine & _
                        "                     Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And Nvl(R.弃用结果,0)=0 And /*Nvl(R.是否检验, 0) = 1 And*/ Q.质控品id + 0 = [1] And" & vbNewLine & _
                        "                           R.检验项目id + 0 = [2] And (1 = [3] Or Nvl(T.标记, 0) <> 2) And" & vbNewLine & _
                        "                           (Q.检验时间 Between To_Date([4], 'yyyy-MM-dd') And To_Date([5], 'yyyy-MM-dd')) And" & vbNewLine & _
                        "                           (Q.检验时间 Between X.开始日期 And NVL(X.结束日期,M.结束日期)) And " & vbNewLine & _
                        "                            Q.质控品id=M.id And M.id=X.质控品id And X.项目ID = [2] And " & vbNewLine & _
                        "                            Instr(';'||[7]||';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
                        "                    )" & vbNewLine & _
                        "              Order By 结果))" & vbNewLine & _
                        "Where 序号 Between [6] / 2 And [6] / 2 + 1"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, lngItemID, intFact, strFromDate, strToDate, intCounts, mstr质控品期限)
            Else
                gstrSql = "Select Avg(结果) As 中位数" & vbNewLine & _
                        "From (Select Rownum As 序号, 结果" & vbNewLine & _
                        "       From (Select 结果" & vbNewLine & _
                        "              From (Select zl_Lis_ToNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果" & vbNewLine & _
                        "                     From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T,检验质控品 M,检验质控均值 X" & vbNewLine & _
                        "                     Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And Nvl(R.弃用结果,0)=0  And /*Nvl(R.是否检验, 0) = 1 And*/ Q.质控品id + 0 = [1] And" & vbNewLine & _
                        "                           R.检验项目id + 0 = [2] And (1 = [3] Or Nvl(T.标记, 0) <> 2) And" & vbNewLine & _
                        "                           (Q.检验时间 between trunc(M.开始日期) and  To_Date([4], 'yyyy-MM-dd') ) And" & vbNewLine & _
                        "                           (Q.检验时间 Between X.开始日期 And NVL(X.结束日期,M.结束日期)) And " & vbNewLine & _
                        "                            Q.质控品id=M.id And M.id=X.质控品id  And  X.项目ID = [2] And " & vbNewLine & _
                        "                            Instr(';'||[6]||';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
                        "             ) Order By 结果))" & vbNewLine & _
                        "Where 序号 Between [5] / 2 And [5] / 2 + 1"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, lngItemID, intFact, strToDate, intCounts, mstr质控品期限)
            End If
            If rsTemp.RecordCount > 0 Then .TextMatrix(lngCount, mCol.中位数) = "" & rsTemp.Fields(0).Value
            If Left(.TextMatrix(lngCount, mCol.中位数), 1) = "." Then .TextMatrix(lngCount, mCol.中位数) = "0" & .TextMatrix(lngCount, mCol.中位数)
        Next
        Call .AutoSize(mCol.质控品, .Cols - 1)
        .Redraw = flexRDDirect
    End With
    Me.Tag = ""
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
    
    Dim lngResId As Long, strLable As String
    Dim aryX() As Variant, aryY() As Variant
    
    lngResId = Val(Me.vfgData.TextMatrix(Me.vfgData.Row, mCol.ID))
    
    '将序列数字设置为0，清除图形显示
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    With Me.chtThis.Header
        .Text = Me.vfgData.TextMatrix(Me.vfgData.Row, mCol.质控品) & "：频数分布图"
        .Font.Size = 16
        .Font.Bold = True
    End With
    If lngResId = 0 Then Exit Sub
    
    '设置图形的基本形态
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypeArea
        .Styles(oc2dTypePlot).Symbol.Shape = oc2dShapeBox
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 3
            .NumPoints(1) = 4
        End With
        .Styles(1).Line.COLOR = RGB(255, 192, 128)
        .Styles(2).Line.COLOR = RGB(255, 192, 128)
        .Styles(3).Line.COLOR = RGB(255, 192, 128)
    End With
    With Me.chtThis.ChartArea
        .Axes("Y").MajorGrid.Spacing.IsDefault = True
        .Axes("Y").AnnotationMethod = oc2dAnnotateValues
        .Axes("Y").Title.Text = "值个数"
        .Axes("X").MajorGrid.Spacing.IsDefault = True
        .Axes("X").AnnotationMethod = oc2dAnnotateValueLabels   '横坐标显示值提示
        .Axes("X").Title.Text = "测定值"
    End With
    
    '数据组织
    Dim rsTemp As New adodb.Recordset, strValTab As String
    ReDim aryX(12)
    ReDim aryY(12, 2)
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    
    Err = 0: On Error GoTo ErrHand


    gstrSql = "Select Nvl(Max(最小), 0) As 最小, Nvl(Max(最大), 0) As 最大, Nvl(Max(等差), 0) As 等差," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 0), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 1), -1, 1, 0))) As A," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 1), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 2), -1, 1, 0))) As B," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 2), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 3), -1, 1, 0))) As C," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 3), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 4), -1, 1, 0))) As D," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 4), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 5), -1, 1, 0))) As E," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 5), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 6), -1, 1, 0))) As F," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 6), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 7), -1, 1, 0))) As G," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 7), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 8), -1, 1, 0))) As H," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 8), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 9), -1, 1, 0))) As I," & vbNewLine & _
            "       Sum(Decode(Sign(结果 - 最小 - 等差 * 9), -1, 0, Decode(Sign(结果 - 最小 - 等差 * 10), 1, 0, 1))) As J, 0 As K" & vbNewLine & _
            "From "
    gstrSql = gstrSql & _
            "(Select Min(结果) As 最小, Max(结果) As 最大, (Max(结果) - Min(结果)) / 10 As 等差" & vbNewLine & _
            "       From (Select zl_Lis_toNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果" & vbNewLine & _
            "              From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T,检验质控品 M,检验质控均值 X" & vbNewLine & _
            "              Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And /*Nvl(R.是否检验, 0) = 1 And*/ Q.质控品id + 0 = [1] And" & vbNewLine & _
            "                    R.检验项目id + 0 = [2] And" & vbNewLine & _
            "                    (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')) And " & vbNewLine & _
            "                    (Q.检验时间 Between X.开始日期 And NVL(X.结束日期,M.结束日期)) And " & vbNewLine & _
            "                     Q.质控品id=M.id And M.id=X.质控品id  And  X.项目ID = [2] And " & vbNewLine & _
            "                     Instr(';'||[5]||';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "             ) ) S," & vbNewLine & _
            "     (Select zl_Lis_toNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果" & vbNewLine & _
            "       From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T,检验质控品 M,检验质控均值 X" & vbNewLine & _
            "       Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) And /*Nvl(R.是否检验, 0) = 1 And*/ Q.质控品id + 0 = [1] And" & vbNewLine & _
            "             R.检验项目id + 0 = [2] And" & vbNewLine & _
            "             (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')) And " & _
            "             (Q.检验时间 Between X.开始日期 And NVL(X.结束日期,M.结束日期)) And " & vbNewLine & _
            "              Q.质控品id=M.id And M.id=X.质控品id  And  X.项目ID = [2] And " & vbNewLine & _
            "             Instr(';'||[5]||';',';' || X.质控品id||'='||To_char(X.开始日期,'yyyy-MM-dd')||','||to_char(Nvl(X.结束日期, M.结束日期),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "      ) D"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstrFromDate, mstrToDate, mstr质控品期限)
    With rsTemp
        If !等差 = 0 Then Exit Sub
        For lngCount = LBound(aryX) To UBound(aryX)
            aryX(lngCount) = lngCount
            If lngCount > LBound(aryX) And lngCount < UBound(aryX) Then
                strLable = CStr(Round(!最小 + !等差 * (lngCount - 1), 3))
                Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, strLable
                Select Case (lngCount Mod 3)
                Case 1
                    aryY(lngCount, 0) = .Fields(Chr(64 + lngCount)).Value
                    aryY(lngCount, 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    If lngCount = LBound(aryX) + 1 Then
                        aryY(lngCount, 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        aryY(lngCount, 2) = .Fields(Chr(64 + lngCount - 1)).Value
                    End If
                Case 2
                    aryY(lngCount, 0) = .Fields(Chr(64 + lngCount - 1)).Value
                    aryY(lngCount, 1) = .Fields(Chr(64 + lngCount)).Value
                    aryY(lngCount, 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                Case 0
                    aryY(lngCount, 1) = .Fields(Chr(64 + lngCount - 1)).Value
                    aryY(lngCount, 2) = .Fields(Chr(64 + lngCount)).Value
                    aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                End Select
            Else
                aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(lngCount, 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
            End If
        Next
    End With

    '变更刷新内部数据
    Me.chtThis.IsBatched = True
    Me.chtThis.ChartGroups(1).Data.NumPoints(1) = UBound(aryX) + 1
    Call Me.chtThis.ChartGroups(1).Data.CopyXVectorIn(1, aryX)
    Call Me.chtThis.ChartGroups(1).Data.CopyYArrayIn(aryY)
    Me.chtThis.ChartArea.Axes("Y").Min = 0
    Me.chtThis.IsBatched = False
    Me.chtThis.AllowUserChanges = False
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
            End If
            .Refresh
        End With
    End If
End Sub

'--------------------------------------------
'以下为控件事件处理
'--------------------------------------------
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.vfgData
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop
        If .RowHeightMin * .Rows < Me.ScaleHeight * 3 / 5 Then
            .Height = .RowHeightMin * .Rows
        Else
            .Height = Me.ScaleHeight * 3 / 5
        End If
    End With
    With Me.chtThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.vfgData.Top + Me.vfgData.Height + Screen.TwipsPerPixelY
        .Height = Me.ScaleHeight - .Top
    End With
End Sub

Private Sub vfgData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Me.vfgData.Rows <= OldRow Or Me.vfgData.Rows <= NewRow Then Exit Sub
    If Me.vfgData.TextMatrix(NewRow, mCol.ID) = Me.vfgData.TextMatrix(OldRow, mCol.ID) Then Exit Sub
    If Me.Tag = "不刷新" Then Exit Sub
    Call RefChart
End Sub

Public Function ZLGetFQ_QCID() As Long
    '功能       得到当前使用的质控品的ID
    ZLGetFQ_QCID = Me.vfgData.TextMatrix(Me.vfgData.Row, mCol.ID)
End Function
