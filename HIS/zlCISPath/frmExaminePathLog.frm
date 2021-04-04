VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Begin VB.Form frmExaminePathLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "路径审核详情"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8070
   Icon            =   "frmExaminePathLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8070
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptLog 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _Version        =   589884
      _ExtentX        =   5318
      _ExtentY        =   1931
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6120
      TabIndex        =   1
      Top             =   4920
      Width           =   1100
   End
End
Attribute VB_Name = "frmExaminePathLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPathID As Long
Private mlngVersion As Long

Private Enum COL_LIST_LOG
    LOG_类型 = 0
    LOG_操作说明
    LOG_操作人员
    LOG_操作时间
End Enum

Public Sub ShowMe(ByRef objFrmMain As Object, ByVal lngPathID As Long, ByVal lngVersion As Long)
    mlngPathID = lngPathID
    mlngVersion = lngVersion
    Me.Show 1, objFrmMain
End Sub

Private Sub InitReportColumnLog()
    Dim objCol As ReportColumn

    With rptLog
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)或ItemIndex查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(LOG_类型, "审核结果", 100, True)
        Set objCol = .Columns.Add(LOG_操作说明, "审核说明", 200, True)
        Set objCol = .Columns.Add(LOG_操作人员, "审核人", 80, True)
        Set objCol = .Columns.Add(LOG_操作时间, "审核时间", 140, True)

        For Each objCol In .Columns
            objCol.Editable = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的审核内容..."
        End With

        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False    '会引发SelectionChanged事件
'        .SetImageList Me.img16
        
'        .GroupsOrder.Add .Columns(LOG_类型)
'        .GroupsOrder(0).SortAscending = True    '分组之后,如果分组列不显示,分组列的排序是不变的
'
'        '分组之后可能失去记录集中的顺序,因此强行加入排序列
'        .SortOrder.Add .Columns(LOG_类型)
'        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(LOG_操作时间)
        .SortOrder(0).SortAscending = False
    End With
End Sub

Private Sub LoadAduit(ByVal lngPathID As Long, ByVal lngVersion As Long)

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord   As ReportRecord
    Dim objItem     As ReportRecordItem
    If lngPathID = 0 Then
        rptLog.Records.DeleteAll
        rptLog.Populate
        Exit Sub
    End If

    If gbln双审核 Then
        strSql = "Select Decode(操作状态, 1, '医务科审核通过', 2, '医务科审核未过', 3, '药剂科审核通过', 4, '药剂科审核未过') As 操作状态, NVL(操作说明,'未填写') As 操作说明, 操作人员, 操作时间" & vbNewLine & _
            "From 临床路径审核" & vbNewLine & _
            "Where 路径id = [1] And 版本号 = [2]" & vbNewLine & _
            "Order By 操作时间 Desc"
    Else
        strSql = "Select Decode(操作状态, 1, '审核通过', 2, '审核未过', 3, '药剂科审核通过', 4, '药剂科审核未过') As 操作状态, NVL(操作说明,'未填写') As 操作说明, 操作人员, 操作时间" & vbNewLine & _
            "From 临床路径审核" & vbNewLine & _
            "Where 路径id = [1] And 版本号 = [2]" & vbNewLine & _
            "Order By 操作时间 Desc"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, lngVersion)
    
    rptLog.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptLog.Records.Add()
        Set objItem = objRecord.AddItem(rsTmp!操作状态 & "")
        Set objItem = objRecord.AddItem(rsTmp!操作说明 & "")
        Set objItem = objRecord.AddItem(rsTmp!操作人员 & "")
        Set objItem = objRecord.AddItem(Format(rsTmp!操作时间 & "", "YYYY-MM-DD HH:MM:SS"))
        rsTmp.MoveNext
    Loop

    rptLog.Populate
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitReportColumnLog
    Call LoadAduit(mlngPathID, mlngVersion)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.rptLog.Move 60, 60, Me.ScaleWidth - 120, Me.ScaleHeight - 700
    cmdOK.Move Me.ScaleWidth - cmdOK.Width - 240, Me.ScaleHeight - cmdOK.Height - 120
End Sub
