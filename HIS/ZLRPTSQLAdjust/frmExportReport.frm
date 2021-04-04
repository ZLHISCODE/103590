VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportReport 
   BackColor       =   &H80000005&
   Caption         =   "报表导出备份"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   ControlBox      =   0   'False
   Icon            =   "frmExportReport.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmExportReport.frx":628A
   ScaleHeight     =   5925
   ScaleWidth      =   9330
   Begin VB.Frame fraRPTList 
      Caption         =   "SQL中涉及""病人费用记录""的报表清单"
      Height          =   4545
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   5415
      Begin MSComctlLib.ListView lvwReport 
         Height          =   4200
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   7408
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "系统"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "编号"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "说明"
            Object.Width           =   7938
         EndProperty
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出(&E)…"
      Height          =   350
      Left            =   7920
      TabIndex        =   3
      Top             =   120
      Width           =   1100
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   9270
      TabIndex        =   0
      Top             =   5388
      Visible         =   0   'False
      Width           =   9324
      Begin MSComctlLib.ProgressBar pgbState 
         Height          =   180
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正在执行"
         Height          =   180
         Left            =   135
         TabIndex        =   2
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.Frame fraFloder 
      Caption         =   "目标文件夹"
      Height          =   4620
      Left            =   5640
      TabIndex        =   5
      Top             =   720
      Width           =   3495
      Begin VB.DriveListBox div 
         Height          =   300
         Left            =   96
         TabIndex        =   7
         Top             =   276
         Width           =   3285
      End
      Begin VB.DirListBox dirFloder 
         Height          =   3870
         Left            =   96
         TabIndex        =   6
         Top             =   576
         Width           =   3285
      End
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   8640
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":6783
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":6A9D
            Key             =   "Publish"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":6DB7
            Key             =   "Fixed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":70D1
            Key             =   "PubFixed"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":73EB
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":7705
            Key             =   "BillPublish"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   8040
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":7A1F
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":7B79
            Key             =   "Publish"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":7CD3
            Key             =   "Fixed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":7E2D
            Key             =   "PubFixed"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":7F87
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportReport.frx":80E1
            Key             =   "BillPublish"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmExportReport.frx":823B
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   480
      Left            =   840
      TabIndex        =   4
      Top             =   60
      Width           =   7200
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   120
      Picture         =   "frmExportReport.frx":82C5
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmExportReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjFile As New FileSystemObject
Private mobjText As TextStream
Public Event StatusTextUpdate(ByVal strMSG As String) '要求更新主窗体状态栏文字
 
Private Sub ShowStatusInfor(ByVal strMSG As String)
    RaiseEvent StatusTextUpdate(strMSG)
End Sub

Public Sub RefreshList()
    Call LoadReportList
End Sub
 
Private Sub cmdExport_Click()
    Dim strPath As String, strFile As String, i As Long, k As Long
    Dim curDate As Date
    
    strPath = dirFloder.List(dirFloder.ListIndex)
    k = lvwReport.ListItems.Count
    If MsgBox("即将导出 " & k & " 张报表到 " & strPath & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    If CheckRPTIndex = False Then Exit Sub
    
    pgbState.Max = k
    picStatus.Visible = True
    Call Form_Resize
    Me.Refresh
    
    lblStatus.Caption = "正在创建报表日志..."
    If CreateLog = False Then Exit Sub
    
    curDate = Currentdate
    For i = 1 To k
        lblStatus.Caption = "正在导出第" & i & "张:" & lvwReport.ListItems(i).Text & ".ZLR"
        pgbState.Value = i

        strFile = "[" & lvwReport.ListItems(i).SubItems(2) & "]" & lvwReport.ListItems(i).Text & ".ZLR"
        If Not ExportReport(CLng(Mid(lvwReport.ListItems(i).Key, 2)), strPath & "\" & strFile, curDate) Then
            If MsgBox("导出报表时出现错误，要继续导出下一张报表吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        End If
        'If i = 10 Then Exit For    '调试用
    Next
        
    picStatus.Visible = False
    Call Form_Resize
    Call ShowStatusInfor("共导出了" & i & "张报表。")
End Sub

Private Function CreateLog() As Boolean
    Dim rstmp As ADODB.Recordset, strSQL As String, i As Long, blnT As Boolean
    Dim strOut As String, strIn As String, strDefault As String
    
    CreateLog = True
    
    '建表
    On Error GoTo errHandle
    strSQL = "Select 1 From All_Tables Where Table_Name = Upper('zlrptadjustlog') And Owner = 'ZLTOOLS'"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "查找日志表")
    If rstmp.RecordCount = 0 Then
        strSQL = "create table zltools.zlrptadjustlog(报表ID NUMBER(18),数据源 VARCHAR2(20),源id NUMBER(18),序号 NUMBER(2),缺省 NUMBER(1),更改 NUMBER(1),全表扫描 NUMBER(1))"
        On Error Resume Next
        gcnOracle.Execute strSQL
        If Err.Number <> 0 Then
            MsgBox "创建报表调整记录表(zltools.zlrptadjustlog)出错。", vbInformation, gstrSysName
            CreateLog = False
            Exit Function
        End If
    End If
    
    '写数据
    On Error GoTo errHandle
    strSQL = "Select 1 From zltools.zlrptadjustlog Where rownum<2"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "查找日志表")
    
    If rstmp.RecordCount = 0 Then
        strOut = ",ZL1_BILL_1111,ZL1_BILL_1111_1,ZL1_BILL_1120,ZL1_BILL_1121_1,ZL1_BILL_1121_2,ZL1_BILL_1121_3,ZL1_BILL_1122" & _
                ",ZL1_INSIDE_1111_1,ZL1_INSIDE_1121_1,ZL1_INSIDE_1260_1,ZL1_INSIDE_1862,ZL1_REPORT_1123,ZL1_REPORT_1421" & _
                ",ZL1_REPORT_1876,ZL1_REPORT_1877,ZL1_REPORT_1882,ZL1_REPORT_1883,ZL1_SUB_1420_3,ZL1_SUB_1432_1" & _
                ",ZL1_SUB_1432_2,ZL1_SUB_1875_3,ZL1_SUB_1880_1,ZL1_SUB_1880_2,ZL1_SUB_1880_3,"
        strIn = ",ZL1_BILL_1133,ZL1_BILL_1134,ZL1_BILL_1135,ZL1_INSIDE_1102,ZL1_INSIDE_1102_1,ZL1_INSIDE_1139_2,ZL1_INSIDE_1342_1,ZL1_INSIDE_1605,"
        
        strSQL = "Select A.编号,B.报表id, B.数据源, B.序号" & vbNewLine & _
                "From zltools.zlReports A," & vbNewLine & _
                "     (Select Distinct 报表id, 名称 数据源, -null As 序号" & vbNewLine & _
                "       From zltools.zlRPTDatas" & vbNewLine & _
                "       Where 对象 Like '%病人费用记录%'" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select A.报表id, A.名称 数据源, B.序号" & vbNewLine & _
                "       From zltools.zlRPTDatas A, zltools.zlRPTPars B" & vbNewLine & _
                "       Where A.Id = B.源id And B.对象 Like '%病人费用记录%') B" & vbNewLine & _
                "Where A.Id = B.报表id" & vbNewLine & _
                "Order By 程序id, 编号"
        Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "查找报表")
        With rstmp
            gcnOracle.BeginTrans: blnT = True
            For i = 1 To .RecordCount
                If InStr(strOut, "," & !编号 & ",") > 0 Then
                    strDefault = "1"
                ElseIf InStr(strIn, "," & !编号 & ",") > 0 Then
                    strDefault = "2"
                Else
                    strDefault = "0"
                End If
                strSQL = "Insert into zltools.zlrptadjustlog(报表ID,数据源,序号,缺省) values(" & !报表ID & ",'" & !数据源 & "'," & IIf(IsNull(!序号), "Null", !序号) & "," & strDefault & ")"
                gcnOracle.Execute strSQL
                .MoveNext
            Next
            gcnOracle.CommitTrans: blnT = False
        End With
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    If blnT Then gcnOracle.RollbackTrans
    CreateLog = False
End Function

Private Function CheckRPTIndex() As Boolean
    Dim rstmp As ADODB.Recordset, strSQL As String, strIndex As String
        
    strIndex = "ZLRPTITEMS_IX_报表ID"
    strSQL = "Select 1 From All_Indexes Where Index_Name = [1] And Owner='ZLTOOLS'"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "查找索引", strIndex)
    If rstmp.RecordCount = 0 Then
        MsgBox "缺省索引[" & strIndex & "]，导出报表将非常慢，请先根据安装脚本[zlServer.Sql]创建该索引。", vbInformation, gstrSysName
        Exit Function
    End If
    
    strIndex = "ZLRPTITEMS_IX_上级ID"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "查找索引", strIndex)
    If rstmp.RecordCount = 0 Then
        MsgBox "缺省索引[" & strIndex & "]，导出报表将非常慢，请先根据安装脚本[zlServer.Sql]创建该索引。", vbInformation, gstrSysName
        Exit Function
    End If
    CheckRPTIndex = True
End Function

Public Function GetExpField(objFld As ADODB.Field) As String
'功能：导出报表时用
    Dim strTmp As String
    
    If IsNull(objFld.Value) Then
        Exit Function
    ElseIf InStr(",系统,程序ID,功能,发布时间,", "," & objFld.Name & ",") > 0 Then
        Exit Function
    ElseIf objFld.Name = "编号" Then
        GetExpField = "[编号]" '导入时取当前时间
    ElseIf objFld.Name = "修改时间" Then
        GetExpField = "Sysdate" '导入时取当前时间
    ElseIf objFld.Name = "ID" Then
        GetExpField = "[NextVal]" '导入时取"当前表_ID.NextVal"
    ElseIf objFld.Name = "上级ID" Then
        GetExpField = "[CurrVal-X]" '导入时取"当前表_ID.CurrVal-X",X为上级ID不为空的开始数
    ElseIf objFld.Name = "报表ID" Then
        GetExpField = "[zlReports_ID.CurrVal]" '导入时取"zlReports_ID.CurrVal"
    ElseIf objFld.Name = "源ID" Then
        GetExpField = "[zlRPTDatas_ID.CurrVal]" '导入时取"zlRPTDatas_ID.CurrVal"
    ElseIf objFld.Name = "对象" Then
        GetExpField = Replace(UCase(objFld.Value), UCase(gstrDBUser) & ".", "USER.")
    Else '导入时根据数据类型转换取值
        Select Case objFld.Type
            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                GetExpField = objFld.Value
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                GetExpField = objFld.Value
            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                If Format(objFld.Value, "HH:mm:ss") = "00:00:00" Then
                    GetExpField = Format(objFld.Value, "yyyy-MM-dd")
                Else
                    GetExpField = Format(objFld.Value, "yyyy-MM-dd HH:mm:ss")
                End If
            Case adBinary, adVarBinary, adLongVarBinary
                '暂时不支持图片的处理
        End Select
    End If
End Function

Public Function ExportReport(ByVal lngRPTID As Long, ByVal strFile As String, ByVal curDate As Date) As Boolean
'功能：导出一张自定义报表
'参数：lngRPTID=报表ID
'      strFile=文件名
'返回：导出是否成功。
'说明：
'      1.对于已发布的报表,导出成为非发布报表
'      2.目前不支持图片元素内容的导出
    
    Dim rstmp As ADODB.Recordset
    Dim rsSub As ADODB.Recordset
    Dim rsSQL As ADODB.Recordset
    Dim objFld As ADODB.Field
    Dim i As Integer, j As Integer
    Dim blnOpen As Boolean, blnSub As Boolean
    Dim strSQL As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    strSQL = "Select * From zltools.zlReports Where ID=[1]"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    If rstmp.EOF Then
        MsgBox "没有发现指定报表的数据！", vbInformation, App.Title
        Exit Function
    End If
    
    '打开磁盘文件
    If mobjFile.FileExists(strFile) Then Call mobjFile.DeleteFile(strFile, True)
    Set mobjText = mobjFile.CreateTextFile(strFile, True)
    blnOpen = True
    
    '产生报表表头
    Call mobjText.WriteLine("[HEAD]")
    Call mobjText.WriteLine("报表编号=" & rstmp!编号)
    Call mobjText.WriteLine("报表名称=" & rstmp!名称)
    Call mobjText.WriteLine("报表说明=" & IIf(IsNull(rstmp!说明), "", rstmp!说明))
    Call mobjText.WriteLine("导出用户=" & gstrDBUser)
    Call mobjText.WriteLine("导出时间=" & Format(curDate, "yyyy-MM-dd HH:mm:ss"))
    
    '报表:ZLReport,以分号为行结束；以分号为一个字段结束,单分号为一条记录结束
    Call mobjText.WriteLine("[ZLREPORTS]")
    Call mobjText.WriteLine(";")
    For Each objFld In rstmp.Fields
        Call mobjText.WriteLine(objFld.Name & "=" & GetExpField(objFld) & ";")
    Next
    
    '报表格式
    'Set rsTmp = New ADODB.Recordset
    strSQL = "Select * From zltools.zlRPTFmts Where 报表ID=[1]"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    If Not rstmp.EOF Then
        Call mobjText.WriteLine("[ZLRPTFMTS]")
        For i = 1 To rstmp.RecordCount
            Call mobjText.WriteLine(";")
            For Each objFld In rstmp.Fields
                Call mobjText.WriteLine(objFld.Name & "=" & GetExpField(objFld) & ";")
            Next
            rstmp.MoveNext
        Next
    End If
    
    '报表元素
    'Set rsTmp = New ADODB.Recordset
    strSQL = "Select * From zltools.zlRPTItems Where 报表ID=[1] Start With 上级ID is NULL Connect by Prior ID=上级ID"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    If Not rstmp.EOF Then
        Call mobjText.WriteLine("[ZLRPTITEMS]")
        For i = 1 To rstmp.RecordCount
            Call mobjText.WriteLine(";")
            For Each objFld In rstmp.Fields
                Call mobjText.WriteLine(objFld.Name & "=" & GetExpField(objFld) & ";")
            Next
            rstmp.MoveNext
        Next
    End If
    
    '报表数据,'数据参数
    'Set rsTmp = New ADODB.Recordset
    strSQL = "Select * From zltools.zlRPTDatas Where 报表ID=[1]"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    'Set rsSQL = New ADODB.Recordset
    strSQL = "Select B.* From zltools.zlRPTDatas A,zlRPTSQLs B Where A.ID=B.源ID And A.报表ID=[1]"
    Set rsSQL = zlDatabase.OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    'Set rsSub = New ADODB.Recordset
    strSQL = "Select B.* From zltools.zlRPTDatas A,zlRPTPars B Where A.ID=B.源ID And A.报表ID=[1]"
    Set rsSub = zlDatabase.OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    If Not rstmp.EOF Then
        Call mobjText.WriteLine("[ZLRPTDATAS]")
        For i = 1 To rstmp.RecordCount
            If blnSub Then Call mobjText.WriteLine("[ZLRPTDATAS]")
            
            Call mobjText.WriteLine(";")
            For Each objFld In rstmp.Fields
                Call mobjText.WriteLine(objFld.Name & "=" & GetExpField(objFld) & ";")
            Next
            
            blnSub = False
            
            rsSQL.Filter = "源ID=" & rstmp!Id
            If Not rsSQL.EOF Then
                blnSub = True
                Call mobjText.WriteLine("[ZLRPTSQLS]")
                For j = 1 To rsSQL.RecordCount
                    Call mobjText.WriteLine(";")
                    For Each objFld In rsSQL.Fields
                        Call mobjText.WriteLine(objFld.Name & "=" & GetExpField(objFld) & ";")
                    Next
                    rsSQL.MoveNext
                Next
            End If
           
            rsSub.Filter = "源ID=" & rstmp!Id
            If Not rsSub.EOF Then
                blnSub = True
                Call mobjText.WriteLine("[ZLRPTPARS]")
                For j = 1 To rsSub.RecordCount
                    Call mobjText.WriteLine(";")
                    For Each objFld In rsSub.Fields
                        Call mobjText.WriteLine(objFld.Name & "=" & GetExpField(objFld) & ";")
                    Next
                    rsSub.MoveNext
                Next
            End If
            
            rstmp.MoveNext
        Next
    End If
    
    rstmp.Close
    rsSub.Close
    rsSQL.Close
    mobjText.Close
    Screen.MousePointer = 0
    
    ExportReport = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
    If blnOpen Then mobjText.Close
End Function


Private Sub div_Change()
    dirFloder.Path = div.Drive
End Sub

Private Sub Form_Load()
    Call LoadReportList
    lvwReport.ColumnHeaders(2).Position = 1 '系统
    lvwReport.ColumnHeaders(3).Position = 2 '编号
    
    cmdExport.Enabled = lvwReport.ListItems.Count > 0
    Call ShowStatusInfor("共" & lvwReport.ListItems.Count & "张报表")
End Sub

Private Sub LoadReportList()
    Dim rstmp As ADODB.Recordset
    Dim i As Long, objItem As ListItem
    
    lvwReport.ListItems.Clear
    Set rstmp = GetReportList()
    If Not rstmp Is Nothing Then
        For i = 1 To rstmp.RecordCount
            Set objItem = lvwReport.ListItems.Add(, "_" & rstmp!Id, rstmp!名称, "Report", "Report")
            
            objItem.SubItems(1) = IIf(IsNull(rstmp!系统), "共享", rstmp!系统)
            objItem.SubItems(2) = rstmp!编号
            objItem.SubItems(3) = Nvl(rstmp!说明)
            
            rstmp.MoveNext
        Next
    End If
End Sub

Private Function GetReportList() As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select A.Id, A.编号, A.名称, A.系统, A.程序id, A.说明" & vbNewLine & _
            "From zltools.zlReports A," & vbNewLine & _
            "     (Select Distinct 报表id" & vbNewLine & _
            "       From zltools.zlRPTDatas" & vbNewLine & _
            "       Where 对象 Like '%病人费用记录%'" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select A.报表id From zltools.zlRPTDatas A, zltools.zlRPTPars B Where A.Id = B.源id And B.对象 Like '%病人费用记录%') B" & vbNewLine & _
            "Where A.Id = B.报表id" & vbNewLine & _
            "Order By 系统,程序id, 编号"

    On Error GoTo errH
    Set GetReportList = zlDatabase.OpenSQLRecord(strSQL, "读取报表")

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Dim sngWidth As Long '最小宽度
    
    sngWidth = IIf(ScaleWidth < 5600, 5600, ScaleWidth)
    cmdExport.Left = Me.ScaleLeft + sngWidth - cmdExport.Width - 200
    
    With fraFloder  '固定宽度
        .Left = Me.ScaleLeft + sngWidth - fraFloder.Width - 200
        .Height = Me.ScaleHeight - .Top - 100 - IIf(picStatus.Visible, picStatus.Height, 0)
    End With
    dirFloder.Height = fraFloder.Height - dirFloder.Top - 50
    
    fraRPTList.Width = sngWidth - fraFloder.Width - 400
    fraRPTList.Height = fraFloder.Height
    lvwReport.Width = fraRPTList.Width - 300
    lvwReport.Height = fraRPTList.Height - 400
    
    pgbState.Width = picStatus.ScaleWidth - 200
 End Sub
Private Sub Form_Unload(Cancel As Integer)
    If picStatus.Visible Then Cancel = 1
End Sub

Private Sub lvwReport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim blnDesc As Boolean
    
    If ColumnHeader.Tag = "1" Then
        blnDesc = True
        ColumnHeader.Tag = ""
    Else
        blnDesc = False
        ColumnHeader.Tag = "1"
    End If
    lvwReport.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwReport.SortOrder = lvwDescending
    Else
        lvwReport.SortOrder = lvwAscending
    End If
    lvwReport.Sorted = True
    
    If Not lvwReport.SelectedItem Is Nothing Then lvwReport.SelectedItem.EnsureVisible
End Sub

Private Sub picStatus_Resize()
    On Error Resume Next
    pgbState.Width = picStatus.ScaleWidth - pgbState.Left * 2
End Sub
Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub subPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

 
 
 


