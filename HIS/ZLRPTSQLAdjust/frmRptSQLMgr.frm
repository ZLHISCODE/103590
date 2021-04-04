VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRptSQLMgr 
   BackColor       =   &H80000005&
   Caption         =   "报表SQL调整工具"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   ControlBox      =   0   'False
   Icon            =   "frmRptSQLMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmRptSQLMgr.frx":5E12
   ScaleHeight     =   7005
   ScaleWidth      =   10830
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "按上下键改变当前行"
      Top             =   1320
      Width           =   4905
      _Version        =   589884
      _ExtentX        =   8652
      _ExtentY        =   9763
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.Frame fraCmd 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   7080
      TabIndex        =   12
      Top             =   525
      Width           =   3615
      Begin VB.CommandButton cmdSaveAll 
         Caption         =   "按缺省方式更改全部报表"
         Height          =   350
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   2160
      End
      Begin VB.CommandButton cmdDesign 
         Caption         =   "报表设计(&D)"
         Height          =   350
         Left            =   2400
         TabIndex        =   8
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存更改(&S)"
         Height          =   350
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1200
      End
   End
   Begin RichTextLib.RichTextBox rtbExplan 
      Height          =   1575
      Left            =   5400
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   20000
      TextRTF         =   $"frmRptSQLMgr.frx":630B
   End
   Begin VB.Frame fraMode 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   350
      Left            =   5040
      TabIndex        =   5
      Top             =   3600
      Width           =   5950
      Begin VB.OptionButton optMode 
         BackColor       =   &H80000005&
         Caption         =   "全部费用(&0)"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   19
         Top             =   0
         Value           =   -1  'True
         Width           =   1400
      End
      Begin VB.OptionButton optMode 
         BackColor       =   &H80000005&
         Caption         =   "住院费用(&2)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   18
         Top             =   0
         Width           =   1400
      End
      Begin VB.OptionButton optMode 
         BackColor       =   &H80000005&
         Caption         =   "门诊费用(&1)"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   0
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H80000005&
         Caption         =   "(按空格切换)"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4850
         TabIndex        =   20
         Top             =   30
         Width           =   1335
      End
      Begin VB.Label lblMode 
         BackColor       =   &H80000005&
         Caption         =   "更改为"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   30
         Width           =   615
      End
   End
   Begin VB.PictureBox picLR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   4920
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5685
      ScaleWidth      =   45
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1080
      Width           =   45
   End
   Begin RichTextLib.RichTextBox rtbNew 
      Height          =   2895
      Left            =   5040
      TabIndex        =   6
      Top             =   3960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5106
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   20000
      TextRTF         =   $"frmRptSQLMgr.frx":63AD
   End
   Begin RichTextLib.RichTextBox rtbOld 
      Height          =   2175
      Left            =   5040
      TabIndex        =   4
      Top             =   1320
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   20000
      TextRTF         =   $"frmRptSQLMgr.frx":6452
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptSQLMgr.frx":64DF
            Key             =   "签名"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptSQLMgr.frx":6831
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptSQLMgr.frx":6DCB
            Key             =   ""
            Object.Tag             =   "99"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptSQLMgr.frx":7365
            Key             =   ""
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptSQLMgr.frx":78FF
            Key             =   ""
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptSQLMgr.frx":7E99
            Key             =   ""
            Object.Tag             =   "90003"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptSQLMgr.frx":8233
            Key             =   ""
            Object.Tag             =   "90004"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraFind 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   700
      Left            =   120
      TabIndex        =   14
      Top             =   577
      Width           =   5175
      Begin VB.CommandButton cmdExplan 
         Caption         =   "查看执行计划(&X)"
         Height          =   350
         Left            =   3630
         TabIndex        =   25
         Top             =   0
         Width           =   1480
      End
      Begin VB.ComboBox cboModify 
         Height          =   300
         Left            =   2350
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   400
         Width           =   975
      End
      Begin VB.ComboBox cboDefault 
         Height          =   300
         Left            =   800
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   400
         Width           =   975
      End
      Begin VB.CheckBox chkOnlyTableFull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "仅显全表扫描的"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton checkExplan 
         Caption         =   "检查全表扫描"
         Height          =   350
         Left            =   2350
         TabIndex        =   1
         Top             =   0
         Width           =   1280
      End
      Begin VB.TextBox txtFind 
         Height          =   350
         Left            =   800
         TabIndex        =   0
         Top             =   0
         Width           =   1465
      End
      Begin VB.Label lblFact 
         BackColor       =   &H80000005&
         Caption         =   "更改"
         Height          =   255
         Left            =   1900
         TabIndex        =   24
         Top             =   450
         Width           =   375
      End
      Begin VB.Label lblDefault 
         BackColor       =   &H80000005&
         Caption         =   "缺省"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   450
         Width           =   375
      End
      Begin VB.Label lblFind 
         BackColor       =   &H80000005&
         Caption         =   "报表定位"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   75
         Width           =   735
      End
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "下面列出的是SQL中含有“病人费用记录”的数据源，请选择更改方式后执行""保存更改""(F2)。如果需要调整数据源的其它SQL，请执行""报表设计""。"
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
      Left            =   2160
      TabIndex        =   11
      Top             =   60
      Width           =   7125
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报表SQL调整"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   10
      Top             =   120
      Width           =   1320
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   0
      Picture         =   "frmRptSQLMgr.frx":85CD
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmRptSQLMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event StatusTextUpdate(ByVal strMSG As String) '要求更新主窗体状态栏文字

Private mrsSQL As ADODB.Recordset
Private mlngCurRow As Long
Private mblnUnChange As Boolean
Private mclsReport As clsReport
Private Enum mmode
    m未改 = -1
    m全部 = 0
    m门诊 = 1
    m住院 = 2
    m手工 = 3   '使用报表设计器进行了更改
End Enum
Private mlngDBVer As Long


Private Sub ShowStatusInfor(ByVal strMSG As String)
    RaiseEvent StatusTextUpdate(strMSG)
End Sub

Private Function GetSQLPlan(lng源id As Long, lng参数号 As Long, lngSys As Long) As ADODB.Recordset
    Dim strSQL As String, rstmp As ADODB.Recordset
    Dim strOwner As String, strSID As String
    Dim objPars As RPTPars
    
    Set objPars = GetParsObj(lng源id, lng参数号, lngSys)
    strOwner = GetSQLObj(lng源id, lng参数号)
    
    Set rstmp = GetRPTSQL(lng源id, lng参数号)
    strSQL = GetTextByRs(rstmp)
    strSQL = Replace(strSQL, "[系统]", lngSys)
    strSQL = RemoveNote(strSQL)
    strSQL = SQLReplaceOwner(strSQL, strOwner)
    
    If objPars.Count = 0 Then
        strSQL = GetExecSQL(strSQL)
    Else
        strSQL = GetExecSQL(strSQL, objPars)
    End If
        
    If strSQL <> "" Then
        On Error Resume Next
        strSID = lng源id & Time()
          
        strSQL = "explain plan set statement_id = '" & strSID & "' for " & strSQL & ""
        gcnOracle.Execute strSQL
        If Err.Number = 0 Then
            If mlngDBVer >= 100 Then
                strSQL = "Select Plan_Table_Output From Table(DBMS_XPLAN.DISPLAY)"
            Else
                strSQL = "Select Cardinality ||'    '|| LPad(' ', Level - 1) || Operation || ' ' || Options || ' ' || Object_Name as Plan_Table_Output" & vbNewLine & _
                        "From Plan_Table" & vbNewLine & _
                        "Connect By Prior ID = Parent_Id And Prior Statement_Id = Statement_Id" & vbNewLine & _
                        "Start With ID = 0 And Statement_Id = [1]" & vbNewLine & _
                        "Order By ID"
            End If
            Set GetSQLPlan = zlDatabase.OpenSQLRecord(strSQL, "执行计划", strSID)
            If mlngDBVer < 100 Then
                gcnOracle.Execute "Delete plan_table"
            End If
        End If
    End If
End Function

Private Sub cboDefault_Click()
    If mblnUnChange Then Exit Sub
    Call LoadReportList
End Sub

Private Sub cboModify_Click()
    If mblnUnChange Then Exit Sub
    Call LoadReportList
End Sub

Private Sub cmdExplan_Click()
    Dim rstmp As ADODB.Recordset, strPlan As String, i As Long, j As Long, lngLen As Long, strFind As String
    Dim lng源id As Long, str源 As String, lng参数号 As Long, lng报表id As Long, arrtmp As Variant, lngSys As Long
    
    If rtbExplan.Visible = False And mlngCurRow > -1 Then
        If rptList.Rows(mlngCurRow).Childs.Count = 0 Then
            arrtmp = Split(rptList.Rows(mlngCurRow).Record(0).Value, "|SP|")
            str源 = arrtmp(1)
            lng报表id = arrtmp(0)
            lng源id = Get源ID(lng报表id, str源)
            lng参数号 = Val(arrtmp(2))
            arrtmp = Split(rptList.Rows(mlngCurRow).Record.Tag, "|SP|")
            lngSys = Val(arrtmp(0))
        
            Set rstmp = GetSQLPlan(lng源id, lng参数号, lngSys)
            If Not rstmp Is Nothing Then
                For i = 1 To rstmp.RecordCount
                    strPlan = IIf(i = 1, "", strPlan & vbNewLine) & rstmp.Fields(0).Value
                    rstmp.MoveNext
                Next
            End If
            If rtbExplan.Visible = False Then rtbExplan.Visible = True
            rtbExplan.Text = strPlan
        Else
            rtbExplan.Visible = False
        End If
    Else
        rtbExplan.Visible = False
    End If
    
    If strPlan <> "" Then
        lngLen = Len(strPlan)
        strFind = "TABLE ACCESS FULL"
        j = 1
        Do
            i = InStr(j, strPlan, strFind)
            If i <= 0 Then Exit Do
            
            rtbExplan.SelStart = i - 1
            rtbExplan.SelLength = Len(strFind)
            rtbExplan.SelColor = &HFF&     '红
            
            j = i + Len(strFind)
        Loop While j < lngLen
        
        rtbExplan.SelStart = 0
        rtbExplan.SelLength = 0
    End If
    
    If rtbExplan.Visible Then
        cmdExplan.Caption = "查看SQL(&X)"
    Else
        cmdExplan.Caption = "查看计划(&X)"
    End If
    rptList.SetFocus
End Sub


Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdExplan_Click
End Sub

Private Sub checkExplan_Click()
    Dim strSQL As String, i As Long, j As Long, k As Long, blnHave As Boolean
    Dim rstmp As ADODB.Recordset
    Dim lng源id As Long, str源 As String, lng参数号 As Long, lng报表id As Long, arrtmp As Variant, lngSys As Long
    Dim blnUpdaterow As Boolean
    
    blnUpdaterow = checkExplan.Tag = "update"
            
    If rptList.Rows.Count < 1 Then Exit Sub
    On Error Resume Next
    
    If blnUpdaterow = False Then
        If MsgBox("检查列表清单中所有数据源的执行计划，标记执行计划中存在对“病人费用记录”进行全表扫描的行。此操作可能需要1至3分钟，你确定要继续吗？", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
        
        gcnOracle.Execute "Update zltools.zlrptadjustlog Set 全表扫描=Null"
    End If
    
    DoEvents    '为了即时显示状态栏提示
    '第0行不用检查
    For j = 1 To rptList.Rows.Count - 1
        If rptList.Rows(j).Childs.Count = 0 Then
            If blnUpdaterow = False Then
                Call ShowStatusInfor("共" & rptList.Rows.Count & "行,正在检查第" & j + 1 & "行")
            End If
            
            If blnUpdaterow And mlngCurRow = j Or blnUpdaterow = False Then
                arrtmp = Split(rptList.Rows(j).Record(0).Value, "|SP|")
                str源 = arrtmp(1)
                lng报表id = arrtmp(0)
                lng源id = Get源ID(lng报表id, str源)
                lng参数号 = Val(arrtmp(2))
                arrtmp = Split(rptList.Rows(j).Record.Tag, "|SP|")
                lngSys = Val(arrtmp(0))
            
                Set rstmp = GetSQLPlan(lng源id, lng参数号, lngSys)
                rstmp.Filter = "Plan_Table_Output Like '*TABLE ACCESS FULL*'"
                blnHave = False
                For i = 1 To rstmp.RecordCount
                    If rstmp!Plan_Table_Output Like "*病人费用记录*" Or rstmp!Plan_Table_Output Like "*住院费用记录*" Or rstmp!Plan_Table_Output Like "*门诊费用记录*" Then
                        blnHave = True
                        Exit For
                    End If
                    rstmp.MoveNext
                Next
                If blnHave Then
                    gcnOracle.Execute "Update zltools.zlrptadjustlog Set 全表扫描=1 Where 报表id=" & lng报表id & " And 数据源='" & str源 & "' And Nvl(序号,-1)=" & lng参数号
                    k = k + 1
                    rptList.Rows(j).Record(3).Caption = "★"
                Else
                    rptList.Rows(j).Record(3).Caption = ""
                End If
            End If
        End If
    Next
    rptList.Populate
    
    If blnUpdaterow = False Then
        chkOnlyTableFull.Visible = k > 0
        
        Call ShowStatusInfor("共" & k & "条数据源存在对“病人费用记录”表的全表扫描。")
    End If
End Sub

Private Sub chkOnlyTableFull_Click()
    Call RefreshList
End Sub

Private Sub chkUnModify_Click()
    Call LoadReportList
End Sub

Private Sub cmdDesign_Click()
    Dim strNO As String, lngSys As Long, strSQLText As String
    Dim arrtmp As Variant, lng源id As Long, lng参数号 As Long
    
    '选择"系统"行时已禁用了此按钮
    If rptList.SelectedRows.Count = 0 Then Exit Sub
    
    If mclsReport Is Nothing Then
        Set mclsReport = New clsReport
        Call mclsReport.InitOracle(gcnOracle)
    End If
    strNO = rptList.Rows(mlngCurRow).Record.Tag
    lngSys = Val(Split(strNO, "|SP|")(0))
    strNO = Split(strNO, "|SP|")(1)
        
    Call mclsReport.ReportDesign(gcnOracle, lngSys, strNO, frmMDIMain, True)
    
    '如果修改了内容，则填写自定义修改标记,选中报表行时不处理
    If rptList.Rows(mlngCurRow).Childs.Count = 0 Then
         arrtmp = Split(rptList.Rows(mlngCurRow).Record(0).Value, "|SP|")
         lng源id = Get源ID(arrtmp(0), arrtmp(1))
         lng参数号 = Val(arrtmp(2))
        
         Set mrsSQL = GetRPTSQL(lng源id, lng参数号)
         If mrsSQL.RecordCount = 0 Then  '修改数据源后会删除数据重新产生，如果删除了，重新加载报表列表
             Call LoadReportList
             
             '重新标记执行计划
             checkExplan.Tag = "update"
             Call checkExplan_Click
             checkExplan.Tag = ""
         Else
             mlngCurRow = 0
             Call rptList_SelectionChanged
         End If
    End If
End Sub

Private Function LoadReportList() As Boolean
    Dim rstmp As ADODB.Recordset
    Dim objsys As ReportRecord, objrpt As ReportRecord, objdata As ReportRecord, objPar As ReportRecord
    Dim objItem As ReportRecordItem, objItemRpt As ReportRecordItem, objItemData As ReportRecordItem, objItemPar As ReportRecordItem
    Dim i As Long, strOldSys As String, strOldRpt As String
    Dim strOldRow As String, blnHaveTabFull As Boolean
    Dim lngDefault As Long, lngModify As Long
    
    If cboDefault.ListIndex = 0 Then
        lngDefault = -2
    Else
        lngDefault = cboDefault.ListIndex
        If lngDefault = 3 Then lngDefault = 0
    End If
    
    If cboModify.ListIndex = 0 Then
        lngModify = -2
    Else
        lngModify = cboModify.ListIndex
        If lngModify = 3 Then
            lngModify = 0
        ElseIf lngModify = cboModify.ListCount - 1 Then
            lngModify = -1 '最后一个表示"未改"
        End If
    End If
    
    Set rstmp = GetReportList(chkOnlyTableFull.Value = 1, lngDefault, lngModify)
    If rstmp Is Nothing Then Exit Function
           
    
    With rptList
        If .SelectedRows.Count > 0 Then
            If .SelectedRows(0).Childs.Count = 0 Then
                strOldRow = .SelectedRows(0).Record(0).Value
            End If
        End If
        .Records.DeleteAll
        
        For i = 1 To rstmp.RecordCount
            If strOldSys <> rstmp!系统名 Then
                strOldSys = rstmp!系统名: strOldRpt = ""
                Set objsys = .Records.Add()
                objsys.Expanded = True
                
                Set objItem = objsys.AddItem(Val("" & rstmp!系统))
                objItem.Caption = strOldSys
                objItem.BackColor = &HFFC0C0      'frmMDIMain.cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
                objItem.ForeColor = &HFF0000
            End If
            
            If strOldRpt <> rstmp!编号 Then
                strOldRpt = rstmp!编号
                Set objrpt = objsys.Childs.Add()
                objrpt.Expanded = True
                objrpt.Tag = rstmp!系统 & "|SP|" & rstmp!编号
                
                Set objItemRpt = objrpt.AddItem(Val(rstmp!报表ID))
                objItemRpt.Caption = rstmp!编号 & ":" & rstmp!报表
                objItemRpt.BackColor = &HC0FFFF      '&HFFC0C0      'frmMDIMain.cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
            End If
                
            
            Set objdata = objrpt.Childs.Add()
            objdata.Tag = rstmp!系统 & "|SP|" & rstmp!编号
            If Not IsNull(rstmp!参数) Then
                Set objItemData = objdata.AddItem(rstmp!报表ID & "|SP|" & rstmp!数据源 & "|SP|" & rstmp!序号)
                objItemData.Caption = rstmp!参数 & "(" & rstmp!数据源 & "的参数)"
                objItemData.ForeColor = &H40C0&
            Else
                objdata.Expanded = True
                Set objItemData = objdata.AddItem(rstmp!报表ID & "|SP|" & rstmp!数据源 & "|SP|" & "-1")
                objItemData.Caption = rstmp!数据源
            End If
            Set objItemData = objdata.AddItem(Val("" & rstmp!缺省))
            objItemData.Caption = Choose(Val("" & rstmp!缺省) + 1, "全部", "门诊", "住院")
            Set objItemData = Nothing
            
            Set objItemData = objdata.AddItem(IIf(IsNull(rstmp!更改), -1, Val("" & rstmp!更改)))
            objItemData.Caption = IIf(IsNull(rstmp!更改), " ", Choose(Val("" & rstmp!更改) + 1, "全部", "门诊", "住院"))
            
            Set objItemData = objdata.AddItem("")   '是否全表扫描
            objItemData.Caption = IIf(IsNull(rstmp!全表扫描), " ", "★")
            If Not IsNull(rstmp!全表扫描) Then blnHaveTabFull = True
            
            rstmp.MoveNext
        Next
        .Populate
        If blnHaveTabFull Then chkOnlyTableFull.Visible = True
        If strOldRow <> "" Then
            For i = 0 To .Rows.Count - 1
                If .Rows(i).Childs.Count = 0 Then
                    If .Rows(i).Record(0).Value = strOldRow Then
                        Set .FocusedRow = .Rows(i)
                        .Rows(i).EnsureVisible
                        Exit For
                    End If
                End If
            Next
        End If
        If rstmp.RecordCount = 0 Then Call rptList_SelectionChanged
    End With
    
    Call ShowStatusInfor("共" & rstmp.RecordCount & "条记录(数据源和带SQL选择器的参数)！")
    LoadReportList = True
End Function

Private Function CheckModified() As Boolean
    Dim i As Long, lngNewMode As Long, lngOldMode As Long
        
    Call GetTwoMode(lngOldMode, lngNewMode)
    CheckModified = True
    If lngOldMode <> lngNewMode Then
        i = 0
        If MsgBox("当前更改未保存，是否自动保存?", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            CheckModified = SaveData
        Else
            cmdSave.Tag = ""
        End If
    End If
End Function


Private Sub Form_Load()
    Dim strHeadStr As String
    
    mblnUnChange = True
    
    cboDefault.AddItem "0-所有"
    cboDefault.AddItem "1-门诊"
    cboDefault.AddItem "2-住院"
    cboDefault.AddItem "3-全部费用"
    cboDefault.ListIndex = 0
    
    cboModify.AddItem "0-所有"
    cboModify.AddItem "1-门诊"
    cboModify.AddItem "2-住院"
    cboModify.AddItem "3-全部费用"
    cboModify.AddItem "4-未改"
    cboModify.ListIndex = 0
    
    mblnUnChange = False
    
    mlngCurRow = -1
    mblnUnChange = False
    strHeadStr = "标题,230;缺省,30;更改,30;全表扫描,54"
    
    Call InitReportListHead(strHeadStr)
    
    Call LoadReportList
    
    mlngDBVer = GetDBVer
End Sub


Private Sub InitReportListHead(strHeadStr As String)
    Dim arrtmp As Variant, arrItem As Variant, i As Long
    Dim rptCol As ReportColumn
    
    With rptList
        arrtmp = Split(strHeadStr, ";")
        For i = 0 To UBound(arrtmp)
            arrItem = Split(arrtmp(i), ",")
            If UBound(arrItem) > 0 Then
                Set rptCol = .Columns.Add(i, CStr(arrItem(0)), Val(arrItem(1)), True)
                rptCol.Visible = True
                rptCol.Editable = False
                rptCol.Groupable = False
                rptCol.Sortable = False
                rptCol.Alignment = xtpAlignmentLeft
            Else
                Set rptCol = .Columns.Add(i, CStr(arrItem(0)), 0, False)
                rptCol.Visible = False
            End If
        Next
                
        .SetImageList img16
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .AutoColumnSizing = False
        .ShowGroupBox = False
'        With .PaintManager
'            .ColumnStyle = xtpColumnFlat
'            .GridLineColor = RGB(225, 225, 225)
'            .NoGroupByText = "拖动列标题到这里,按该列分组..."
'            .NoItemsText = "没有找到符合条件的病人..."
'            .VerticalGridStyle = xtpGridSolid
'        End With

        .Columns(0).TreeColumn = True
        
'        .PaintManager.TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
'        .GroupsOrder.Add .Columns(0)
    End With
End Sub


Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub subPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

Public Sub RefreshList()
    Call LoadReportList
End Sub

Private Sub cmdSaveAll_Click()
    Dim lngidx As Long, lng缺省 As Long, k As Long
    Dim strErr As String
    
    If rptList.Rows.Count < 2 Then Exit Sub
    If MsgBox("你确定要对当前列表中的所有数据按缺省方式进行更改吗？", vbQuestion + vbOKCancel, Me.Caption) = vbCancel Then
        Exit Sub
    End If
    
    lngidx = 1
    Do
        If rptList.Rows(lngidx).Childs.Count = 0 Then
            Set rptList.FocusedRow = rptList.Rows(lngidx)   '激发事件rptList_SelectionChanged
            lng缺省 = Val("" & rptList.Rows(lngidx).Record(1).Value)
            optMode(lng缺省).Value = True   '激发click事件
            
            If SaveData = False Then
                If MsgBox("第" & lngidx & "行更改失败，是否继续处理下一行数据?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                    Exit Do
                Else
                    strErr = strErr & " , " & Split(rptList.Rows(lngidx).Record.Tag, "|SP|")(1) & "(" & rptList.Rows(lngidx).Record(0).Caption & ")"
                End If
            Else
                k = k + 1
            End If
        End If
        lngidx = lngidx + 1
    Loop While lngidx < rptList.Rows.Count
    
    If strErr <> "" Then
        strErr = Mid(strErr, 4, 1000) & IIf(Len(strErr) > 1000, "......", "")
        MsgBox "以下报表更改失败，请检查(可按Ctrl+C拷贝提示信息)" & vbCrLf & strErr
    End If
    
    Call ShowStatusInfor("共处理了" & k & "行数据。")
End Sub
Private Sub cmdSave_Click()
    If SaveData Then
        If mlngCurRow > 0 Then
            Set rptList.FocusedRow = rptList.Rows(mlngCurRow)
            Call rptList.SetFocus
        End If
    End If
End Sub
Private Function GetPrivsData(ByVal lng报表id As Long, str系统 As String, str程序id As String, str功能 As String) As Boolean
    Dim rstmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select 系统, 程序id, 功能" & vbNewLine & _
        "From zltools.zlReports" & vbNewLine & _
        "Where ID = [1] And 程序id Is Not Null" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select 系统, 程序id, 功能" & vbNewLine & _
        "From zltools.zlRPTPuts" & vbNewLine & _
        "Where 报表id = [1]" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select B.系统, B.程序id, A.功能 From zltools.zlRPTSubs A, zltools.zlRPTGroups B Where A.组id = B.Id And A.报表id = [1]"
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "读取权限数据", lng报表id)
    If rstmp.RecordCount > 0 Then
        str系统 = "" & rstmp!系统
        str程序id = "" & rstmp!程序id
        str功能 = "" & rstmp!功能
    End If
    
    GetPrivsData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOwner(lng源id As Long, lng参数号 As Long, strTab As String) As String
    Dim rstmp As ADODB.Recordset, strSQL As String
    Dim arrtmp As Variant, i As Long, p As Long
    
    If lng参数号 = -1 Then
        strSQL = "Select 对象 From zltools.zlrptdatas Where id=[1]"
    Else
        strSQL = "Select a.对象 From zltools.zlRPTPars Where a.源id=[1] And a.序号=[2]"
    End If
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "读取所有者", lng源id, lng参数号)
    If rstmp.RecordCount > 0 Then
        arrtmp = Split(rstmp!对象, ",")
        For i = 0 To UBound(arrtmp)
            p = InStr(arrtmp(i), ".")
            If Mid(arrtmp(i), p + 1) = strTab Then
                GetOwner = Mid(arrtmp(i), 1, p - 1)
                Exit For
            End If
        Next
    End If
   
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get源ID(ByVal lng报表id As Long, ByVal str源 As String) As Long
    Dim rstmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select ID From zltools.zlrptdatas Where 报表id=[1] And 名称=[2]"
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "读取源id", lng报表id, str源)
    If rstmp.RecordCount > 0 Then Get源ID = rstmp!Id
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetTwoMode(ByRef lngOldMode As Long, ByRef lngNewMode As Long)
    Dim i As Long, strSQL As String
    
    strSQL = RemoveNote(rtbOld.Text)
    i = InStr(strSQL, "病人费用记录")
    If i <= 0 Then
        i = InStr(strSQL, "门诊费用记录")
        If i <= 0 Then
            i = InStr(strSQL, "住院费用记录")
            If i > 0 Then lngOldMode = 2
        Else
            lngOldMode = 1
        End If
    Else
        lngOldMode = 0
    End If
    
    For i = 0 To optMode.UBound
        If optMode(i).Value = True Then lngNewMode = i: Exit For
    Next
End Sub

Private Function SaveData() As Boolean
    '没有选择行，或选择的是“系统”或“报表”行时，以及被外部修改过的记录，是禁用了此按钮的
    Dim arrtmp As Variant, lng参数号 As Long, lng源id As Long, str源 As String
    Dim blnTrans As Boolean, strSQL As String, strObj As String, strSQLContent As String
    Dim i As Long, lngNewMode As Long, lngOldMode As Long
    Dim strTabOld As String, strTabNew As String
    Dim str系统 As String, str程序id As String, str功能 As String
    Dim lng报表id As Long, strOwner As String, blnExists As Boolean, blnNewPrivs As Boolean, blnDel As Boolean
    Dim rsRole As ADODB.Recordset
        
    arrtmp = Split(rptList.Rows(mlngCurRow).Record(0).Value, "|SP|")
    lng报表id = Val(arrtmp(0))
    str源 = arrtmp(1)
    lng参数号 = Val(arrtmp(2))
    
    lng源id = Get源ID(lng报表id, str源)
    If lng源id = 0 Then Exit Function
    
    Call GetTwoMode(lngOldMode, lngNewMode)
    If lngOldMode = lngNewMode Then
        If Trim(rptList.Rows(mlngCurRow).Record(2).Caption) = "" Then
            strSQL = "Update zltools.Zlrptadjustlog Set 更改=" & Choose(lngNewMode + 1, 0, 1, 2) & " Where 报表ID=" & lng报表id & " And 数据源='" & str源 & "' And Nvl(序号,-1)=" & lng参数号
            gcnOracle.Execute strSQL
            rptList.Rows(mlngCurRow).Record(2).Value = lngNewMode
            rptList.Rows(mlngCurRow).Record(2).Caption = Choose(lngNewMode + 1, "全部", "门诊", "住院")
            rptList.Populate
        Else
            Call ShowStatusInfor("当前内容未更改，无需保存。")
        End If
        SaveData = True
        Exit Function
    End If
    strTabOld = Choose(lngOldMode + 1, "病人费用记录", "门诊费用记录", "住院费用记录")
    strTabNew = Choose(lngNewMode + 1, "病人费用记录", "门诊费用记录", "住院费用记录")
    
    If GetPrivsData(lng报表id, str系统, str程序id, str功能) = False Then
        Exit Function
    End If
    
    
    '执行SQL检查语法
    If CheckSQLPhrase(lng源id, lng参数号, Val(str系统), strTabOld, strTabNew) = False Then
        Exit Function
    End If
        
    If str程序id <> "" Then '仅报告单没有程序ID
        strOwner = GetOwner(lng源id, lng参数号, strTabOld)
        If strOwner = "" Then
            Call ShowStatusInfor("保存失败，未找到SQL对象的所有者。")
            Exit Function
        End If
        
        blnExists = ExistTablePrivs(str系统, str程序id, str功能, strTabNew)
        If blnExists = False Then
            blnNewPrivs = ExistOtherTablePrivs(strTabOld)
        Else
            blnDel = Not ExistOtherTablePrivs(strTabOld)
        End If
    End If
    
    strObj = "Replace(对象, '" & strTabOld & "','" & strTabNew & "')"
    
    If str程序id <> "" And str功能 <> "" Then
        If Not blnExists Then
            strSQL = "Select 角色" & vbNewLine & _
                    "From zltools.zlRoleGrant A" & vbNewLine & _
                    "Where Nvl(系统, 0) = [1] And 序号 = [2] And 功能 = [3] And Exists (Select 1 From dba_Role_Privs B Where A.角色 = B.Granted_Role)"
            On Error Resume Next    '如果没有权限访问dba_Role_Privs，则不授权
            Set rsRole = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str系统), str程序id, str功能)
            Err.Clear
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        
        If lng参数号 = -1 Then
            strSQL = "Update zltools.zlRPTDatas Set 对象=" & strObj & " Where ID=" & lng源id
            gcnOracle.Execute strSQL
            
            strSQLContent = "Replace(内容, '" & strTabOld & "','" & strTabNew & "')"
            mrsSQL.Filter = "内容 like '*" & strTabOld & "*'"
            For i = 1 To mrsSQL.RecordCount
                If Not mrsSQL!内容 Like "--*" Then
                    strSQL = "Update zltools.zlrptsqls Set 内容=" & strSQLContent & " Where 源ID=" & lng源id & " And 行号=" & mrsSQL!行号
                    gcnOracle.Execute strSQL
                End If
                mrsSQL.MoveNext
            Next
        Else
            strSQLContent = "Replace(明细SQL, '" & strTabOld & "','" & strTabNew & "')"
            strSQL = "Update zltools.zlRPTPars Set 对象=" & strObj & ",明细SQL=" & strSQLContent & " Where 源ID=" & lng源id & " And 序号=" & lng参数号
            gcnOracle.Execute strSQL
        End If
        
        '仅报告单没有程序ID,功能为空的是异常数据
        '同一报表的不同数据源，如果有不同的表，则插入
        If str程序id <> "" And str功能 <> "" Then
            If Not blnExists Then
                If blnNewPrivs Then
                    strSQL = "Insert into zltools.zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(" & IIf(str系统 = "", "Null", str系统) & "," & str程序id & ",'" & str功能 & "','" & strOwner & "','" & strTabNew & "','SELECT')"
                Else
                    strSQL = "Update zltools.zlProgPrivs Set 对象='" & strTabNew & "' Where Nvl(系统,0)=" & IIf(str系统 = "", "0", str系统) & " And 序号=" & str程序id & " And 功能='" & str功能 & "' And 对象='" & strTabOld & "'"
                End If
                gcnOracle.Execute strSQL
                If Not rsRole Is Nothing Then
                    For i = 1 To rsRole.RecordCount
                        strSQL = "Grant Select on " & strOwner & "." & strTabNew & " to " & rsRole!角色
                        gcnOracle.Execute strSQL
                        rsRole.MoveNext
                    Next
                End If
            Else
                If blnDel Then
                    strSQL = "Delete zltools.zlProgPrivs Where Nvl(系统,0)=" & IIf(str系统 = "", "0", str系统) & " And 序号=" & str程序id & " And 功能='" & str功能 & "' And 对象='" & strTabOld & "'"
                    gcnOracle.Execute strSQL
                End If
            End If
            
        End If
        
        strSQL = "Update zltools.Zlrptadjustlog Set 更改=" & Choose(lngNewMode + 1, 0, 1, 2) & " Where 报表ID=" & lng报表id & " And 数据源='" & str源 & "' And Nvl(序号,-1)=" & lng参数号
        gcnOracle.Execute strSQL
    
    gcnOracle.CommitTrans: blnTrans = False
    
    Call ShowReportSQL(lng源id, lng参数号)
    
    rptList.Rows(mlngCurRow).Record(2).Value = lngNewMode
    rptList.Rows(mlngCurRow).Record(2).Caption = Choose(lngNewMode + 1, "全部", "门诊", "住院")
    rptList.Populate
    
    Call ShowStatusInfor("[" & rptList.Rows(mlngCurRow).Record(0).Caption & "]保存成功！")
    cmdSave.Tag = ""
    SaveData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    If blnTrans Then gcnOracle.RollbackTrans

End Function


Private Function ExistTablePrivs(str系统 As String, str程序id As String, str功能 As String, strTab As String) As Boolean
'功能：判断是否存在新使用的表的权限数据
    Dim strSQL As String
    Dim rstmp As ADODB.Recordset
    
    strSQL = "Select 1 From zltools.zlprogprivs Where 系统=[1] And 序号=[2] And 功能=[3] And 对象=[4]"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "读取对象", str系统, str程序id, str功能, strTab)
    ExistTablePrivs = rstmp.RecordCount > 0
End Function

Private Function ExistOtherTablePrivs(strTabOld As String) As Boolean
'功能：判断当前报表的其它数据源或参数是否涉及到与该数据源不同的表权限
    Dim lngStart As Long, lngLast As Long, strSQL As String
    Dim rstmp As ADODB.Recordset
    Dim i As Long, arrtmp As Variant, lng源id As Long, lng参数号 As Long
    
    If mlngCurRow > 2 Then  '在最开始处,第0行是系统，第1行是报表，第二行是数据源
        For i = mlngCurRow To 0 Step -1
            If rptList.Rows(i).Childs.Count > 0 Then
                lngStart = i + 1
                Exit For
            End If
        Next
    Else
        lngStart = mlngCurRow
    End If
    If mlngCurRow < rptList.Rows.Count - 1 Then
        lngLast = rptList.Rows.Count - 1
        For i = mlngCurRow To rptList.Rows.Count - 1
            If rptList.Rows(i).Childs.Count > 0 Then
                lngLast = i - 1
                Exit For
            End If
        Next
    Else
        lngLast = mlngCurRow
    End If
    
    On Error GoTo errH
    For i = lngStart To lngLast
        If i <> mlngCurRow Then
            If rptList.Rows(i).Record(2).Value <> -1 Then   '还没有改的数据源不检查
                arrtmp = Split(rptList.Rows(i).Record(0).Value, "|SP|")
                lng源id = Get源ID(arrtmp(0), arrtmp(1))
                lng参数号 = Val(arrtmp(2))
                
                If lng参数号 = -1 Then
                    strSQL = "Select 对象 From zltools.zlrptdatas Where id=[1]"
                Else
                    strSQL = "Select 对象 From zltools.zlRPTPars Where 源id=[1] And 序号=[2]"
                End If
                On Error GoTo errH
                Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "读取对象", lng源id, lng参数号)
                If rstmp.RecordCount > 0 Then
                    If rstmp!对象 Like "*" & strTabOld & "*" Then
                        ExistOtherTablePrivs = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetParsObj(ByVal lng源id As Long, ByVal lng参数号 As Long, ByVal lngSys As Long) As RPTPars
    Dim tmpPar As RPTPar, j As Long
    Dim rsPar As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    Set GetParsObj = New RPTPars
    strSQL = "Select * From zlRPTPars Where 源id=[1]"
    Set rsPar = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng源id)
    
    For j = 1 To rsPar.RecordCount
        Set tmpPar = New RPTPar
        tmpPar.组名 = Nvl(rsPar!组名)
        tmpPar.序号 = Nvl(rsPar!序号, 0)
        tmpPar.名称 = Nvl(rsPar!名称)
        tmpPar.类型 = Nvl(rsPar!类型, 0)
        tmpPar.缺省值 = Nvl(rsPar!缺省值)
        tmpPar.格式 = Nvl(rsPar!格式, 0)
        
        tmpPar.值列表 = Nvl(rsPar!值列表)
        tmpPar.分类SQL = Replace(Nvl(rsPar!分类SQL), "[系统]", lngSys)
        tmpPar.明细SQL = Replace(Nvl(rsPar!明细SQL), "[系统]", lngSys)
        tmpPar.分类字段 = Nvl(rsPar!分类字段)
        tmpPar.明细字段 = Nvl(rsPar!明细字段)
        tmpPar.对象 = Nvl(rsPar!对象)
        
        '！！！以参数序号为关键字加入集合
        GetParsObj.Add tmpPar.组名, tmpPar.序号, tmpPar.名称, tmpPar.类型, tmpPar.缺省值, tmpPar.格式, tmpPar.值列表, tmpPar.分类SQL, tmpPar.明细SQL, tmpPar.分类字段, tmpPar.明细字段, tmpPar.对象, "_" & tmpPar.序号
        
        rsPar.MoveNext
    Next
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSQLObj(ByVal lng源id As Long, ByVal lng参数号 As Long) As String
    Dim rstmp As ADODB.Recordset, strSQL As String
    
    If lng参数号 = -1 Then
        strSQL = "Select 对象 From zltools.zlrptdatas Where id=[1]"
    Else
        strSQL = "Select 对象 From zltools.zlRPTPars Where 源id=[1] And 序号=[2]"
    End If
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "读取所有者", lng源id, lng参数号)
    If rstmp.RecordCount > 0 Then
        GetSQLObj = "" & rstmp!对象
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckSQLPhrase(ByVal lng源id As Long, ByVal lng参数号 As Long, ByVal lngSys As Long, _
    ByVal strTabOld As String, ByVal strTabNew As String) As Boolean

    Dim strR As String, strFields As String, strOwner As String
    Dim objPars As RPTPars, strSQL As String
    
   
    Set objPars = GetParsObj(lng源id, lng参数号, lngSys)
    strOwner = GetSQLObj(lng源id, lng参数号)
    strOwner = Replace(strOwner, strTabOld, strTabNew)
    
    strSQL = rtbNew.Text
    strSQL = Replace(strSQL, "[系统]", lngSys)
    strSQL = RemoveNote(strSQL)
    strSQL = SQLReplaceOwner(strSQL, strOwner)
    
    If objPars.Count = 0 Then
        strFields = CheckSQL(strSQL, strR)
    Else
        strFields = CheckSQL(strSQL, strR, objPars)
    End If
    If strFields = "" Then
        MsgBox "SQL语句校验失败！" & vbCrLf & vbCrLf & _
            "错误 " & strR & vbCrLf & vbCrLf & _
            "请检查是否正确书写,或参数是否正确设置！", vbInformation, App.Title
        CheckSQLPhrase = False
    Else
        CheckSQLPhrase = True
    End If
       
End Function

Private Function GetReportList(blnTableFull As Boolean, lngDefault As Long, lngModify As Long) As ADODB.Recordset
'功能：获取报表数据源列表
'参数：blnUnModied：只显示未修改的记录,blnTableFull:只显示全表扫描的
    Dim strSQL As String, strIF As String
 
    strIF = IIf(blnTableFull, " And A.全表扫描=1", "")
    strIF = strIF & IIf(lngDefault = -2, "", " And A.缺省 = [1]")
    strIF = strIF & IIf(lngModify = -2, "", " And Nvl(A.更改,-1) = [2]")
    strSQL = "Select *" & vbNewLine & _
            "From (Select A.报表id, A.序号, A.缺省, A.更改, Nvl(E.名称, '共享') 系统名, B.编号, B.名称 报表, C.名称 数据源, Null 参数, B.系统, A.全表扫描" & vbNewLine & _
            "       From Zltools.Zlrptadjustlog A, Zltools.Zlreports B, Zltools.Zlrptdatas C, Zltools.Zlsystems E" & vbNewLine & _
            "       Where A.报表id = B.Id And A.报表id = C.报表id And A.数据源 = C.名称 And B.系统 = E.编号(+)" & strIF & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select A.报表id, A.序号, A.缺省, A.更改, Nvl(E.名称, '共享') 系统名, B.编号, B.名称 报表, C.名称 数据源, D.名称 参数, B.系统, A.全表扫描" & vbNewLine & _
            "       From Zltools.Zlrptadjustlog A, Zltools.Zlreports B, Zltools.Zlrptdatas C, Zltools.Zlrptpars D, Zltools.Zlsystems E" & vbNewLine & _
            "       Where A.报表id = B.Id And A.报表id = C.报表id And A.数据源 = C.名称 And C.Id = D.源id And A.序号 = D.序号 And B.系统 = E.编号(+)" & strIF & ")" & vbNewLine & _
            "Order By 系统, 编号, 数据源, Nvl(序号, 0)"

    On Error GoTo errH
    Set GetReportList = zlDatabase.OpenSQLRecord(strSQL, "读取报表", lngDefault, lngModify)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Unload(Cancel As Integer)
    Set mclsReport = Nothing
    Set mrsSQL = Nothing
End Sub

Private Sub optMode_Click(Index As Integer)
    If mblnUnChange Then Exit Sub
    
    If fraMode.Tag <> CStr(Index) Then
        cmdSave.Tag = "待保存"
        fraMode.Tag = CStr(Index)
        Call SetNewText(True)
    End If
End Sub

Private Sub picLR_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If rptList.Width + x < 1000 Then Exit Sub
        picLR.Left = picLR.Left + x

        rptList.Width = rptList.Width + x
        rtbOld.Left = rptList.Left + rptList.Width + 100
        rtbOld.Width = rtbOld.Width - x
        fraMode.Left = rtbOld.Left
        rtbNew.Left = rtbOld.Left
        rtbNew.Width = rtbOld.Width
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngLeft As Long
    
    lngLeft = Me.ScaleLeft + Me.ScaleWidth - 100
    fraCmd.Left = lngLeft - fraCmd.Width
    
    rptList.Height = Me.ScaleHeight - rptList.Top - 100
    picLR.Left = rptList.Left + rptList.Width
    picLR.Top = rptList.Top
    picLR.Height = rptList.Height
    
    rtbOld.Top = rptList.Top
    rtbOld.Left = rptList.Left + rptList.Width + 100
    rtbOld.Height = (rptList.Height - fraMode.Height - 200) / 2
    rtbOld.Width = lngLeft - rtbOld.Left
    
    fraMode.Top = rtbOld.Top + rtbOld.Height + 100
    fraMode.Left = rtbOld.Left
    
    rtbNew.Left = rtbOld.Left
    rtbNew.Width = rtbOld.Width
    rtbNew.Top = fraMode.Top + fraMode.Height
    rtbNew.Height = rptList.Top + rptList.Height - rtbNew.Top
    
    rtbExplan.Top = rtbOld.Top + rtbOld.Height + 30
    rtbExplan.Left = rtbOld.Left
    rtbExplan.Width = rtbOld.Width
    rtbExplan.Height = rtbNew.Top + rtbNew.Height - rtbExplan.Top

 End Sub

Private Sub ClearSQLText()
'功能：清空当前SQL调整区域数据及状态设置
    rtbOld.Text = ""
    rtbNew.Text = ""
    mblnUnChange = True:    optMode(m全部).Value = True:    mblnUnChange = False
    fraMode.Enabled = False
    
    cmdSave.Enabled = False
    cmdExplan.Enabled = False
End Sub

Private Sub cmdNext_Click()
    Dim lngidx As Long
    If rptList.SelectedRows.Count = 0 Then Exit Sub
    
    lngidx = mlngCurRow
    Do
        If lngidx >= rptList.Rows.Count - 1 Then
            lngidx = 1
        Else
            lngidx = lngidx + 1
        End If
        If rptList.Rows(lngidx).Childs.Count = 0 Then Exit Do
    Loop While 1 = 1
    
    Set rptList.FocusedRow = rptList.Rows(lngidx)
End Sub

Private Sub cmdPrevious_Click()
    Dim lngidx As Long
    If rptList.SelectedRows.Count = 0 Then Exit Sub
    
    lngidx = mlngCurRow
    Do
        If lngidx <= 1 Then
            lngidx = rptList.Rows.Count - 1
        Else
            lngidx = lngidx - 1
        End If
        If rptList.Rows(lngidx).Childs.Count = 0 Then Exit Do
    Loop While 1 = 1
    
    Set rptList.FocusedRow = rptList.Rows(lngidx)
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
        If rptList.SelectedRows(0).Childs.Count > 0 Or mlngCurRow = rptList.Rows.Count - 1 Then KeyCode = 0: cmdNext_Click
    ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
        If rptList.SelectedRows(0).Childs.Count > 0 Then KeyCode = 0: cmdPrevious_Click
    ElseIf KeyCode = vbKeySpace Then
       Call ChangeMode
    ElseIf KeyCode = vbKeyF2 Then
        If cmdSave.Enabled Then Call cmdSave_Click
    End If
End Sub
Private Sub ChangeMode()
    Dim i As Long, j As Long
    If fraMode.Enabled = False Or rtbExplan.Visible Then Exit Sub
    For i = 0 To optMode.UBound
        If optMode(i).Value = True Then
            If i = optMode.UBound Then
                j = 0
            Else
                j = i + 1
            End If
            optMode(j).Value = True
            Exit For
        End If
    Next
End Sub


Private Sub rptList_SelectionChanged()
    Dim lng源id As Long, lng参数号 As Long, arrtmp As Variant
    Dim lng缺省 As Long, lng更改 As Long
    
    If rptList.SelectedRows.Count = 0 Then '未选择时
        Call ClearSQLText
        mlngCurRow = -1
        Exit Sub
    End If
    If mlngCurRow = rptList.SelectedRows(0).Index Then Exit Sub
    
    If rtbExplan.Visible Then rtbExplan.Visible = False: rtbExplan.Text = "": cmdExplan.Caption = "查看计划(&X)"
    If mlngCurRow > 0 And cmdSave.Tag <> "" And mlngCurRow < rptList.Rows.Count Then
        If rptList.Rows(mlngCurRow).Childs.Count = 0 Then
            If CheckModified = False Then
                Set rptList.FocusedRow = rptList.Rows(mlngCurRow)
                rptList.SetFocus
                Exit Sub
            Else
                rptList.SetFocus
            End If
        End If
    End If
    
    Call ShowStatusInfor("")
    fraMode.Tag = ""
    mlngCurRow = rptList.SelectedRows(0).Index
    cmdDesign.Enabled = rptList.SelectedRows(0).Record.Tag <> ""    '选中"系统"行时
    cmdSave.Enabled = cmdDesign.Enabled
    cmdExplan.Enabled = cmdDesign.Enabled
        
    If rptList.SelectedRows(0).Childs.Count > 0 Then
        Call ClearSQLText
    Else
        lng缺省 = Val("" & rptList.SelectedRows(0).Record(1).Value)
        lng更改 = Val("" & rptList.SelectedRows(0).Record(2).Value)
        
        mblnUnChange = True
        If lng更改 = m未改 Then
            optMode(lng缺省).Value = True
        ElseIf lng更改 <> m手工 Then
            optMode(lng更改).Value = True
        End If
        mblnUnChange = False
                
        arrtmp = Split(rptList.SelectedRows(0).Record(0).Value, "|SP|")
        lng源id = Get源ID(arrtmp(0), arrtmp(1))
        lng参数号 = Val(arrtmp(2))
        Call ShowReportSQL(lng源id, lng参数号)
            
        If lng更改 = m未改 Then
            '如果是通过外部工具或报表编辑器修改了数据源，没有填写"更改"字段，则此时填写
            If InStr(rtbOld.Text, "门诊费用记录") > 0 Or InStr(rtbOld.Text, "住院费用记录") > 0 Then
                On Error GoTo errH
                gcnOracle.Execute "Update zltools.Zlrptadjustlog Set 更改=" & m手工 & " Where 源ID=" & lng源id & " And Nvl(序号,-1)=" & lng参数号
                
                rptList.SelectedRows(0).Record(2).Value = m手工
                rptList.SelectedRows(0).Record(2).Caption = "手工"
                lng更改 = m手工
            End If
        End If
        If lng更改 = m手工 Then rtbNew.Text = ""
        
        fraMode.Enabled = lng更改 <> m手工
        cmdSave.Enabled = lng更改 <> m手工
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowReportSQL(lng源id As Long, lng参数号 As Long)
    Dim strSQLText As String
    
    Set mrsSQL = GetRPTSQL(lng源id, lng参数号)
    If mrsSQL Is Nothing Or lng源id = 0 Then
        Call ClearSQLText
    Else
        strSQLText = GetTextByRs(mrsSQL)
        Call SetOldText(strSQLText)
        Call SetNewText(False)
    End If
End Sub

Private Sub SetOldText(ByVal strSQLText As String)
    Dim i As Long, j As Long, lngMode As Long
    
    rtbOld.Text = strSQLText
    j = 1
    Do
        i = InStr(j, rtbOld.Text, "病人费用记录")
        If i <= 0 Then
            i = InStr(j, rtbOld.Text, "门诊费用记录")
            If i <= 0 Then
                i = InStr(j, rtbOld.Text, "住院费用记录")
                If i > 0 Then lngMode = 2
            Else
                lngMode = 1
            End If
            If i <= 0 Then Exit Do
        End If
        
        rtbOld.SelStart = i - 1
        rtbOld.SelLength = 6
        If lngMode = 1 Then
            rtbOld.SelColor = &HC000&   '绿
        ElseIf lngMode = 2 Then
            rtbOld.SelColor = &HFF&     '红
        Else
            rtbOld.SelColor = &H8000000D    '蓝
        End If
        j = i + 6
    Loop While j < Len(rtbOld.Text)
    
    rtbOld.SelStart = 0
    rtbOld.SelLength = 0
End Sub

Private Function GetRPTSQL(ByVal lng源id As Long, ByVal lng参数号 As Long) As ADODB.Recordset
'功能：获取数据源或参数的SQL
'参数：lng参数号:-1表示获取数据源SQL，否则获取参数的SQL
    Dim strSQL As String
    On Error GoTo errH
    
    If lng参数号 = -1 Then
        strSQL = "Select 行号,内容 From zltools.zlRPTSQLs Where 源id = [1] Order By 行号"
        Set GetRPTSQL = zlDatabase.OpenSQLRecord(strSQL, "读取SQL", lng源id)
    Else
        strSQL = "Select 0 as 行号,明细sql as 内容 From zltools.zlRPTPars Where 源id = [1] And 序号 = [2]"
        Set GetRPTSQL = zlDatabase.OpenSQLRecord(strSQL, "读取SQL", lng源id, lng参数号)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetTextByRs(ByRef rstmp As ADODB.Recordset) As String
    Dim i As Long, strTmp As String
    
    If rstmp.RecordCount > 0 Then rstmp.MoveFirst
    For i = 1 To rstmp.RecordCount
        strTmp = IIf(i = 1, "", strTmp & vbNewLine) & rstmp!内容
        rstmp.MoveNext
    Next
    GetTextByRs = strTmp
End Function

Private Sub SetNewText(blnChange As Boolean)
    Dim strTmp As String, lngNewMode As Long, lngOldMode As Long
    Dim i As Long, j As Long, lngLen As Long
    Dim strTabOld As String, strTabNew As String
    
    Call GetTwoMode(lngOldMode, lngNewMode)
        
    strTabOld = Choose(lngOldMode + 1, "病人费用记录", "门诊费用记录", "住院费用记录")
    strTabNew = Choose(lngNewMode + 1, "病人费用记录", "门诊费用记录", "住院费用记录")
        
   
    If lngOldMode = lngNewMode Then
        strTmp = rtbOld.Text
    Else
        strTmp = Replace(rtbOld.Text, strTabOld, strTabNew)
         '床号,病人病区ID,多病人单这几个字段因为其它表可能有同名字段，所以不能自动替换，需要人工修改
    End If
    rtbNew.Text = strTmp
    
    lngLen = Len(strTmp)
    j = 1
    Do
        i = InStr(j, strTmp, strTabNew)
        If i <= 0 Then Exit Do
        
        rtbNew.SelStart = i - 1
        rtbNew.SelLength = Len(strTabNew)
        If lngNewMode = 1 Then
            rtbNew.SelColor = &HC000&   '绿
        ElseIf lngNewMode = 2 Then
            rtbNew.SelColor = &HFF&     '红
        Else
            rtbNew.SelColor = &H8000000D    '蓝
        End If
        j = i + Len(strTabNew)
    Loop While j < lngLen
    
    rtbNew.SelStart = 0
    rtbNew.SelLength = 0
    
    If blnChange Then
        Call ShowStatusInfor("已调整为[" & strTabNew & "]")
    End If
End Sub

Private Sub rtbNew_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call ChangeMode
    ElseIf KeyCode = vbKeyF2 Then
        If cmdSave.Enabled Then Call cmdSave_Click
    End If
End Sub

Private Sub rtbOld_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call ChangeMode
    ElseIf KeyCode = vbKeyF2 Then
        If cmdSave.Enabled Then Call cmdSave_Click
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim lngStart As Long, i As Long
        
        If txtFind.Tag <> txtFind.Text Or mlngCurRow < 0 Then
            txtFind.Tag = txtFind.Text
            lngStart = 1
        Else
            '查找下一个
            lngStart = rptList.Rows(mlngCurRow).Index
        End If
        
        For i = lngStart To rptList.Rows.Count - 1
            With rptList.Rows(i)
                If .Childs.Count > 0 Then
                    
                    If .Record.Item(0).Caption Like "*" & txtFind.Text & "*" Then
                        If i + 1 < rptList.Rows.Count - 2 Then i = i + 1
                        Set rptList.FocusedRow = rptList.Rows(i)
                        
                        Exit For
                    End If
                End If
            End With
        Next
    End If
End Sub
