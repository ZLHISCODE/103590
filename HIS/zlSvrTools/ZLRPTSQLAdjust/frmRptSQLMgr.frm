VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRptSQLMgr 
   BackColor       =   &H80000005&
   Caption         =   "����SQL��������"
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
      ToolTipText     =   "�����¼��ı䵱ǰ��"
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
         Caption         =   "��ȱʡ��ʽ����ȫ������"
         Height          =   350
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   2160
      End
      Begin VB.CommandButton cmdDesign 
         Caption         =   "�������(&D)"
         Height          =   350
         Left            =   2400
         TabIndex        =   8
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�������(&S)"
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
         Caption         =   "ȫ������(&0)"
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
         Caption         =   "סԺ����(&2)"
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
         Caption         =   "�������(&1)"
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
         Caption         =   "(���ո��л�)"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4850
         TabIndex        =   20
         Top             =   30
         Width           =   1335
      End
      Begin VB.Label lblMode 
         BackColor       =   &H80000005&
         Caption         =   "����Ϊ"
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
            Key             =   "ǩ��"
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
         Caption         =   "�鿴ִ�мƻ�(&X)"
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
         Caption         =   "����ȫ��ɨ���"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton checkExplan 
         Caption         =   "���ȫ��ɨ��"
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
         Caption         =   "����"
         Height          =   255
         Left            =   1900
         TabIndex        =   24
         Top             =   450
         Width           =   375
      End
      Begin VB.Label lblDefault 
         BackColor       =   &H80000005&
         Caption         =   "ȱʡ"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   450
         Width           =   375
      End
      Begin VB.Label lblFind 
         BackColor       =   &H80000005&
         Caption         =   "����λ"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   75
         Width           =   735
      End
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "�����г�����SQL�к��С����˷��ü�¼��������Դ����ѡ����ķ�ʽ��ִ��""�������""(F2)�������Ҫ��������Դ������SQL����ִ��""�������""��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����SQL����"
      BeginProperty Font 
         Name            =   "����"
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
Public Event StatusTextUpdate(ByVal strMSG As String) 'Ҫ�����������״̬������

Private mrsSQL As ADODB.Recordset
Private mlngCurRow As Long
Private mblnUnChange As Boolean
Private mclsReport As clsReport
Private Enum mmode
    mδ�� = -1
    mȫ�� = 0
    m���� = 1
    mסԺ = 2
    m�ֹ� = 3   'ʹ�ñ�������������˸���
End Enum
Private mlngDBVer As Long


Private Sub ShowStatusInfor(ByVal strMSG As String)
    RaiseEvent StatusTextUpdate(strMSG)
End Sub

Private Function GetSQLPlan(lngԴid As Long, lng������ As Long, lngSys As Long) As ADODB.Recordset
    Dim strSQL As String, rstmp As ADODB.Recordset
    Dim strOwner As String, strSID As String
    Dim objPars As RPTPars
    
    Set objPars = GetParsObj(lngԴid, lng������, lngSys)
    strOwner = GetSQLObj(lngԴid, lng������)
    
    Set rstmp = GetRPTSQL(lngԴid, lng������)
    strSQL = GetTextByRs(rstmp)
    strSQL = Replace(strSQL, "[ϵͳ]", lngSys)
    strSQL = RemoveNote(strSQL)
    strSQL = SQLReplaceOwner(strSQL, strOwner)
    
    If objPars.Count = 0 Then
        strSQL = GetExecSQL(strSQL)
    Else
        strSQL = GetExecSQL(strSQL, objPars)
    End If
        
    If strSQL <> "" Then
        On Error Resume Next
        strSID = lngԴid & Time()
          
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
            Set GetSQLPlan = zlDatabase.OpenSQLRecord(strSQL, "ִ�мƻ�", strSID)
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
    Dim lngԴid As Long, strԴ As String, lng������ As Long, lng����id As Long, arrtmp As Variant, lngSys As Long
    
    If rtbExplan.Visible = False And mlngCurRow > -1 Then
        If rptList.Rows(mlngCurRow).Childs.Count = 0 Then
            arrtmp = Split(rptList.Rows(mlngCurRow).Record(0).Value, "|SP|")
            strԴ = arrtmp(1)
            lng����id = arrtmp(0)
            lngԴid = GetԴID(lng����id, strԴ)
            lng������ = Val(arrtmp(2))
            arrtmp = Split(rptList.Rows(mlngCurRow).Record.Tag, "|SP|")
            lngSys = Val(arrtmp(0))
        
            Set rstmp = GetSQLPlan(lngԴid, lng������, lngSys)
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
            rtbExplan.SelColor = &HFF&     '��
            
            j = i + Len(strFind)
        Loop While j < lngLen
        
        rtbExplan.SelStart = 0
        rtbExplan.SelLength = 0
    End If
    
    If rtbExplan.Visible Then
        cmdExplan.Caption = "�鿴SQL(&X)"
    Else
        cmdExplan.Caption = "�鿴�ƻ�(&X)"
    End If
    rptList.SetFocus
End Sub


Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdExplan_Click
End Sub

Private Sub checkExplan_Click()
    Dim strSQL As String, i As Long, j As Long, k As Long, blnHave As Boolean
    Dim rstmp As ADODB.Recordset
    Dim lngԴid As Long, strԴ As String, lng������ As Long, lng����id As Long, arrtmp As Variant, lngSys As Long
    Dim blnUpdaterow As Boolean
    
    blnUpdaterow = checkExplan.Tag = "update"
            
    If rptList.Rows.Count < 1 Then Exit Sub
    On Error Resume Next
    
    If blnUpdaterow = False Then
        If MsgBox("����б��嵥����������Դ��ִ�мƻ������ִ�мƻ��д��ڶԡ����˷��ü�¼������ȫ��ɨ����С��˲���������Ҫ1��3���ӣ���ȷ��Ҫ������", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
        
        gcnOracle.Execute "Update zltools.zlrptadjustlog Set ȫ��ɨ��=Null"
    End If
    
    DoEvents    'Ϊ�˼�ʱ��ʾ״̬����ʾ
    '��0�в��ü��
    For j = 1 To rptList.Rows.Count - 1
        If rptList.Rows(j).Childs.Count = 0 Then
            If blnUpdaterow = False Then
                Call ShowStatusInfor("��" & rptList.Rows.Count & "��,���ڼ���" & j + 1 & "��")
            End If
            
            If blnUpdaterow And mlngCurRow = j Or blnUpdaterow = False Then
                arrtmp = Split(rptList.Rows(j).Record(0).Value, "|SP|")
                strԴ = arrtmp(1)
                lng����id = arrtmp(0)
                lngԴid = GetԴID(lng����id, strԴ)
                lng������ = Val(arrtmp(2))
                arrtmp = Split(rptList.Rows(j).Record.Tag, "|SP|")
                lngSys = Val(arrtmp(0))
            
                Set rstmp = GetSQLPlan(lngԴid, lng������, lngSys)
                rstmp.Filter = "Plan_Table_Output Like '*TABLE ACCESS FULL*'"
                blnHave = False
                For i = 1 To rstmp.RecordCount
                    If rstmp!Plan_Table_Output Like "*���˷��ü�¼*" Or rstmp!Plan_Table_Output Like "*סԺ���ü�¼*" Or rstmp!Plan_Table_Output Like "*������ü�¼*" Then
                        blnHave = True
                        Exit For
                    End If
                    rstmp.MoveNext
                Next
                If blnHave Then
                    gcnOracle.Execute "Update zltools.zlrptadjustlog Set ȫ��ɨ��=1 Where ����id=" & lng����id & " And ����Դ='" & strԴ & "' And Nvl(���,-1)=" & lng������
                    k = k + 1
                    rptList.Rows(j).Record(3).Caption = "��"
                Else
                    rptList.Rows(j).Record(3).Caption = ""
                End If
            End If
        End If
    Next
    rptList.Populate
    
    If blnUpdaterow = False Then
        chkOnlyTableFull.Visible = k > 0
        
        Call ShowStatusInfor("��" & k & "������Դ���ڶԡ����˷��ü�¼�����ȫ��ɨ�衣")
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
    Dim arrtmp As Variant, lngԴid As Long, lng������ As Long
    
    'ѡ��"ϵͳ"��ʱ�ѽ����˴˰�ť
    If rptList.SelectedRows.Count = 0 Then Exit Sub
    
    If mclsReport Is Nothing Then
        Set mclsReport = New clsReport
        Call mclsReport.InitOracle(gcnOracle)
    End If
    strNO = rptList.Rows(mlngCurRow).Record.Tag
    lngSys = Val(Split(strNO, "|SP|")(0))
    strNO = Split(strNO, "|SP|")(1)
        
    Call mclsReport.ReportDesign(gcnOracle, lngSys, strNO, frmMDIMain, True)
    
    '����޸������ݣ�����д�Զ����޸ı��,ѡ�б�����ʱ������
    If rptList.Rows(mlngCurRow).Childs.Count = 0 Then
         arrtmp = Split(rptList.Rows(mlngCurRow).Record(0).Value, "|SP|")
         lngԴid = GetԴID(arrtmp(0), arrtmp(1))
         lng������ = Val(arrtmp(2))
        
         Set mrsSQL = GetRPTSQL(lngԴid, lng������)
         If mrsSQL.RecordCount = 0 Then  '�޸�����Դ���ɾ���������²��������ɾ���ˣ����¼��ر����б�
             Call LoadReportList
             
             '���±��ִ�мƻ�
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
            lngModify = -1 '���һ����ʾ"δ��"
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
            If strOldSys <> rstmp!ϵͳ�� Then
                strOldSys = rstmp!ϵͳ��: strOldRpt = ""
                Set objsys = .Records.Add()
                objsys.Expanded = True
                
                Set objItem = objsys.AddItem(Val("" & rstmp!ϵͳ))
                objItem.Caption = strOldSys
                objItem.BackColor = &HFFC0C0      'frmMDIMain.cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
                objItem.ForeColor = &HFF0000
            End If
            
            If strOldRpt <> rstmp!��� Then
                strOldRpt = rstmp!���
                Set objrpt = objsys.Childs.Add()
                objrpt.Expanded = True
                objrpt.Tag = rstmp!ϵͳ & "|SP|" & rstmp!���
                
                Set objItemRpt = objrpt.AddItem(Val(rstmp!����ID))
                objItemRpt.Caption = rstmp!��� & ":" & rstmp!����
                objItemRpt.BackColor = &HC0FFFF      '&HFFC0C0      'frmMDIMain.cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
            End If
                
            
            Set objdata = objrpt.Childs.Add()
            objdata.Tag = rstmp!ϵͳ & "|SP|" & rstmp!���
            If Not IsNull(rstmp!����) Then
                Set objItemData = objdata.AddItem(rstmp!����ID & "|SP|" & rstmp!����Դ & "|SP|" & rstmp!���)
                objItemData.Caption = rstmp!���� & "(" & rstmp!����Դ & "�Ĳ���)"
                objItemData.ForeColor = &H40C0&
            Else
                objdata.Expanded = True
                Set objItemData = objdata.AddItem(rstmp!����ID & "|SP|" & rstmp!����Դ & "|SP|" & "-1")
                objItemData.Caption = rstmp!����Դ
            End If
            Set objItemData = objdata.AddItem(Val("" & rstmp!ȱʡ))
            objItemData.Caption = Choose(Val("" & rstmp!ȱʡ) + 1, "ȫ��", "����", "סԺ")
            Set objItemData = Nothing
            
            Set objItemData = objdata.AddItem(IIf(IsNull(rstmp!����), -1, Val("" & rstmp!����)))
            objItemData.Caption = IIf(IsNull(rstmp!����), " ", Choose(Val("" & rstmp!����) + 1, "ȫ��", "����", "סԺ"))
            
            Set objItemData = objdata.AddItem("")   '�Ƿ�ȫ��ɨ��
            objItemData.Caption = IIf(IsNull(rstmp!ȫ��ɨ��), " ", "��")
            If Not IsNull(rstmp!ȫ��ɨ��) Then blnHaveTabFull = True
            
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
    
    Call ShowStatusInfor("��" & rstmp.RecordCount & "����¼(����Դ�ʹ�SQLѡ�����Ĳ���)��")
    LoadReportList = True
End Function

Private Function CheckModified() As Boolean
    Dim i As Long, lngNewMode As Long, lngOldMode As Long
        
    Call GetTwoMode(lngOldMode, lngNewMode)
    CheckModified = True
    If lngOldMode <> lngNewMode Then
        i = 0
        If MsgBox("��ǰ����δ���棬�Ƿ��Զ�����?", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            CheckModified = SaveData
        Else
            cmdSave.Tag = ""
        End If
    End If
End Function


Private Sub Form_Load()
    Dim strHeadStr As String
    
    mblnUnChange = True
    
    cboDefault.AddItem "0-����"
    cboDefault.AddItem "1-����"
    cboDefault.AddItem "2-סԺ"
    cboDefault.AddItem "3-ȫ������"
    cboDefault.ListIndex = 0
    
    cboModify.AddItem "0-����"
    cboModify.AddItem "1-����"
    cboModify.AddItem "2-סԺ"
    cboModify.AddItem "3-ȫ������"
    cboModify.AddItem "4-δ��"
    cboModify.ListIndex = 0
    
    mblnUnChange = False
    
    mlngCurRow = -1
    mblnUnChange = False
    strHeadStr = "����,230;ȱʡ,30;����,30;ȫ��ɨ��,54"
    
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
'            .NoGroupByText = "�϶��б��⵽����,�����з���..."
'            .NoItemsText = "û���ҵ����������Ĳ���..."
'            .VerticalGridStyle = xtpGridSolid
'        End With

        .Columns(0).TreeColumn = True
        
'        .PaintManager.TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
'        .GroupsOrder.Add .Columns(0)
    End With
End Sub


Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub subPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

Public Sub RefreshList()
    Call LoadReportList
End Sub

Private Sub cmdSaveAll_Click()
    Dim lngidx As Long, lngȱʡ As Long, k As Long
    Dim strErr As String
    
    If rptList.Rows.Count < 2 Then Exit Sub
    If MsgBox("��ȷ��Ҫ�Ե�ǰ�б��е��������ݰ�ȱʡ��ʽ���и�����", vbQuestion + vbOKCancel, Me.Caption) = vbCancel Then
        Exit Sub
    End If
    
    lngidx = 1
    Do
        If rptList.Rows(lngidx).Childs.Count = 0 Then
            Set rptList.FocusedRow = rptList.Rows(lngidx)   '�����¼�rptList_SelectionChanged
            lngȱʡ = Val("" & rptList.Rows(lngidx).Record(1).Value)
            optMode(lngȱʡ).Value = True   '����click�¼�
            
            If SaveData = False Then
                If MsgBox("��" & lngidx & "�и���ʧ�ܣ��Ƿ����������һ������?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
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
        MsgBox "���±������ʧ�ܣ�����(�ɰ�Ctrl+C������ʾ��Ϣ)" & vbCrLf & strErr
    End If
    
    Call ShowStatusInfor("��������" & k & "�����ݡ�")
End Sub
Private Sub cmdSave_Click()
    If SaveData Then
        If mlngCurRow > 0 Then
            Set rptList.FocusedRow = rptList.Rows(mlngCurRow)
            Call rptList.SetFocus
        End If
    End If
End Sub
Private Function GetPrivsData(ByVal lng����id As Long, strϵͳ As String, str����id As String, str���� As String) As Boolean
    Dim rstmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select ϵͳ, ����id, ����" & vbNewLine & _
        "From zltools.zlReports" & vbNewLine & _
        "Where ID = [1] And ����id Is Not Null" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select ϵͳ, ����id, ����" & vbNewLine & _
        "From zltools.zlRPTPuts" & vbNewLine & _
        "Where ����id = [1]" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select B.ϵͳ, B.����id, A.���� From zltools.zlRPTSubs A, zltools.zlRPTGroups B Where A.��id = B.Id And A.����id = [1]"
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡȨ������", lng����id)
    If rstmp.RecordCount > 0 Then
        strϵͳ = "" & rstmp!ϵͳ
        str����id = "" & rstmp!����id
        str���� = "" & rstmp!����
    End If
    
    GetPrivsData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOwner(lngԴid As Long, lng������ As Long, strTab As String) As String
    Dim rstmp As ADODB.Recordset, strSQL As String
    Dim arrtmp As Variant, i As Long, p As Long
    
    If lng������ = -1 Then
        strSQL = "Select ���� From zltools.zlrptdatas Where id=[1]"
    Else
        strSQL = "Select a.���� From zltools.zlRPTPars Where a.Դid=[1] And a.���=[2]"
    End If
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������", lngԴid, lng������)
    If rstmp.RecordCount > 0 Then
        arrtmp = Split(rstmp!����, ",")
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

Private Function GetԴID(ByVal lng����id As Long, ByVal strԴ As String) As Long
    Dim rstmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select ID From zltools.zlrptdatas Where ����id=[1] And ����=[2]"
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԴid", lng����id, strԴ)
    If rstmp.RecordCount > 0 Then GetԴID = rstmp!Id
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetTwoMode(ByRef lngOldMode As Long, ByRef lngNewMode As Long)
    Dim i As Long, strSQL As String
    
    strSQL = RemoveNote(rtbOld.Text)
    i = InStr(strSQL, "���˷��ü�¼")
    If i <= 0 Then
        i = InStr(strSQL, "������ü�¼")
        If i <= 0 Then
            i = InStr(strSQL, "סԺ���ü�¼")
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
    'û��ѡ���У���ѡ����ǡ�ϵͳ���򡰱�����ʱ���Լ����ⲿ�޸Ĺ��ļ�¼���ǽ����˴˰�ť��
    Dim arrtmp As Variant, lng������ As Long, lngԴid As Long, strԴ As String
    Dim blnTrans As Boolean, strSQL As String, strObj As String, strSQLContent As String
    Dim i As Long, lngNewMode As Long, lngOldMode As Long
    Dim strTabOld As String, strTabNew As String
    Dim strϵͳ As String, str����id As String, str���� As String
    Dim lng����id As Long, strOwner As String, blnExists As Boolean, blnNewPrivs As Boolean, blnDel As Boolean
    Dim rsRole As ADODB.Recordset
        
    arrtmp = Split(rptList.Rows(mlngCurRow).Record(0).Value, "|SP|")
    lng����id = Val(arrtmp(0))
    strԴ = arrtmp(1)
    lng������ = Val(arrtmp(2))
    
    lngԴid = GetԴID(lng����id, strԴ)
    If lngԴid = 0 Then Exit Function
    
    Call GetTwoMode(lngOldMode, lngNewMode)
    If lngOldMode = lngNewMode Then
        If Trim(rptList.Rows(mlngCurRow).Record(2).Caption) = "" Then
            strSQL = "Update zltools.Zlrptadjustlog Set ����=" & Choose(lngNewMode + 1, 0, 1, 2) & " Where ����ID=" & lng����id & " And ����Դ='" & strԴ & "' And Nvl(���,-1)=" & lng������
            gcnOracle.Execute strSQL
            rptList.Rows(mlngCurRow).Record(2).Value = lngNewMode
            rptList.Rows(mlngCurRow).Record(2).Caption = Choose(lngNewMode + 1, "ȫ��", "����", "סԺ")
            rptList.Populate
        Else
            Call ShowStatusInfor("��ǰ����δ���ģ����豣�档")
        End If
        SaveData = True
        Exit Function
    End If
    strTabOld = Choose(lngOldMode + 1, "���˷��ü�¼", "������ü�¼", "סԺ���ü�¼")
    strTabNew = Choose(lngNewMode + 1, "���˷��ü�¼", "������ü�¼", "סԺ���ü�¼")
    
    If GetPrivsData(lng����id, strϵͳ, str����id, str����) = False Then
        Exit Function
    End If
    
    
    'ִ��SQL����﷨
    If CheckSQLPhrase(lngԴid, lng������, Val(strϵͳ), strTabOld, strTabNew) = False Then
        Exit Function
    End If
        
    If str����id <> "" Then '�����浥û�г���ID
        strOwner = GetOwner(lngԴid, lng������, strTabOld)
        If strOwner = "" Then
            Call ShowStatusInfor("����ʧ�ܣ�δ�ҵ�SQL����������ߡ�")
            Exit Function
        End If
        
        blnExists = ExistTablePrivs(strϵͳ, str����id, str����, strTabNew)
        If blnExists = False Then
            blnNewPrivs = ExistOtherTablePrivs(strTabOld)
        Else
            blnDel = Not ExistOtherTablePrivs(strTabOld)
        End If
    End If
    
    strObj = "Replace(����, '" & strTabOld & "','" & strTabNew & "')"
    
    If str����id <> "" And str���� <> "" Then
        If Not blnExists Then
            strSQL = "Select ��ɫ" & vbNewLine & _
                    "From zltools.zlRoleGrant A" & vbNewLine & _
                    "Where Nvl(ϵͳ, 0) = [1] And ��� = [2] And ���� = [3] And Exists (Select 1 From dba_Role_Privs B Where A.��ɫ = B.Granted_Role)"
            On Error Resume Next    '���û��Ȩ�޷���dba_Role_Privs������Ȩ
            Set rsRole = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strϵͳ), str����id, str����)
            Err.Clear
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        
        If lng������ = -1 Then
            strSQL = "Update zltools.zlRPTDatas Set ����=" & strObj & " Where ID=" & lngԴid
            gcnOracle.Execute strSQL
            
            strSQLContent = "Replace(����, '" & strTabOld & "','" & strTabNew & "')"
            mrsSQL.Filter = "���� like '*" & strTabOld & "*'"
            For i = 1 To mrsSQL.RecordCount
                If Not mrsSQL!���� Like "--*" Then
                    strSQL = "Update zltools.zlrptsqls Set ����=" & strSQLContent & " Where ԴID=" & lngԴid & " And �к�=" & mrsSQL!�к�
                    gcnOracle.Execute strSQL
                End If
                mrsSQL.MoveNext
            Next
        Else
            strSQLContent = "Replace(��ϸSQL, '" & strTabOld & "','" & strTabNew & "')"
            strSQL = "Update zltools.zlRPTPars Set ����=" & strObj & ",��ϸSQL=" & strSQLContent & " Where ԴID=" & lngԴid & " And ���=" & lng������
            gcnOracle.Execute strSQL
        End If
        
        '�����浥û�г���ID,����Ϊ�յ����쳣����
        'ͬһ����Ĳ�ͬ����Դ������в�ͬ�ı������
        If str����id <> "" And str���� <> "" Then
            If Not blnExists Then
                If blnNewPrivs Then
                    strSQL = "Insert into zltools.zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(" & IIf(strϵͳ = "", "Null", strϵͳ) & "," & str����id & ",'" & str���� & "','" & strOwner & "','" & strTabNew & "','SELECT')"
                Else
                    strSQL = "Update zltools.zlProgPrivs Set ����='" & strTabNew & "' Where Nvl(ϵͳ,0)=" & IIf(strϵͳ = "", "0", strϵͳ) & " And ���=" & str����id & " And ����='" & str���� & "' And ����='" & strTabOld & "'"
                End If
                gcnOracle.Execute strSQL
                If Not rsRole Is Nothing Then
                    For i = 1 To rsRole.RecordCount
                        strSQL = "Grant Select on " & strOwner & "." & strTabNew & " to " & rsRole!��ɫ
                        gcnOracle.Execute strSQL
                        rsRole.MoveNext
                    Next
                End If
            Else
                If blnDel Then
                    strSQL = "Delete zltools.zlProgPrivs Where Nvl(ϵͳ,0)=" & IIf(strϵͳ = "", "0", strϵͳ) & " And ���=" & str����id & " And ����='" & str���� & "' And ����='" & strTabOld & "'"
                    gcnOracle.Execute strSQL
                End If
            End If
            
        End If
        
        strSQL = "Update zltools.Zlrptadjustlog Set ����=" & Choose(lngNewMode + 1, 0, 1, 2) & " Where ����ID=" & lng����id & " And ����Դ='" & strԴ & "' And Nvl(���,-1)=" & lng������
        gcnOracle.Execute strSQL
    
    gcnOracle.CommitTrans: blnTrans = False
    
    Call ShowReportSQL(lngԴid, lng������)
    
    rptList.Rows(mlngCurRow).Record(2).Value = lngNewMode
    rptList.Rows(mlngCurRow).Record(2).Caption = Choose(lngNewMode + 1, "ȫ��", "����", "סԺ")
    rptList.Populate
    
    Call ShowStatusInfor("[" & rptList.Rows(mlngCurRow).Record(0).Caption & "]����ɹ���")
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


Private Function ExistTablePrivs(strϵͳ As String, str����id As String, str���� As String, strTab As String) As Boolean
'���ܣ��ж��Ƿ������ʹ�õı��Ȩ������
    Dim strSQL As String
    Dim rstmp As ADODB.Recordset
    
    strSQL = "Select 1 From zltools.zlprogprivs Where ϵͳ=[1] And ���=[2] And ����=[3] And ����=[4]"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", strϵͳ, str����id, str����, strTab)
    ExistTablePrivs = rstmp.RecordCount > 0
End Function

Private Function ExistOtherTablePrivs(strTabOld As String) As Boolean
'���ܣ��жϵ�ǰ�������������Դ������Ƿ��漰���������Դ��ͬ�ı�Ȩ��
    Dim lngStart As Long, lngLast As Long, strSQL As String
    Dim rstmp As ADODB.Recordset
    Dim i As Long, arrtmp As Variant, lngԴid As Long, lng������ As Long
    
    If mlngCurRow > 2 Then  '���ʼ��,��0����ϵͳ����1���Ǳ����ڶ���������Դ
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
            If rptList.Rows(i).Record(2).Value <> -1 Then   '��û�иĵ�����Դ�����
                arrtmp = Split(rptList.Rows(i).Record(0).Value, "|SP|")
                lngԴid = GetԴID(arrtmp(0), arrtmp(1))
                lng������ = Val(arrtmp(2))
                
                If lng������ = -1 Then
                    strSQL = "Select ���� From zltools.zlrptdatas Where id=[1]"
                Else
                    strSQL = "Select ���� From zltools.zlRPTPars Where Դid=[1] And ���=[2]"
                End If
                On Error GoTo errH
                Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", lngԴid, lng������)
                If rstmp.RecordCount > 0 Then
                    If rstmp!���� Like "*" & strTabOld & "*" Then
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

Private Function GetParsObj(ByVal lngԴid As Long, ByVal lng������ As Long, ByVal lngSys As Long) As RPTPars
    Dim tmpPar As RPTPar, j As Long
    Dim rsPar As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    Set GetParsObj = New RPTPars
    strSQL = "Select * From zlRPTPars Where Դid=[1]"
    Set rsPar = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngԴid)
    
    For j = 1 To rsPar.RecordCount
        Set tmpPar = New RPTPar
        tmpPar.���� = Nvl(rsPar!����)
        tmpPar.��� = Nvl(rsPar!���, 0)
        tmpPar.���� = Nvl(rsPar!����)
        tmpPar.���� = Nvl(rsPar!����, 0)
        tmpPar.ȱʡֵ = Nvl(rsPar!ȱʡֵ)
        tmpPar.��ʽ = Nvl(rsPar!��ʽ, 0)
        
        tmpPar.ֵ�б� = Nvl(rsPar!ֵ�б�)
        tmpPar.����SQL = Replace(Nvl(rsPar!����SQL), "[ϵͳ]", lngSys)
        tmpPar.��ϸSQL = Replace(Nvl(rsPar!��ϸSQL), "[ϵͳ]", lngSys)
        tmpPar.�����ֶ� = Nvl(rsPar!�����ֶ�)
        tmpPar.��ϸ�ֶ� = Nvl(rsPar!��ϸ�ֶ�)
        tmpPar.���� = Nvl(rsPar!����)
        
        '�������Բ������Ϊ�ؼ��ּ��뼯��
        GetParsObj.Add tmpPar.����, tmpPar.���, tmpPar.����, tmpPar.����, tmpPar.ȱʡֵ, tmpPar.��ʽ, tmpPar.ֵ�б�, tmpPar.����SQL, tmpPar.��ϸSQL, tmpPar.�����ֶ�, tmpPar.��ϸ�ֶ�, tmpPar.����, "_" & tmpPar.���
        
        rsPar.MoveNext
    Next
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSQLObj(ByVal lngԴid As Long, ByVal lng������ As Long) As String
    Dim rstmp As ADODB.Recordset, strSQL As String
    
    If lng������ = -1 Then
        strSQL = "Select ���� From zltools.zlrptdatas Where id=[1]"
    Else
        strSQL = "Select ���� From zltools.zlRPTPars Where Դid=[1] And ���=[2]"
    End If
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������", lngԴid, lng������)
    If rstmp.RecordCount > 0 Then
        GetSQLObj = "" & rstmp!����
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckSQLPhrase(ByVal lngԴid As Long, ByVal lng������ As Long, ByVal lngSys As Long, _
    ByVal strTabOld As String, ByVal strTabNew As String) As Boolean

    Dim strR As String, strFields As String, strOwner As String
    Dim objPars As RPTPars, strSQL As String
    
   
    Set objPars = GetParsObj(lngԴid, lng������, lngSys)
    strOwner = GetSQLObj(lngԴid, lng������)
    strOwner = Replace(strOwner, strTabOld, strTabNew)
    
    strSQL = rtbNew.Text
    strSQL = Replace(strSQL, "[ϵͳ]", lngSys)
    strSQL = RemoveNote(strSQL)
    strSQL = SQLReplaceOwner(strSQL, strOwner)
    
    If objPars.Count = 0 Then
        strFields = CheckSQL(strSQL, strR)
    Else
        strFields = CheckSQL(strSQL, strR, objPars)
    End If
    If strFields = "" Then
        MsgBox "SQL���У��ʧ�ܣ�" & vbCrLf & vbCrLf & _
            "���� " & strR & vbCrLf & vbCrLf & _
            "�����Ƿ���ȷ��д,������Ƿ���ȷ���ã�", vbInformation, App.Title
        CheckSQLPhrase = False
    Else
        CheckSQLPhrase = True
    End If
       
End Function

Private Function GetReportList(blnTableFull As Boolean, lngDefault As Long, lngModify As Long) As ADODB.Recordset
'���ܣ���ȡ��������Դ�б�
'������blnUnModied��ֻ��ʾδ�޸ĵļ�¼,blnTableFull:ֻ��ʾȫ��ɨ���
    Dim strSQL As String, strIF As String
 
    strIF = IIf(blnTableFull, " And A.ȫ��ɨ��=1", "")
    strIF = strIF & IIf(lngDefault = -2, "", " And A.ȱʡ = [1]")
    strIF = strIF & IIf(lngModify = -2, "", " And Nvl(A.����,-1) = [2]")
    strSQL = "Select *" & vbNewLine & _
            "From (Select A.����id, A.���, A.ȱʡ, A.����, Nvl(E.����, '����') ϵͳ��, B.���, B.���� ����, C.���� ����Դ, Null ����, B.ϵͳ, A.ȫ��ɨ��" & vbNewLine & _
            "       From Zltools.Zlrptadjustlog A, Zltools.Zlreports B, Zltools.Zlrptdatas C, Zltools.Zlsystems E" & vbNewLine & _
            "       Where A.����id = B.Id And A.����id = C.����id And A.����Դ = C.���� And B.ϵͳ = E.���(+)" & strIF & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select A.����id, A.���, A.ȱʡ, A.����, Nvl(E.����, '����') ϵͳ��, B.���, B.���� ����, C.���� ����Դ, D.���� ����, B.ϵͳ, A.ȫ��ɨ��" & vbNewLine & _
            "       From Zltools.Zlrptadjustlog A, Zltools.Zlreports B, Zltools.Zlrptdatas C, Zltools.Zlrptpars D, Zltools.Zlsystems E" & vbNewLine & _
            "       Where A.����id = B.Id And A.����id = C.����id And A.����Դ = C.���� And C.Id = D.Դid And A.��� = D.��� And B.ϵͳ = E.���(+)" & strIF & ")" & vbNewLine & _
            "Order By ϵͳ, ���, ����Դ, Nvl(���, 0)"

    On Error GoTo errH
    Set GetReportList = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", lngDefault, lngModify)

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
        cmdSave.Tag = "������"
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
'���ܣ���յ�ǰSQL�����������ݼ�״̬����
    rtbOld.Text = ""
    rtbNew.Text = ""
    mblnUnChange = True:    optMode(mȫ��).Value = True:    mblnUnChange = False
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
    Dim lngԴid As Long, lng������ As Long, arrtmp As Variant
    Dim lngȱʡ As Long, lng���� As Long
    
    If rptList.SelectedRows.Count = 0 Then 'δѡ��ʱ
        Call ClearSQLText
        mlngCurRow = -1
        Exit Sub
    End If
    If mlngCurRow = rptList.SelectedRows(0).Index Then Exit Sub
    
    If rtbExplan.Visible Then rtbExplan.Visible = False: rtbExplan.Text = "": cmdExplan.Caption = "�鿴�ƻ�(&X)"
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
    cmdDesign.Enabled = rptList.SelectedRows(0).Record.Tag <> ""    'ѡ��"ϵͳ"��ʱ
    cmdSave.Enabled = cmdDesign.Enabled
    cmdExplan.Enabled = cmdDesign.Enabled
        
    If rptList.SelectedRows(0).Childs.Count > 0 Then
        Call ClearSQLText
    Else
        lngȱʡ = Val("" & rptList.SelectedRows(0).Record(1).Value)
        lng���� = Val("" & rptList.SelectedRows(0).Record(2).Value)
        
        mblnUnChange = True
        If lng���� = mδ�� Then
            optMode(lngȱʡ).Value = True
        ElseIf lng���� <> m�ֹ� Then
            optMode(lng����).Value = True
        End If
        mblnUnChange = False
                
        arrtmp = Split(rptList.SelectedRows(0).Record(0).Value, "|SP|")
        lngԴid = GetԴID(arrtmp(0), arrtmp(1))
        lng������ = Val(arrtmp(2))
        Call ShowReportSQL(lngԴid, lng������)
            
        If lng���� = mδ�� Then
            '�����ͨ���ⲿ���߻򱨱�༭���޸�������Դ��û����д"����"�ֶΣ����ʱ��д
            If InStr(rtbOld.Text, "������ü�¼") > 0 Or InStr(rtbOld.Text, "סԺ���ü�¼") > 0 Then
                On Error GoTo errH
                gcnOracle.Execute "Update zltools.Zlrptadjustlog Set ����=" & m�ֹ� & " Where ԴID=" & lngԴid & " And Nvl(���,-1)=" & lng������
                
                rptList.SelectedRows(0).Record(2).Value = m�ֹ�
                rptList.SelectedRows(0).Record(2).Caption = "�ֹ�"
                lng���� = m�ֹ�
            End If
        End If
        If lng���� = m�ֹ� Then rtbNew.Text = ""
        
        fraMode.Enabled = lng���� <> m�ֹ�
        cmdSave.Enabled = lng���� <> m�ֹ�
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowReportSQL(lngԴid As Long, lng������ As Long)
    Dim strSQLText As String
    
    Set mrsSQL = GetRPTSQL(lngԴid, lng������)
    If mrsSQL Is Nothing Or lngԴid = 0 Then
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
        i = InStr(j, rtbOld.Text, "���˷��ü�¼")
        If i <= 0 Then
            i = InStr(j, rtbOld.Text, "������ü�¼")
            If i <= 0 Then
                i = InStr(j, rtbOld.Text, "סԺ���ü�¼")
                If i > 0 Then lngMode = 2
            Else
                lngMode = 1
            End If
            If i <= 0 Then Exit Do
        End If
        
        rtbOld.SelStart = i - 1
        rtbOld.SelLength = 6
        If lngMode = 1 Then
            rtbOld.SelColor = &HC000&   '��
        ElseIf lngMode = 2 Then
            rtbOld.SelColor = &HFF&     '��
        Else
            rtbOld.SelColor = &H8000000D    '��
        End If
        j = i + 6
    Loop While j < Len(rtbOld.Text)
    
    rtbOld.SelStart = 0
    rtbOld.SelLength = 0
End Sub

Private Function GetRPTSQL(ByVal lngԴid As Long, ByVal lng������ As Long) As ADODB.Recordset
'���ܣ���ȡ����Դ�������SQL
'������lng������:-1��ʾ��ȡ����ԴSQL�������ȡ������SQL
    Dim strSQL As String
    On Error GoTo errH
    
    If lng������ = -1 Then
        strSQL = "Select �к�,���� From zltools.zlRPTSQLs Where Դid = [1] Order By �к�"
        Set GetRPTSQL = zlDatabase.OpenSQLRecord(strSQL, "��ȡSQL", lngԴid)
    Else
        strSQL = "Select 0 as �к�,��ϸsql as ���� From zltools.zlRPTPars Where Դid = [1] And ��� = [2]"
        Set GetRPTSQL = zlDatabase.OpenSQLRecord(strSQL, "��ȡSQL", lngԴid, lng������)
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
        strTmp = IIf(i = 1, "", strTmp & vbNewLine) & rstmp!����
        rstmp.MoveNext
    Next
    GetTextByRs = strTmp
End Function

Private Sub SetNewText(blnChange As Boolean)
    Dim strTmp As String, lngNewMode As Long, lngOldMode As Long
    Dim i As Long, j As Long, lngLen As Long
    Dim strTabOld As String, strTabNew As String
    
    Call GetTwoMode(lngOldMode, lngNewMode)
        
    strTabOld = Choose(lngOldMode + 1, "���˷��ü�¼", "������ü�¼", "סԺ���ü�¼")
    strTabNew = Choose(lngNewMode + 1, "���˷��ü�¼", "������ü�¼", "סԺ���ü�¼")
        
   
    If lngOldMode = lngNewMode Then
        strTmp = rtbOld.Text
    Else
        strTmp = Replace(rtbOld.Text, strTabOld, strTabNew)
         '����,���˲���ID,�ಡ�˵��⼸���ֶ���Ϊ�����������ͬ���ֶΣ����Բ����Զ��滻����Ҫ�˹��޸�
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
            rtbNew.SelColor = &HC000&   '��
        ElseIf lngNewMode = 2 Then
            rtbNew.SelColor = &HFF&     '��
        Else
            rtbNew.SelColor = &H8000000D    '��
        End If
        j = i + Len(strTabNew)
    Loop While j < lngLen
    
    rtbNew.SelStart = 0
    rtbNew.SelLength = 0
    
    If blnChange Then
        Call ShowStatusInfor("�ѵ���Ϊ[" & strTabNew & "]")
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
            '������һ��
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
