VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportReport 
   BackColor       =   &H80000005&
   Caption         =   "����������"
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
      Caption         =   "SQL���漰""���˷��ü�¼""�ı����嵥"
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
            Text            =   "����"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ϵͳ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "���"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "˵��"
            Object.Width           =   7938
         EndProperty
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "����(&E)��"
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
         Caption         =   "����ִ��"
         Height          =   180
         Left            =   135
         TabIndex        =   2
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.Frame fraFloder 
      Caption         =   "Ŀ���ļ���"
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
Public Event StatusTextUpdate(ByVal strMSG As String) 'Ҫ�����������״̬������
 
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
    If MsgBox("�������� " & k & " �ű��� " & strPath & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    If CheckRPTIndex = False Then Exit Sub
    
    pgbState.Max = k
    picStatus.Visible = True
    Call Form_Resize
    Me.Refresh
    
    lblStatus.Caption = "���ڴ���������־..."
    If CreateLog = False Then Exit Sub
    
    curDate = Currentdate
    For i = 1 To k
        lblStatus.Caption = "���ڵ�����" & i & "��:" & lvwReport.ListItems(i).Text & ".ZLR"
        pgbState.Value = i

        strFile = "[" & lvwReport.ListItems(i).SubItems(2) & "]" & lvwReport.ListItems(i).Text & ".ZLR"
        If Not ExportReport(CLng(Mid(lvwReport.ListItems(i).Key, 2)), strPath & "\" & strFile, curDate) Then
            If MsgBox("��������ʱ���ִ���Ҫ����������һ�ű�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        End If
        'If i = 10 Then Exit For    '������
    Next
        
    picStatus.Visible = False
    Call Form_Resize
    Call ShowStatusInfor("��������" & i & "�ű���")
End Sub

Private Function CreateLog() As Boolean
    Dim rstmp As ADODB.Recordset, strSQL As String, i As Long, blnT As Boolean
    Dim strOut As String, strIn As String, strDefault As String
    
    CreateLog = True
    
    '����
    On Error GoTo errHandle
    strSQL = "Select 1 From All_Tables Where Table_Name = Upper('zlrptadjustlog') And Owner = 'ZLTOOLS'"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "������־��")
    If rstmp.RecordCount = 0 Then
        strSQL = "create table zltools.zlrptadjustlog(����ID NUMBER(18),����Դ VARCHAR2(20),Դid NUMBER(18),��� NUMBER(2),ȱʡ NUMBER(1),���� NUMBER(1),ȫ��ɨ�� NUMBER(1))"
        On Error Resume Next
        gcnOracle.Execute strSQL
        If Err.Number <> 0 Then
            MsgBox "�������������¼��(zltools.zlrptadjustlog)����", vbInformation, gstrSysName
            CreateLog = False
            Exit Function
        End If
    End If
    
    'д����
    On Error GoTo errHandle
    strSQL = "Select 1 From zltools.zlrptadjustlog Where rownum<2"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "������־��")
    
    If rstmp.RecordCount = 0 Then
        strOut = ",ZL1_BILL_1111,ZL1_BILL_1111_1,ZL1_BILL_1120,ZL1_BILL_1121_1,ZL1_BILL_1121_2,ZL1_BILL_1121_3,ZL1_BILL_1122" & _
                ",ZL1_INSIDE_1111_1,ZL1_INSIDE_1121_1,ZL1_INSIDE_1260_1,ZL1_INSIDE_1862,ZL1_REPORT_1123,ZL1_REPORT_1421" & _
                ",ZL1_REPORT_1876,ZL1_REPORT_1877,ZL1_REPORT_1882,ZL1_REPORT_1883,ZL1_SUB_1420_3,ZL1_SUB_1432_1" & _
                ",ZL1_SUB_1432_2,ZL1_SUB_1875_3,ZL1_SUB_1880_1,ZL1_SUB_1880_2,ZL1_SUB_1880_3,"
        strIn = ",ZL1_BILL_1133,ZL1_BILL_1134,ZL1_BILL_1135,ZL1_INSIDE_1102,ZL1_INSIDE_1102_1,ZL1_INSIDE_1139_2,ZL1_INSIDE_1342_1,ZL1_INSIDE_1605,"
        
        strSQL = "Select A.���,B.����id, B.����Դ, B.���" & vbNewLine & _
                "From zltools.zlReports A," & vbNewLine & _
                "     (Select Distinct ����id, ���� ����Դ, -null As ���" & vbNewLine & _
                "       From zltools.zlRPTDatas" & vbNewLine & _
                "       Where ���� Like '%���˷��ü�¼%'" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select A.����id, A.���� ����Դ, B.���" & vbNewLine & _
                "       From zltools.zlRPTDatas A, zltools.zlRPTPars B" & vbNewLine & _
                "       Where A.Id = B.Դid And B.���� Like '%���˷��ü�¼%') B" & vbNewLine & _
                "Where A.Id = B.����id" & vbNewLine & _
                "Order By ����id, ���"
        Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "���ұ���")
        With rstmp
            gcnOracle.BeginTrans: blnT = True
            For i = 1 To .RecordCount
                If InStr(strOut, "," & !��� & ",") > 0 Then
                    strDefault = "1"
                ElseIf InStr(strIn, "," & !��� & ",") > 0 Then
                    strDefault = "2"
                Else
                    strDefault = "0"
                End If
                strSQL = "Insert into zltools.zlrptadjustlog(����ID,����Դ,���,ȱʡ) values(" & !����ID & ",'" & !����Դ & "'," & IIf(IsNull(!���), "Null", !���) & "," & strDefault & ")"
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
        
    strIndex = "ZLRPTITEMS_IX_����ID"
    strSQL = "Select 1 From All_Indexes Where Index_Name = [1] And Owner='ZLTOOLS'"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��������", strIndex)
    If rstmp.RecordCount = 0 Then
        MsgBox "ȱʡ����[" & strIndex & "]�����������ǳ��������ȸ��ݰ�װ�ű�[zlServer.Sql]������������", vbInformation, gstrSysName
        Exit Function
    End If
    
    strIndex = "ZLRPTITEMS_IX_�ϼ�ID"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��������", strIndex)
    If rstmp.RecordCount = 0 Then
        MsgBox "ȱʡ����[" & strIndex & "]�����������ǳ��������ȸ��ݰ�װ�ű�[zlServer.Sql]������������", vbInformation, gstrSysName
        Exit Function
    End If
    CheckRPTIndex = True
End Function

Public Function GetExpField(objFld As ADODB.Field) As String
'���ܣ���������ʱ��
    Dim strTmp As String
    
    If IsNull(objFld.Value) Then
        Exit Function
    ElseIf InStr(",ϵͳ,����ID,����,����ʱ��,", "," & objFld.Name & ",") > 0 Then
        Exit Function
    ElseIf objFld.Name = "���" Then
        GetExpField = "[���]" '����ʱȡ��ǰʱ��
    ElseIf objFld.Name = "�޸�ʱ��" Then
        GetExpField = "Sysdate" '����ʱȡ��ǰʱ��
    ElseIf objFld.Name = "ID" Then
        GetExpField = "[NextVal]" '����ʱȡ"��ǰ��_ID.NextVal"
    ElseIf objFld.Name = "�ϼ�ID" Then
        GetExpField = "[CurrVal-X]" '����ʱȡ"��ǰ��_ID.CurrVal-X",XΪ�ϼ�ID��Ϊ�յĿ�ʼ��
    ElseIf objFld.Name = "����ID" Then
        GetExpField = "[zlReports_ID.CurrVal]" '����ʱȡ"zlReports_ID.CurrVal"
    ElseIf objFld.Name = "ԴID" Then
        GetExpField = "[zlRPTDatas_ID.CurrVal]" '����ʱȡ"zlRPTDatas_ID.CurrVal"
    ElseIf objFld.Name = "����" Then
        GetExpField = Replace(UCase(objFld.Value), UCase(gstrDBUser) & ".", "USER.")
    Else '����ʱ������������ת��ȡֵ
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
                '��ʱ��֧��ͼƬ�Ĵ���
        End Select
    End If
End Function

Public Function ExportReport(ByVal lngRPTID As Long, ByVal strFile As String, ByVal curDate As Date) As Boolean
'���ܣ�����һ���Զ��屨��
'������lngRPTID=����ID
'      strFile=�ļ���
'���أ������Ƿ�ɹ���
'˵����
'      1.�����ѷ����ı���,������Ϊ�Ƿ�������
'      2.Ŀǰ��֧��ͼƬԪ�����ݵĵ���
    
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
        MsgBox "û�з���ָ����������ݣ�", vbInformation, App.Title
        Exit Function
    End If
    
    '�򿪴����ļ�
    If mobjFile.FileExists(strFile) Then Call mobjFile.DeleteFile(strFile, True)
    Set mobjText = mobjFile.CreateTextFile(strFile, True)
    blnOpen = True
    
    '���������ͷ
    Call mobjText.WriteLine("[HEAD]")
    Call mobjText.WriteLine("������=" & rstmp!���)
    Call mobjText.WriteLine("��������=" & rstmp!����)
    Call mobjText.WriteLine("����˵��=" & IIf(IsNull(rstmp!˵��), "", rstmp!˵��))
    Call mobjText.WriteLine("�����û�=" & gstrDBUser)
    Call mobjText.WriteLine("����ʱ��=" & Format(curDate, "yyyy-MM-dd HH:mm:ss"))
    
    '����:ZLReport,�Էֺ�Ϊ�н������Էֺ�Ϊһ���ֶν���,���ֺ�Ϊһ����¼����
    Call mobjText.WriteLine("[ZLREPORTS]")
    Call mobjText.WriteLine(";")
    For Each objFld In rstmp.Fields
        Call mobjText.WriteLine(objFld.Name & "=" & GetExpField(objFld) & ";")
    Next
    
    '�����ʽ
    'Set rsTmp = New ADODB.Recordset
    strSQL = "Select * From zltools.zlRPTFmts Where ����ID=[1]"
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
    
    '����Ԫ��
    'Set rsTmp = New ADODB.Recordset
    strSQL = "Select * From zltools.zlRPTItems Where ����ID=[1] Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
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
    
    '��������,'���ݲ���
    'Set rsTmp = New ADODB.Recordset
    strSQL = "Select * From zltools.zlRPTDatas Where ����ID=[1]"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    'Set rsSQL = New ADODB.Recordset
    strSQL = "Select B.* From zltools.zlRPTDatas A,zlRPTSQLs B Where A.ID=B.ԴID And A.����ID=[1]"
    Set rsSQL = zlDatabase.OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    'Set rsSub = New ADODB.Recordset
    strSQL = "Select B.* From zltools.zlRPTDatas A,zlRPTPars B Where A.ID=B.ԴID And A.����ID=[1]"
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
            
            rsSQL.Filter = "ԴID=" & rstmp!Id
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
           
            rsSub.Filter = "ԴID=" & rstmp!Id
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
    lvwReport.ColumnHeaders(2).Position = 1 'ϵͳ
    lvwReport.ColumnHeaders(3).Position = 2 '���
    
    cmdExport.Enabled = lvwReport.ListItems.Count > 0
    Call ShowStatusInfor("��" & lvwReport.ListItems.Count & "�ű���")
End Sub

Private Sub LoadReportList()
    Dim rstmp As ADODB.Recordset
    Dim i As Long, objItem As ListItem
    
    lvwReport.ListItems.Clear
    Set rstmp = GetReportList()
    If Not rstmp Is Nothing Then
        For i = 1 To rstmp.RecordCount
            Set objItem = lvwReport.ListItems.Add(, "_" & rstmp!Id, rstmp!����, "Report", "Report")
            
            objItem.SubItems(1) = IIf(IsNull(rstmp!ϵͳ), "����", rstmp!ϵͳ)
            objItem.SubItems(2) = rstmp!���
            objItem.SubItems(3) = Nvl(rstmp!˵��)
            
            rstmp.MoveNext
        Next
    End If
End Sub

Private Function GetReportList() As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select A.Id, A.���, A.����, A.ϵͳ, A.����id, A.˵��" & vbNewLine & _
            "From zltools.zlReports A," & vbNewLine & _
            "     (Select Distinct ����id" & vbNewLine & _
            "       From zltools.zlRPTDatas" & vbNewLine & _
            "       Where ���� Like '%���˷��ü�¼%'" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select A.����id From zltools.zlRPTDatas A, zltools.zlRPTPars B Where A.Id = B.Դid And B.���� Like '%���˷��ü�¼%') B" & vbNewLine & _
            "Where A.Id = B.����id" & vbNewLine & _
            "Order By ϵͳ,����id, ���"

    On Error GoTo errH
    Set GetReportList = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����")

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Dim sngWidth As Long '��С���
    
    sngWidth = IIf(ScaleWidth < 5600, 5600, ScaleWidth)
    cmdExport.Left = Me.ScaleLeft + sngWidth - cmdExport.Width - 200
    
    With fraFloder  '�̶����
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
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub subPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

 
 
 


