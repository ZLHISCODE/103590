VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImp 
   BackColor       =   &H80000005&
   Caption         =   "���ݵ���"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmImp.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "������ѡ��"
      ForeColor       =   &H80000008&
      Height          =   1665
      Index           =   1
      Left            =   4980
      TabIndex        =   21
      Top             =   1950
      Width           =   1935
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ɾ��ԭ�б�(&T)"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   150
         TabIndex        =   22
         Top             =   300
         Width           =   1575
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��ձ�����(&A)"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   420
         TabIndex        =   23
         Top             =   690
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ʹԼ����Ч(&D)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   420
         TabIndex        =   24
         Top             =   1110
         Value           =   1  'Checked
         Width           =   1485
      End
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   630
      Index           =   2
      Left            =   1020
      Locked          =   -1  'True
      MaxLength       =   256
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   4770
      Width           =   5925
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "ִ��(&E)"
      Height          =   350
      Left            =   5820
      TabIndex        =   25
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "��"
      Height          =   300
      Index           =   1
      Left            =   6570
      TabIndex        =   8
      Top             =   1560
      Width           =   300
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   2070
      MaxLength       =   256
      TabIndex        =   7
      Top             =   1560
      Width           =   4485
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "����ѡ��"
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   0
      Left            =   1020
      TabIndex        =   9
      Top             =   1920
      Width           =   3795
      Begin VB.TextBox txtBuffer 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   20
         Text            =   "300"
         ToolTipText     =   "��ע���ֵ��Ҫ������ǰ�����ڴ��С"
         Top             =   2040
         Width           =   555
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         MaxLength       =   256
         TabIndex        =   18
         Top             =   1680
         Width           =   2235
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���Զ��󴴽�����(&R)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   390
         TabIndex        =   16
         Top             =   1350
         Value           =   1  'Checked
         Width           =   2025
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����ÿ���ύ(&M)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2100
         TabIndex        =   15
         Top             =   1005
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ֻ��ʾ����(&W)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   390
         TabIndex        =   14
         Top             =   1005
         Width           =   1515
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����Ȩ��(&G)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   2100
         TabIndex        =   13
         Top             =   660
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��������(&I)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   11
         Top             =   330
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����Լ��(&C)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   12
         Top             =   660
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���������(&R)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   10
         Top             =   330
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݻ�������С(��λ:KB)(&B)"
         Height          =   180
         Index           =   5
         Left            =   360
         TabIndex        =   19
         Top             =   2100
         Width           =   2340
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����û�(&U)"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   17
         Top             =   1740
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "��"
      Height          =   300
      Index           =   0
      Left            =   6570
      TabIndex        =   5
      Top             =   1140
      Width           =   300
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   2070
      MaxLength       =   256
      TabIndex        =   4
      Top             =   1155
      Width           =   4485
   End
   Begin MSComDlg.CommonDialog cmmFile 
      Left            =   5220
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������ı�"
      Height          =   180
      Index           =   3
      Left            =   1020
      TabIndex        =   28
      Top             =   4500
      Width           =   900
   End
   Begin VB.Label lbl˵�� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1020
      TabIndex        =   26
      Top             =   5640
      Width           =   6825
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��¼��־(&L)"
      Height          =   180
      Index           =   2
      Left            =   1020
      TabIndex        =   6
      Top             =   1590
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ϵͳ(&S)"
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   780
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ļ�(&F)"
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   3
      Top             =   1200
      Width           =   990
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   240
      Picture         =   "frmImp.frx":04F9
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݵ���"
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
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
End
Attribute VB_Name = "frmImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mrsSystem As New ADODB.Recordset
Dim mstr������ As String '���浱ǰϵͳ����������
Dim mstrVer As String

Private Enum conCheck
    Rows = 0
    Indexes = 1
    Constraints = 2
    Grants = 3
    OnlyShow = 4
    Commit = 5
    Ignore = 6
    Clear = 7
    Disable = 8
    Drop = 9
End Enum

Private Function GetCommand() As String
    Dim strFromUser As String
    
    strFromUser = Trim(txtUser.Text)
    If strFromUser = "" Then strFromUser = mstr������
    
    GetCommand = "IMP" & mstrVer & " USERID=" & gstrUserName & "/" & IIf(gstrUserName <> gstrLoginUserName, "*****", gstrPassword) & IIf(gstrServer = "", "", "@" & gstrServer) _
        & " FROMUSER=(" & strFromUser & ")  TOUSER=(" & mstr������ & ") BUFFER=" & IIf(IsNumeric(txtBuffer.Text), CStr(Val(txtBuffer.Text) * 1024), "30720") _
        & " FILE=" & Trim(txtFile(0).Text) & IIf(Trim(txtFile(1).Text) = "", "", " LOG=" & Trim(txtFile(1).Text)) _
        & " ROWS=" & IIf(chk(Rows).value = 1, "Y", "N") & " INDEXES=" & IIf(chk(Indexes).value = 1, "Y", "N") _
        & IIf(chk(Constraints).Enabled, " CONSTRAINTS =" & IIf(chk(Constraints).value = 1, "Y", "N"), "") & " GRANTS =" & IIf(chk(Grants).value = 1, "Y", "N") _
        & " SHOW=" & IIf(chk(OnlyShow).value = 1, "Y", "N" & " COMMIT =" & IIf(chk(Commit).value = 1, "Y", "N") _
                                                       & " IGNORE=" & IIf(chk(Ignore).value = 1, "Y", "N"))
End Function

Private Sub cmdExecute_Click()
    Dim strDMPFile As String
    Dim strLogFile As String
    Dim lngProcess As Long
    Dim lngTemp As Long
    Dim strCommand As String
    Dim varTime As Variant
    Dim rsTemp As New ADODB.Recordset
    Dim rsCons As New ADODB.Recordset
    Dim blnSuccess As Boolean
    Dim intVer As Integer
    Dim strNote As String
    
    intVer = GetOracleVersion
    
    '���ļ����ĺϷ��Խ����ж�
    strDMPFile = Trim(txtFile(0).Text)
    strLogFile = Trim(txtFile(1).Text)
    If strDMPFile = "" Then
        txtFile(0).SetFocus
        MsgBox "�����뵼���ļ�����", vbExclamation, gstrSysName
        Exit Sub
    End If
    If strLogFile = strDMPFile Then
        txtFile(1).SetFocus
        MsgBox "�����ļ�����־��¼�ļ�������ͬһ���ļ���", vbExclamation, gstrSysName
        Exit Sub
    End If
    If Dir(strDMPFile) = "" Then
        MsgBox "������һ����ȷ�ĵ����ļ�����", vbExclamation, gstrSysName
        txtFile(0).SetFocus
        Exit Sub
    End If
    If Dir(strLogFile) <> "" And strLogFile <> "" Then
        If MsgBox("��¼��־�Ѿ����ڣ��Ƿ񸲸ǣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            txtFile(1).SetFocus
            Exit Sub
        End If
    End If
    If strLogFile <> "" Then
        On Error Resume Next
        lngTemp = FreeFile
        '��¼��־�ļ�����Ϊ��
        Open strLogFile For Binary As lngTemp
        If err <> 0 Then
            MsgBox "��¼��־�ļ����Ƿ���", vbExclamation, gstrSysName
            txtFile(1).SetFocus
            Exit Sub
        End If
        Close lngTemp
    End If
    
    'ִ�е������
    
    If MsgBox("���Ҫ���е��������" & vbCrLf & "�������е����ݿ�������Ӱ��ġ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error GoTo errHandle
    SetEnable False
    
    frmWait.BeginWait "��ʼʱ��:" & Now() & ".����������ݡ���"
    strCommand = GetCommand()
    If gstrUserName <> gstrLoginUserName Then
        strCommand = Replace(strCommand, "*****", gstrPassword)
    End If
    
    varTime = Now() '��¼�¿�ʼ������ʱ��
    
    On Error Resume Next
    rsTemp.CursorLocation = adUseClient
    rsCons.CursorLocation = adUseClient
    If chk(Drop).value = 1 Then
        'ɾ�����б�
        gstrSQL = "select TABLE_NAME from all_tables where OWNER='" & mstr������ & "' And Instr(Table_NAME,'BIN$')<=0 "
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            '--- 2007-03-07 ɾ����ʱ,�����10g,�����Purge �ؼ���,���������վ
            gcnOracle.Execute "Drop Table " & mstr������ & "." & rsTemp("TABLE_NAME") & " cascade constraints" & IIf(intVer >= 100, " Purge", "")
            rsTemp.MoveNext
        Loop
        
        If rsTemp.State = adStateOpen Then rsTemp.Close
        
        'ɾ��������ͼ
        gstrSQL = "Select View_name From All_Views Where Owner = '" & mstr������ & "'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            gcnOracle.Execute "Drop View " & mstr������ & "." & rsTemp("View_name")
            rsTemp.MoveNext
        Loop
        
        'ɾ����������
        gstrSQL = "select SEQUENCE_NAME from all_sequences where SEQUENCE_OWNER='" & mstr������ & "'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            gcnOracle.Execute "Drop Sequence " & mstr������ & "." & rsTemp("SEQUENCE_NAME")
            rsTemp.MoveNext
        Loop
    Else
        If chk(Disable).value = 1 Then
            gstrSQL = "select CONSTRAINT_NAME,CONSTRAINT_TYPE,TABLE_NAME from all_constraints where OWNER='" & mstr������ & "' And Instr(Table_NAME,'BIN$')<=0"
            rsCons.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
            '����ʹ���Լ����Ч
            rsCons.Filter = "CONSTRAINT_TYPE='R'"
            Do Until rsCons.EOF
                gcnOracle.Execute "Alter Table " & mstr������ & "." & rsCons("TABLE_NAME") & " disable constraint " & rsCons("CONSTRAINT_NAME")
                rsCons.MoveNext
            Loop
            '��ʹ�������͵�Լ����Ч
            rsCons.Filter = "CONSTRAINT_TYPE<>'R'"
            Do Until rsCons.EOF
                gcnOracle.Execute "Alter Table " & mstr������ & "." & rsCons("TABLE_NAME") & " disable constraint " & rsCons("CONSTRAINT_NAME")
                rsCons.MoveNext
            Loop
        End If
        If chk(Clear).value = 1 Then
            gstrSQL = "select TABLE_NAME from all_tables where OWNER='" & mstr������ & "' And Instr(Table_NAME,'BIN$')<=0"
            rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
            Do Until rsTemp.EOF
                gcnOracle.Execute "truncate Table " & mstr������ & "." & rsTemp("TABLE_NAME") & "  drop storage"
                rsTemp.MoveNext
            Loop
        End If
    End If
    
    'ִ��Import����
    '��ʾ���ڵ�������
    frmWait.lbl���� = Replace(frmWait.lbl����, "���", "����")
    err.Clear
    lngTemp = Shell(strCommand, vbHide)
    If err <> 0 Then
        err.Clear
        MsgBox "Ŀǰ��ϵͳ������ȷ������ݻָ������飺" & _
            vbCrLf & "   1�� �Ƿ����imp" & mstrVer & ".exe�ļ���" & _
            vbCrLf & "   2�� Set Path�Ƿ�ָ������ڵ�Ŀ¼��" & _
            vbCrLf & "   3�� �����ļ�����ͬ�汾��Export���򵼳��ġ�", vbExclamation, gstrSysName
        frmWait.EndWait
        SetEnable True
        Exit Sub
    End If
    
    On Error GoTo errHandle
        
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngTemp
        DoEvents
    Loop While lngTemp = Still_Active
    CloseHandle lngProcess
    Call AdjustSequence(mstr������, gcnOracle)
    
    If lngTemp <> 0 And lngTemp <> 1 Then
        frmWait.EndWait
        MsgBox "���ݵ����������ʧ�ܡ������Ҫ������ϸ������Ϣ���������С��������ı�����" & vbCrLf & _
            "���飺" & vbCrLf & _
            "   1�� ��ѡ�������ļ��Ƿ���Ч�ĵ����ļ���" & vbCrLf & _
            "   2�� �����ļ���Ҫ��������ݿ�汾��ͬ�����ܽ�8.0�������ļ����뵽8i�У�" & vbCrLf & _
            "   3�� �����û���Ȩ���Ƿ����Ҫ�󣬲��ܽ�DBA�û���������������ͨ�û����룻" & vbCrLf & _
            "   4�� �����û����Ƿ���ȷ������뵱ǰ�û�����ͬ�����ڡ������û���������ȷ���û�����", vbExclamation, gstrSysName
        SetEnable True
        Exit Sub
    End If
    'ִ�����
errHandle:
    If err = 0 Then blnSuccess = True
    
    On Error Resume Next
    strLogFile = "" '���������������δ�ɹ��ָ���Լ��
    If chk(Disable).value = 1 And chk(Drop).value = 0 Then
        '����ʹ�������͵�Լ����Ч
        rsCons.MoveFirst
        rsCons.Filter = "CONSTRAINT_TYPE<>'R'"
        Do Until rsCons.EOF
            gcnOracle.Execute "Alter Table " & mstr������ & "." & rsCons("TABLE_NAME") & " Enable constraint " & rsCons("CONSTRAINT_NAME")
            If err <> 0 Then
                err.Clear
                strLogFile = strLogFile & vbCrLf & rsCons("CONSTRAINT_NAME")
            End If
            rsCons.MoveNext
        Loop
        'Ȼ��ʹ���Լ����Ч
        rsCons.Filter = "CONSTRAINT_TYPE='R'"
        Do Until rsCons.EOF
            gcnOracle.Execute "Alter Table " & mstr������ & "." & rsCons("TABLE_NAME") & " Enable constraint " & rsCons("CONSTRAINT_NAME")
            If err <> 0 Then
                err.Clear
                strLogFile = strLogFile & vbCrLf & rsCons("CONSTRAINT_NAME")
            End If
            rsCons.MoveNext
        Loop
    End If
    '�ָ�����
    Call AdjustSequence(mstr������, gcnOracle)
    frmWait.EndWait
    If blnSuccess = True Then
        MsgBox "���ݻָ���ɣ�" & vbCrLf & vbCrLf & _
            "����ʱ" & Format(CDate(Now - varTime), "hhʱmm��ss�롣") & _
            IIf(strLogFile = "", "", "������Լ�������������ã�" & strLogFile), vbExclamation, gstrSysName
        If chk(Rows).value = 1 Then strNote = ",������"
        If chk(Indexes).value = 1 Then strNote = ",����"
        If chk(Constraints).value = 1 Then strNote = ",Լ��"
        If chk(Grants).value = 1 Then strNote = ",Ȩ��"
        '������Ҫ������־
        Call SaveAuditLog(2, "ִ��", "�ɹ��������ļ�" & Right(strDMPFile, Len(strDMPFile) - InStrRev(strDMPFile, "\")) & "�е�" & Mid(strNote, 2) & "���뵽��" & Split(cmbSystem.Text, " ")(0) & "����")
    Else
        MsgBox "���ݵ���ʧ�ܡ�" & IIf(strLogFile = "", "", "��������Լ�������������ã�" & strLogFile), vbExclamation, gstrSysName
    End If
    SetEnable True
End Sub

Private Sub SetEnable(ByVal blnEnable As Boolean)
    frmMDIMain.Enabled = blnEnable
    cmbSystem.Enabled = blnEnable
    cmdExecute.Enabled = blnEnable
    fra(0).Enabled = blnEnable
    fra(1).Enabled = blnEnable
End Sub

Private Sub chk_Click(Index As Integer)
    Dim i As Integer
    If Index = OnlyShow Then
        '���������ֻ�����ʾʱ����Щѡ���ǲ����õ�
        If chk(Index).value = 1 Then
            For i = 5 To 9
                chk(i).Enabled = False
                chk(i).value = 0
            Next
        Else
            For i = 5 To 9
                chk(i).Enabled = True
            Next
        End If
    ElseIf Index = Drop Then
        '���������Ҫɾ��ԭ�б�ʱ����Щѡ���ǲ����õ�
        If chk(Index).value = 1 Then
            For i = 7 To 8
                chk(i).Enabled = False
                chk(i).value = 0
            Next
        Else
            For i = 7 To 8
                chk(i).Enabled = True
            Next
        End If
    ElseIf Index = Clear Then
        If chk(Index).value = 1 Then
            chk(Disable).Enabled = False
            chk(Disable).value = 1
        Else
            chk(Disable).Enabled = True
        End If
    End If
    txtFile(2).Text = GetCommand()
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    SendKeys "{TAB}"
End Sub

Private Sub cmdFile_Click(Index As Integer)
    cmmFile.FileName = txtFile(Index).Text
    If Index = 0 Then
        cmmFile.Filter = "�����ļ�(*.dmp)|*.dmp"
        cmmFile.ShowOpen
    Else
        cmmFile.Filter = "��¼��־(*.log)|*.log"
        cmmFile.ShowSave
    End If
    If cmmFile.FileName <> "" Then txtFile(Index).Text = cmmFile.FileName
End Sub

Private Sub cmbSystem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtFile(0).SetFocus
End Sub


Private Sub txtBuffer_Change()
    txtFile(2).Text = GetCommand()
End Sub

Private Sub txtBuffer_GotFocus()
    txtBuffer.SelStart = 0
    txtBuffer.SelLength = Len(txtBuffer.Text)
End Sub

Private Sub txtBuffer_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBuffer_LostFocus()
    If Not IsNumeric(txtBuffer) Then
        txtBuffer.SetFocus
    End If
End Sub


Private Sub txtFile_Change(Index As Integer)
    If Index <> 2 Then
        txtFile(2).Text = GetCommand()
    End If
End Sub

Private Sub txtFile_GotFocus(Index As Integer)
    'txtFile(Index).SetFocus
End Sub

Private Sub txtFile_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            txtFile(1).SetFocus
        Else
            chk(Rows).SetFocus
        End If
    End If

End Sub

Private Sub Form_Load()
    lbl˵��.Caption = "��ʾ��" & vbCrLf & _
                    "     ��������Ҫ����һ���������Ĺ��̲�����ɡ������ʱ���ڷ������Կͻ�����Ӧ���óٶۣ��������ڷ���������ʱ��ɱ�������" & vbCrLf & _
                    "     �ڵ�������л�����ɺ󣬿���ͨ����¼��־�ļ��˽⵼���һЩ���������" & vbCrLf & _
                    "     ������Ե���������Ϥ��Ҳ����ֱ����Windows���д���ִ���������ı���"
    
    Dim intVer As Integer
    
    intVer = GetOracleVersion
    
    If intVer < 80 Then
        MsgBox "��Oracle�汾�������ڹ��ɣ���������ܲ����������У�" & vbCr _
            & "�뿼�ǽ�BINĿ¼�е�[IMP+�汾��.EXE]��Ϊ[IMP.EXE]��ִ�С�", vbExclamation, gstrSysName
        mstrVer = ""
    ElseIf intVer = 80 Then            'Oracle8.0
        mstrVer = "80"
        chk(Constraints).value = 0
        chk(Constraints).Enabled = False
    Else
        mstrVer = ""
    End If
    Call FillSystem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsSystem.State = 1 Then mrsSystem.Close
    Set mrsSystem = Nothing
    mstr������ = ""
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    
    On Error Resume Next
    sngTemp = IIf(ScaleWidth > 5000, ScaleWidth, 5000)
    cmbSystem.Width = sngTemp - cmbSystem.Left - 200
    cmdFile(0).Left = sngTemp - cmdFile(0).Width - 200
    cmdFile(1).Left = cmdFile(0).Left
    txtFile(0).Width = cmdFile(0).Left - 15 - txtFile(1).Left
    txtFile(1).Width = txtFile(0).Width
    txtFile(2).Width = cmbSystem.Left + cmbSystem.Width - txtFile(2).Left
    
    lbl˵��.Width = ScaleWidth - 200 - lbl˵��.Left
    lbl˵��.Height = ScaleHeight - 200 - lbl˵��.Top
    
End Sub

Private Sub cmbSystem_Click()
    If cmbSystem.ItemData(cmbSystem.ListIndex) = -1 Then
        cmdExecute.Enabled = True
        mstr������ = "ZLTOOLS"
    Else
        mrsSystem.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
        If mrsSystem.RecordCount = 0 Then
            cmdExecute.Enabled = False
        Else
            cmdExecute.Enabled = True
            mstr������ = mrsSystem("������")
        End If
    End If
    txtFile(2).Text = GetCommand()
End Sub

Private Sub FillSystem()
    '��ʾ���п���ʾ��ϵͳ
    On Error GoTo errHandle
    If gblnDBA = True Then
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", UCase(gstrUserName))
    End If
    
    Do Until mrsSystem.EOF
        cmbSystem.AddItem mrsSystem("����") & " v" & mrsSystem("�汾��") & "��" & mrsSystem("���") & "��"
        cmbSystem.ItemData(cmbSystem.NewIndex) = mrsSystem("���")
        mrsSystem.MoveNext
    Loop
    If gblnDBA = True Then
        cmbSystem.AddItem "������"
        cmbSystem.ItemData(cmbSystem.NewIndex) = -1
    End If
    If mrsSystem.RecordCount > 0 Then
        cmbSystem.ListIndex = 0
    Else
        cmdExecute.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

Private Sub txtUser_Change()
    txtFile(2).Text = GetCommand()
End Sub
