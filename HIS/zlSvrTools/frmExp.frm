VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExp 
   BackColor       =   &H80000005&
   Caption         =   "���ݵ���"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmExp.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   870
      Index           =   2
      Left            =   1020
      Locked          =   -1  'True
      MaxLength       =   256
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   3600
      Width           =   5850
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "ִ��(&E)"
      Height          =   350
      Left            =   5760
      TabIndex        =   14
      Top             =   3165
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
      Height          =   1215
      Left            =   1020
      TabIndex        =   9
      Top             =   1920
      Width           =   5850
      Begin VB.CheckBox chkGrant 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����Ȩ��(&G)"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2100
         TabIndex        =   13
         Top             =   780
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chkIndex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��������(&I)"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2100
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chkConstraint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����Լ��(&C)"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   390
         TabIndex        =   12
         Top             =   780
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chkRows 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����������(&R)"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   390
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.PictureBox picXp 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   3690
         ScaleHeight     =   690
         ScaleWidth      =   1515
         TabIndex        =   18
         Top             =   390
         Width           =   1515
         Begin VB.OptionButton optPath 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��ͳ��ʽ(&V)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.OptionButton optPath 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ֱ��·��(&D)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   420
            Width           =   1305
         End
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
      TabIndex        =   17
      Top             =   3330
      Width           =   900
   End
   Begin VB.Label lbl˵�� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   990
      TabIndex        =   15
      Top             =   4650
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
      Picture         =   "frmExp.frx":04F9
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
Attribute VB_Name = "frmExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mrsSystem As New ADODB.Recordset
Dim mstr������ As String '���浱ǰϵͳ����������
Dim mstrVer As String
Private mblnLoad As Boolean

Private Sub Form_Activate()
    If mblnLoad Then
        txtFile(2).Text = GetCommand()
        mblnLoad = False
    End If
End Sub

Private Function GetCommand() As String
    GetCommand = "EXP" & mstrVer & " USERID=" & gstrUserName & "/" & IIf(gstrUserName <> gstrLoginUserName, "*****", gstrPassword) & IIf(gstrServer = "", "", "@" & gstrServer) _
        & " BUFFER=4096 FILE=" & Trim(txtFile(0).Text) & IIf(Trim(txtFile(1).Text) = "", "", " LOG=" & Trim(txtFile(1).Text)) _
        & " OWNER=(" & mstr������ & ") ROWS=" & IIf(chkRows.value = 1, "Y", "N") _
        & " INDEXES =" & IIf(chkIndex.value = 1, "Y", "N") & " CONSTRAINTS =" & IIf(chkConstraint.value = 1, "Y", "N") _
        & " GRANTS=" & IIf(chkGrant.value = 1, "Y", "N") & " DIRECT=" & IIf(optPath(1).value = True, "Y", "N")
End Function

Private Sub cmdExecute_Click()
    Dim strDMPFile As String
    Dim strLogFile As String
    Dim lngProcess As Long
    Dim lngTemp As Long
    Dim strCommand As String
    Dim varTime As Variant
    Dim strNote As String
    
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
    If Dir(strDMPFile) <> "" Then
        If MsgBox("�����ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            txtFile(0).SetFocus
            Exit Sub
        End If
    End If
    If Dir(strLogFile) <> "" And strLogFile <> "" Then
        If MsgBox("��¼��־�Ѿ����ڣ��Ƿ񸲸ǣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            txtFile(1).SetFocus
            Exit Sub
        End If
    End If
    
    lngTemp = FreeFile
    On Error Resume Next
    Open strDMPFile For Binary As lngTemp
    If err <> 0 Then
        MsgBox "�����ļ����Ƿ���", vbExclamation, gstrSysName
        txtFile(0).SetFocus
        Exit Sub
    End If
    Close lngTemp
    If strLogFile <> "" Then
        '��¼��־�ļ�����Ϊ��
        Open strLogFile For Binary As lngTemp
        If err <> 0 Then
            MsgBox "��¼��־�ļ����Ƿ���", vbExclamation, gstrSysName
            txtFile(1).SetFocus
            Exit Sub
        End If
        Close lngTemp
    End If
    
    'ִ�е�������
    On Error GoTo ErrHandle
    frmMDIMain.Enabled = False
    cmbSystem.Enabled = False
    cmdExecute.Enabled = False
    fra.Enabled = False
    
    frmWait.BeginWait "���ڵ������ݡ���"
    strCommand = GetCommand()
    
    If gstrUserName <> gstrLoginUserName Then
        strCommand = Replace(strCommand, "*****", gstrPassword)
    End If
    err = 0
    On Error Resume Next
    varTime = Now() '��¼�¿�ʼ������ʱ��
    lngTemp = Shell(strCommand, vbHide)
    err = 0
    On Error GoTo ErrHandle
        
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngTemp
        DoEvents
    Loop While lngTemp = Still_Active
    CloseHandle lngProcess
            
    frmWait.EndWait
    If Len(Dir(strDMPFile)) <> 0 And lngTemp = 0 Then
        MsgBox "���ݱ�����ɣ�" & vbCrLf & vbCrLf & "����ʱ" & Format(CDate(Now - varTime), "hhʱmm��ss��") & "��", vbExclamation, gstrSysName
        If chkRows.value = 1 Then strNote = ",������"
        If chkIndex.value = 1 Then strNote = strNote & ",����"
        If chkConstraint.value = 1 Then strNote = strNote & ",Լ��"
        If chkGrant.value = 1 Then strNote = strNote & ",Ȩ��"
        '������Ҫ������־
        Call SaveAuditLog(2, "ִ��", "�ɹ�����" & Split(cmbSystem.Text, " ")(0) & "���е�" & Mid(strNote, 2) & "�������ļ�" & Right(strDMPFile, Len(strDMPFile) - InStrRev(strDMPFile, "\")) & "��")
    Else
        MsgBox "Ŀǰ��ϵͳ������ȷ������ݱ��ݣ����飺" & _
            vbCr & "   1) �Ƿ����exp" & mstrVer & ".exe�ļ���" & _
            vbCr & "   2) Set Path�Ƿ�ָ������ڵ�Ŀ¼��" & _
            vbCr & "   3) ����ͬ�����Ҳ���д�ĵ����ļ�����־�ļ���", vbExclamation, gstrSysName
    End If
    
    
    frmMDIMain.Enabled = True
    cmbSystem.Enabled = True
    cmdExecute.Enabled = True
    fra.Enabled = True
    Exit Sub
ErrHandle:
    frmWait.EndWait
    MsgBox "���ݵ���ʧ�ܡ�", vbExclamation, gstrSysName
    frmMDIMain.Enabled = True
    cmbSystem.Enabled = True
    cmdExecute.Enabled = True
    fra.Enabled = True
End Sub

Private Sub chkRows_Click()
    txtFile(2).Text = GetCommand()
End Sub

Private Sub chkIndex_Click()
    txtFile(2).Text = GetCommand()
End Sub

Private Sub chkConstraint_Click()
    txtFile(2).Text = GetCommand()
End Sub

Private Sub chkGrant_Click()
    txtFile(2).Text = GetCommand()
End Sub

Private Sub optPath_Click(Index As Integer)
    txtFile(2).Text = GetCommand()
End Sub

Private Sub chkRows_KeyPress(KeyAscii As Integer)
    chkIndex.SetFocus
End Sub

Private Sub chkIndex_KeyPress(KeyAscii As Integer)
    chkConstraint.SetFocus
End Sub

Private Sub chkConstraint_KeyPress(KeyAscii As Integer)
    chkGrant.SetFocus
End Sub

Private Sub chkGrant_KeyPress(KeyAscii As Integer)
    If optPath(0).value = True Then
        optPath(0).SetFocus
    Else
        optPath(1).SetFocus
    End If
End Sub

Private Sub cmdFile_Click(Index As Integer)
    If Index = 0 Then
        cmmFile.Filter = "�����ļ�(*.dmp)|*.dmp"
    Else
        cmmFile.Filter = "��¼��־(*.log)|*.log"
    End If
    cmmFile.FileName = txtFile(Index).Text
    cmmFile.ShowSave
    If cmmFile.FileName <> "" Then txtFile(Index).Text = cmmFile.FileName
End Sub

Private Sub optPath_KeyPress(Index As Integer, KeyAscii As Integer)
    cmdExecute.SetFocus
End Sub

Private Sub cmbSystem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtFile(0).SetFocus
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
            chkRows.SetFocus
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim intVer As Integer
    
    mblnLoad = True
    lbl˵��.Caption = "��ʾ��" & vbCrLf & _
                    "     ��������Ҫ����һ���������Ĺ��̲�����ɡ������ʱ���ڷ������Կͻ�����Ӧ���óٶۣ��������ڷ���������ʱ��ɱ�������" & vbCrLf & _
                    "     �ڵ��������л򵼳���ɺ󣬿���ͨ����¼��־�ļ��˽⵼����һЩ���������" & vbCrLf & _
                    "     ������Ե���������Ϥ��Ҳ����ֱ����Windows���д���ִ���������ı���"
    
    intVer = GetOracleVersion
    
    If intVer < 80 Then
        MsgBox "��Oracle�汾�������ڹ��ɣ���������ܲ����������У�" & vbCr _
            & "�뿼�ǽ�BINĿ¼�е�[EXP+�汾��.EXE]��Ϊ[EXP.EXE]��ִ�С�", vbExclamation, gstrSysName
        mstrVer = ""
    ElseIf intVer = 80 Then      'Oracle8.0
        mstrVer = "80"
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
        'ֻ����������
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
    
End Sub

Private Sub FillSystem()
    '��ʾ���п���ʾ��ϵͳ
    On Error GoTo ErrHandle
    mrsSystem.CursorLocation = adUseClient
    cmbSystem.Clear
    If gblnDBA = True Then
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", UCase(gstrUserName))
    End If
    Do Until mrsSystem.EOF
        cmbSystem.addItem mrsSystem("����") & " v" & mrsSystem("�汾��") & "��" & mrsSystem("���") & "��"
        cmbSystem.ItemData(cmbSystem.NewIndex) = mrsSystem("���")
        mrsSystem.MoveNext
    Loop
    If gblnDBA = True Then
        cmbSystem.addItem "������"
        cmbSystem.ItemData(cmbSystem.NewIndex) = -1
    End If
    If mrsSystem.RecordCount > 0 Then
        cmbSystem.ListIndex = 0
    Else
        cmdExecute.Enabled = False
    End If
    Exit Sub
ErrHandle:
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

