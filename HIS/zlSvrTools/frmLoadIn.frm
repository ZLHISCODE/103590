VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoadIn 
   BackColor       =   &H80000005&
   Caption         =   "���ݵ���"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmLoadIn.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "����ǰ��ձ�����(&L)"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3690
      TabIndex        =   13
      Top             =   4590
      Value           =   1  'Checked
      Width           =   2085
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Left            =   2310
      TabIndex        =   10
      Top             =   5040
      Width           =   1155
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   1020
      TabIndex        =   9
      Top             =   5040
      Width           =   1155
   End
   Begin VB.DriveListBox DriveBak 
      Height          =   300
      Left            =   3675
      TabIndex        =   6
      Top             =   1290
      Width           =   2880
   End
   Begin VB.DirListBox DirBak 
      Appearance      =   0  'Flat
      Height          =   2820
      Left            =   3675
      TabIndex        =   7
      Top             =   1620
      Width           =   2880
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "ִ��(&E)"
      Height          =   350
      Left            =   5460
      TabIndex        =   8
      Top             =   5040
      Width           =   1155
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5760
      Width           =   5625
   End
   Begin VB.ComboBox cmbSystem 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   4485
   End
   Begin MSComctlLib.ListView lvwTabs 
      Height          =   3675
      Left            =   1020
      TabIndex        =   4
      Top             =   1290
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   6482
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblTabs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݱ�ѡ��(T)"
      Height          =   180
      Left            =   1020
      TabIndex        =   3
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ļ����Ŀ¼(&D)"
      Height          =   180
      Left            =   3690
      TabIndex        =   5
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Loader�����˽�Ļ���Ҳ���ֹ�����ִ����������"
      Height          =   180
      Index           =   3
      Left            =   1050
      TabIndex        =   11
      Top             =   5490
      Width           =   4500
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
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   240
      Picture         =   "frmLoadIn.frx":04F9
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
Attribute VB_Name = "frmLoadIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mrsTable As New ADODB.Recordset
Dim mstr������ As String '���浱ǰϵͳ����������
Dim mstrVer As String

Private Sub cmbSystem_Click()
    Call DirBak_Change
End Sub

Private Sub cmdSelect_Click()
    Dim i As Integer
    For i = 1 To lvwTabs.ListItems.Count
        lvwTabs.ListItems(i).Checked = True
    Next
    Call lvwTabs_ItemCheck(Nothing)
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    For i = 1 To lvwTabs.ListItems.Count
        lvwTabs.ListItems(i).Checked = False
    Next
    Call lvwTabs_ItemCheck(Nothing)
End Sub

Private Sub DirBak_Change()
    Dim strFile As String
    Dim strPath As String
    Dim lst As ListItem
    
    strPath = DirBak.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    lvwTabs.ListItems.Clear
    txtFile.Text = ""
    strFile = UCase(Dir(strPath))
    Do Until strFile = ""
        If Right(strFile, 4) = ".LDR" Then
            mrsTable.Filter = "Table_name='" & Left(strFile, Len(strFile) - 4) & "'"
            If Not mrsTable.EOF Then
                Set lst = lvwTabs.ListItems.Add(, , mrsTable("TABLE_NAME"))
                lst.Checked = True
                '��ʾ��ǰ��ִ�е�����
                txtFile.Text = txtFile.Text & "SQLLDR" & mstrVer & " " & gstrUserName & "/" & gstrPassword & IIf(gstrServer = "", "", "@" & gstrServer) _
                    & " CONTROL=" & strPath & strFile & " DIRECT=TRUE" & vbCrLf
            End If
        End If
        strFile = UCase(Dir())
    Loop
    cmdExecute.Enabled = (txtFile.Text <> "")
End Sub

Private Sub lvwTabs_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim strFile As String
    Dim strPath As String
    Dim lst As ListItem
    
    strPath = DirBak.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    txtFile.Text = ""
    strFile = UCase(Dir(strPath))
    For Each lst In lvwTabs.ListItems
        If lst.Checked = True Then
            '��ʾ��ǰ��ִ�е�����
            txtFile.Text = txtFile.Text & "SQLLDR" & mstrVer & " " & gstrUserName & "/" & gstrPassword & IIf(gstrServer = "", "", "@" & gstrServer) _
                & " CONTROL=" & strPath & lst.Text & ".LDR DIRECT=TRUE" & vbCrLf
        End If
    Next
    cmdExecute.Enabled = (txtFile.Text <> "")
End Sub

Private Sub DriveBak_Change()
    Dim strDrive As String
        
    On Error GoTo UnDo
    strDrive = DirBak.Path
    DirBak.Path = DriveBak.Drive
    Exit Sub
UnDo:
    MsgBox "������δ׼����", vbExclamation, gstrSysName
    DriveBak.Drive = strDrive
End Sub

Private Sub cmdExecute_Click()
    Dim strPath As String
    Dim strErrTabs As String
    Dim strErrCons As String
    Dim strCommand As String
    Dim strTable As String
    Dim lngTemp As Long
    Dim lngProcess As Long
    Dim varTime As Variant
    Dim lst As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    strPath = DirBak.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    strErrTabs = ""
    strErrCons = ""
    
    If MsgBox("����������һ���������Ĺ��̣���Ҫ�ƻ��������ݣ�" & vbCrLf & "��׼��������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    SetEnable False
    frmWait.BeginWait "���ڵ������ݡ���"
    varTime = Now() '��¼�¿�ʼ������ʱ��
    
    '��������loader
    For Each lst In lvwTabs.ListItems
        If lst.Checked Then
            strTable = lst.Text
            If chkDelete.value = 1 Then
                '������ñ�����Լ��disable
                gstrSQL = "select 'ALTER TABLE '||D.table_name||' DISABLE CONSTRAINT '||D.constraint_name" & _
                        " from user_constraints U,user_constraints D" & _
                        " where U.table_name='" & strTable & "' and U.constraint_type in('P','U')" & _
                        "       and U.constraint_name=D.r_constraint_name"
                With rsTemp
                    If .State = adStateOpen Then .Close
                    .Open gstrSQL, gcnOracle, adOpenKeyset
                    Do While Not .EOF
                        gcnOracle.Execute CStr(.Fields(0).value)
                        .MoveNext
                    Loop
                End With
                gcnOracle.Execute "truncate table " & mstr������ & "." & strTable & " drop storage"
            End If
            
            strCommand = "SQLLDR" & mstrVer & " " & gstrUserName & "/" & gstrPassword & IIf(gstrServer = "", "", "@" & gstrServer) _
                & " CONTROL=" & strPath & strTable & ".LDR DIRECT=TRUE"
            On Error Resume Next
            frmMDIMain.stbThis.Panels(2).Text = strCommand
            lngTemp = Shell(strCommand, vbHide)
            If err = 0 Then
                On Error GoTo 0
                lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
                Do
                    Sleep 100
                    GetExitCodeProcess lngProcess, lngTemp
                    DoEvents
                Loop While lngTemp = Still_Active
                CloseHandle lngProcess
                
                If lngTemp <> 0 And lngTemp <> 1 Then
                    frmWait.EndWait
                    MsgBox "���ݵ����������ʧ�ܡ�", vbCritical, gstrSysName
                    SetEnable True
                    Exit Sub
                End If
                
            Else
                strErrTabs = strErrTabs & vbCrLf & "��" & strTable
            End If
        End If
    Next
    frmMDIMain.stbThis.Panels(2).Text = ""
    '�ָ�����ʧЧԼ��(enable)
    gstrSQL = "select 'ALTER TABLE '||table_name||' ENABLE CONSTRAINT '||constraint_name,constraint_name" & _
            " from user_constraints" & _
            " where STATUS='DISABLED'"
    With rsTemp
        If .State = adStateOpen Then .Close
        .Open gstrSQL, gcnOracle, adOpenKeyset
        On Error Resume Next
        Do Until .EOF
            gcnOracle.Execute CStr(.Fields(0).value)
            If err <> 0 Then
                strErrCons = strErrCons & vbCrLf & "��" & .Fields(1).value
            End If
            .MoveNext
        Loop
    End With
    
    
    '�ָ�����
    Call AdjustSequence(mstr������, gcnOracle)
    frmWait.EndWait
    SetEnable True
    '�ܽ�
    If strErrTabs <> "" Then
        strErrTabs = vbCrLf & "�����ļ������������ݱ��޷�ִ��Loader:" & vbCrLf & strErrTabs
    End If
    If strErrCons <> "" Then
        strErrTabs = strErrTabs & vbCrLf & "��������ԭ������Լ��������Ч������:" & vbCrLf & strErrCons
    End If
    MsgBox "���ݵ�����ϣ�" & vbCrLf & vbCrLf & _
        "����ʱ" & Format(CDate(Now - varTime), "hhʱmm��ss�롣") & _
        IIf(strErrTabs = "", "", "����" & strErrTabs), vbExclamation, gstrSysName
End Sub

Private Sub SetEnable(ByVal blnEnable As Boolean)
    frmMDIMain.Enabled = blnEnable
    lvwTabs.Enabled = blnEnable
    DirBak.Enabled = blnEnable
    DriveBak.Enabled = blnEnable
    chkDelete.Enabled = blnEnable
    cmdClear.Enabled = blnEnable
    CmdSelect.Enabled = blnEnable
    cmdExecute.Enabled = blnEnable
End Sub

Private Sub Form_Load()
    Dim intVer As Integer
    
    intVer = GetOracleVersion
    
    If intVer < 80 Then
        MsgBox "��Oracle�汾�������ڹ��ɣ���������ܲ����������У�" & vbCr _
            & "�뿼�ǽ�BINĿ¼�е�[IMP+�汾��.EXE]��Ϊ[IMP.EXE]��ִ�С�", vbExclamation, gstrSysName
        mstrVer = ""
    ElseIf intVer = 80 Then            'Oracle8.0
        mstrVer = "80"
    Else
        mstrVer = ""
    End If
    Call FillSystem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsTable.State = 1 Then mrsTable.Close
    Set mrsTable = Nothing
    mstr������ = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtFile.Width = ScaleWidth - 200 - txtFile.Left
    txtFile.Height = ScaleHeight - 200 - txtFile.Top
    
End Sub

Private Sub FillSystem()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    '��ʾ����ʾ��ϵͳ
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", UCase(gstrUserName))
    
    If Not rsTemp.EOF Then
        cmbSystem.AddItem rsTemp("����") & " v" & rsTemp("�汾��") & "��" & rsTemp("���") & "��"
        mstr������ = UCase(gstrUserName)
        
        mrsTable.CursorLocation = adUseClient
        mrsTable.Open "select table_name from all_tables where owner='" & mstr������ & "'", gcnOracle, adOpenStatic, adLockReadOnly
        
        cmbSystem.ListIndex = 0
    Else
        cmbSystem.Enabled = False
        cmdExecute.Enabled = False
        DriveBak.Enabled = False
        DirBak.Enabled = False
        lvwTabs.Enabled = False
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

