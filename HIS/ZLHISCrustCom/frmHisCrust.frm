VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHisCrust 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Զ�����"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   ControlBox      =   0   'False
   Icon            =   "frmHisCrust.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   7050
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdLog 
      Caption         =   "�鿴��־(&C)"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "���(&O)"
      Height          =   375
      Left            =   5325
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ListView lvwMan 
      Height          =   2430
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   4286
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "������Ϣ"
         Object.Width           =   9402
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHisCrust.frx":030A
            Key             =   "Ok"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHisCrust.frx":08A4
            Key             =   "Err"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHisCrust.frx":0E3E
            Key             =   "List"
         EndProperty
      EndProperty
   End
   Begin zlHisCrustCom.UsrProgressBar prgPross 
      Height          =   264
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   6792
      _ExtentX        =   11986
      _ExtentY        =   450
      Color           =   16750899
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ�����������,���Ժ�..."
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1740
      TabIndex        =   4
      Top             =   480
      Width           =   4020
   End
   Begin VB.Image imgUpdate 
      Height          =   720
      Left            =   240
      Picture         =   "frmHisCrust.frx":13D8
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ע�Ჿ��"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   1530
      Width           =   1080
   End
End
Attribute VB_Name = "frmHisCrust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOP                          As Long = 0
Private Const SWP_NOSIZE                        As Long = &H1
Private mblnOperateCompleted                    As Boolean      '�����Ƿ����
Private mblnOK                                  As Boolean      '�Ƿ�û�д���
Private mintColumn                              As Integer      '��ǰ��
Private mintVB6                                 As Integer      'VB6���̵Ĵ������,0-��δ����1-����VB6 2-ɱ��VB6
Private mlngTimes                               As Long         'Tmrִ�д���
Private Enum ErrListCol
    ELC_���� = 0 '��������
    ELC_������Ϣ = 1
End Enum

Private Enum ControlType
    CT_KillProc = 0         'ֻɱ������
    CT_KillProcAndSvr = 1   'ֹͣ����ɱ������
    CT_StartSvr = 2         '���÷���
End Enum

Private Enum UpdateCheck
    UC_IgnorUp = 7          '�����з���ռ�ã���������
    UC_SvrMD5Null = 6       '������MD5Ϊ��
    UC_NotExists = 5        'Ĭ�ϱ��ز�����
    UC_Normal = 4           '��������
    UC_AddtionUp = 3        '���¸���·���ļ�
    UC_RegAgain = 2         '�ϴ�ע��δ�ɹ�����Ҫ�ٴ�ע��
    UC_Update = 1           '��Ҫ���ظ���
    UC_NewDown = 0          '�����أ����ز�����
End Enum
Private mcllOldComs         As Collection '�ϵĲ�����û�����嵥�д���

Private Sub cmdLog_Click()
    Dim lngRet As Long
    Dim strNotPad As String
    
    On Error Resume Next
    strNotPad = gstrSystemPath & "\notepad.exe"
    If gobjFSO.FileExists(strNotPad) Then
        lngRet = ShellExecute(0&, "open", strNotPad, gobjTrace.LogFile, gobjTrace.LogFile, 5)    'SW_SHOW
        If lngRet = 31 Then
           If Not gblnHelperMain Then MsgBox "û���ҵ��ʵ��ĳ���������,�밲װ��Ч�ĳ���!", vbInformation, "�ͻ����Զ�����"
        End If
    Else
        If Not gblnHelperMain Then MsgBox "����û�а�װ���±�����,���ܴ���־�ļ�!" & vbCrLf & "���ֹ������������,���±�·��Ϊ:" & vbCrLf & gobjTrace.LogFile, vbInformation, "�ͻ����Զ�����"
    End If
End Sub

Private Sub cmdOK_Click()
    If Not gblnHelperMain Then Call CallHISEXE(mblnOK)
    Call gobjTrace.CloseLog
    Call gobjMe.ExitApp
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call gobjMe.ExitApp
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    imgUpdate.Width = Me.ScaleWidth
    
    cmdOK.Top = ScaleHeight - cmdOK.Height - 50
    cmdOK.Left = ScaleWidth - cmdOK.Width - 100

    cmdLog.Top = ScaleHeight - cmdLog.Height - 50
    cmdLog.Left = ScaleWidth - cmdLog.Width - cmdOK.Width - 200
End Sub

Private Sub lvwMan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mblnOperateCompleted = False Then Exit Sub
    
    On Error Resume Next
    If mintColumn = ColumnHeader.Index - 1 Then
        lvwMan.SortOrder = IIf(lvwMan.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMan.SortKey = mintColumn
        lvwMan.SortOrder = lvwAscending
    End If
End Sub

Public Sub tmrStart_Timer()
    Dim lstTmp      As ListItem
    Call SetWindowPos(Me.hwnd, HWND_TOP, ((Screen.Width - Me.Width) / 2) / 15, ((Screen.Height - Me.Height) / 2) / 15, 0, 0, SWP_NOSIZE)
    Me.cmdOK.Caption = "ȡ��(&C)"
    If gotCurType = OT_PreUpgrade Or gotCurType = OT_CheckFile Then
        Me.Hide
    Else
        Me.Show
    End If
    mlngTimes = 0
    prgPross.Value = 0
    prgPross.Min = 0
    prgPross.Max = 100
    lblInfor.Caption = "���ذ�װ�嵥����..."
    
    mintVB6 = 0
    'ɱ��HIS��ؽ���
    Call ControlProcAndSvr
    MousePointer = 11
    If FileUpgrade = False Then
        gobjTrace.WriteSection "��������", SL_LevelOne
        gobjTrace.WriteInfo "��������", "�������", "����һ����������������ע��δ�ɹ�"
        cmdOK.Caption = "ȡ��(&C)"
        MousePointer = 0
        If gblnHelperMain Then Call cmdOK_Click
    Else
        cmdOK.Caption = "���(&O)"
        cmdOK.Visible = False
        '����������˳�
        mblnOK = True
        gobjTrace.WriteSection "��������", SL_LevelOne
        gobjTrace.WriteInfo "��������", "�������", "�����ɹ�"
        Call cmdOK_Click
        MousePointer = 0
    End If
End Sub

Private Function ControlProcAndSvr(Optional ByVal lngCurPro As Long, Optional ByVal lngCurIncPro As Long, Optional ByVal ctType As ControlType = CT_KillProc) As Boolean
'���ܣ����ƽ�������񣬳�����ɱ��ֹͣ
'������lngCurPro=��ǰ����
'      lngCurIncPro=��ǰ����ִ����Ϻ����������
'      ctType=0-�����з�����,1-���н��̵�ֹͣ�����н���ɱ��,2-ֻ���÷���
'���أ�����������Ԥ֪������ΪFalse,����ΪTrue
    Dim lngHwnd     As Long, lngZlhisHwnd   As Long, lngVBHwnd      As Long
    Dim lngProcess  As Long, lngPid         As Long
    Dim i           As Long, lngTotal       As Long
    Dim strReturn   As String, strErr       As String
    Dim objShell    As New clsShell
    
    On Error GoTo ErrH
    '���Ԥ����,���˳���
    If gotCurType = OT_PreUpgrade Or gotCurType = OT_CheckFile Then
        ControlProcAndSvr = True
        prgPross.Value = lngCurPro + lngCurIncPro
        Exit Function
    End If
    If ctType = CT_KillProcAndSvr Then
        grsFileUpgrade.Filter = "�Զ�ע��=" & RegFileType.RFT_NETServer
        lngTotal = grsFileUpgrade.RecordCount
        For i = 1 To lngTotal
            prgPross.Value = lngCurPro + lngCurIncPro * 0.2 * (i / lngTotal)
            If gobjFSO.FileExists(grsFileUpgrade!ʵ��·��) Then
                gobjTrace.WriteInfo "ControlProcAndSvr", "ֹͣ����", gobjFSO.GetBaseName(grsFileUpgrade!ʵ��·��)
                lblInfor.Caption = "����ֹͣ����" & gobjFSO.GetBaseName(grsFileUpgrade!ʵ��·��)
                If objShell.Run("NET STOP " & gobjFSO.GetBaseName(grsFileUpgrade!ʵ��·��), strReturn, strErr, 30000) Then
                End If
                gobjTrace.WriteInfo "ControlProcAndSvr", "�������", strReturn, "������Ϣ", strErr
            End If
            grsFileUpgrade.MoveNext
        Next
    End If
    If ctType <> CT_StartSvr Then
        Do While True
            lblInfor.Caption = "���ڼ�������̣�ZLHIS+��VB6.EXE"
            prgPross.Value = lngCurPro + lngCurIncPro * 0.3
            lngHwnd = FindWindow(vbNullString, "����̨")
            If lngHwnd = 0 Then
               lngHwnd = FindWindow(vbNullString, "ҽԺ��Ϣϵͳ")
               If lngHwnd = 0 Then
                   Exit Do
               End If
            End If
            If lngHwnd <> 0 Then
                '�����Ƿ���VB�ڵ��õ���̨���ǳ���ֱ��ִ�е���̨
                lngZlhisHwnd = FindExitsProcess("ZLHIS+.EXE")
                If lngZlhisHwnd <> 0 Then
                    gobjTrace.WriteInfo "KillHisProcess", "ɱ������", "ZLHISEXE"
                    Call TerminateProcess(lngZlhisHwnd, 1&)
                Else
                    lngVBHwnd = FindExitsProcess("VB6.EXE")
                    If lngVBHwnd <> 0 Then
                        If mintVB6 = 0 Then
                            If Not gblnHelperMain Then
                                If MsgBox("���������⵽VB6�����˿��ܻ������Ĳ���." & vbCrLf & "Ϊ�˱�֤ϵͳ��������,�Ƿ�ر�VB6����!", vbQuestion + vbYesNo, "�ͻ����Զ�����") = vbYes Then
                                    Call GetWindowThreadProcessId(lngVBHwnd, lngPid)
                                    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                                    Call TerminateProcess(lngProcess, 1&)
                                    gobjTrace.WriteInfo "KillHisProcess", "ɱ������", "VB6EXE"
                                    mintVB6 = 2
                                Else
                                    mintVB6 = 1
                                    Exit Do
                                End If
                            Else
                                mintVB6 = 1
                                Exit Do
                            End If
                        ElseIf mintVB6 = 2 Then
                            Call GetWindowThreadProcessId(lngVBHwnd, lngPid)
                            lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                            Call TerminateProcess(lngProcess, 1&)
                            gobjTrace.WriteInfo "KillHisProcess", "ɱ������", "VB6EXE"
                        Else
                            Exit Do
                        End If
                    Else
                        Call GetWindowThreadProcessId(lngHwnd, lngPid)
                        lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                        Call TerminateProcess(lngProcess, 1&)
                        gobjTrace.WriteInfo "KillHisProcess", "ɱ������", "VB6EXE"
                    End If
                End If
            End If
        Loop
        '���ڼ��ZLHISCrust.exe����������
        lblInfor.Caption = "���ڼ�������̣�ZLHISCRUST.EXE"
        prgPross.Value = lngCurPro + lngCurIncPro * 0.4
        lngHwnd = FindExitsProcess("ZLHISCRUST.EXE", GetCurrentProcessId)
        If lngHwnd <> 0 Then
            gobjTrace.WriteInfo "KillHisProcess", "ɱ������", "ZLHISCRUST.EXE(����)"
            Call TerminateProcess(lngHwnd, 1&)
        End If
        lngTotal = UBound(garrKillProcess) + 1
        For i = LBound(garrKillProcess) To UBound(garrKillProcess)
            If garrKillProcess(i) <> "VB6.EXE" And garrKillProcess(i) <> "ZLHISCRUST.EXE" Then
                lblInfor.Caption = "���ڼ�������̣�" & garrKillProcess(i)
                prgPross.Value = lngCurPro + lngCurIncPro * 0.4 + lngCurIncPro * 0.6 * (i + 1) / lngTotal
                lngHwnd = FindExitsProcess(garrKillProcess(i))
                If lngHwnd <> 0 Then
                    gobjTrace.WriteInfo "KillHisProcess", "ɱ������", garrKillProcess(i)
                    Call TerminateProcess(lngHwnd, 1&)
                End If
            End If
        Next
    End If
    If ctType = CT_StartSvr Then
        grsFileUpgrade.Filter = "�Զ�ע��=" & RegFileType.RFT_NETServer
        lngTotal = grsFileUpgrade.RecordCount
        For i = 1 To lngTotal
            prgPross.Value = lngCurPro + lngCurIncPro * (i / lngTotal)
            If gobjFSO.FileExists(grsFileUpgrade!ʵ��·��) Then
                gobjTrace.WriteInfo "ControlProcAndSvr", "��������", gobjFSO.GetBaseName(grsFileUpgrade!ʵ��·��)
                lblInfor.Caption = "������������" & gobjFSO.GetBaseName(grsFileUpgrade!ʵ��·��)
                If objShell.Run("NET STOP " & gobjFSO.GetBaseName(grsFileUpgrade!ʵ��·��), strReturn, strErr, 30000) Then
                End If
                gobjTrace.WriteInfo "ControlProcAndSvr", "�������", strReturn, "������Ϣ", strErr
            End If
            grsFileUpgrade.MoveNext
        Next
    End If
    prgPross.Value = lngCurPro + lngCurIncPro
    ControlProcAndSvr = True
    Exit Function
ErrH:
    gobjTrace.WriteInfo "ControlProcAndSvr", "������̿��Ƴ���", Err.Description
    Err.Clear
End Function

Private Function FileUpgrade() As Boolean
''���ܣ������������߼�����
    Dim strErr          As String, lngRet   As Long
    Dim lngCurPro       As Long, i          As Long, lngTotal       As Long
    Dim blnOperateOK    As Boolean
    
    
    blnOperateOK = False
    On Error GoTo ErrH
    gobjTrace.WriteSection "�ļ�����", SL_LevelTwo
    '����ļ�����
    If Not CheckUpdate(prgPross.Value, IIf(gotCurType = OT_PreUpgrade, 10, 10)) Then GoTo ErrEnd
    If gotCurType <> OT_CheckFile Then
        '�ļ����ؽ�ѹ
        If Not DownAndDecFiles(prgPross.Value, IIf(gotCurType = OT_PreUpgrade, 60, 60)) Then GoTo ErrEnd
    End If
    
    prgPross.Value = 70
    'Ԥ��������Ҫ��Щ
    If gotCurType <> OT_PreUpgrade And gotCurType <> OT_CheckFile Then
        'ɱ������
        Call ControlProcAndSvr(prgPross.Value, 3)
        'ɾ�������ļ�
        If Not DeleteExpiredFile(prgPross.Value, 3) Then GoTo ErrEnd
        If gotCurType = OT_Repair Then
            grsFileUpgrade.Filter = "(����<" & UC_Normal & " And ������Ϣ=NULL) Or (����=" & UC_Normal & " And �����ļ�·��<>NULL) OR  (����=" & UC_Normal & " And �����ļ�·��=NULL And �ļ�����<>" & FT_System & ")"
        Else
            grsFileUpgrade.Filter = "(����<" & UC_Normal & " And ������Ϣ=NULL) Or (����=" & UC_Normal & " And �����ļ�·��<>NULL)"
        End If
        If grsFileUpgrade.RecordCount <> 0 Then
            'ɱ������
            Call ControlProcAndSvr(prgPross.Value, 3, CT_KillProcAndSvr)
            '�ļ���װע��
            If Not SetupFiles(prgPross.Value, 10) Then GoTo ErrEnd
            '���÷���
            Call ControlProcAndSvr(prgPross.Value, 3, CT_StartSvr)
            '�������ļ�ִ��
            If Not ExecBatFile(prgPross.Value, 3) Then GoTo ErrEnd
        Else
            gobjTrace.WriteInfo "FileUpgrade", "��װ������", "��������Ҫ������߰�װ���ļ�"
        End If
    ElseIf gotCurType = OT_CheckFile Then
        If Not ReportCheckInfo(blnOperateOK) Then GoTo ErrEnd
    End If
    
    If gotCurType <> OT_CheckFile Then
        prgPross.Value = 95
        'ɱ��7Z.exe����
        lblInfor.Caption = "��������7Z.EXE����"
        lngRet = FindExitsProcess("7Z.EXE")
        If lngRet <> 0 Then Call TerminatePID(lngRet)
        lblInfor.Caption = "���ڹر��ļ�����������"
        Call gclsConnect.CloseConnect
    End If
    
    If gotCurType <> OT_CheckFile And gotCurType <> OT_PreUpgrade Then
        If Not LoadDetailList(prgPross.Value, 3) Then GoTo ErrEnd
        blnOperateOK = lvwMan.ListItems.Count = 0
        prgPross.Value = 99
        lblInfor.Caption = "����������ʱĿ¼"
        Call ClearFolder(gstrSetupPath & "\ZLUPTMP", blnOperateOK)

        '�ٴ���������̨ʱ�����¼�ⰲװ����
        If gblnReCheckComs Then
            gobjTrace.WriteInfo "FileUpgrade", "���¼�ⰲװ����", True
            SaveSetting "ZLSOFT", "ע����Ϣ", "��������", ""
        End If
    End If
    
    If Not blnOperateOK Then
        Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:" & Decode(gotCurType, OT_OfficialUpgrade, "����ʧ��", OT_Repair, "�޸�ʧ��", OT_PreUpgrade, "Ԥ����ʧ��", OT_CheckFile, "������,��Ҫ����") & " ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        lblInfor.Caption = "���ڱ�Ǳ�������״̬"
        Call SetOperateProcess(gotCurType, IIf(gotCurType = OT_CheckFile, OS_Failure, OS_NotInProcessing), SumErrMsg, glngFileBatch)  '��ʶ�������
    Else
        Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:" & Decode(gotCurType, OT_OfficialUpgrade, "�����ɹ�", OT_Repair, "�޸��ɹ�", OT_PreUpgrade, "Ԥ����ʧ��", OT_CheckFile, "������,��������") & " ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        lblInfor.Caption = "���ڱ�Ǳ�������״̬"
        Call SetOperateProcess(gotCurType, OS_Completed, SumErrMsg, glngFileBatch)  '��ʶ�������
        
        If gotCurType = OT_OfficialUpgrade Then
            lblInfor.Caption = "���ڼ�ⰲװOO4O���"
            gobjTrace.WriteSection "OO4O��װ", SL_LevelThree
            If Not InstallOO4O(strErr) Then
                gobjTrace.WriteInfo "InstallOO4O", "��װ��װOO4O���", "ʧ�ܣ�" & strErr
            Else
                gobjTrace.WriteInfo "InstallOO4O", "��װ��װOO4O���", "�ɹ���" & strErr
            End If
        End If
    End If
    prgPross.Value = 100
    lblInfor.Caption = "��������"
    cmdOK.Visible = True
    mblnOperateCompleted = True
    If gotCurType <> OT_CheckFile And gotCurType <> OT_PreUpgrade Then
        FileUpgrade = blnOperateOK
    Else
        FileUpgrade = True
    End If
    Exit Function
ErrH:
    prgPross.Value = 100
    gobjTrace.WriteInfo "FileUpgrade", "�������̷�����������", Err.Description
    If Not gblnHelperMain Then MsgBox "�������̷���������������ϵ����Ա����Ϣ��" & Err.Description, vbInformation, App.Title
    Err.Clear
ErrEnd:
    Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:����ʧ�� ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
    Call SetOperateProcess(gotCurType, IIf(gotCurType = OT_CheckFile, OS_Failure, OS_NotInProcessing), SumErrMsg, glngFileBatch)       '��ʶ��������
    FileUpgrade = False
    cmdOK.Visible = True
    mblnOperateCompleted = True
End Function

Private Function LoadDetailList(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
    Dim strFile     As String, strErr       As String
    Dim lstTmp      As ListItem, intKey     As Integer
    Dim intIndex    As Integer
    Dim lngTotal    As Long, i              As Long
    
    On Error GoTo ErrH
    gobjTrace.WriteSection "������Ҫ", SL_LevelTwo
    grsFileUpgrade.Filter = ""
    grsFileUpgrade.Sort = "����,�ļ���"
    lngTotal = grsFileUpgrade.RecordCount
    lblInfor.Caption = "���ڼ��ش����嵥"
    lvwMan.ListItems.Clear
    With grsFileUpgrade
        Do While Not .EOF
            i = i + 1
            prgPross.Value = lngCurPro + lngCurIncPro * 0.2 * (i / lngTotal)
            If !���� <= UC_Normal Then
                If !������Ϣ & "" <> "" Then
                    intKey = intKey + 1
                    Set lstTmp = lvwMan.ListItems.Add(, "K" & intKey, !�ļ���, "List", "List")
                    lstTmp.SubItems(ELC_������Ϣ) = !������Ϣ & ""
                    lstTmp.SmallIcon = "Err"
                    gobjTrace.WriteInfo "LoadDetailList", !�ļ���, "ʧ��", "��Ϣ", !������Ϣ
                ElseIf !���� < UC_Normal Then
                    gobjTrace.WriteInfo "LoadDetailList", !�ļ���, "�ɹ�"
                Else
                    gobjTrace.WriteInfo "LoadDetailList", !�ļ���, "�������", "��Ϣ", !�����Ϣ
                End If
            ElseIf !���� > UC_SvrMD5Null Then
                gobjTrace.WriteInfo "LoadDetailList", !�ļ���, "�ɹ����Ǵ��ھ���", "����", !�����Ϣ
            Else
                gobjTrace.WriteInfo "LoadDetailList", !�ļ���, "�������", "��Ϣ", !�����Ϣ
            End If
            .MoveNext
        Loop
    End With
    If lvwMan.ListItems.Count <> 0 Then
        Me.Height = 5445
        Me.Refresh
    End If
    prgPross.Value = lngCurPro + lngCurIncPro
    LoadDetailList = True
    Exit Function
ErrH:
    strErr = Err.Description
    gobjTrace.WriteInfo "CheckUpdate", "���ش����嵥��������", strErr
    If Not gblnHelperMain Then MsgBox "���ش����嵥������������ϵ����Ա����Ϣ��" & strErr, vbInformation, App.Title
    Call RecordErrMsg(MT_ChcekUpdate, "���ش����嵥��������", strErr)
    Err.Clear
End Function

Private Function CheckUpdate(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
'���ܣ���鲢��ȡ�����ļ�
'������lngCurPro=��ǰ����
'      lngCurIncPro=��ǰ����ִ����Ϻ����������
'���أ�����������Ԥ֪������ΪFalse,����ΪTrue
    Dim rsFileList   As ADODB.Recordset, ucUpdate       As UpdateCheck, intPreDown  As Integer
    Dim lngRecCount     As Long, i                      As Long, lngBeach           As Integer
    Dim strFile         As String, strWrongFile         As String, strAddSetFile    As String
    Dim arrComs         As Variant
    Dim rsTmp           As ADODB.Recordset
    Dim strlocVersion   As String, strLocModiTime       As String, strLocMd5        As String
    Dim lngTotal        As Long, lngLoop                As Long
    Dim strTmpErr       As String, lngSort               As Long, strNoSubfix       As String
    Dim strOldFile      As String
    
    On Error GoTo ErrH
    Set mcllOldComs = New Collection
    gobjTrace.WriteSection "���¼��(0)", SL_LevelThree
    lblInfor.Caption = "���ڽ����ļ����¼��..."
    '��ȡԤ�������ص����������ļ�
    If gobjFSO.FileExists(gstrPreTempPath & "\ZLList.adtg") And (gotCurType = OT_OfficialUpgrade Or gotCurType = OT_Repair) Then
        Set rsFileList = New ADODB.Recordset
        rsFileList.Open gstrPreTempPath & "\ZLList.adtg", , adOpenStatic, adLockOptimistic, adCmdFile
        On Error Resume Next
        rsFileList.Sort = "�ļ���"
        If Err.Number = 0 Then
            lngRecCount = rsFileList.RecordCount
            gobjTrace.WriteInfo "CheckUpdate", "Ԥ�����嵥��¼", lngRecCount
        Else
            gobjTrace.WriteInfo "CheckUpdate", "Ԥ�����嵥��¼", Err.Description
            Err.Clear
        End If
        On Error GoTo ErrH
    End If
    
    With grsFileUpgrade
        '�Ƚ���·��ת�������жϱ����ļ��Ƿ����,�����м򵥵������ж�
        lngTotal = grsFileUpgrade.RecordCount
        For lngLoop = 1 To lngTotal
            gobjTrace.WriteSection "-", SL_LevelThree
            lblInfor.Caption = "���ڼ���ļ���" & !�ļ���
            prgPross.Value = lngCurPro + lngCurIncPro * 0.75 * (lngLoop / lngTotal)
            strFile = GetActualPath(!��װ·�� & "", Val(!�ļ����� & ""), !�ļ���)
            lngSort = Decode(Val(!�ļ����� & ""), FT_Apply, 5, FT_Public, 4, FT_System, 3, FT_AdditionFile, 0, FT_Help, 0, FT_Other, 2, 1)
            strNoSubfix = !��׼�ļ���
            If InStr(strNoSubfix, ".") > 0 Then
                strNoSubfix = Mid(strNoSubfix, 1, InStrRev(strNoSubfix, ".") - 1)
            End If
            gobjTrace.WriteInfo "CheckUpdate", "����", !�ļ���, "��װ·��", !��װ·��, _
                            "�Զ�ע��", Decode(!�Զ�ע��, RFT_NotReg, "��ע��", RFT_NormalReg, "�Զ�ʶ��ע��", RFT_NETGAC, "NETȫ�ֳ��򼯻���ע��", RFT_NETServer, "NETϵͳ������ע��", RFT_NETComReg, "NETCOMע��"), _
                            "�ļ�����", Decode(!�ļ�����, FT_Apply, "ҵ�񲿼�", FT_Public, "��������", FT_System, "ϵͳ�ļ�", FT_AdditionFile, "�����ļ�", FT_Help, "�����ļ�", FT_Other, "�����ļ�", "δʶ����ļ�"), _
                            "ǿ�Ƹ���", !ǿ�Ƹ���, "ҵ�񲿼�", !ҵ�񲿼�, "MD5", !MD5, "�޸�����", !�޸�����, "�ļ��汾", !�汾��, "���Ӱ�װ·��", !���Ӱ�װ·��
            intPreDown = 0: ucUpdate = UC_NotExists: lngBeach = 0: strTmpErr = "": strWrongFile = "": strAddSetFile = ""
            If InStr(",ZLMIPCLIENTSHELL.EXE,ZLIDKIND.OCX,ZLIDCARD.DLL,ZLICCARD.DLL,ZLREGISTER.DLL,ZL9COMLIB.DLL,ZLLOGIN.DLL,ZLHIS+.EXE,", "," & !��׼�ļ��� & ",") > 0 Then
                lngBeach = -1
            ElseIf !�Զ�ע�� = RFT_NotReg Then
                lngBeach = -2
            End If
            If lngRecCount <> 0 Then
                rsFileList.Filter = "��׼�ļ���='" & !��׼�ļ��� & "'"
                If Not rsFileList.EOF Then
                    If !MD5 = rsFileList!MD5 Then
                        If gobjFSO.FileExists(gstrPreTempPath & "\" & rsFileList!�ļ��� & ".7z") Then
                            ucUpdate = rsFileList!���� '��Ԥ������ȡ�ļ�
                            intPreDown = 1
                            lngBeach = rsFileList!�ж�����
                            strTmpErr = rsFileList!�����Ϣ & ""
                        End If
                    End If
                    Call rsFileList.Delete
                    Call rsFileList.UpdateBatch
                    lngRecCount = lngRecCount - 1
                End If
            End If
            '�ж��ܷ�����
            If ucUpdate = UC_NotExists Then
                If Not IsNull(!MD5) Then
                    If gobjFSO.FileExists(strFile) Then
                        gobjTrace.WriteInfo "CheckUpdate", "���ڱ����ļ�", strFile
                        If !�ļ����� = FT_System And !ǿ�Ƹ��� = 0 Then
                            ucUpdate = UC_Normal
                            strTmpErr = "���ļ���ϵͳ�ļ������ش����Ҳ���Ҫǿ�Ƹ���,��������"
                        Else
                            strLocMd5 = FileMD5(strFile)
                            If !MD5 = strLocMd5 Then
                                ucUpdate = UC_Normal
                                strTmpErr = "���غͷ�����MD5��ͬ,��������"
                            Else
                                ucUpdate = UC_Update
                                strTmpErr = "���غͷ�����MD5����ͬ,��Ҫ����"
                            End If
                        End If
                    ElseIf !�ļ����� = FT_Apply Then
                        If IsNull(!ҵ�񲿼�) Then 'Ӧ�ò�����ҵ�񲿼�Ϊ�ղ�����
                            ucUpdate = UC_NotExists
                            strTmpErr = "���ļ���Ӧ�ò���,���ز����ڵ�ҵ�񲿼�Ϊ��,��������"
                        ElseIf UCase(!ҵ�񲿼�) = !��׼�ļ��� Then 'Ӧ�ò�����ҵ�񲿼���������ǿ������
                            ucUpdate = UC_NewDown
                            strTmpErr = "���ļ���Ӧ�ò���,���ز����ڵ�ҵ�񲿼���������,��Ҫ����"
                        End If
                    Else '��ͨ�ļ������ھ�����
                        ucUpdate = UC_NewDown
                        strTmpErr = "���ļ��Ƿ�Ӧ�ò���,���ز���,��Ҫ����"
                    End If
                Else '������MD5Ϊ�ղ���������
                    ucUpdate = UC_SvrMD5Null '��ʾ������MD5Ϊ��
                    strTmpErr = "ZLFilesUpgrade��û�и��ļ���MD5��Ϣ���޷����в����������"
                End If
            End If
            '��ȡ����·���ļ�
            '�����ļ�·�����ܰ������Ӱ�װ·���е�·��,�˴������Ƿ�ע�ᣬʹ���߷ֿ���
            If !�Զ�ע�� > RFT_NotReg Then '��Ҫע����ļ����������·��
                strWrongFile = GetWrongFiles(!�ļ���, strFile)
                If strWrongFile <> "" Then '���ڴ���·���ļ��������ļ������أ����Զ����Ϊ���ظ���
                    If ucUpdate = UC_NotExists Then
                        ucUpdate = UC_NewDown
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "�ļ�·������(����·����" & strWrongFile & "),�����Ҫ��������"
                    ElseIf ucUpdate <> UC_SvrMD5Null Then
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "�ļ�·������(����·����" & strWrongFile & "),�����Ҫ����ע��"
                    End If
                End If
            Else
                strAddSetFile = GetAdditionSetup(!�ļ���, !MD5 & "", !���Ӱ�װ·�� & "")
                If strAddSetFile <> "" Then '���ڸ��Ӱ�װ·���������ļ������أ����Զ����Ϊ���ظ���
                    If ucUpdate = UC_NotExists Then
                        ucUpdate = UC_NewDown
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "��Ҫ���Ӱ�װ�ļ�(���Ӱ�װ·����" & strAddSetFile & "),�����Ҫ��������"
                    ElseIf ucUpdate <> UC_SvrMD5Null Then
                        If ucUpdate = UC_Normal Then ucUpdate = UC_AddtionUp
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "��Ҫ���Ӱ�װ�ļ�(���Ӱ�װ·����" & strAddSetFile & "),�����Ҫ���°�װ�����ļ�"
                    Else
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "��Ҫ���Ӱ�װ�ļ�(���Ӱ�װ·����" & strAddSetFile & "),���Ƿ������ļ�MD5Ϊ�գ��޷�����"
                    End If
                End If
            End If
            grsFileUpgrade.Update Array("����", "ʵ��·��", "�����ļ�·��", "����ʵ��·��", "Ԥ��������", "�ж�����", "�����Ϣ", "�޺�׺�ļ���", "��������"), _
                                Array(ucUpdate, strFile, IIf(strWrongFile = "", Null, strWrongFile), IIf(strAddSetFile = "", Null, strAddSetFile), intPreDown, lngBeach, IIf(strTmpErr = "", Null, strTmpErr), strNoSubfix, lngSort)
            gobjTrace.WriteInfo "CheckUpdate", "����", ucUpdate, "Ԥ��������", intPreDown, "�ж�����", lngBeach, "ʵ��·��", strFile, "�����ļ�·��", strWrongFile, "����ʵ��·��", strAddSetFile, "�޺�׺�ļ���", strNoSubfix, "��������", lngSort
            gobjTrace.WriteInfo "CheckUpdate", "���˵��(0)", strTmpErr
            grsFileUpgrade.MoveNext
        Next
        '��ѯ������ÿ�ζ�û���ж������ұ��ز����ڵĴ���ҵ�񲿼����õ�Ӧ�ò��������ж�
        '�жϵĲο�����ÿ���жϱ䶯���ļ���
        grsFileUpgrade.Filter = ""
        lngRecCount = 0: lngBeach = 0
        Set rsTmp = CopyNewRec(grsFileUpgrade)
        grsFileUpgrade.Filter = "����=" & UC_NotExists & " And ҵ�񲿼�<>NULL And �ļ�����=" & FT_Apply
        Do While lngRecCount <> grsFileUpgrade.RecordCount
            lngRecCount = grsFileUpgrade.RecordCount
            lngBeach = lngBeach + 1
            If lngRecCount > 0 Then gobjTrace.WriteSection "���¼��(" & lngBeach & ")", SL_LevelThree
            For lngLoop = 1 To lngRecCount
                gobjTrace.WriteSection "-", SL_LevelThree
                lblInfor.Caption = "���ڼ���ļ���" & !�ļ���
                prgPross.Value = lngCurPro + lngCurIncPro * 0.75 + lngCurIncPro * 0.25 * IIf(lngBeach > 3, 1, (lngBeach / 3) * (lngLoop / lngRecCount))
                'ҵ�񲿼��ж�
                ucUpdate = UC_NotExists: strTmpErr = !�����Ϣ & ""
                arrComs = Split(UCase(grsFileUpgrade!ҵ�񲿼�), ",")
                For i = LBound(arrComs) To UBound(arrComs)
                    If arrComs(i) Like "*.*" Then 'ҵ�񲿼�����׺
                        rsTmp.Filter = "��׼�ļ���='" & arrComs(i) & "'"
                    Else
                        rsTmp.Filter = "�޺�׺�ļ��� = '" & arrComs(i) & "'"
                    End If
                    If Not rsTmp.EOF Then
                        If rsTmp!���� < UC_NotExists Then     '��Ҫ���»��Ȿ�ش����Ҳ���Ҫ����
                            ucUpdate = UC_NewDown
                            If rsTmp!���� = UC_Update Then
                                strTmpErr = "ҵ�񲿼�""" & rsTmp!�ļ��� & """���ش�������Ҫ���£���˸ò�����Ҫ����"
                            ElseIf rsTmp!���� = UC_NewDown Then
                                strTmpErr = "ҵ�񲿼�""" & rsTmp!�ļ��� & """�Ѿ����Ϊ���أ���˸ò�����Ҫ����"
                            Else 'UC_normal,���ش����Ҳ���Ҫ����
                                strTmpErr = "ҵ�񲿼�""" & rsTmp!�ļ��� & """���ش��ڵ��ǲ���Ҫ���£���˸ò�����Ҫ����"
                            End If
                        ElseIf rsTmp!���� = UC_SvrMD5Null And gobjFSO.FileExists(rsTmp!ʵ��·��) Then
                            ucUpdate = UC_NewDown
                            strTmpErr = "ҵ�񲿼�""" & rsTmp!�ļ��� & """���ش��ھ���ҵ�񲿼�������MD5Ϊ�գ���˸ò�����Ҫ����"
                        End If
                    ElseIf IsOldComponentExists(arrComs(i), strOldFile) Then
                        ucUpdate = UC_NewDown
                        strTmpErr = "ҵ�񲿼�""" & strOldFile & """���ش��ڣ��������Ѿ�����ʹ�õ�ҵ�񲿼��������嵥�в����ڸ�ҵ�񲿼�������˸ò�����Ҫ����"
                    End If
                    If ucUpdate = UC_NewDown Then Exit For
                Next
                If ucUpdate = UC_NewDown Then
                    grsFileUpgrade.Update Array("����", "�ж�����", "�����Ϣ"), Array(ucUpdate, lngBeach, IIf(strTmpErr = "", Null, strTmpErr))
                    gobjTrace.WriteInfo "CheckUpdate", "����", !�ļ���, "��װ·��", !��װ·��, _
                                    "�Զ�ע��", Decode(!�Զ�ע��, RFT_NotReg, "��ע��", RFT_NormalReg, "�Զ�ʶ��ע��", RFT_NETGAC, "NETȫ�ֳ��򼯻���ע��", RFT_NETServer, "NETϵͳ������ע��", RFT_NETComReg, "NETCOMע��"), _
                                    "�ļ�����", Decode(!�ļ�����, FT_Apply, "ҵ�񲿼�", FT_Public, "��������", FT_System, "ϵͳ�ļ�", FT_AdditionFile, "�����ļ�", FT_Help, "�����ļ�", FT_Other, "�����ļ�", "δʶ����ļ�"), _
                                    "ǿ�Ƹ���", !ǿ�Ƹ���, "ҵ�񲿼�", !ҵ�񲿼�, "MD5", !MD5, "�޸�����", !�޸�����, "�ļ��汾", !�汾��
                    gobjTrace.WriteInfo "CheckUpdate", "����", ucUpdate, "�ж�����", lngBeach
                    gobjTrace.WriteInfo "CheckUpdate", "���˵��(" & lngBeach & ")", strTmpErr
                Else
                    gobjTrace.WriteInfo "CheckUpdate", "����", !�ļ���, "���˵��(" & lngBeach & ")", "����"
                End If
                grsFileUpgrade.MoveNext
            Next
            grsFileUpgrade.Filter = "�ж�����=" & lngBeach '��ȡ�����ж���Ҫ���صĲ���
            If grsFileUpgrade.EOF Then Exit Do
            Set rsTmp = CopyNewRec(grsFileUpgrade)
            grsFileUpgrade.Filter = "����=" & UC_NotExists & " And ҵ�񲿼�<>NULL And �ļ�����=" & FT_Apply
        Loop
        '�ϴ�ע����
        '��ȡԤ�������ص����������ļ�
        If gobjFSO.FileExists(gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg") And (gotCurType = OT_OfficialUpgrade Or gotCurType = OT_Repair) Then
            Set rsFileList = New ADODB.Recordset
            rsFileList.Open gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg", , adOpenStatic, adLockOptimistic, adCmdFile
            On Error Resume Next
            rsFileList.Sort = "�ļ���"
            If Err.Number = 0 Then '��Ϊ����Ч���м����ܱ����ء�
                lngRecCount = rsFileList.RecordCount
                gobjTrace.WriteInfo "CheckUpdate", "�ϴ�ע��ʧ�ܼ�¼", lngRecCount
                lngBeach = lngBeach + 1
                gobjTrace.WriteSection "���¼��(" & lngBeach & ")", SL_LevelThree
                Do While Not rsFileList.EOF
                    strTmpErr = ""
                    .Filter = "��׼�ļ���='" & rsFileList!��׼�ļ��� & "'"
                    If Not .EOF Then
                        '�����ػ��߸��»��ߴ����ļ���Ϊ�յĶ�����Ҫ����ע��ġ�
                        If (!���� = UC_Normal Or !���� = UC_SvrMD5Null And gobjFSO.FileExists(!ʵ��·��)) And !�Զ�ע�� <> RFT_NotReg Then
                            strTmpErr = !�����Ϣ & ""
                            strTmpErr = strTmpErr & IIf(strTmpErr = "", "", ";") & "�ϴ�����ע�᲻�ɹ�����Ҫ����ע��"
                            .Update Array("����", "�����Ϣ"), Array(UC_RegAgain, strTmpErr)
                            gobjTrace.WriteInfo "CheckUpdate", "����", !�ļ���, "��װ·��", !��װ·��, _
                                            "�Զ�ע��", Decode(!�Զ�ע��, RFT_NotReg, "��ע��", RFT_NormalReg, "�Զ�ʶ��ע��", RFT_NETGAC, "NETȫ�ֳ��򼯻���ע��", RFT_NETServer, "NETϵͳ������ע��", RFT_NETComReg, "NETCOMע��"), _
                                            "�ļ�����", Decode(!�ļ�����, FT_Apply, "ҵ�񲿼�", FT_Public, "��������", FT_System, "ϵͳ�ļ�", FT_AdditionFile, "�����ļ�", FT_Help, "�����ļ�", FT_Other, "�����ļ�", "δʶ����ļ�"), _
                                            "ǿ�Ƹ���", !ǿ�Ƹ���, "ҵ�񲿼�", !ҵ�񲿼�, "MD5", !MD5, "�޸�����", !�޸�����, "�ļ��汾", !�汾��
                            gobjTrace.WriteInfo "CheckUpdate", "����", ucUpdate
                            gobjTrace.WriteInfo "CheckUpdate", "���˵��(" & lngBeach & ")", strTmpErr
                        End If
                    End If
                    rsFileList.MoveNext
                Loop
            Else
                gobjTrace.WriteInfo "CheckUpdate", "�ϴ�ע��ʧ�ܼ�¼", Err.Description
                Err.Clear
            End If
            On Error GoTo ErrH
        End If
    End With
    '��δ�ж��Ĳ�������Ϊ��������
    Call UpdateRec(grsFileUpgrade, "����=" & UC_NotExists, "�����Ϣ", "���ز������Ҳ���Ҫ����")
    Call UpdateRec(grsFileUpgrade, "�Զ�ע��=" & RFT_NETServer, "�ж�����", lngBeach + 1)
    prgPross.Value = lngCurPro + lngCurIncPro
    CheckUpdate = True
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    strTmpErr = Err.Description
    gobjTrace.WriteInfo "CheckUpdate", "����ļ����·�����������", strTmpErr
    If Not gblnSilence And Not gblnHelperMain Then MsgBox "����ļ����·���������������ϵ����Ա����Ϣ��" & strTmpErr, vbInformation, App.Title
    Call RecordErrMsg(MT_ChcekUpdate, "����ļ����·�����������", strTmpErr)
    Err.Clear
End Function

Private Function DownAndDecFiles(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
'���ܣ����ز���ѹ�ļ�
'������lngCurPro=��ǰ����
'      lngCurIncPro=��ǰ����ִ����Ϻ����������
'���أ�����������Ԥ֪������ΪFalse,����ΪTrue
    Dim lngTotal        As Long, lngLoop                As Long
    Dim strTmpPath      As String, strErrTmp            As String
    Dim strlocVersion   As String, strLocModiTime       As String, strLocMd5    As String
    Dim strErrInfo      As String
    Dim rsTmp           As ADODB.Recordset
    Dim blnZip          As Boolean
    
    On Error GoTo ErrH
    lblInfor.Caption = "���������ļ�..."
    gobjTrace.WriteSection "���ؽ�ѹ", SL_LevelThree
    With grsFileUpgrade
        .Filter = "����< " & UC_RegAgain
        lngTotal = .RecordCount
        gblnReCheckComs = lngTotal > 0
        For lngLoop = 1 To lngTotal
            gobjTrace.WriteSection "-", SL_LevelThree
            lblInfor.Caption = IIf(gotCurType <> OT_PreUpgrade, IIf(!Ԥ�������� = 0, "�������ز���ѹ�ļ�:", "���ڽ�ѹ�ļ�:"), "���������ļ�:") & !�ļ���
            prgPross.Value = lngCurPro + lngCurIncPro * 0.9 * (lngLoop / lngTotal)
            strErrInfo = ""
            If gotCurType = OT_PreUpgrade Or !Ԥ�������� = 1 Then
                strTmpPath = gstrPreTempPath
            Else
                strTmpPath = gstrTempPath
            End If
            '�Ѿ�������ļ���������Ӻ�׺�Լ���ѹ���ڰ�װ���ѹ���������尲װ���֮���ٽ��е����İ�װ����
            blnZip = !��׼�ļ��� Like "*.7Z"
            If !Ԥ�������� = 0 Then
                If gclsConnect.IsServerFileExists(!��׼�ļ��� & IIf(blnZip, "", ".7z")) Then
                    DoEvents
                    If gclsConnect.DownloadFile(!��׼�ļ��� & IIf(blnZip, "", ".7z"), strTmpPath, strErrTmp) Then
                    Else
                        strErrInfo = "�ļ�����ʧ�ܣ�" & strErrTmp & "(��������" & gclsConnect.ServerPath & ")"
                    End If
                Else
                    strErrInfo = "�������ļ�������(��������" & gclsConnect.ServerPath & ")"
                End If
            End If
            
            If gotCurType <> OT_PreUpgrade And strErrInfo = "" Then
                DoEvents
                If Not blnZip Then
                    If Not gobj7zZip.UnZipFile(strTmpPath & "\" & !�ļ��� & ".7z", strTmpPath & "\" & !�ļ���, , strErrTmp) Then
                        If strErrTmp = "" Then
                            strErrInfo = "��ѹ���ļ�" & strTmpPath & "\" & !�ļ��� & "������,���ܱ�ɱ�����ɱ��"
                        Else
                            strErrInfo = "�ļ���ѹʧ�ܣ�" & strErrTmp
                        End If
                    End If
                End If
                
                If strErrInfo = "" Then
                    If gobjFSO.FileExists(strTmpPath & "\" & !�ļ���) Then
                        strLocMd5 = FileMD5(strTmpPath & "\" & !�ļ���)
                        If !MD5 <> strLocMd5 Then
                            If gblnMD5Check Then
                                strErrInfo = "�������ļ�����(�������ļ����ռ���MD5��ƥ��)(��������" & gclsConnect.ServerPath & ")"
                            Else
                                Call RecordErrMsg(MT_DownAndDec, !�ļ���, "�������ļ�����(�������ļ����ռ���MD5��ƥ��)(��������" & gclsConnect.ServerPath & ")")
                                gobjTrace.WriteInfo "DownAndDecFiles", "�������ļ�����(�������ļ����ռ���MD5��ƥ��)(��������" & gclsConnect.ServerPath & ")"
                            End If
                        End If
                    End If
                End If
            End If
            If strErrInfo <> "" Then
                grsFileUpgrade.Update "������Ϣ", strErrInfo
                Call RecordErrMsg(MT_DownAndDec, !�ļ���, strErrInfo)
                gobjTrace.WriteInfo "DownAndDecFiles", IIf(gotCurType <> OT_PreUpgrade, IIf(!Ԥ�������� = 0, "���ز���ѹ�ļ�", "��ѹ�ļ�"), "�����ļ�"), !�ļ���, "ʧ����Ϣ", strErrInfo
            Else
                gobjTrace.WriteInfo "DownAndDecFiles", IIf(gotCurType <> OT_PreUpgrade, IIf(!Ԥ�������� = 0, "���ز���ѹ�ļ�", "��ѹ�ļ�"), "�����ļ�"), !�ļ���
            End If
            grsFileUpgrade.MoveNext
        Next
        '����Ԥ�����ļ��嵥
        If gotCurType = OT_PreUpgrade Then
            lblInfor.Caption = "���ڱ���Ԥ�����ļ��嵥"
            .Filter = "����<" & UC_RegAgain
            If .RecordCount > 0 Then
                Set rsTmp = CopyNewRec(grsFileUpgrade)
                rsTmp.Sort = "�ļ���"
                If gobjFSO.FileExists(gstrPreTempPath & "\ZLList.adtg") Then
                    gobjFSO.DeleteFile gstrPreTempPath & "\ZLList.adtg", True
                End If
                rsTmp.Save gstrPreTempPath & "\ZLList.adtg", adPersistADTG
                rsTmp.Close
            End If
        End If
        prgPross.Value = lngCurPro + lngCurIncPro
    End With
    DownAndDecFiles = True
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "DownAndDecFiles", "���ؽ�ѹ�ļ�������������", strErrInfo
    Call RecordErrMsg(MT_DownAndDec, "���ؽ�ѹ�ļ�������������", strErrInfo)
    If Not gblnHelperMain Then MsgBox "���ؽ�ѹ�ļ�����������������ϵ����Ա����Ϣ��" & strErrInfo, vbInformation, App.Title
    Err.Clear
End Function

Private Function DeleteExpiredFile(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
'���ܣ�ɾ�������ļ�
    Dim strSQL      As String, rsTmp        As ADODB.Recordset
    Dim rsSys       As ADODB.Recordset
    Dim strFile     As String
    Dim i           As Integer, lngCount    As Long, strErr     As String

    On Error Resume Next
    gobjTrace.WriteSection "���������ļ�", SL_LevelThree
    strSQL = "Select �ļ���,Upper(�ļ���) ��׼�ļ���,��װ·��,ϵͳ���,ϵͳ�汾 From zlFilesExpired"
    Set rsTmp = OpenSQLRecord(strSQL, "�����ļ�")
    lblInfor.Caption = "���ڼ�������ļ�..."
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrH
    If Not rsTmp Is Nothing Then '���ܸñ�����
        If Not rsTmp.EOF Then
            strSQL = "Select ���� �汾��, 0 ��� From Zlreginfo Where ��Ŀ = '�汾��' Union All Select �汾��, ��� From Zlsystems"
            Set rsSys = OpenSQLRecord(strSQL, "��ȡϵͳ�汾")
        End If
        lngCount = rsTmp.RecordCount
        For i = 1 To lngCount
            prgPross.Value = lngCurPro + lngCurIncPro * (i / lngCount)
            lblInfor.Caption = "���ڼ�������ļ���" & rsTmp!�ļ���
            rsSys.Filter = "���=" & Val(rsTmp!ϵͳ��� & "")
            gobjTrace.WriteInfo "DeleteExpiredFile", "����ļ�", rsTmp!�ļ���, "���汾", rsTmp!ϵͳ�汾, "���·��", rsTmp!��װ·��
            If Not rsSys.EOF Then
                '�ļ����ð汾С�ڵ�ǰϵͳ�汾�Ϳ���������
                If VerFull(rsTmp!ϵͳ�汾) <= VerFull(rsSys!�汾��) Then
                    On Error Resume Next
                    strFile = gcllSetPath("K_" & UCase(rsTmp!��װ·��)) & "\" & rsTmp!�ļ���
                    If Err.Number <> 0 Then
                        gobjTrace.WriteInfo "DeleteExpiredFile", "�޷�����", "�ļ�·���޷�ת����" & rsTmp!��װ·��
                        Err.Clear
                        On Error GoTo ErrH
                    Else
                        On Error GoTo ErrH
                        'ֻ�����������ļ��嵥�е��ļ�����Ϊ
                        If gobjFSO.FileExists(strFile) Then
                            grsFileUpgrade.Filter = "��׼�ļ���='" & rsTmp!��׼�ļ��� & "'"
                            If grsFileUpgrade.EOF Then
                                gobjTrace.WriteInfo "DeleteExpiredFile", "�����ļ�", strFile
                                On Error Resume Next
                                If FileSystem.GetAttr(strFile) <> vbNormal Then
                                    FileSystem.SetAttr strFile, vbNormal
                                End If
                                Call gclsRegCom.UnRegCom(strFile)
                                Call gobjFSO.DeleteFile(strFile, True)
                                If Err.Number <> 0 Then Err.Clear
                            Else
                                gobjTrace.WriteInfo "DeleteExpiredFile", "��������", "�ļ��Ѿ�������Ǩ�ļ��б���"
                            End If
                        Else
                            gobjTrace.WriteInfo "DeleteExpiredFile", "��������", "�����ļ�������"
                        End If
                    End If
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    prgPross.Value = lngCurPro + lngCurIncPro
    DeleteExpiredFile = True
    Exit Function
ErrH:
    strErr = Err.Description
    gobjTrace.WriteInfo "DeleteExpiredFile", "���������ļ�������������", strErr
    Call RecordErrMsg(MT_SetUp, "���������ļ�������������", strErr)
    If Not gblnHelperMain Then MsgBox "���������ļ�����������������ϵ����Ա����Ϣ��" & strErr, vbInformation, App.Title
    Err.Clear
End Function

Private Function SetupFiles(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
    Dim arrComs     As Variant, i       As Integer
    Dim lngLoop     As Long, lngTotal   As Long
    Dim strErrInfo  As String, blnCanUp As Boolean, strErrTmp As String
    Dim blnRegErr   As Boolean
    Dim rsTmp           As ADODB.Recordset
    
    gobjTrace.WriteSection "��װע���ļ�", SL_LevelThree
    On Error GoTo ErrH
    With grsFileUpgrade
        If gotCurType = OT_Repair Then
            grsFileUpgrade.Filter = "(����<" & UC_Normal & " And ������Ϣ=NULL) Or (����=" & UC_Normal & " And �����ļ�·��<>NULL) OR  (����=" & UC_Normal & " And �����ļ�·��=NULL And �ļ�����<>" & FT_System & ")"
        Else
            grsFileUpgrade.Filter = "(����<" & UC_Normal & " And ������Ϣ=NULL) Or (����=" & UC_Normal & " And �����ļ�·��<>NULL)"
        End If
        .Sort = "�ж�����,�Զ�ע��,��������"
        lngTotal = .RecordCount
        For lngLoop = 1 To .RecordCount
            gobjTrace.WriteSection "-", SL_LevelThree
            prgPross.Value = lngCurPro + lngCurIncPro * (lngLoop / lngTotal)
            If !���� > 0 Then
                gobjTrace.WriteInfo "SetupFiles", "��װע���ļ�", !�ļ���
            Else
                gobjTrace.WriteInfo "SetupFiles", "��������ļ�", !�ļ���
            End If
            strErrInfo = "": blnCanUp = True: blnRegErr = False
            If Not IsNull(!�����ļ�·��) Then
                On Error Resume Next
                lblInfor.Caption = "���������ļ�:" & !�ļ���
                arrComs = Split(!�����ļ�·��, "|")
                For i = LBound(arrComs) To UBound(arrComs)
                    '�ļ�����,�������
                    If FileSystem.GetAttr(arrComs(i)) <> vbNormal Then
                        FileSystem.SetAttr arrComs(i), vbNormal
                    End If
                    Call gclsRegCom.UnRegCom(arrComs(i), , !�Զ�ע��)
                    Call gobjFSO.DeleteFile(arrComs(i), True)
                    If Err.Number <> 0 Then Err.Clear
                Next
                On Error GoTo ErrH
            End If
            If !���� < UC_RegAgain Then
                lblInfor.Caption = "���ڰ�װ�ļ�:" & !�ļ���
                If SetupOneFile(!��׼�ļ���, IIf(!Ԥ�������� = 1, gstrPreTempPath, gstrTempPath) & "\" & !�ļ���, !ʵ��·��, !ǿ�Ƹ��� = 1, strErrInfo) Then
                    '���ܱ�ռ�ã��Һ���
                    If strErrInfo <> "" Then
                    ElseIf gobjFSO.FileExists(!ʵ��·��) Then
                        gobjTrace.WriteInfo "SetupFiles", "�ɹ���װ�ļ�", !�ļ���
                        '����7zѹ���ļ������Զ���ѹ������ɾ��ѹ���ļ���������һ���жϡ�
                        If !��׼�ļ��� Like "*.7Z" Then
                            If Not gobj7zZip.UnZipFile(!ʵ��·��, Mid(!ʵ��·��, 1, Len(!ʵ��·��) - 3), False, strErrTmp, True) Then
                                gobjTrace.WriteInfo "SetupFiles", "ѹ������ѹʧ��", !�ļ��� & ":" & strErrTmp
                            Else
                                gobjTrace.WriteInfo "SetupFiles", "ѹ������ѹ�ɹ�", !�ļ���
                            End If
                        End If
                    Else
                        blnCanUp = False
                        strErrInfo = "�ļ�" & !ʵ��·�� & "��װ�󲻴���,���ܱ�ɱ�����ɱ��"
                        gobjTrace.WriteInfo "SetupFiles", "��װ�ļ�ʧ��", "�ļ�" & !ʵ��·�� & "��װ�󲻴���,���ܱ�ɱ�����ɱ��"
                    End If
                Else
                    blnCanUp = False
                    gobjTrace.WriteInfo "SetupFiles", "��װ�ļ�ʧ��", strErrInfo
                End If
            End If
            If blnCanUp And strErrInfo = "" And Not IsNull(!����ʵ��·��) Then
                lblInfor.Caption = "���ڽ��и��Ӱ�װ:" & !�ļ���
                arrComs = Split(!����ʵ��·��, "|")
                For i = LBound(arrComs) To UBound(arrComs)
                    If SetupOneFile(!��׼�ļ���, !ʵ��·��, arrComs(i), !ǿ�Ƹ��� = 1, strErrTmp) Then
                        If strErrTmp <> "" Then
                        ElseIf gobjFSO.FileExists(arrComs(i)) Then
                            gobjTrace.WriteInfo "SetupFiles", "�ɹ���װ���Ӱ�װ�ļ�", arrComs(i)
                        Else
                            blnCanUp = False
                            strErrTmp = "���Ӱ�װ�ļ�" & arrComs(i) & "��װ�󲻴���,���ܱ�ɱ�����ɱ��"
                            gobjTrace.WriteInfo "SetupFiles", "��װ���Ӱ�װ�ļ�ʧ��", "�ļ�" & arrComs(i) & "��װ�󲻴���,���ܱ�ɱ�����ɱ��"
                        End If
                    Else
                        blnCanUp = False
                        gobjTrace.WriteInfo "SetupFiles", "��װ���Ӱ�װ�ļ�ʧ��", strErrTmp
                    End If
                    If strErrTmp <> "" Then
                        strErrInfo = strErrInfo & ";" & strErrInfo
                    End If
                Next
            End If
            If strErrInfo = "" And NVL(!�Զ�ע��, 0) <> 0 Then
                lblInfor.Caption = "����ע���ļ�:" & !�ļ���
                If Not gclsRegCom.RegCom(!ʵ��·��, strErrInfo, !�Զ�ע��) Then
                    blnCanUp = False: blnRegErr = True
                    strErrInfo = "ע��ʧ��(" & strErrInfo & ")"
                    gobjTrace.WriteInfo "SetupFiles", "ע���ļ�ʧ��", strErrInfo
                Else
                    If strErrInfo <> "" Then 'ֻ�Ǿ�����Ϣ
                        gobjTrace.WriteInfo "SetupFiles", "ע���ļ�ʧ��", strErrInfo
                    Else
                        gobjTrace.WriteInfo "SetupFiles", "�ɹ�ע���ļ�", !�ļ���
                    End If
                End If
            End If
            
            If strErrInfo <> "" Then
                If blnCanUp Then
                    grsFileUpgrade.Update Array("����", "�����Ϣ"), Array(UC_IgnorUp, strErrInfo)
                Else
                    grsFileUpgrade.Update Array("������Ϣ", "ע�����"), Array(strErrInfo, IIf(blnRegErr, 1, 0))
                End If
                Call RecordErrMsg(MT_SetUp, !�ļ���, strErrInfo)
            End If
            .MoveNext
        Next
        lblInfor.Caption = "���ڱ���ע��ʧ���ļ��嵥"
        .Filter = "ע�����=1"
        If .RecordCount > 0 Then
            Set rsTmp = CopyNewRec(grsFileUpgrade)
            rsTmp.Sort = "�ļ���"
            If gobjFSO.FileExists(gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg") Then
                gobjFSO.DeleteFile gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg", True
            End If
            rsTmp.Save gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg", adPersistADTG
            rsTmp.Close
        Else
            If gobjFSO.FileExists(gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg") Then
                gobjFSO.DeleteFile gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg", True
            End If
        End If
    End With
    
    prgPross.Value = lngCurPro + lngCurIncPro
    SetupFiles = True
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "SetupFiles", "��װע���ļ�������������", strErrInfo
    Call RecordErrMsg(MT_SetUp, "��װע���ļ�������������", strErrInfo)
    If Not gblnHelperMain Then MsgBox "��װע���ļ�����������������ϵ����Ա����Ϣ��" & strErrInfo, vbInformation, App.Title
    Err.Clear
    lblInfor.Caption = "���ڱ���ע��ʧ���ļ��嵥"
    grsFileUpgrade.Filter = "ע�����=1"
    If grsFileUpgrade.RecordCount > 0 Then
        Set rsTmp = CopyNewRec(grsFileUpgrade)
        rsTmp.Sort = "�ļ���"
        If gobjFSO.FileExists(gstrSetupPath & "ZLUPTMP\ZLRegErr.adtg") Then
            gobjFSO.DeleteFile gstrSetupPath & "ZLUPTMP\ZLRegErr.adtg", True
        End If
        rsTmp.Save gstrSetupPath & "ZLUPTMP\ZLRegErr.adtg", adPersistADTG
        rsTmp.Close
    End If
End Function

Public Function SetupOneFile(ByVal strSTFileName As String, ByVal strTmpFile As String, ByVal strSetupFile As String, Optional ByVal blnForceCover As Boolean, Optional ByRef strErrReturn As String) As Boolean
'���ܣ��������ļ����ڰ�װ·��
'˵�����ù��ܶ�����������Ϊ�ù��̴��ڽ϶�Goto���
    Dim blnGoto  As Boolean, sgResult           As VbMsgBoxResult
    Dim cllProcess  As New Collection   '���̼�array(����,Exe�ļ���,ģ�����)
    Dim i           As Long, strMsgBox          As String
    Dim blnReturn   As Boolean, strErr          As String
    
    On Error Resume Next
    If gobjFSO.FileExists(strSetupFile) Then
        If FileSystem.GetAttr(strSetupFile) <> vbNormal Then
            FileSystem.SetAttr strSetupFile, vbNormal
        End If
    End If
    blnGoto = False
SartSetup:
     blnReturn = True: strErrReturn = ""
    If Err.Number <> 0 Then Err.Clear
    If Not gobjFSO.FileExists(strTmpFile) Then
        strErrReturn = "��ѹ����ļ������ڣ����ܱ�ɱ�������ɱ"
        Call RecordErrMsg(MT_SetUp, gobjFSO.GetFileName(strTmpFile), strErrReturn)
        gobjTrace.WriteInfo "SetupFile", "��װʧ��", strErrReturn
        Exit Function
    End If
    '2����ʼ�������Լ��������������������
    gobjFSO.CopyFile strTmpFile, strSetupFile, True
    If Err.Number <> 0 Then '�ܾ�Ȩ���ȸ���
        gobjTrace.WriteInfo "SetupFile", "������װ�ļ�ʧ��", Err.Number & "-" & Err.Description
        If Not strSTFileName Like "ZL*" Then 'ϵͳ�ļ�
            strErrReturn = Err.Description
            Err.Clear '�������
            If blnForceCover Then  'ǿ�Ƹ���
                If gobjFSO.FileExists(strSetupFile & "_old") Then Kill (strSetupFile & "_old")
                Call Kill(strSetupFile & "_old")
                Name strSetupFile As strSetupFile & "_old"
                Call Kill(strSetupFile & "_old")
                '���¿����ļ�
                If Err.Number <> 0 Then Err.Clear
                If Not blnGoto Then
                    blnGoto = True
                    GoTo SartSetup
                End If
            End If
        Else
            '��������,�϶������ļ���ֻ���򱻶�ռ�򿪻���ִ��
            If Err.Number <> 70 And Err.Number <> 70 - 2146828288 Then
                If Not gblnHelperMain Then
                    sgResult = MsgBox("ע�⣺" & vbCrLf & _
                                "     �ļ���" & strSetupFile & "������������ ,ԭ�����£�" & vbCrLf & Err.Number & "-" & Err.Description & vbCrLf & _
                                "�����ԡ���ʾ�ֹ��Ѿ������ش�������ִ��������" & vbCrLf & _
                                "��ȡ������ʾȡ������������", vbQuestion + vbRetryCancel + vbDefaultButton1, "�Զ�����")
                Else
                    sgResult = vbRetry
                End If
                If sgResult = vbRetry Then
                    '����ִ��һ�ο���
                    If Not blnGoto Then
                        blnGoto = True
                        GoTo SartSetup
                    Else
                        blnReturn = False
                        strErrReturn = strSetupFile & "��װʧ�ܣ�" & Err.Number & "-" & Err.Description & ")"
                    End If
                Else
                    blnReturn = False
                    strErrReturn = strSetupFile & "��װʧ�ܣ�" & Err.Number & "-" & Err.Description & ")"
                End If
            Else
                Call zlGetFileProcess(strSetupFile, cllProcess)
                If strSTFileName Like "*.EXE" Then
                    If Not gblnHelperMain Then
                        sgResult = MsgBox("ע�⣺" & vbCrLf & _
                               "     �ļ���" & strSetupFile & "������ִ�У�����������" & vbCrLf & _
                               "����ֹ����ʾȡ��������������" & vbCrLf & _
                               "�����ԡ���ʾ��ֹ�����еĳ�������ִ��������" & vbCrLf & _
                               "�����ԡ���ʾ���β�����������", vbQuestion + vbAbortRetryIgnore, "�Զ�����")    'vbAbortRetryIgnore
                    Else
                        sgResult = vbRetry
                    End If
                ElseIf strSTFileName Like "*.OCX" Or strSTFileName Like "*.DLL" Then
                    strMsgBox = ""
                    For i = 1 To cllProcess.Count
                        If UCase(cllProcess(i)(1)) = "ZLHISCRUST.EXE" Then
                            Err.Clear
                            strErrReturn = strSetupFile & "��װʧ�ܣ�ZLHISCRUST.EXE����ռ�ã��Ѿ����ԣ�"
                            Exit Function
                        End If
                        If i > 2 Then
                            strMsgBox = strMsgBox & Space(5) & cllProcess(i)(0) & "��" & cllProcess(i)(1) & vbCrLf & Space(5) & "...."
                            Exit For
                        Else
                            strMsgBox = strMsgBox & Space(5) & cllProcess(i)(0) & "��" & cllProcess(i)(1) & vbCrLf
                        End If
                    Next
                    If Not gblnHelperMain Then
                        sgResult = MsgBox("ע�⣺" & vbCrLf & _
                                "     �ļ���" & strSetupFile & "���������³������ã��������� ��" & vbCrLf & _
                                strMsgBox & vbCrLf & _
                                "����ֹ����ʾȡ��������������" & vbCrLf & _
                                "�����ԡ���ʾ��ֹ�����еĳ�������ִ��������" & vbCrLf & _
                                "�����ԡ���ʾ���β�����������", vbQuestion + vbAbortRetryIgnore, "�Զ�����")    'vbAbortRetryIgnore
                    Else
                        sgResult = vbRetry
                    End If
                Else
                    If Not gblnHelperMain Then
                        sgResult = MsgBox("ע�⣺" & vbCrLf & _
                                "     �ļ���" & strSetupFile & "�����������ļ���վ�򿪣��������� ��" & vbCrLf & _
                                "����ֹ����ʾȡ��������������" & vbCrLf & _
                                "�����ԡ���ʾ�ֹ��Ѿ������վ���еĳ�������ִ��������" & vbCrLf & _
                                "�����ԡ���ʾ���β�����������", vbQuestion + vbAbortRetryIgnore, "�Զ�����")
                    Else
                        sgResult = vbRetry
                    End If
                End If
                If sgResult = vbAbort Then
                    blnReturn = False
                    If strErrReturn = "" Then strErrReturn = strSetupFile & "��װʧ�ܣ����������ļ�ռ��)"
                ElseIf sgResult = vbRetry Then
                    If strSTFileName Like "*.EXE" Or strSTFileName Like "*.OCX" Or strSTFileName Like "*.DLL" Then
                        For i = 1 To cllProcess.Count
                            Call TerminatePID(cllProcess(i)(0))
                        Next
                    End If
                    If Not blnGoto Then
                        blnGoto = True
                        GoTo SartSetup
                    End If
                ElseIf sgResult = vbIgnore Then
                    If strErrReturn = "" Then strErrReturn = strSetupFile & "��װʧ�ܣ����������ļ�ռ��,�Ѿ�����)"
                End If
            End If
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
    SetupOneFile = blnReturn
End Function

Private Function ExecBatFile(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
'���ܣ�ִ���������ļ�
    Dim strAutoRun      As String, strAutoRunBat As String
    Dim lngRet          As Long
    Dim lngTaskID       As Double, lngProcID As Long
    Dim objBat          As TextStream
    Dim strErrInfo      As String
    
    On Error GoTo ErrH
    gobjTrace.WriteSection "-", SL_LevelThree
    '��ǰ��������Ҫִ��������
    If gotCurType <> OT_PreUpgrade Then
        'ִ���������ļ�
        lblInfor.Caption = "���ڼ��������zlAutoRun.bat"
        strAutoRun = gstrSetupPath & "\zlAutoRun.ini"
        strAutoRunBat = gstrSetupPath & "\zlAutoRun.bat"
        '�޸�ģʽ���Զ�����������
        If gotCurType = OT_Repair And Not (gobjFSO.FileExists(strAutoRun) Or gobjFSO.FileExists(strAutoRunBat)) Then
            Set objBat = gobjFSO.CreateTextFile(strAutoRun, True)
            objBat.WriteLine gstrSetupPath & "\PUBLIC\zlMipClientShell.exe /regserver"
            objBat.WriteLine "IF ""%1""=="""" for %%c in (" & gstrSetupPath & "\PUBLIC\*.dll) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF ""%1""=="""" for %%c in (" & gstrSetupPath & "\PUBLIC\*.ocx) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF not ""%1""=="""" for %%c in (" & gstrSetupPath & "\PUBLIC\*.dll) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF not ""%1""=="""" for %%c in (" & gstrSetupPath & "\PUBLIC\*.ocx) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF ""%1""=="""" for %%c in (" & gstrSetupPath & "\apply\*.dll) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF ""%1""=="""" for %%c in (" & gstrSetupPath & "\apply\*.ocx) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF not ""%1""=="""" for %%c in (" & gstrSetupPath & "\apply\*.dll) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF not ""%1""=="""" for %%c in (" & gstrSetupPath & "\apply\*.ocx) do regsvr32.exe /s %%c"
            objBat.Close
            Set objBat = Nothing
        End If
        If gobjFSO.FileExists(strAutoRun) Or gobjFSO.FileExists(strAutoRunBat) Then
            On Error Resume Next
            If gobjFSO.FileExists(strAutoRun) Then
                If gobjFSO.FileExists(strAutoRunBat) Then Call gobjFSO.DeleteFile(strAutoRunBat, True)
                Name strAutoRun As gstrSetupPath & "\zlAutoRun.bat"
            End If
            Call Kill(strAutoRun)
            lngTaskID = Shell(gstrSetupPath & "\zlAutoRun.bat", vbHide)  'SW_SHOW
            If lngTaskID <> 0 Then
                lngProcID = OpenProcess(SYNCHRONIZE, False, lngTaskID)
                If lngProcID <> 0 Then
                    DoEvents
                    lngRet = WaitForSingleObject(lngProcID, INFINITE)
                    lngRet = CloseHandle(lngProcID)
                End If
                gobjTrace.WriteInfo "ExecBatFile", "�������ļ�ִ��", "�ɹ�"
            Else
                gobjTrace.WriteInfo "ExecBatFile", "�������ļ�ִ��", "ʧ��"
                Call RecordErrMsg(MT_ExeBat, "�������ļ�ִ��", "�������ļ�ִ��ʧ��")
                prgPross.Value = lngCurPro + lngCurIncPro
                Me.Refresh
                Exit Function
            End If
        End If
    End If
    prgPross.Value = lngCurPro + lngCurIncPro
    Me.Refresh
    ExecBatFile = True
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "SetupFiles", "�������ļ�ִ�з�����������", strErrInfo
    Call RecordErrMsg(MT_SetUp, "�������ļ�ִ�з�����������", strErrInfo)
    If Not gblnHelperMain Then MsgBox "�������ļ�ִ�з���������������ϵ����Ա����Ϣ��" & strErrInfo, vbInformation, App.Title
    Err.Clear
End Function

Private Function IsOldComponentExists(ByVal strName As String, Optional ByRef strExistsFile As String) As Boolean
'���ܣ��������嵥������ʱ���ж��Ƿ����ϵĲ�����û���ڲ����嵥��
'������strName=�����������ܲ�����׺
    Dim varItem         As Variant, strFileTmp              As String
    Dim strTmp As String
    
    On Error Resume Next
    If mcllOldComs Is Nothing Then
        Set mcllOldComs = New Collection
    End If
    strName = UCase(strName)
    strExistsFile = ""
    strExistsFile = mcllOldComs("K_" & strName)
    If Err.Number <> 0 Then
        Err.Clear
        If strName Like "*.DLL" Or strName Like "*.EXE" Then
            For Each varItem In gcllSetPath
                strFileTmp = varItem & "\" & strName
                If gobjFSO.FileExists(strFileTmp) Then
                    strExistsFile = strFileTmp
                    mcllOldComs.Add strFileTmp, "K_" & strName
                    mcllOldComs.Add strFileTmp, "K_" & Mid(strName, 1, Len(strName) - 4)
                    IsOldComponentExists = True
                    Exit For
                End If
            Next
        Else
            For Each varItem In gcllSetPath
                strFileTmp = varItem & "\" & strName & ".DLL"
                If gobjFSO.FileExists(strFileTmp) Then
                    strExistsFile = strFileTmp
                    mcllOldComs.Add strFileTmp, "K_" & strName & ".DLL"
                    mcllOldComs.Add strFileTmp, "K_" & strName
                    IsOldComponentExists = True
                    Exit For
                End If
                strFileTmp = varItem & "\" & strName & ".EXE"
                If gobjFSO.FileExists(strFileTmp) Then
                    strExistsFile = strFileTmp
                    mcllOldComs.Add strFileTmp, "K_" & strName & ".EXE"
                    mcllOldComs.Add strFileTmp, "K_" & strName
                    IsOldComponentExists = True
                    Exit For
                End If
            Next
        End If
    Else
        IsOldComponentExists = strExistsFile <> ""
    End If
End Function

Private Function ReportCheckInfo(ByRef blnOprateOK As Boolean) As Boolean
    Dim lngLoop     As Long
    
    gobjTrace.WriteSection "��װע���ļ�", SL_LevelThree
    On Error GoTo ErrH
    With grsFileUpgrade
        grsFileUpgrade.Filter = "(����<" & UC_Normal & " And ������Ϣ=NULL) Or (����=" & UC_Normal & " And �����ļ�·��<>NULL)"
        .Sort = "�ж�����,�Զ�ע��,��������"
        blnOprateOK = .RecordCount = 0
        For lngLoop = 1 To .RecordCount
            gobjTrace.WriteSection "-", SL_LevelThree
            If !���� > 0 Then
                gobjTrace.WriteInfo "SetupFiles", "��װע���ļ�", !�ļ���
            Else
                gobjTrace.WriteInfo "SetupFiles", "��������ļ�", !�ļ���
            End If
            If !���� > 0 Then
                Call ReportInfo("������" & !�ļ��� & "(" & !ʵ��·�� & ")��Ҫ����")
            ElseIf !���� = UC_NewDown Then
                Call ReportInfo("������" & !�ļ��� & "(" & !ʵ��·�� & ")ȱʧ����Ҫ����")
            End If
            
            If Not IsNull(!�����ļ�·��) Then
                Call ReportInfo("������" & !�ļ��� & "(" & !ʵ��·�� & ")���ڴ���·���ļ���" & Replace(!�����ļ�·��, "|", ","))
            End If
            
            If Not IsNull(!����ʵ��·��) Then
                Call ReportInfo("������" & !�ļ��� & "(" & !ʵ��·�� & ")�����¸��Ӱ�װ·����Ҫ���£�" & Replace(!����ʵ��·��, "|", ","))
            End If
            .MoveNext
        Next
    End With
    ReportCheckInfo = True
    Exit Function
ErrH:
    Call RecordErrMsg(MT_ChcekUpdate, "�ϴ���Ϣ�����ִ���", Err.Description)
End Function

