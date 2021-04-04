VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�û���¼"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4140
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "aqa"
      Top             =   600
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   -120
      TabIndex        =   7
      Top             =   1440
      Width           =   4335
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ ��(&C)"
         Height          =   350
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ ��(&O)"
         Height          =   350
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.ComboBox cboServerName 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1920
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Text            =   "ZLHIS"
      Top             =   120
      Width           =   1920
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   653
      Width           =   420
   End
   Begin VB.Label lblServerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   173
      Width           =   540
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserLogin.frx":058A
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFirst As Boolean  'ΪTrue��ʾ�Ѿ�������ʾ��
Private mintTimes As Integer  '��¼���Դ���
Private mblnת�� As Boolean     '��ʾ����������Ƿ�Ϊ���ݿ����룬�Ƿ���Ҫ��ת��
Private mcolServer As New Collection  '������������б�


Private Sub cboServerName_Change()
On Error GoTo errHandle

    Call ClearComponent
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cboServerName_Click()
On Error GoTo errHandle

    Call ClearComponent

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errHandle

 Set gcnOracle = Nothing
    Unload Me

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle

    Dim strNote As String
    Dim strUserName As String
    Dim strServerName As String
    Dim strPassword As String
    On Error GoTo InputError
    
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(txtUserName.Text)
    strPassword = Trim(txtPwd.Text)

    strServerName = Trim(cboServerName.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txtUserName)) = 0 Then
        strNote = "�������û���"
        txtUserName.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUserName.SetFocus
            strNote = "�û�������"
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtPwd.Enabled Then txtPwd.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            cboServerName.SetFocus
            GoTo InputError
        End If
    End If
    
    '�����ַ���
    Dim intPos As Integer
    intPos = InStr(strUserName, "@")
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(strUserName, "/")
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(strPassword, "@")
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    
    If Len(Trim(strPassword)) = 0 Then
        strNote = "����������"
        GoTo InputError
    End If
    
    If Not OraDataOpen(strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strPassword, IIf(mblnת��, TranPasswd(strPassword), strPassword))) Then
        txtPwd.Text = ""
        If txtPwd.Enabled Then txtPwd.SetFocus

        Exit Sub
    End If
      
    
    '�޸�ע���
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", strServerName
    
    '������ݷ�ʽ��
    SaveSetting "ZLSOFT", "����ȫ��", "����·��", App.Path & "\" & App.EXEName & ".exe"
    
    Unload Me
    
    Call frmMain.Show

InputError:
    If mintTimes > 3 Then
        MsgBox "�������ε�¼ʧ�ܣ�ϵͳ���Զ��˳�", vbExclamation, gstrSysName
        cmdCancel_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        Exit Sub
    End If

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim ArrCommand
    Dim i As Integer
    Call LoadServer
    mblnת�� = True
    On Error GoTo errH
    mblnFirst = False
    mintTimes = 1
    cboServerName.Text = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="SERVER", Default:="")
    
    If InStr(Command(), "=") > 0 Then Me.Hide
    '��������в��������û��������룬����䲢ִ��
    If Command() <> "" Then
        
        ArrCommand = Split(Command(), " ")
        
        If UBound(ArrCommand) >= 1 Then
            If InStr(ArrCommand(0), "=") <= 0 Then
                Me.txtUserName.Text = ArrCommand(0)
                Me.txtPwd.Text = ArrCommand(1)
            End If
        ElseIf UBound(ArrCommand) = 0 Then
            '�������/����ʾͬʱ�������û��������룬�������벻��Ҫ����ת��
            If InStr(1, ArrCommand(0), "/") <> 0 Then
                Me.txtUserName.Text = Split(ArrCommand(0), "/")(0)
                Me.txtPwd.Text = Split(ArrCommand(0), "/")(1)
                mblnת�� = False
            End If
        End If
    End If
    Exit Sub
errH:
    If CStr(Command()) <> "" Then MsgBox CStr(Erl()) & "�г��ִ������ֶ���¼��" & vbNewLine & Err.Description, vbQuestion
End Sub


Private Sub LoadServer()
'���ܣ��������صķ������б�
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    
    cboServerName.Clear
    
    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
    If Not gobjFile.FolderExists(strPath) Then '10G
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORA_CRS_HOME")
    End If
    If Not gobjFile.FolderExists(strPath) Then '10Gr2
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home1", "ORACLE_HOME")
    End If
    If Not gobjFile.FolderExists(strPath) Then '10Gr2
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home2", "ORACLE_HOME")
    End If
    If Not gobjFile.FolderExists(strPath) Then    '10G ��ҵ��
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraClient10g_home1", "ORACLE_HOME")
    End If
    
    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i����
    If Not gobjFile.FileExists(strFile) Then
        strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
        If Not gobjFile.FileExists(strFile) Then Exit Sub
    End If
    
    lngFile = FreeFile()
    Open strFile For Input Access Read As lngFile
    
    Set mcolServer = Nothing
    Do Until EOF(lngFile)
        Input #lngFile, strLine
        
        strLine = Trim(strLine)
        If strLine <> "" And Left(strLine, 1) <> "#" Then
            '��ע���л����
            If InStr(strLine, "(") = 0 And InStr(strLine, ")") = 0 Then
                '���е����ݾ��Ƿ��������ˣ����������ݶ���ʼ��
                strServer = Trim(Mid(strLine, 1, InStr(strLine, "=") - 1))
                strComputer = ""
                strSID = ""
            ElseIf InStr(strLine, "(ADDRESS") > 0 Then
                '���е�������������
                If InStr(strLine, "PROTOCOL = TCP") > 0 And InStr(strLine, "PORT = 1521") > 0 Then
                    '�������ǵĳ���Ҫ��
                    strComputer = Mid(strLine, InStr(strLine, "HOST =") + Len("HOST ="))
                    strComputer = Trim(Mid(strComputer, 1, InStr(strComputer, ")") - 1))
                End If
            Else
                lngPos = InStr(strLine, "(SID")
                If lngPos = 0 Then
                    lngPos = InStr(strLine, "(SERVICE_NAME")
                End If
                
                If lngPos > 0 Then
                    '���е�������ʵ����
                    strSID = Mid(strLine, InStr(lngPos, strLine, "=") + 1)
                    strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
                    
                    If strServer <> "" And strComputer <> "" And strSID <> "" Then
                        '�Ѿ��õ�������Ҫ������
                        mcolServer.Add Array(strServer, strComputer, strSID)
                        cboServerName.AddItem strServer
                    End If
                End If
            End If
        End If
    Loop
End Sub

Private Sub ClearComponent()
'���ܣ�--���ע���[��������]--��Ϊ��ͬ�����ݿ����ʹ�õ�ϵͳ�Ͱ汾��ͬ
    If mblnFirst = True Then '����ʱ�Կؼ��ĸ�ֵ����������
        SaveSetting "ZLSOFT", "ע����Ϣ", "��������", ""
    End If
End Sub
