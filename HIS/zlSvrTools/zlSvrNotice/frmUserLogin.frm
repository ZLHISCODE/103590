VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ע��"
   ClientHeight    =   2595
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4470
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdSet 
      Caption         =   "���÷�����"
      Height          =   350
      Left            =   150
      TabIndex        =   10
      ToolTipText     =   "����Oracle�����ַ������ó���"
      Top             =   2115
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "��"
      Height          =   300
      Left            =   3720
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "ѡ����ڵķ������б�"
      Top             =   1455
      Width           =   300
   End
   Begin VB.TextBox txt���ݿ� 
      Height          =   300
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1455
      Width           =   1785
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1050
      Width           =   2115
   End
   Begin VB.TextBox txt�û� 
      Height          =   300
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   2
      Top             =   645
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3135
      TabIndex        =   9
      Top             =   2115
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1875
      TabIndex        =   8
      Top             =   2115
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -150
      TabIndex        =   11
      Top             =   1860
      Width           =   4965
   End
   Begin VB.Label lblNote 
      Caption         =   "    ֻ�о������ݿ�DBA��ɫ�����ϵͳ�������߲���ʹ�ñ����ߡ�"
      Height          =   375
      Left            =   990
      TabIndex        =   0
      Top             =   105
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1485
      TabIndex        =   3
      Top             =   1110
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1305
      TabIndex        =   1
      Top             =   705
      Width           =   540
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   1305
      TabIndex        =   5
      Top             =   1515
      Width           =   540
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   180
      Picture         =   "frmUserLogin.frx":1CFA
      Top             =   105
      Width           =   720
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intTimes As Integer
Dim strNote As String
Dim strUserName As String
Dim strServerName As String
Dim strPassword As String
Private mblnFirst As Boolean
Dim mcolServer As New Collection

Private Sub cmdOK_Click()
    intTimes = intTimes + 1
    
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(txt�û�.Text)
    strPassword = Trim(txt����.Text)
    strServerName = Trim(txt���ݿ�.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txt�û�)) = 0 Then
        strNote = "�������û�����"
        txt�û�.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txt�û�.SetFocus
            strNote = "�û�������"
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            txt����.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            txt���ݿ�.SetFocus
            GoTo InputError
        End If
    End If
    
    '�����ַ���
    Dim intPos As Integer
    intPos = InStr(1, strUserName, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUserName, "/", vbTextCompare)
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPassword, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    If Len(Trim(strPassword)) = 0 Then
        strNote = "δ�������룬����ע�ᡣ"
        txt����.SetFocus
        GoTo InputError
    End If
    
    strUserName = UCase(strUserName)
    If strUserName <> "SYSTEM" And strUserName <> "SYS" Then
        strPassword = TranPasswd(strPassword)
    End If
    
    If Not OraDataOpen(strServerName, strUserName, strPassword) Then
        txt����.Text = ""
        Exit Sub
    End If
    
    '�޸�ע���
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "MANAGER", strUserName
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", strServerName
    
    Unload Me
    Exit Sub

InputError:
    If intTimes > 3 Then
        MsgBox "��������ע��ʧ�ܣ�ϵͳ���Զ��˳���", vbExclamation, gstrSysName
        cmdCancel_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        Exit Sub
    End If

End Sub

Private Sub cmdCancel_Click()
    Set gcnOracle = Nothing
    Unload Me
End Sub


Private Sub CmdSelect_Click()
    Dim strServer As String
    Dim p As POINTAPI
    
    p.x = txt���ݿ�.Left / Screen.TwipsPerPixelX
    p.Y = (cmdSelect.Top + cmdSelect.Height) / Screen.TwipsPerPixelY
    ClientToScreen Me.hWnd, p
    
    strServer = frmServerSelect.GetServer(mcolServer, p.x * Screen.TwipsPerPixelX, p.Y * Screen.TwipsPerPixelY, txt���ݿ�.Text)
    If strServer <> "" Then
        txt���ݿ�.Text = strServer
        txt���ݿ�.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnFirst Then
        Dim LngStyle As Long
        LngStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
        LngStyle = LngStyle Or WinStyle
        Call SetWindowLong(hWnd, GWL_EXSTYLE, LngStyle)
        
        ShowWindow Me.hWnd, 0 '������
        ShowWindow Me.hWnd, 1 '����ʾ
    
        If Trim(txt�û�.Text) = "" Then
            cmdOK.Default = False
            txt�û�.SetFocus
        Else
            If txt����.Enabled Then
                txt����.SetFocus
            Else
                cmdOK.SetFocus
            End If
        End If
        
        mblnFirst = False
        
        If Trim(txt�û�.Text) <> "" And Trim(txt����.Text) <> "" Then Call cmdOK_Click
    
    End If
End Sub

Private Sub Form_Load()
    Dim ArrCommand
    
    txt�û�.Text = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "MANAGER", "")
    txt���ݿ�.Text = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    mblnFirst = True
    
    Call LoadServer
    
    Call ApplyOEM_Picture(Me, "Icon")
    
    '��������в��������û��������룬����䲢ִ��
    If Command() <> "" Then
        ArrCommand = Split(Command(), " ")
        If UBound(ArrCommand) >= 1 Then
            txt�û�.Text = ArrCommand(0)
            txt����.Text = ArrCommand(1)
        ElseIf UBound(ArrCommand) = 0 Then
            '�������/����ʾͬʱ�������û��������룬�������벻��Ҫ����ת��
            If InStr(1, ArrCommand(0), "/") <> 0 Then
                txt�û�.Text = Split(ArrCommand(0), "/")(0)
                txt����.Text = Split(ArrCommand(0), "/")(1)
            End If
        End If
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Set gcnOracle = Nothing
    End If
End Sub

Private Sub txt���ݿ�_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '�س������д���
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub txt�û�_GotFocus()
    SelAll txt�û�
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
End Sub

Private Sub txt���ݿ�_GotFocus()
    SelAll txt���ݿ�
End Sub

Private Sub cmdSet_Click()
    Dim strPath As String   'Oracle��װĿ¼
    Dim strCommond As String
    
    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
    If strPath = "" Then
        MsgBox "������Oracle�Ƿ�������װ�����顣", vbInformation
        Exit Sub
    End If
    
    'ִ��Oracle 8 ��Net Easy���õĳ���
    strCommond = strPath & "\BIN\N8SW.EXE"
    If ExecuteCommand(strCommond) = True Then
        '�Ѿ��ɹ�
        Exit Sub
    End If
    
    'ִ��Oracle 8i��Net Easy���õĳ���
    strCommond = strPath & "\BIN\launch.exe """ & strPath & "\network\tools"" " & strPath & "\network\tools\netca.cl"
    If ExecuteCommand(strCommond) = True Then
        '�Ѿ��ɹ�
        Exit Sub
    End If
    
End Sub

Private Function ExecuteCommand(ByVal strCommand As String) As Boolean
'���ܣ�ִ��ָ������
    Dim lngShell As Long, lngProcess As Long
    
    On Error Resume Next
    lngShell = Shell(strCommand, vbNormalFocus)
    
    If Err <> 0 Then
        Exit Function
    End If
    
    ExecuteCommand = True
End Function

Private Sub LoadServer()
'���ܣ��������صķ������б�
    Dim objSys As New Scripting.FileSystemObject
    Dim txtStream As Scripting.TextStream
    Dim strPath As String, strFile As String
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    
    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
    
    '��������Oracle 8i�������ļ��ڷ�
    strFile = strPath & "\network\ADMIN\tnsnames.ora"
    If objSys.FileExists(strFile) = False Then
        '������Oracle 8�������ļ��ڷ�
        strFile = strPath & "\NET80\ADMIN\tnsnames.ora"
        If objSys.FileExists(strFile) = False Then
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    
    Set mcolServer = Nothing
    Set txtStream = objSys.OpenTextFile(strFile)
    Do Until txtStream.AtEndOfStream
        strLine = Trim(txtStream.ReadLine)
        
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
                    End If
                End If
            End If
        End If
    Loop
End Sub

Private Sub AppendText(KeyAscii As Integer)
'���ܣ���TextBox�ؼ���Text׷�����ݣ������ݵ�ǰText��ֵ���б��м������õ�������Ŀ
'������KeyAscii    ��ǰ�İ���
    Dim strTemp As String
    Dim strInput As String
    Dim lngIndex As Long, lngStart As Long
    Dim varItem As Variant
    
    '���ȵ�ǰ�û�������ַ�
    If KeyAscii < 0 Or InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.", UCase(Chr(KeyAscii))) > 0 Then
        '�����ַ�ֻ�������֡�Ӣ�ĺͺ���
        strInput = Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    With txt���ݿ�
        '��¼�ϴεĲ����λ��
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)
        '���ŵõ��û�������ɺ��ı����г��ֵ�����
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)
    End With
    '���ݼ�������ݵõ����ܵ��б���
    strTemp = ""
    For Each varItem In mcolServer
        If UCase(varItem(0)) Like UCase(strInput & "*") Then
            strTemp = varItem(0)
        End If
    Next
    If strTemp <> "" Then
        txt���ݿ�.Text = strTemp
        txt���ݿ�.SelStart = Len(strInput)
        txt���ݿ�.SelLength = 100
    Else
        txt���ݿ�.Text = strInput
        txt���ݿ�.SelStart = lngStart
    End If

End Sub

