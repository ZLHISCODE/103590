VERSION 5.00
Begin VB.Form frmUserLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Ա��¼"
   ClientHeight    =   2250
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4065
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdSelect 
      Caption         =   "��"
      Height          =   300
      Left            =   3360
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "ѡ����ڵķ������б�"
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txt���ݿ� 
      Height          =   300
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2055
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   1515
   End
   Begin VB.TextBox txt�û� 
      Height          =   300
      Left            =   2055
      MaxLength       =   30
      TabIndex        =   1
      Top             =   195
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   50
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Width           =   4725
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   240
      Picture         =   "frmUserLogin.frx":1CFA
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "��  ��"
      Height          =   180
      Left            =   1440
      TabIndex        =   2
      Top             =   660
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1440
      TabIndex        =   0
      Top             =   255
      Width           =   540
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   1440
      TabIndex        =   4
      Top             =   1065
      Width           =   540
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
Dim mcolServer As New Collection


Private Sub cmdOK_Click()

    intTimes = intTimes + 1
    
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    gstrUserName = Trim(txt�û�.Text)
    gstrPassword = Trim(txt����.Text)
    gstrServer = Trim(txt���ݿ�.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txt�û�)) = 0 Then
        strNote = "�������û�����"
        txt�û�.SetFocus
        GoTo InputError
    End If
    
    If Len(gstrUserName) <> 1 Then
        If Mid(gstrUserName, 1, 1) = "/" Or Mid(gstrUserName, 1, 1) = "@" Or Mid(gstrUserName, Len(gstrUserName) - 1, 1) = "/" Or Mid(gstrUserName, Len(gstrUserName) - 1, 1) = "@" Then
            txt�û�.SetFocus
            strNote = "�û�������"
            Exit Sub
        End If
    End If
    If Trim(gstrPassword) <> "" And Len(gstrPassword) <> 1 Then
        If Mid(gstrPassword, Len(gstrPassword) - 1, 1) = "/" Or Mid(gstrPassword, Len(gstrPassword) - 1, 1) = "@" Or Mid(gstrPassword, 1, 1) = "/" Or Mid(gstrPassword, 1, 1) = "@" Then
            txt����.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(gstrServer) <> "" Then
        If Mid(gstrServer, Len(gstrServer) - 1, 1) = "/" Or Mid(gstrServer, Len(gstrServer) - 1, 1) = "@" Or Mid(gstrServer, 1, 1) = "/" Or Mid(gstrServer, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            txt���ݿ�.SetFocus
            GoTo InputError
        End If
    End If
    
    '�����ַ���
    Dim intPos As Integer
    intPos = InStr(1, gstrUserName, "@", vbTextCompare)
    If intPos > 0 Then
        gstrServer = Mid(gstrUserName, intPos + 1)
        gstrUserName = Mid(gstrUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, gstrUserName, "/", vbTextCompare)
    If intPos > 0 Then
        gstrPassword = Mid(gstrUserName, intPos + 1)
        gstrUserName = Mid(gstrUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, gstrPassword, "@", vbTextCompare)
    If intPos > 0 Then
        gstrServer = Mid(gstrPassword, intPos + 1)
        gstrPassword = Mid(gstrPassword, 1, intPos - 1)
    End If
    
    If Len(Trim(gstrPassword)) = 0 Then
        strNote = "δ�������룬����ע�ᡣ"
        txt����.SetFocus
        GoTo InputError
    End If
    
    gstrUserName = UCase(gstrUserName)
    If gstrUserName <> "SYSTEM" And gstrUserName <> "SYS" Then
        gstrPassword = TranPasswd(gstrPassword)
  
    End If
    
    If Not OraDataOpen(gstrServer, gstrUserName, gstrPassword, gcnOracle) Then
        txt����.Text = ""
        Exit Sub
    End If
    
    '��ʼ����������
''    Call InitCommon(gcnOracle)
    
    Set gmobjCommon = CreateObject("ZL9ComLib.clsComLib")
'    Set gmobjCommon = New zl9ComLib.clsComLib
    Call gmobjCommon.InitCommon(gcnOracle)
  
'    If Not gmobjCommon.RegCheck Then
'        MsgBox "ZL9ComLibע��ʧ��!", vbExclamation, "��ʾ!"
'        Exit Sub
'    End If
    
    Call SaveSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "����������", txt���ݿ�.Text)
    
    Unload Me
    
    '��ʾ������
'    Me.Hide
    frmMain.Show
    
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
    On Error Resume Next
    Set gcnOracle = Nothing
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    Dim strServer As String
    Dim p As POINTAPI
    
    p.X = txt���ݿ�.Left / Screen.TwipsPerPixelX
    p.Y = (cmdSelect.Top + cmdSelect.Height) / Screen.TwipsPerPixelY
    ClientToScreen Me.Hwnd, p
    
    strServer = frmServerSelect.GetServer(mcolServer, p.X * Screen.TwipsPerPixelX, p.Y * Screen.TwipsPerPixelY, txt���ݿ�.Text)
    If strServer <> "" Then
        txt���ݿ�.Text = strServer
        txt���ݿ�.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Dim lngStyle As Long
    lngStyle = GetWindowLong(Hwnd, GWL_EXSTYLE)
    lngStyle = lngStyle Or WinStyle
    Call SetWindowLong(Hwnd, GWL_EXSTYLE, lngStyle)
    
    ShowWindow Me.Hwnd, 0 '������
    ShowWindow Me.Hwnd, 1 '����ʾ
    
    If Len(txt�û�) <> 0 Then
        txt����.SetFocus
    End If
End Sub

Private Sub Form_Load()
    txt�û�.Text = "zlhis" ''GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "MANAGER", "")
    txt���ݿ�.Text = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "����������", "")
    intTimes = 0
    
    Call LoadServer
   ' Call ApplyOEM_Picture(Me, "Icon")
    
 
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
    Me.txt�û�.SelStart = 0: Me.txt�û�.SelLength = 100
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt���ݿ�_GotFocus()
    Me.txt���ݿ�.SelStart = 0: Me.txt���ݿ�.SelLength = 100
End Sub

Public Sub LoadServer()
'���ܣ��������صķ������б�
    Dim objSys As New Scripting.FileSystemObject
    Dim txtStream As Scripting.TextStream
    Dim strPath As String, strFile As String
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    
    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
    
    '��ȡ10gOracleHomeĿ¼
    If strPath = "" Then
       strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home1", "ORACLE_HOME")
    End If
    
    '��������Oracle 8i�������ļ��ڷ�
    strFile = strPath & "\network\ADMIN\tnsnames.ora"
    If objFso.FileExists(strFile) = False Then
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

