VERSION 5.00
Begin VB.Form frmUserLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��¼"
   ClientHeight    =   2205
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmUserLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4170
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cmb���ݿ� 
      Height          =   300
      Left            =   1950
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1050
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   8
      Top             =   1455
      Width           =   5025
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2865
      TabIndex        =   7
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton CMDȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1755
      TabIndex        =   6
      Top             =   1710
      Width           =   1100
   End
   Begin VB.TextBox TXT���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   630
      Width           =   1920
   End
   Begin VB.TextBox txt�û� 
      Height          =   300
      Left            =   1950
      TabIndex        =   1
      Top             =   195
      Width           =   1920
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserLogin.frx":6852
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Lbl������ 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   1320
      TabIndex        =   4
      Top             =   1110
      Width           =   540
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1500
      TabIndex        =   2
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Lbl�û��� 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1320
      TabIndex        =   0
      Top             =   255
      Width           =   540
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�����и�ʽ��
'zlhis90.exe �˵�
'zlhis90.exe �û���/����        �����������Ҫ��������ת��
'zlhis90.exe �û��� ����
'zlhis90.exe �û��� ���� �˵�
Private mblnFirst As Boolean  'ΪTrue��ʾ�Ѿ�������ʾ��
Private mintTimes As Integer  '��¼���Դ���
Private mblnת�� As Boolean     '��ʾ����������Ƿ�Ϊ���ݿ����룬�Ƿ���Ҫ��ת��
Private mcolServer As New Collection  '������������б�

Private Sub CMDȷ��_Click()
    Dim strNote As String
    Dim strUserName As String
    Dim strServerName As String
    Dim strPassword As String
    
    SetConState False
    mintTimes = mintTimes + 1
    
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(txt�û�.Text)
    strPassword = Trim(TXT����.Text)
    strServerName = Trim(cmb���ݿ�.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txt�û�)) = 0 Then
        strNote = "�������û���"
        txt�û�.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txt�û�.SetFocus
            strNote = "�û�������"
            Call SetConState: Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If TXT����.Enabled Then TXT����.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            cmb���ݿ�.SetFocus
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
        strNote = "����������"
        GoTo InputError
    End If
    
    If Not OraDataOpen(strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strPassword, IIf(mblnת��, TranPasswd(strPassword), strPassword))) Then
        TXT����.Text = ""
        If TXT����.Enabled Then TXT����.SetFocus
        Call SetConState: Exit Sub
    End If
    
    '�޸�ע���
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "MANAGER", strUserName
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", strServerName
    gstrDBUser = UCase(strUserName)
    
    Unload Me
    Exit Sub
InputError:
    If mintTimes > 3 Then
        CMD����_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, App.Title
        End If
        Call SetConState: Exit Sub
    End If
End Sub

Private Sub CMD����_Click()
    gstrDBUser = ""
    Set gcnOracle = Nothing
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim LngStyle As Long
    If mblnFirst = False Then
        LngStyle = GetWindowLong(Me.hwnd, (-20))
        LngStyle = LngStyle Or &H40000
        Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
        
        ShowWindow Me.hwnd, 0 '������
        ShowWindow Me.hwnd, 1 '����ʾ

        If Trim(txt�û�.Text) = "" Then
            CMDȷ��.Default = False
            txt�û�.SetFocus
        Else
            If TXT����.Enabled Then
                TXT����.SetFocus
            Else
                CMDȷ��.SetFocus
            End If
        End If
        mblnFirst = True
    
        If Trim(txt�û�.Text) <> "" And Trim(TXT����.Text) <> "" Then Call CMDȷ��_Click
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "TXT����" Then
            Call CMDȷ��_Click
        Else
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim ArrCommand
    Call LoadServer
    
    mblnת�� = True
    mblnFirst = False
    mintTimes = 1
    txt�û�.Text = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "MANAGER", "")
    cmb���ݿ�.Text = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    
    '��������в��������û��������룬����䲢ִ��
    If Command() <> "" Then
        ArrCommand = Split(Command(), " ")
        If UBound(ArrCommand) >= 1 Then
            Me.txt�û�.Text = ArrCommand(0)
            Me.TXT����.Text = ArrCommand(1)
        ElseIf UBound(ArrCommand) = 0 Then
            '�������/����ʾͬʱ�������û��������룬�������벻��Ҫ����ת��
            If InStr(1, ArrCommand(0), "/") <> 0 Then
                Me.txt�û�.Text = Split(ArrCommand(0), "/")(0)
                Me.TXT����.Text = Split(ArrCommand(0), "/")(1)
                mblnת�� = False
            End If
        End If
    End If
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub cmb���ݿ�_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '�س������д���
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub txt�û�_Change()
    If Not mblnFirst Then Exit Sub
    CMDȷ��.Default = False
End Sub

Private Sub txt�û�_GotFocus()
    GetFocus txt�û�
End Sub

Private Sub TXT����_GotFocus()
    GetFocus TXT����
End Sub

Private Sub cmb���ݿ�_GotFocus()
    With cmb���ݿ�
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    CMD����.Enabled = BlnState
    CMDȷ��.Enabled = BlnState
End Sub

Private Sub LoadServer()
'���ܣ��������صķ������б�
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    
    cmb���ݿ�.Clear
    
    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
    If strPath = "" Then '10G
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORA_CRS_HOME")
    End If
    If strPath = "" Then '10Gr2
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home1", "ORACLE_HOME")
    End If
    If strPath = "" Then '10Gr2
        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home2", "ORACLE_HOME")
    End If
    
    lngFile = FreeFile()
    On Error Resume Next
    '��������Oracle 8i�������ļ��ڷ�
    strFile = strPath & "\network\ADMIN\tnsnames.ora"
    Open strFile For Input Access Read As lngFile
    If Err <> 0 Then
        '�ļ�������
        Err.Clear
        '������Oracle 8�������ļ��ڷ�
        strFile = strPath & "\NET80\ADMIN\tnsnames.ora"
        Open strFile For Input Access Read As lngFile
        
        If Err <> 0 Then
            Err.Clear
            Exit Sub
        End If
    End If
    
    
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
                        cmb���ݿ�.AddItem strServer
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
    
    With cmb���ݿ�
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
        cmb���ݿ�.Text = strTemp
        cmb���ݿ�.SelStart = Len(strInput)
        cmb���ݿ�.SelLength = 100
    Else
        cmb���ݿ�.Text = strInput
        cmb���ݿ�.SelStart = lngStart
    End If
End Sub
