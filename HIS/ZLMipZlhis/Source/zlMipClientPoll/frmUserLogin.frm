VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Ա��¼"
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
   Begin VB.ComboBox cboServer 
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
      TabIndex        =   9
      Top             =   1455
      Width           =   5025
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸�����(&M)"
      Height          =   350
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "�����˴��޸�����"
      Top             =   1710
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2865
      TabIndex        =   7
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1755
      TabIndex        =   6
      Top             =   1710
      Width           =   1100
   End
   Begin VB.TextBox txtPassWord 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   630
      Width           =   1920
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1950
      TabIndex        =   1
      Top             =   195
      Width           =   1920
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserLogin.frx":1CFA
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

Private Sub cmdOK_Click()
    Dim strNote As String
    Dim strUserName As String
    Dim strServerName As String
    Dim strPassword As String
    Dim strSQL As String, rsData As New ADODB.Recordset
    Dim arrTmp As Variant, lngLen As Long, i As Long, intChr As Integer
    Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
    Dim blnPwdLen As Boolean, intPwdMin As Integer, intPwdMax As Integer
    Dim blnComplex As Boolean, strOterChrs As String
    Dim blnTransPassword As Boolean

    On Error GoTo InputError
    SetConState False
    
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(txtUser.Text)
    strPassword = Trim(txtPassWord.Text)
    strServerName = Trim(cboServer.Text)
    '��Ч�ַ���Ч��
    If Len(Trim(txtUser)) = 0 Then
        strNote = "�������û���"
        txtUser.SetFocus
        GoTo InputError
    End If
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "�û�������"
            SetConState
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            txtPassWord.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            cboServer.SetFocus
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
    
    mintTimes = mintTimes + 1
    
    If Not gclsMsgOracle.OraDataOpen(strServerName, strUserName, strPassword, True) Then
        txtPassWord.Text = ""
        txtPassWord.SetFocus
        SetConState
        Exit Sub
    Else
        gstrDbUser = UCase(strUserName)
    End If
    
    Call gclsBusiness.InitBusiness(gclsMsgOracle, "", gstrDbUser)
    '------------------------------------------------------------------------------------------------------------------
    '����Ƿ�Ϊ��Ʒ�����ߵ�¼��������ǣ����ֹ���˳�
    If gclsBusiness.IsOwner = False Then
        MsgBox "��¼��ֻ��Ϊ��Ʒ�����ߣ�", vbInformation
        Exit Sub
    End If
    
    If strUserName = strPassword Then
        MsgBox "��¼�û�����������ͬ��������ϵͳ��ȫҪ�����������޸����롣", vbInformation
        cmdModify_Click
        SetConState
        Exit Sub
    End If

    strSQL = "Select ������,Nvl(����ֵ,ȱʡֵ) ����ֵ From zlOptions Where ������ in (20,21,22,23)"
    Set rsData = gclsMsgOracle.OpenSQLRecord(strSQL, Me.Caption)
    blnPwdLen = False: intPwdMin = 0: intPwdMax = 0
    blnComplex = False: strOterChrs = ""
    Do While Not rsData.EOF
        Select Case rsData!������
            Case 20 '�Ƿ�������볤��
                blnPwdLen = Val(rsData!����ֵ & "") = 1
            Case 21 '���볤������
                intPwdMin = Val(rsData!����ֵ & "")
            Case 22 '���볤������
                intPwdMax = Val(rsData!����ֵ & "")
            Case 23 '�Ƿ�������븴�Ӷ�
                blnComplex = Val(rsData!����ֵ & "") = 1
        End Select
        rsData.MoveNext
    Loop
    If blnPwdLen Then
        lngLen = Len(strPassword)
        If Not (lngLen >= intPwdMin And lngLen <= intPwdMax) Then
           If intPwdMin = intPwdMax Then
               MsgBox "�������Ϊ" & intPwdMax & " λ�ַ���������ϵͳ��ȫҪ�����������޸����룡", vbInformation, gstrSysName
              cmdModify_Click
              SetConState
              Exit Sub
           Else
               MsgBox "�������Ϊ" & intPwdMin & "��" & intPwdMax & " λ�ַ���������ϵͳ��ȫҪ�����������޸����룡", vbInformation, gstrSysName
              cmdModify_Click
              SetConState
              Exit Sub
           End If
       End If
    End If
    For i = 1 To Len(strPassword)
        intChr = Asc(UCase(Mid(strPassword, i, 1)))
        If intChr >= 32 And intChr < 127 Then
            
            Select Case intChr
                Case 48 To 57 '����
                    blnHaveNum = True
                Case 65 To 90 '��ĸ
                    blnAlpha = True
                Case 32, 34, 47, 64  '�ո�,˫����,/,@
                    strOterChrs = strOterChrs & Chr(intChr)
                Case Is < 48, 58 To 64, 91 To 96, Is > 122
                    blnChar = True
            End Select
        Else
            strOterChrs = strOterChrs & Chr(intChr)
        End If
    Next
    If strOterChrs <> "" Then
        MsgBox "���벻�����������ַ���" & strOterChrs & "��������ϵͳ��ȫҪ�����������޸����룡", vbInformation, gstrSysName
       cmdModify_Click
       SetConState
       Exit Sub
    ElseIf Not (blnHaveNum And blnAlpha And blnChar) And blnComplex Then
        MsgBox "����������һ�����֡�һ����ĸ��һ�������ַ���ɣ���ǰ���벻����ϵͳ��ȫҪ�����������޸����룡��", vbInformation, gstrSysName
       cmdModify_Click
       SetConState
       Exit Sub
    End If
    '�Ƿ������������
    If CheckPwdExpiry = True Then
        cmdModify_Click
        SetConState
        Exit Sub
    End If
    
    
    '����SQL Trace
    '-----------------------------------------------
    strNote = gclsMsgOracle.SetSQLTrace(gstrUserName, strServerName)
    If strNote <> "" Then
        MsgBox "������SQL Trace����!" & vbCrLf & "���ٽ���ļ�:" & strNote & vbCrLf & _
                "�����Oracle������udumpĿ¼��,����10M��ֹͣд��.", vbInformation, "��ʾ"
    End If
    
    If UCase(strServerName) = "RBO" Then
        gclsMsgOracle.SetRunWithRBO
    End If
    
    '�޸�ע���
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "USER", strUserName
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", strServerName
    '������ݷ�ʽ��
    SaveSetting "ZLSOFT", "����ȫ��", "����·��", App.Path & "\" & App.EXEName & ".exe"
    
    Unload Me
    Exit Sub
InputError:
    If mintTimes > 3 Then
        MsgBox "�������ε�¼ʧ�ܣ�ϵͳ���Զ��˳�", vbExclamation, gstrSysName
        cmdCancel_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        SetConState
        Exit Sub
    End If
End Sub

Private Function CheckPwdExpiry() As Boolean
    Dim strSQL As String
    Dim rsData As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim dtExpiryDate As Date
    Dim dtNow As Date
    Dim intDiff As Integer
    
    strSQL = "Select EXPIRY_DATE From User_Users Where UserName=User"
    Set rsData = gclsMsgOracle.OpenSQLRecord(strSQL, Me.Caption)
    If rsData.BOF = False Then
        If IsNull(rsData("EXPIRY_DATE").Value) = True Then
            CheckPwdExpiry = False
            Exit Function
        End If
        dtExpiryDate = Format(rsData("EXPIRY_DATE").Value, "YYYY-MM-DD HH:MM:SS")
        '�жϹ��������뵱ǰ�����������
        strSQL = "SELECT SYSDATE FROM DUAL"
        Set rsTemp = gclsMsgOracle.OpenSQLRecord(strSQL, Me.Caption)
        If rsTemp.BOF = False Then
            dtNow = Format(rsTemp("SYSDATE").Value, "YYYY-MM-DD HH:MM:SS")
        End If
        intDiff = DateDiff("d", dtNow, dtExpiryDate)
        
        If intDiff > 7 Then
            CheckPwdExpiry = False
            Exit Function
        End If
        
        If intDiff > 3 And intDiff <= 7 Then
            '��ʾ�޸�����
            If MsgBox("������Ч�ڻ���" & intDiff & "��,�Ƿ������޸�����?", vbQuestion + vbYesNo, "�����������") = vbYes Then
                CheckPwdExpiry = True
            Else
                CheckPwdExpiry = False
                Exit Function
            End If
        ElseIf intDiff <= 3 Then
            CheckPwdExpiry = True
            MsgBox "������Ч�ڻ���" & intDiff & "�죬���������޸����롣", vbInformation
        Else
            CheckPwdExpiry = False
            Exit Function
        End If
    End If
End Function

Private Sub cboServer_Change()
    Call ClearComponent
End Sub

Private Sub cboServer_Click()
    Call ClearComponent
End Sub

Private Sub cmdCancel_Click()
    Set gobjRegister = Nothing
    Unload Me
End Sub

Private Sub cmdModify_Click()
    Dim strUserName As String
    Dim strPassword As String
    Dim strServerName As String
    Dim strNote As String
    
    On Error GoTo InputError
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(txtUser.Text)
    strPassword = Trim(txtPassWord.Text)
    strServerName = Trim(cboServer.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txtUser.Text)) = 0 Then
        strNote = "�������û���"
        txtUser.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "�û�������"
            SetConState
            Exit Sub
        End If
    End If
    
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtPassWord.Enabled Then txtPassWord.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            cboServer.SetFocus
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
    
    If FrmChangePass.ShowMe(Me, strUserName, strPassword, strServerName, mblnת��) Then
        txtPassWord.Text = strPassword
        cboServer.Text = strServerName
        If cmdOK.Enabled Then cmdOK.SetFocus
    Else
        txtPassWord.SetFocus
    End If
    Exit Sub
InputError:
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        Exit Sub
End Sub

Private Sub Form_Activate()
    Dim LngStyle As Long
    If mblnFirst = False Then
        LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
        LngStyle = LngStyle Or WinStyle
        Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
        If InStr(Command(), "=") <= 0 And InStr(Command(), "&") <= 0 Then
            ShowWindow Me.hwnd, 0 '������
            ShowWindow Me.hwnd, 1 '����ʾ
        End If
        
        If Trim(txtUser.Text) = "" Then
            cmdOK.Default = False
            txtUser.SetFocus
        Else
            txtPassWord.SetFocus
        End If
        mblnFirst = True
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then Call cmdOK_Click
    End If
    If InStr(Command(), "=") > 0 And InStr(Command(), "&") = 0 Then Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "txtPassWord" Then
            Call cmdOK_Click
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
    Dim i As Integer
    
    Call LoadServer
    On Error GoTo errH
    mblnת�� = True
    mblnFirst = False
    mintTimes = 1
    txtUser.Text = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="USER", Default:="")
    cboServer.Text = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="SERVER", Default:="")
    Call zlComLib.ApplyOEM_Picture(Me, "Icon")

    If InStr(Command(), "=") > 0 And InStr(Command(), "&") = 0 Then Me.Hide
    '��������в��������û��������룬����䲢ִ��
    If Command() <> "" And InStr(Command(), "&") = 0 Then
        
        ArrCommand = Split(Command(), " ")
        If UBound(ArrCommand) >= 1 Then
            If InStr(ArrCommand(0), "=") <= 0 Then
                Me.txtUser.Text = ArrCommand(0)
                Me.txtPassWord.Text = ArrCommand(1)
            End If
        ElseIf UBound(ArrCommand) = 0 Then
            '�������/����ʾͬʱ�������û��������룬�������벻��Ҫ����ת��
            If InStr(1, ArrCommand(0), "/") <> 0 And InStr(1, ArrCommand(0), ",") = 0 Then
                Me.txtUser.Text = Split(ArrCommand(0), "/")(0)
                Me.txtPassWord.Text = Split(ArrCommand(0), "/")(1)
                mblnת�� = False
            End If
        End If
    End If
    Exit Sub
errH:
    If CStr(Command()) <> "" Then MsgBox CStr(Erl()) & "�г��ִ������ֶ���¼��" & vbNewLine & Err.Description, vbQuestion
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub cboServer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '�س������д���
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub txtUser_Change()
    If Not mblnFirst Then Exit Sub
    cmdOK.Default = False
End Sub

Private Sub txtUser_GotFocus()
    GetFocus txtUser
    OpenIme False
End Sub

Private Sub txtPassWord_GotFocus()
    GetFocus txtPassWord
End Sub

Private Sub cboServer_GotFocus()
    With cboServer
        .SelStart = 0
        .SelLength = Len(.Text)
        OpenIme False
    End With
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    cmdCancel.Enabled = BlnState
    cmdModify.Enabled = BlnState
    cmdOK.Enabled = BlnState
End Sub

Private Sub LoadServer()
'���ܣ��������صķ������б�
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    Dim arrTmp As Variant
    Dim rsOraHome As ADODB.Recordset
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer
    Dim i As Long, blnRead As Boolean

    cboServer.Clear
    Set rsOraHome = New ADODB.Recordset
    With rsOraHome
        .Fields.Append "Name", adVarChar, 256 'Name
        .Fields.Append "VerSion", adInteger  '�汾
        .Fields.Append "Times", adInteger '�ڼ��ΰ�װ
        .Fields.Append "Server", adInteger '1-������,2-�ͻ���
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '1:��ȡ64λ��32Ŀ¼���Զ���λ��SOFTWARE\Wow6432Node\Oracle 2����ȡ32λ��32λĿ¼
        arrTmp = GetAllSubKey("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle")
        If TypeName(arrTmp) = "Empty" Then
            If Is64bit Then
                cboServer.ToolTipText = "û���ҵ�ע�����HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Oracle��"
            Else
                cboServer.ToolTipText = "û���ҵ�ע�����HKEY_LOCAL_MACHINE\SOFTWARE\Oracle��"
            End If
        Else
            For i = LBound(arrTmp) To UBound(arrTmp)
                If UCase(arrTmp(i)) Like "KEY_ORA*HOME*" Then
                    intVersion = 0: intTimes = 0:  intServer = 1
                    If GetOraInfoByRegKey(arrTmp(i), intVersion, intTimes, intServer) Then
                        .AddNew Array("Name", "VerSion", "Times", "Server"), Array("\" & arrTmp(i), intVersion, intTimes, intServer)
                        .Update
                    End If
                End If
            Next
            If UBound(arrTmp) <> -1 Then ''����Ŀ¼������Oracle_Home��Ϣ��Ĭ�϶�ȡ���
                .AddNew Array("Name", "VerSion", "Times", "Server"), Array("", 0, 0, 1): .Update
            End If
            .Sort = "VerSion Desc,Times Desc,Server"
            Do While Not .EOF
                strPath = ""
                blnRead = Not GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle" & !Name, "ORACLE_HOME", strPath)
                blnRead = blnRead Or strPath = "" And !Name & "" = ""
                If blnRead Then
                    Call GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle", "ORA_CRS_HOME", strPath)
                End If
                If strPath <> "" Then
                    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i����
                    If Dir(strFile) <> "" Then Exit Do
                    strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
                    If Dir(strFile) <> "" Then Exit Do
                End If
                strFile = ""
                .MoveNext
            Loop
        End If
    End With
    If strFile = "" Then Exit Sub
    cboServer.ToolTipText = "�������б���Դ:" & strFile
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
                If InStr(strLine, "PROTOCOL = TCP") > 0 And InStr(strLine, "PORT = ") > 0 Then
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
                        cboServer.AddItem strServer
                    End If
                End If
            End If
        End If
    Loop
    Close #lngFile
End Sub

Private Function GetOraInfoByRegKey(ByVal strOraHome As String, ByRef intVer As Integer, ByRef intTimes As Integer, ByRef intServer As Integer) As Boolean
'����:ͨ��OracleHome����ȡOracle��Ϣ
    Dim arrTmp As Variant
    Dim i As Long, blnRetrun As Boolean
    'KEY_OraDb11g_home1_32bit
    'Key_Ora*�汾Home_32Bit
    'Key_Ora*�汾_Home*
    arrTmp = Split(UCase(strOraHome), "_")
    For i = 1 To UBound(arrTmp)
        If arrTmp(i) Like "HOME*" Then
            intTimes = ValEx(arrTmp(2))
            blnRetrun = True
        ElseIf arrTmp(i) Like "*HOME*" Then
            intTimes = Val(Mid(arrTmp(1), InStr(UCase(arrTmp(1)), "HOME") + 4))
            blnRetrun = True
        End If
        If arrTmp(i) Like "ORADB*" Then
            intVer = ValEx(Mid(arrTmp(1), 6))
            intServer = 1
            blnRetrun = True
        ElseIf arrTmp(i) Like "ORACLIENT*" Then
            intVer = ValEx(Mid(arrTmp(1), 10))
            intServer = 2
            blnRetrun = True
        ElseIf arrTmp(i) Like "*CLIENT*" Then
            intServer = 2
            intVer = ValEx(arrTmp(i))
            blnRetrun = True
        End If
    Next
    GetOraInfoByRegKey = blnRetrun
End Function

Public Function ValEx(ByVal varInput As Variant) As Variant
'���ܣ�����Valֻ�������ֿ�ͷʶ��ValEx�Ե�һ�����ֽ���ʶ��
    Dim arrTmp As Variant, lngPos As Long
    If Val(varInput) = 0 Then
        varInput = varInput & ""
        If Trim(varInput) = "" Then ValEx = 0: Exit Function
        For lngPos = 1 To Len(varInput)
            If IsNumeric(Mid(varInput, lngPos, 1)) Then Exit For
        Next
        If lngPos = Len(varInput) + 1 Then
            ValEx = 0
        Else
            ValEx = Val(Mid(varInput, lngPos))
        End If
    Else
        ValEx = Val(varInput)
    End If
End Function

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
    
    With cboServer
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
        cboServer.Text = strTemp
        cboServer.SelStart = Len(strInput)
        cboServer.SelLength = 100
    Else
        cboServer.Text = strInput
        cboServer.SelStart = lngStart
    End If

End Sub

Private Sub ClearComponent()
'���ܣ�--���ע���[��������]--��Ϊ��ͬ�����ݿ����ʹ�õ�ϵͳ�Ͱ汾��ͬ
    If mblnFirst = True Then '����ʱ�Կؼ��ĸ�ֵ����������
        SaveSetting "ZLSOFT", "ע����Ϣ", "��������", ""
    End If
End Sub

Public Sub Docmd(ByVal strCmd As String)
    Dim ArrCommand
    Dim i As Integer

    ArrCommand = Split(strCmd, " ")
    
    If InStr(ArrCommand(0), "=") > 0 And InStr(ArrCommand(0), "&") = 0 Then
        '�������������õ���̨��¼�ĸ�ʽ
        For i = LBound(ArrCommand) To UBound(ArrCommand)
            If UCase(ArrCommand(i)) Like "USER=*" Then
                Me.txtUser.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "PASS=*" Then
                Me.txtPassWord.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "SERVER=*" Then
                Me.cboServer.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "ONLYONE=*" Then
                If Split(ArrCommand(i), "=")(1) = "1" Then
                    If App.PrevInstance = True Then
                        MsgBox "�����ظ������������"
                        End
                    End If
                End If
            End If
        Next
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then Call cmdOK_Click
    End If
End Sub

Private Sub txtUser_LostFocus()
    Call UpdateUser
End Sub

Private Sub txtUser_Validate(Cancel As Boolean)
    Call UpdateUser
End Sub

Private Sub UpdateUser()
On Error GoTo errH
    If IsNumeric(txtUser.Text) Then
        txtUser.Text = "U" & txtUser.Text
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
End Sub

Private Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
'����:���������뷨����ر����뷨
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String, blnNotCloseIme As Boolean
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    blnNotCloseIme = True
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '��Ҫ�����뷨�������ж��Ƿ��������뷨
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
                End If
            End If
        ElseIf blnOpen = False Then
            '�������뷨��������Ӧ�˹ر����뷨������
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0

    If blnNotCloseIme And blnOpen = False Then
        '����windows Vistaϵͳ��Ӣ�����뷨��ImmIsIME���Գ���true�����뷨,���,��Ҫ��������.
        '���˺�:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function


