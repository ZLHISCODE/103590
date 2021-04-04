VERSION 5.00
Begin VB.Form frmUserCheckLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户连接"
   ClientHeight    =   2700
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4875
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboServer 
      Height          =   300
      Left            =   1793
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1716
      Width           =   2592
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Left            =   0
      TabIndex        =   8
      Top             =   1992
      Width           =   4965
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2025
      TabIndex        =   6
      Top             =   2256
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3285
      TabIndex        =   7
      Top             =   2256
      Width           =   1100
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1793
      MaxLength       =   30
      TabIndex        =   1
      Text            =   "sys"
      Top             =   900
      Width           =   2592
   End
   Begin VB.TextBox txtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1793
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1308
      Width           =   2592
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   210
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      Caption         =   "服务器"
      Height          =   180
      Left            =   1092
      TabIndex        =   4
      Top             =   1776
      Width           =   540
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1092
      TabIndex        =   0
      Top             =   960
      Width           =   540
   End
   Begin VB.Label lblPWD 
      AutoSize        =   -1  'True
      Caption         =   "口令"
      Height          =   180
      Left            =   1272
      TabIndex        =   2
      Top             =   1368
      Width           =   360
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "需要连接到指定的实例来终止会话，请输入实例""""ORCL""""的DBA用户连接信息。"
      Height          =   360
      Left            =   1140
      TabIndex        =   9
      Top             =   240
      Width           =   3555
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUserCheckLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcnOracle As ADODB.Connection  '返回其他实例的连接
Private mblnFirst As Boolean  '为True表示已经正常显示出

Private mcolServer As New Collection
Private mblnOk As Boolean

Private mlngINST_ID As Long
Private mstrInstance As String



Public Function ShowLogin(ByRef cnOracle As ADODB.Connection, ByVal lngINST_ID As Long, ByVal strUname, ByVal strPassword) As Boolean
'功能：验证用户登录
'参数：lngThis_INST_ID=登录到指定的实例号
'          cnOracle=返回的连接
    mlngINST_ID = lngINST_ID
    
    Me.Show 1
    
    txtUser.Text = strUname
    txtPWD.Text = strPassword
    Set cnOracle = mcnOracle
   
    ShowLogin = mblnOk
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Set mcnOracle = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strNote As String
    Dim strUser As String, strPwd As String, strServer As String
    Dim intPos As Integer

    
    SetConState False
    
    '------检验用户是否oracle合法用户----------------
    strUser = Trim(txtUser.Text)
    strPwd = Trim(txtPWD.Text)
    strServer = Trim(cboServer.Text)
    
    '有效字符串效验
    If Len(Trim(txtUser.Text)) = 0 Then
        strNote = "请输入用户名。"
        txtUser.SetFocus
        GoTo InputError
    End If
    
    If Len(strUser) <> 1 Then
        If Mid(strUser, 1, 1) = "/" Or Mid(strUser, 1, 1) = "@" Or Mid(strUser, Len(strUser) - 1, 1) = "/" Or Mid(strUser, Len(strUser) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "用户名错误。"
            Exit Sub
        End If
    End If
    If Trim(strPwd) <> "" And Len(strPwd) <> 1 Then
        If Mid(strPwd, Len(strPwd) - 1, 1) = "/" Or Mid(strPwd, Len(strPwd) - 1, 1) = "@" Or Mid(strPwd, 1, 1) = "/" Or Mid(strPwd, 1, 1) = "@" Then
            txtPWD.SetFocus
            strNote = "口令错误。"
            GoTo InputError
        End If
    End If
    If Trim(strServer) <> "" Then
        If Mid(strServer, Len(strServer) - 1, 1) = "/" Or Mid(strServer, Len(strServer) - 1, 1) = "@" Or Mid(strServer, 1, 1) = "/" Or Mid(strServer, 1, 1) = "@" Then
            strNote = "主机连接串错误。"
            cboServer.SetFocus
            GoTo InputError
        End If
    End If
    
    '分离字符串
    intPos = InStr(1, strUser, "@")
    If intPos > 0 Then
        strServer = Mid(strUser, intPos + 1)
        strUser = Mid(strUser, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUser, "/")
    If intPos > 0 Then
        strPwd = Mid(strUser, intPos + 1)
        strUser = Mid(strUser, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPwd, "@")
    If intPos > 0 Then
        strServer = Mid(strPwd, intPos + 1)
        strPwd = Mid(strPwd, 1, intPos - 1)
    End If
    
    If Len(Trim(strPwd)) = 0 Then
        strNote = "请输入密码"
        txtPWD.SetFocus
        GoTo InputError
    End If
    
    strUser = UCase(strUser)
    
    If Not OpenConnection(strServer, strUser, strPwd, mcnOracle) Then
        If txtPWD.Enabled Then txtPWD.SetFocus
        SetConState
        
        Exit Sub
    ElseIf CheckIsDBA(mcnOracle) = False Then
        MsgBox "不是数据库DBA用户，请重新输入。", vbExclamation, gstrSysName
        If txtUser.Enabled Then txtUser.SetFocus
        
        SetConState
        Exit Sub
    End If
    
    If CheckThisInstance = False Then
        mcnOracle.Close
        
        MsgBox "请重新选择服务名，连接到指定的实例""" & mstrInstance & """。", vbExclamation, gstrSysName
        If cboServer.Enabled Then cboServer.SetFocus
        
        SetConState
        Exit Sub
    End If
    
    
    mblnOk = True
    Unload Me
    Exit Sub
    
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbExclamation, gstrSysName
    End If
End Sub

Private Function CheckThisInstance() As Boolean
'功能：检查当前实例是否为指定的实例
    Dim rstmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    strSQL = "Select Instance_Name From V$instance Where Instance_Name = '" & mstrInstance & "'"
    rstmp.Open strSQL, mcnOracle
    
    CheckThisInstance = rstmp.RecordCount > 0
    Exit Function
    
errh:
    MsgBox Err.Description, vbExclamation
End Function

Private Sub Form_Activate()
    If mblnFirst = False Then
        If Trim(txtUser.Text) = "" Then
            cmdOK.Default = False
            If txtUser.Enabled Then txtUser.SetFocus
        Else
            If txtPWD.Enabled Then
                txtPWD.SetFocus
            Else
                cmdOK.SetFocus
            End If
        End If
        mblnFirst = True
        If Trim(txtUser.Text) <> "" And Trim(txtPWD.Text) <> "" Then Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "txtPWD" Then
            Call cmdOK_Click
        Else
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cboServer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '回车键另行处理
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii, cboServer, mcolServer)
        End If
    End If
End Sub

Private Sub txtUser_GotFocus()
    If Me.ActiveControl Is txtUser Then
        SelAll txtUser
        OpenIme False
    End If
End Sub

Private Sub txtPwd_GotFocus()
    SelAll txtPWD
End Sub

Private Sub cboServer_GotFocus()
    If Me.ActiveControl Is cboServer Then
        SelAll cboServer
        OpenIme False
    End If
End Sub

Private Sub Form_Load()
    Dim varItem As Variant
    
    mblnFirst = False
    Set mcnOracle = New ADODB.Connection
 
    Set mcolServer = LoadServer(cboServer.ToolTipText)
    For Each varItem In mcolServer
        cboServer.AddItem varItem(0)
    Next
    
    mstrInstance = Get_INST_Name(mlngINST_ID)
    lblNote.Caption = "需要连接到指定的实例来终止会话，请输入实例""" & mstrInstance & """的DBA用户帐户信息。"
    
End Sub


Private Function Get_INST_Name(ByVal lngINST_ID As Long) As String
'功能：根据实例ID获取实例名
    Dim rstmp As ADODB.Recordset, strSQL As String

    strSQL = "Select Instance_Name From Gv$instance Where Inst_Id = [1]"
    
    On Error GoTo errh
    Set rstmp = OpenSQLRecord(strSQL, Me.Caption, lngINST_ID)
    If rstmp.RecordCount > 0 Then
        Get_INST_Name = rstmp!Instance_Name
    End If
    
    Exit Function
errh:
    MsgBox Err.Description, vbExclamation
End Function


Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    cmdOK.Enabled = BlnState
    cmdCancel.Enabled = BlnState
End Sub


Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub


Public Function OpenIme(Optional blnOpen As Boolean = False, Optional strImeName As String) As Boolean
'功能:打开中文输入法，或关闭输入法
'参数：strImeName-打开指定的输入法
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    
 
    '用户没进行设置，就不处理
    If blnOpen Then
        If strImeName <> "" Then
            strIme = strImeName
        End If
        If strIme = "" Then Exit Function                  '要求打开输入法，但是又没有设置
    End If
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))

    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '需要打开输入法。接着判断是否指定输入法
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then
                        OpenIme = True
                        Exit Function
                    End If
                End If
            End If
        ElseIf blnOpen = False Then
            '不是中文输入法，正好是应了关闭输入法的请求
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnOpen = False Then
        '由于windows Vista系统的英文输入法用ImmIsIME测试出是1的输入法,因此,需要单独处理.
        '刘兴宏:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function



Public Function LoadServer(ByRef strFileInfo As String) As Collection
'功能：读出本地的服务器列表
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    Dim arrTmp As Variant
    Dim rsOraHome As ADODB.Recordset
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer
    Dim i As Long
    Dim colServer As New Collection

    Set rsOraHome = New ADODB.Recordset
    With rsOraHome
        .Fields.Append "Name", adVarChar, 256 'Name
        .Fields.Append "VerSion", adInteger  '版本
        .Fields.Append "Times", adInteger '第几次安装
        .Fields.Append "Server", adInteger '1-服务器,2-客户端
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '1:读取64位下32目录会自动定位到SOFTWARE\Wow6432Node\Oracle 2：读取32位下32位目录
        arrTmp = GetAllSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle")
        If TypeName(arrTmp) = "Empty" Then
            If Is64bit Then
                strFileInfo = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Oracle！"
            Else
                strFileInfo = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Oracle！"
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
            If UBound(arrTmp) <> -1 Then ''顶级目录可能有Oracle_Home信息，默认读取这个
                .AddNew Array("Name", "VerSion", "Times", "Server"), Array("", 0, 0, 1): .Update
            End If
            .Sort = "VerSion Desc,Times Desc,Server"
            Do While Not .EOF
                strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle" & !Name, "ORACLE_HOME")
                If strPath = "" And !Name & "" = "" Then
                    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle", "ORA_CRS_HOME")
                End If
                If strPath <> "" Then
                    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i以上
                    If gobjFile.FileExists(strFile) Then Exit Do
                    strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
                    If gobjFile.FileExists(strFile) Then Exit Do
                End If
                strFile = ""
                .MoveNext
            Loop
        End If
    End With
    If strFile = "" Then Exit Function
    strFileInfo = "服务器列表来源:" & strFile
    lngFile = FreeFile()
    Open strFile For Input Access Read As lngFile
    Do Until EOF(lngFile)
        Input #lngFile, strLine
        strLine = Trim(strLine)
        If strLine <> "" And Left(strLine, 1) <> "#" Then
            '非注释行或空行
            If InStr(strLine, "(") = 0 And InStr(strLine, ")") = 0 Then
                '该行的内容就是服务器名了，把所有内容都初始化
                strServer = Trim(Mid(strLine, 1, InStr(strLine, "=") - 1))
                strComputer = ""
                strSID = ""
            ElseIf InStr(strLine, "(ADDRESS") > 0 Then
                '该行的内容是主机名
                If InStr(strLine, "PROTOCOL = TCP") > 0 And InStr(strLine, "PORT = ") > 0 Then
                    '符合我们的程序要求
                    strComputer = Mid(strLine, InStr(strLine, "HOST =") + Len("HOST ="))
                    strComputer = Trim(Mid(strComputer, 1, InStr(strComputer, ")") - 1))
                End If
            Else
                lngPos = InStr(strLine, "(SID")
                If lngPos = 0 Then
                    lngPos = InStr(strLine, "(SERVICE_NAME")
                End If
                
                If lngPos > 0 Then
                    '该行的内容是实例名
                    strSID = Mid(strLine, InStr(lngPos, strLine, "=") + 1)
                    strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
                    
                    If strServer <> "" And strComputer <> "" And strSID <> "" Then
                        '已经得到所有需要的内容
                        colServer.Add Array(strServer, strComputer, strSID)
                    End If
                End If
            End If
        End If
    Loop
    Close #lngFile
    
    Set LoadServer = colServer
End Function


 Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '功能：是否是64位系统
    '返回：
    '******************************************************************************************************************
    Dim handle As Long
    Dim bolFunc As Boolean
        
    bolFunc = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), bolFunc
    End If
    Is64bit = bolFunc
End Function


Private Function GetOraInfoByRegKey(ByVal strOraHome As String, ByRef intVer As Integer, ByRef intTimes As Integer, ByRef intServer As Integer) As Boolean
'功能:通过OracleHome键获取Oracle信息
    Dim arrTmp As Variant
    Dim i As Long, blnRetrun As Boolean
    'KEY_OraDb11g_home1_32bit
    'Key_Ora*版本Home_32Bit
    'Key_Ora*版本_Home*
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
