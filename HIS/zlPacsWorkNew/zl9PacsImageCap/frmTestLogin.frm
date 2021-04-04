VERSION 5.00
Begin VB.Form frmTestLogin 
   Caption         =   "�ɼ����Ե�¼"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTestLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4755
   StartUpPosition =   1  '����������
   Begin VB.Timer tmerOpenVideo 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   120
      Top             =   1695
   End
   Begin VB.PictureBox picDepartment 
      BorderStyle     =   0  'None
      Height          =   2670
      Left            =   570
      ScaleHeight     =   2670
      ScaleWidth      =   3705
      TabIndex        =   8
      Top             =   210
      Visible         =   0   'False
      Width           =   3705
      Begin VB.CommandButton Command1 
         Caption         =   "ȷ ��(&S)"
         Height          =   435
         Left            =   2160
         TabIndex        =   13
         Top             =   1305
         Width           =   1245
      End
      Begin VB.ComboBox cbxDepartment 
         Height          =   330
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   735
         Width           =   2220
      End
      Begin VB.Label Label1 
         Caption         =   "��¼���ң�"
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   9
         Top             =   765
         Width           =   1110
      End
   End
   Begin VB.ComboBox cbxSystem 
      Height          =   330
      ItemData        =   "frmTestLogin.frx":000C
      Left            =   1725
      List            =   "frmTestLogin.frx":0016
      TabIndex        =   11
      Text            =   "1291-Ӱ��ɼ�"
      Top             =   300
      Width           =   2220
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   435
      Left            =   2670
      TabIndex        =   4
      Top             =   2415
      Width           =   1245
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "�� ¼(&E)"
      Height          =   435
      Left            =   840
      TabIndex        =   3
      Top             =   2415
      Width           =   1245
   End
   Begin VB.TextBox txtServer 
      Height          =   360
      Left            =   1710
      TabIndex        =   2
      Top             =   1770
      Width           =   2220
   End
   Begin VB.TextBox txtPwd 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1725
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1290
      Width           =   2205
   End
   Begin VB.TextBox txtUser 
      Height          =   360
      Left            =   1725
      TabIndex        =   0
      Text            =   "zlhis"
      Top             =   780
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "ϵͳ���ͣ�"
      Height          =   315
      Index           =   4
      Left            =   600
      TabIndex        =   12
      Top             =   330
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "  ��������"
      Height          =   315
      Index           =   2
      Left            =   585
      TabIndex        =   7
      Top             =   1815
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "  ��  �룺"
      Height          =   315
      Index           =   1
      Left            =   585
      TabIndex        =   6
      Top             =   1305
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "  �û�����"
      Height          =   315
      Index           =   0
      Left            =   585
      TabIndex        =   5
      Top             =   795
      Width           =   1110
   End
End
Attribute VB_Name = "frmTestLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents mobjVideo As clsPacsCapture
Attribute mobjVideo.VB_VarHelpID = -1
Private mcnOracle As ADODB.Connection
Private mstrPrivs As String
Private WithEvents mHotCapture As clsHookKey
Attribute mHotCapture.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLogin_Click()
On Error GoTo errHandle
    If OraDataOpen(txtServer.Text, txtUser.Text, TranPasswd(txtPwd.Text)) = False Then Exit Sub
    
    glngSys = 100
    glngModule = Val(cbxSystem.Text)
    
    
    Call InitCommonLib(mcnOracle)
    mstrPrivs = GetInsidePrivs(glngModule)
    
    If InitDepts Then
        picDepartment.Visible = True
    End If
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    str��Դ = "1,2,3"
    
    If InStr(mstrPrivs, "���п���") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & zlCL_GetNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')> 0 And B.�������� IN('���')" & _
            " Order by A.����"
    Else
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & zlCL_GetNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')>0  And B.�������� IN('���')" & _
            " Order by A.����"
    End If
   

    
    Set rsTmp = zlCL_GetDBObj.OpenSQLRecord(strSQL, Me.Caption, CStr("," & str��Դ & ","))
    
    If rsTmp.EOF Then
        MsgboxCus "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    Else
        Do Until rsTmp.EOF
            cbxDepartment.AddItem Val(Nvl(rsTmp!ID)) & "-" & Nvl(rsTmp!����)
            rsTmp.MoveNext
        Loop
        
        InitDepts = True
    End If
End Function



Private Sub Command1_Click()
    tmerOpenVideo.Enabled = True
End Sub

Private Sub Form_Load()
BUGEX "TestLogin Form_Load 1"
    Set mcnOracle = New ADODB.Connection
     
    txtUser.SelStart = 0
    txtUser.SelLength = 250
    
BUGEX "TestLogin Form_Load 2"

    gstrHotKeyTest = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")

    Set mHotCapture = New clsHookKey
    
    If Trim(gstrHotKeyTest) <> "" Then Call mHotCapture.EnableHook(WM_KEYDOWN)
    
    
BUGEX "TestLogin Form_Load End"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mHotCapture.FreeHook
    
    Set mHotCapture = Nothing
    Set mobjVideo = Nothing
    Set mcnOracle = Nothing
    
    If Not gobjComLib Is Nothing Then
        Call zlCL_CloseWindow
        Set gobjComLib = Nothing
    End If
End Sub



Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String

    On Error Resume Next
    err = 0
    DoEvents
    With mcnOracle
        If .State = 1 Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '���������Ϣ
            MsgboxCus err.Description, vbInformation, G_STR_HINT_TITLE

            OraDataOpen = False
            Exit Function
        End If
    End With

    err = 0
    On Error GoTo errhand

    OraDataOpen = True
    Exit Function

errhand:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    OraDataOpen = False
    err = 0
End Function

Private Sub mHotCapture_OnKeyBoardLHook(ByVal lngMsg As Long, ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
On Error GoTo errHandle

    If GetKeyAliasEx(lngVkCode) = gstrHotKeyTest Then
        If Not mobjVideo Is Nothing Then
            Call mobjVideo.zlCaptureImg
        End If
    End If
Exit Sub
errHandle:
    MsgBox Me, err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub mobjVideo_OnDockClose()
    Me.Show
End Sub

Private Sub tmerOpenVideo_Timer()
On Error GoTo errHandle
    tmerOpenVideo.Enabled = False
    
    Set mobjVideo = New clsPacsCapture
    
    glngDepartId = Val(cbxDepartment.Text)

'    mobjVideo.VideoWindow.ShowInTaskbar = True
    
    Call mobjVideo.zlInitModule(mcnOracle, glngSys, glngModule, mstrPrivs, glngDepartId, Me.hWnd, Me, True)
    Call mobjVideo.zlShowPopupVideo
    
    picDepartment.Visible = False
    
    Call Me.Hide
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function

