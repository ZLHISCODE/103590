VERSION 5.00
Begin VB.Form frmTwoUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ò����û�"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   Icon            =   "frmTwoUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   3720
      TabIndex        =   14
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "����"
      Height          =   350
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdSame 
      Caption         =   "ͳһ"
      Height          =   350
      Left            =   600
      TabIndex        =   13
      Top             =   2400
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "����ҽ��"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   2535
      Begin VB.TextBox txtUserID 
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   1500
      End
      Begin VB.TextBox txtPassWord 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label lblUserName 
         Caption         =   "�û���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "�û���"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "����"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ҽ��"
      Height          =   1935
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      Begin VB.TextBox txtPassWord 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   1500
      End
      Begin VB.TextBox txtUserID 
         Height          =   270
         Index           =   1
         Left            =   840
         TabIndex        =   1
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "����"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "�û���"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblUserName 
         Caption         =   "�û���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmTwoUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnOk As Boolean                 '�Ƿ�ȷ��
Public intDBState As Integer            '1--ͳһ��2--������
Public blnCnOracleIsNew As Boolean      '��¼�Ƿ����ݿ������Ƿ�Ϊ����̨HIS����
Public cnOracle As New ADODB.Connection
Public strUserNameHIS As String
Public strUserNameNew As String
Public strUserIDNew As String
Public strUserIDHIS As String

Private mstrUserIDNew As String
Private mstrPassWord As String


Private Sub cmdCancel_Click()
    blnOk = False
    Unload Me
End Sub

Private Sub cmdChange_Click()
    Dim strServerName As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If blnCnOracleIsNew = True Then
        mstrUserIDNew = Trim(txtUserID(0).Text)
        mstrPassWord = Trim(txtPassWord(0).Text)
    Else
        mstrUserIDNew = Trim(txtUserID(1).Text)
        mstrPassWord = Trim(txtPassWord(1).Text)
    End If
                        
    strServerName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="SERVER", Default:="")
    '�������ݿ�
    If Not OraDataOpen(strServerName, UCase(mstrUserIDNew), IIf(UCase(mstrUserIDNew) = "SYS" Or UCase(mstrUserIDNew) = "SYSTEM", mstrPassWord, TranPasswd(mstrPassWord))) Then
        intDBState = 1
        Exit Sub
    End If
    '�����û���
    strSql = _
        " Select A.ID,C.����ID,A.���,A.����,A.����,B.�û���" & _
        " From ��Ա�� A,�ϻ���Ա�� B,������Ա C" & _
        " Where A.ID = B.��ԱID And A.ID = C.��ԱID And C.ȱʡ = 1 And B.�û��� = USER" & _
            " and (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)"
    Set rsTemp = cnOracle.Execute(strSql)
    If rsTemp.EOF Then
        MsgBoxD Me, "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�"
        intDBState = 1
        Exit Sub
    Else
        strUserNameNew = rsTemp!����
        strUserIDNew = rsTemp!�û���
    End If
        
    blnCnOracleIsNew = Not blnCnOracleIsNew
    intDBState = 2
    blnOk = True
    Unload Me
End Sub

Private Sub cmdSame_Click()
    intDBState = 1
    mstrUserIDNew = ""
    mstrPassWord = ""
    strUserNameNew = ""
    strUserIDNew = ""
    blnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    '��ʼ������
    blnOk = False
    
    If mstrUserIDNew = "" Or intDBState = 1 Then   '����״̬��ͳһ�����ߵ�һ�ν���
        strUserNameNew = strUserNameHIS
        strUserIDNew = strUserIDHIS
        lblUserName(0).Caption = strUserNameHIS
        lblUserName(1).Caption = strUserNameNew
        txtUserID(0).Enabled = False
        txtPassWord(0).Enabled = False
        blnCnOracleIsNew = False
        
    Else    '���һ���˳��ǽ���״̬
        If blnCnOracleIsNew = False Then         '��������ͨ������̨��½��,���ҽ�����µ�½��
            lblUserName(0).Caption = strUserNameHIS
            lblUserName(1).Caption = strUserNameNew
            txtUserID(1).Text = mstrUserIDNew
            txtPassWord(1).Text = mstrPassWord
            
            '����ؼ�״̬
            txtUserID(0).Text = ""
            txtPassWord(0).Text = ""
            txtUserID(0).Enabled = False
            txtPassWord(0).Enabled = False
            txtUserID(1).Enabled = True
            txtPassWord(1).Enabled = True
        Else                                    '���������µ�½�ģ����ҽ����ͨ������̨��½��
            lblUserName(0).Caption = strUserNameNew
            lblUserName(1).Caption = strUserNameHIS
            txtUserID(0).Text = mstrUserIDNew
            txtPassWord(0).Text = mstrPassWord
            
            '����ؼ�״̬
            txtUserID(1).Text = ""
            txtPassWord(1).Text = ""
            txtUserID(1).Enabled = False
            txtPassWord(1).Enabled = False
            txtUserID(0).Enabled = True
            txtPassWord(0).Enabled = True
        End If
    End If
End Sub

Private Function TranPasswd(strOld As String) As String
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


Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSql As String
    Dim strError As String
    
    
    On Error Resume Next
    err = 0
    DoEvents
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '���������Ϣ
            strError = err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBoxD Me, "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBoxD Me, "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBoxD Me, "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBoxD Me, "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBoxD Me, "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBoxD Me, "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBoxD Me, "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
            Else
                MsgBoxD Me, strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    err = 0
    On Error GoTo errHand
        
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    err = 0
End Function


