VERSION 5.00
Begin VB.Form frmConnectionEdit 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6150
   ControlBox      =   0   'False
   Icon            =   "frmConnectionEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtLinkName 
      Height          =   300
      Left            =   1065
      MaxLength       =   30
      TabIndex        =   0
      Top             =   345
      Width           =   1725
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2445
      MaxLength       =   3
      TabIndex        =   5
      Tag             =   "IP��ַ"
      Top             =   945
      Width           =   315
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1995
      MaxLength       =   3
      TabIndex        =   4
      Tag             =   "IP��ַ"
      Top             =   945
      Width           =   315
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1545
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "IP��ַ"
      Top             =   945
      Width           =   315
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1095
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "IP��ַ"
      Top             =   945
      Width           =   315
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Left            =   1065
      MaxLength       =   20
      TabIndex        =   25
      Tag             =   "IP"
      Text            =   "   ��   ��   ��"
      Top             =   900
      Width           =   1725
   End
   Begin VB.TextBox txtNotes 
      Height          =   1320
      Left            =   1065
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1965
      Width           =   4695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4665
      TabIndex        =   13
      Top             =   3885
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "����(&T)"
      Height          =   350
      Left            =   420
      TabIndex        =   11
      Top             =   3885
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3570
      TabIndex        =   12
      Top             =   3885
      Width           =   1100
   End
   Begin VB.Frame fraMain 
      Height          =   30
      Index           =   1
      Left            =   -195
      TabIndex        =   18
      Top             =   3540
      Width           =   6570
   End
   Begin VB.TextBox txtDatabase 
      Height          =   300
      Left            =   4260
      MaxLength       =   50
      TabIndex        =   1
      Top             =   345
      Width           =   1500
   End
   Begin VB.TextBox txtPort 
      Height          =   300
      Left            =   4260
      MaxLength       =   5
      TabIndex        =   7
      Top             =   885
      Width           =   1500
   End
   Begin VB.TextBox txtPasswd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4260
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1425
      Width           =   1500
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1065
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1425
      Width           =   1725
   End
   Begin VB.Label lblLinkName 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   405
      Width           =   720
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   2820
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblNotes 
      AutoSize        =   -1  'True
      Caption         =   "˵  ��"
      Height          =   180
      Left            =   420
      TabIndex        =   24
      Top             =   1965
      Width           =   540
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   5790
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   3
      Left            =   5790
      TabIndex        =   22
      Top             =   780
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   2820
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   5790
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   2820
      TabIndex        =   19
      Top             =   765
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      Caption         =   "ʵ����"
      Height          =   180
      Left            =   3585
      TabIndex        =   17
      Top             =   405
      Width           =   540
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      Caption         =   "�˿ں�"
      Height          =   180
      Left            =   3585
      TabIndex        =   16
      Top             =   945
      Width           =   540
   End
   Begin VB.Label lblIp 
      AutoSize        =   -1  'True
      Caption         =   "IP��ַ"
      Height          =   180
      Left            =   420
      TabIndex        =   15
      Top             =   945
      Width           =   540
   End
   Begin VB.Label lblPasswd 
      AutoSize        =   -1  'True
      Caption         =   "��  ��"
      Height          =   180
      Left            =   3585
      TabIndex        =   14
      Top             =   1485
      Width           =   540
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   420
      TabIndex        =   6
      Top             =   1485
      Width           =   540
   End
End
Attribute VB_Name = "frmConnectionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ڱ���
Private mstrLinkName As String '��������
Private mstrUser As String  '�û���
Private mstrPasswd As String  '����
Private mstrDatabase As String  '���ݿ�ʵ����
Private mstrNotes As String  '˵��
Private mStrIP As String  'IP��ַ
Private mstrPort As String  '�˿ں�
Private mlngId As Long  '���ӱ�ţ�������Ϊ0ʱ��ʾ��������
Private mblnState As Boolean  '����Ƿ���������Ӳ���
Private mblnSaveClick As Boolean  '����Ƿ񱣴�
Private mclsCiph As clsCipher  '����һ���ӽ���ʵ��������

Private Const lng_���� As Long = 0

Private Enum CheckTag
    CT_�������� = 0
    CT_ʵ���� = 1
    CT_IP��ַ = 2
    CT_�˿ں� = 3
    CT_�û��� = 4
    CT_���� = 5
End Enum

Public Function ShowEdit(lngID As Long, strLinkName As String, strUser As String, strPasswd As String, strIp As String, _
                        strPort As String, strDatabase As String, strNotes As String) As Boolean
    '-------------------------------------------------------------------------------
    '--���ܣ���ʾ�ͱ༭������Ϣ
    '--������strUser:�û���, strIp:������IP, strDatabase:ʵ����, strNotes:��ע˵��, strCaption:����, lngId:���
    '-------------------------------------------------------------------------------
    Set mclsCiph = New clsCipher
    mstrLinkName = strLinkName
    mlngId = lngID
    mStrIP = strIp
    mstrPort = strPort
    mstrUser = strUser
    mstrPasswd = mclsCiph.Decipher(MSTR_DBLINK_KEY, strPasswd)
    mstrDatabase = strDatabase
    mstrNotes = strNotes
    
    Me.Caption = IIf(lngID = lng_����, "������������", "�޸���������")
    Me.Show vbModal, frmMDIMain
    
    If mblnSaveClick Then
        lngID = mlngId
        strUser = mstrUser
        strIp = mStrIP
        strPort = mstrPort
        strDatabase = mstrDatabase
        strNotes = mstrNotes
        strPasswd = mstrPasswd
        strLinkName = mstrLinkName
    End If
    ShowEdit = mblnSaveClick
    '�������
    Call ClearDate
End Function

Private Sub FillData(ByVal strLinkName As String, ByVal strIps As String, ByVal strPort As String, _
                    ByVal strUser As String, ByVal strPasswd As String, ByVal strDatabase As String, ByVal strNotes As String)
    '-------------------------------------------------------------------------------
    '--���ܣ���Ҫ�޸�����ʱ���������������ʾ�ڶ�Ӧλ��
    '-------------------------------------------------------------------------------
    Dim strIp() As String
    Dim i As Long
    
    On Error GoTo errH:
    strIp = Split(strIps, ".")
    For i = 0 To 3
        txtIp(i) = strIp(i)
    Next
    txtLinkName = strLinkName
    txtPort.Text = strPort
    txtDatabase.Text = strDatabase
    txtUser.Text = strUser
    txtPasswd.Text = strPasswd
    txtNotes.Text = strNotes
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub cmdCancel_Click()
    mblnSaveClick = False
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim rsID As ADODB.Recordset
    Dim strSQL As String
    Dim strNote As String
    
    On Error GoTo errH:
    '��ʼ����ǣ���ֹ�ڵ������֮����ȥ�޸���Ϣ�����±���ʱ��������
    mblnState = False
    mblnSaveClick = True
    Call cmdTest_Click
    '�����ݼ������ݿ�
    If mblnState Then
        Set mclsCiph = New clsCipher
        If mlngId = lng_���� Then
            strSQL = "Zl_Zlconnections_Edit(0,Null,'" & Trim(txtLinkName.Text) & "','" & Trim(txtUser.Text) & "','" & _
                                            mclsCiph.Cipher(MSTR_DBLINK_KEY, txtPasswd.Text) & "','" & _
                                            txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & _
                                            "." & txtIp(3).Text & "'," & txtPort.Text & ",'" & txtDatabase.Text & _
                                            "','" & txtNotes.Text & "')"
            Call ExecuteProcedure(strSQL, Me.Caption)
            strSQL = "Select Max(���) As ��� From Zlconnections"
            Set rsID = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ѯ���ID")
            mlngId = rsID!���
            '������Ҫ������־
            Call SaveAuditLog(1, "����", "������ӡ�" & txtLinkName.Text & "��")
        Else
            strSQL = "Zl_Zlconnections_Edit(1," & mlngId & ",'" & Trim(txtLinkName.Text) & "','" & Trim(txtUser.Text) & _
                                            "','" & mclsCiph.Cipher(MSTR_DBLINK_KEY, txtPasswd.Text) & "','" & _
                                            txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & _
                                            "." & txtIp(3).Text & "'," & txtPort.Text & ",'" & txtDatabase.Text & _
                                            "','" & txtNotes.Text & "')"
            Call ExecuteProcedure(strSQL, Me.Caption)
            If mstrLinkName <> Trim(txtLinkName.Text) Then strNote = ",�����ɡ�" & mstrLinkName & "���޸�Ϊ��" & Trim(txtLinkName.Text) & "��"
            If mstrDatabase <> txtDatabase.Text Then strNote = strNote & ",ʵ������" & mstrDatabase & "�޸�Ϊ" & txtDatabase.Text
            If mStrIP <> txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & "." & txtIp(3).Text Then
                strNote = strNote & ",IP��ַ��" & mStrIP & "�޸�Ϊ" & txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & "." & txtIp(3).Text
            End If
            If mstrPort <> txtPort.Text Then strNote = strNote & ",�˿ں���" & mstrPort & "�޸�Ϊ" & txtPort.Text
            If mstrUser <> Trim(txtUser.Text) Then strNote = strNote & ",�û�����" & mstrUser & "�޸�Ϊ" & Trim(txtUser.Text)
            '������Ҫ������־
            If strNote <> "" Then
                Call SaveAuditLog(2, "�޸�", "�޸����ӡ�" & mstrLinkName & "��" & strNote)
            End If
        End If
        
        mstrLinkName = Trim(txtLinkName.Text)
        mstrUser = Trim(txtUser.Text)
        mStrIP = txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & "." & txtIp(3).Text
        mstrPort = txtPort.Text
        mstrDatabase = txtDatabase.Text
        mstrNotes = txtNotes.Text
        mstrPasswd = mclsCiph.Cipher(MSTR_DBLINK_KEY, txtPasswd.Text)
        Unload Me
    Else
        mblnSaveClick = False
    End If
    Exit Sub
errH:
    mblnSaveClick = False
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub cmdTest_Click()
    Dim cnOracle As ADODB.Connection
    Dim strServerName As String
    
    mstrUser = Trim(txtUser.Text)
    mstrPasswd = txtPasswd.Text
    If CheckData Then
        'strServerName = "192.168.2.13:1521/dyyy"
        strServerName = Val(txtIp(0).Text) & "." & Val(txtIp(1).Text) & "." & Val(txtIp(2).Text) & "." & Val(txtIp(3).Text) & _
                        ":" & Val(txtPort.Text) & "/" & Trim(txtDatabase.Text)
        Set cnOracle = gobjRegister.GetConnection(strServerName, mstrUser, mstrPasswd, False, OraOLEDB, , False)
        If cnOracle.State = adStateOpen Then
            mblnState = True
            If mblnSaveClick = False Then MsgBox "���ӿ��ã�", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Function CheckData() As Boolean
    '��������Ƿ���д����
    Dim blnDataType As Boolean
    Dim i As Integer
    
    '���IP��ַ
    For i = 0 To 3
        If txtIp(i).Text = "" Then
            blnDataType = False
        Else
            blnDataType = True
        End If
    Next
    If blnDataType = False Then
        lblCheck(CT_IP��ַ).Visible = True
    Else
        lblCheck(CT_IP��ַ).Visible = False
    End If
    
    '�����������
    If txtLinkName.Text = "" Then
        lblCheck(CT_��������).Visible = True
    Else
        lblCheck(CT_��������).Visible = False
    End If
    
    '���˿�
    If txtPort.Text = "" Then
        lblCheck(CT_�˿ں�).Visible = True
    Else
        lblCheck(CT_�˿ں�).Visible = False
    End If
    
    '���ʵ����
    If txtDatabase.Text = "" Then
        lblCheck(CT_ʵ����).Visible = True
    Else
        lblCheck(CT_ʵ����).Visible = False
    End If
    
    '����û���
    If txtUser.Text = "" Then
        lblCheck(CT_�û���).Visible = True
    Else
        lblCheck(CT_�û���).Visible = False
    End If
    
    '�������
    If txtPasswd.Text = "" Then
        lblCheck(CT_����).Visible = True
    Else
        lblCheck(CT_����).Visible = False
    End If
    
    '����������ƺϷ���
    If InStr(1, txtLinkName.Text, "'") <> 0 Then
        MsgBox "���������ơ��в������뵥����!", vbInformation + vbOKOnly, gstrSysName
        txtLinkName.SetFocus
        Exit Function
    End If
    
    '���˵���Ϸ���
    If InStr(1, txtNotes.Text, "'") <> 0 Then
        MsgBox "��˵�����в������뵥����!", vbInformation + vbOKOnly, gstrSysName
        txtNotes.SetFocus
        Exit Function
    End If
    For i = 0 To 5
        If lblCheck(i).Visible = True Then
            CheckData = False
            Select Case i
                Case CT_��������
                    txtLinkName.SetFocus
                Case CT_IP��ַ
                    txtIp(0).SetFocus
                Case CT_�˿ں�
                    txtPort.SetFocus
                Case CT_ʵ����
                    txtDatabase.SetFocus
                Case CT_�û���
                    txtUser.SetFocus
                Case CT_����
                    txtPasswd.SetFocus
            End Select
            MsgBox "�뽫��Ϣ��д������", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        Else
            CheckData = True
        End If
    Next
End Function

Private Sub ClearDate()
    mstrUser = ""
    mstrPasswd = ""
    mstrDatabase = ""
    mstrNotes = ""
    mStrIP = ""
    mblnState = False
    mblnSaveClick = False
End Sub

Private Sub Form_Load()
    '�������
    If mStrIP <> "" Then
        Call FillData(mstrLinkName, mStrIP, mstrPort, mstrUser, mstrPasswd, mstrDatabase, mstrNotes)
    End If
End Sub

Private Sub txtDatabase_KeyPress(KeyAscii As Integer)
    'ֻ�������Сд��ĸ������
    If Not ((KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtIp_Change(Index As Integer)
    Dim lngLineNo As Long '�к�
    Dim lngColNo As Long  '�к�
    
    Call GetCursorPos(Me.txtIp(Index).hwnd, lngLineNo, lngColNo)
    If lngColNo > 3 Then
        If Index < 3 Then
            If txtIp(Index + 1).Enabled Then txtIp(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtIp_GotFocus(Index As Integer)
    txtIp(Index).SelStart = 0
    txtIp(Index).SelLength = Len(txtIp(Index).Text)
End Sub

Private Sub txtIp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lngLineNo As Long '�к�
    Dim lngColNo As Long  '�к�
    err = 0
    On Error Resume Next

    Call GetCursorPos(Me.txtIp(Index).hwnd, lngLineNo, lngColNo)

    Select Case KeyCode
    Case 37     '<-

        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtIp(Index - 1).Enabled Then
                txtIp(Index - 1).SelStart = Len(txtIp(Index - 1))
                txtIp(Index - 1).SetFocus
            End If
        End If
    Case 39     '->
        If Index < 3 Then
            If lngColNo <= Len(txtIp(Index)) Then Exit Sub
            If txtIp(Index + 1).Enabled Then
                txtIp(Index + 1).SelStart = 0
                txtIp(Index + 1).SetFocus
            End If
        End If
    Case 8     'BACKSPACE
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtIp(Index - 1).Enabled Then
                txtIp(Index - 1).SelStart = Len(txtIp(Index - 1))
                txtIp(Index - 1).SetFocus
            End If
        End If
    Case Else
    End Select

End Sub

Private Sub txtIp_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
        If Chr(KeyAscii) = "." Then
            If Index < 3 And Index >= 0 And Trim(txtIp(Index)) <> "" Then
                If txtIp(Index + 1).Enabled Then txtIp(Index + 1).SetFocus
            End If
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Public Sub GetCursorPos(ByVal hwnd5 As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long
    Dim lParam As Long, wParam As Long
    Dim K As Long
    
    i = SendMessage(hwnd5, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16 'ȡ��Ŀǰ�������λ��ǰ�ж��ٸ�Byte
    LineNo = SendMessage(hwnd5, EM_LINEFROMCHAR, j, 0) 'ȡ�ù��ǰ���ж�����
    LineNo = LineNo + 1
    K = SendMessage(hwnd5, EM_LINEINDEX, -1, 0)
    'ȡ��Ŀǰ���������ǰ���ж��ٸ�Byte
    ColNo = j - K + 1
End Sub

Private Sub txtIp_LostFocus(Index As Integer)
    If txtIp(Index).Text = "" Then Exit Sub
    Select Case Index
    Case 0
        If Val(txtIp(Index).Text) < 1 Or Val(txtIp(Index).Text) > 233 Then
            MsgBox "��" & txtIp(Index).Text & "��������Ч���ָ��һ������1��233���ֵ", vbOKOnly + vbInformation, gstrSysName
            txtIp(Index).SetFocus
            txtIp(Index).Text = 233
        End If
    Case 1, 2, 3
        If (Not IsNumeric(txtIp(Index).Text)) Or Val(txtIp(Index).Text) > 255 Then
            MsgBox "��" & txtIp(Index).Text & "��������Ч���ָ��һ������0��255���ֵ", vbOKOnly + vbInformation, gstrSysName
            txtIp(Index).SetFocus
            txtIp(Index).Text = 255
        End If
    End Select
End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)
    If ActualLen(txtNotes.Text) >= 500 Then
        KeyAscii = 0
        MsgBox "������볤��Ϊ500���ַ�(250����)��", vbOKOnly + vbInformation, gstrSysName
    End If
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    'ֻ�������Сд��ĸ������
    If Not ((KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If Not (InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPort_LostFocus()
    If txtPort.Text = "" Then Exit Sub
    If Not IsNumeric(txtPort.Text) Then
        MsgBox "��" & txtPort.Text & "��������Ч���������ȷ�Ķ˿ںţ�", vbOKOnly + vbInformation, gstrSysName
        txtPort.SetFocus
        txtPort.Text = mstrPort
    End If
    
End Sub
