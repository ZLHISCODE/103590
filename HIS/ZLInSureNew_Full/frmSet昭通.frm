VERSION 5.00
Begin VB.Form frmSet��ͨ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ղ�������"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtCOM 
      Height          =   300
      Left            =   1275
      TabIndex        =   4
      Top             =   1425
      Width           =   915
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "����(&T)"
      Height          =   400
      Left            =   270
      TabIndex        =   5
      Top             =   2145
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   3195
      TabIndex        =   7
      Top             =   2145
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   2085
      TabIndex        =   6
      Top             =   2145
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   4590
   End
   Begin VB.TextBox txtSN 
      Height          =   300
      Left            =   1282
      TabIndex        =   3
      Top             =   1020
      Width           =   3015
   End
   Begin VB.TextBox txtPort 
      Height          =   300
      Left            =   1282
      MaxLength       =   4
      TabIndex        =   2
      Top             =   615
      Width           =   915
   End
   Begin VB.TextBox txtServer 
      Height          =   300
      Left            =   1282
      TabIndex        =   1
      Top             =   210
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�������˿�"
      Height          =   180
      Index           =   3
      Left            =   270
      TabIndex        =   11
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���֤���"
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   9
      Top             =   1095
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�������˿�"
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������ַ"
      Height          =   180
      Index           =   0
      Left            =   277
      TabIndex        =   0
      Top             =   285
      Width           =   900
   End
End
Attribute VB_Name = "frmSet��ͨ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnIsOK As Boolean

Private Sub cmdCancel_Click()
    If UCase(txtServer.Text) <> UCase(txtServer.Tag) Then
        If MsgBox("���ݽ������޸ģ�δ������˳���", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    If UCase(txtSN.Text) <> UCase(txtSN.Tag) Then
        If MsgBox("���ݽ������޸ģ�δ������˳���", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    If UCase(txtPort.Text) <> UCase(txtPort.Tag) Then
        If MsgBox("���ݽ������޸ģ�δ������˳���", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    If UCase(txtCom.Text) <> UCase(txtCom.Tag) Then
        If MsgBox("���ݽ������޸ģ�δ������˳���", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    blnIsOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    If IsValid() = False Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_��ͨ & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_��ͨ & ",null,'��ͨ���֤','" & txtSN.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_��ͨ & ",null,'��ͨ������','" & txtServer.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_��ͨ & ",null,'��ͨ�˿ں�','" & txtPort.Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    '����ǰʹ�õĴ���д��ע���֮��
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", CStr(txtCom.Text)
    blnIsOK = True
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTest_Click()
    txtServer.Text = Trim(txtServer.Text)
    txtPort.Text = Trim(txtPort.Text)
    txtSN.Text = Trim(txtSN.Text)

    If txtServer.Text = "" Or txtPort.Text = "" Or txtSN.Text = "" Then
        MsgBox "�������ò����������ܽ������Ӳ���", vbInformation, "��������"
        Exit Sub
    End If
    If frmConn��ͨ.ConnCenter(txtServer.Text, txtPort.Text, txtSN.Text, UserInfo.ID) = True Then
        MsgBox "���������ӳɹ�", vbInformation, "����"
        frmConn��ͨ.ConnClose
    Else
        MsgBox "����������ʧ��", vbInformation, "����"
    End If
End Sub

Public Function ��������() As Boolean
    Dim rsTemp As New ADODB.Recordset, str����ֵ As String, strCN As String, strServer As String, lngPort As Long
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��ͨ)
        
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "��ͨ���֤"
                strCN = str����ֵ
            Case "��ͨ������"
                strServer = str����ֵ
            Case "��ͨ�˿ں�"
                lngPort = CLng(str����ֵ)
        End Select
        rsTemp.MoveNext
    Loop
    
    txtServer.Text = strServer
    txtServer.Tag = strServer
    txtPort.Text = lngPort
    txtPort.Tag = lngPort
    txtSN.Text = strCN
    txtSN.Tag = strCN
    txtCom.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", "1")
    txtCom.Tag = txtCom.Text
    Me.Show vbModal
    �������� = blnIsOK
End Function

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    txtServer.Text = Trim(txtServer.Text)
    txtPort.Text = Trim(txtPort.Text)
    txtSN.Text = Trim(txtSN.Text)
    txtCom.Text = Trim(txtCom.Text)

    If txtServer.Text = "" Or txtSN.Text = "" Or txtPort.Text = "" Or txtCom.Text = "" Then
        MsgBox "������������ҽ������", vbInformation, "��������"
        IsValid = False
        Exit Function
    End If
    
    '���ж��ַ��ĺϷ���
    If zlCommFun.StrIsValid(txtServer.Text, txtServer.MaxLength) = False Then
        zlControl.TxtSelAll txtServer
        txtServer.SetFocus
        Exit Function
    End If
    
    If zlCommFun.StrIsValid(txtSN.Text, txtSN.MaxLength) = False Then
        zlControl.TxtSelAll txtSN
        txtSN.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtCom.Text) Then
        MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        txtCom.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtPort.Text) Then
        MsgBox "�뽫�������˿ں�����������Ϣ", vbInformation, gstrSysName
        txtPort.SetFocus
        Exit Function
    End If
    
    '�����ӽ��в���
    If frmConn��ͨ.ConnCenter(txtServer.Text, txtPort.Text, txtSN.Text, UserInfo.ID) = False Then
        If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, "����ʧ��") = vbNo Then Exit Function
    Else
        frmConn��ͨ.ConnClose
    End If
    IsValid = True
End Function

Private Sub txtCOM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdTest.SetFocus
End Sub

Private Sub TxtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtSN.SetFocus
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtPort.SetFocus
        
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCom.SetFocus
End Sub
