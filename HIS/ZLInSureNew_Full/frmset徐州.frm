VERSION 5.00
Begin VB.Form frmset���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmset����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraIC 
      Caption         =   "IC������"
      Height          =   810
      Left            =   120
      TabIndex        =   8
      Top             =   1815
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   10
         Text            =   "1"
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����(&D)"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   375
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�Ŵ���"
         Height          =   180
         Index           =   4
         Left            =   1740
         TabIndex        =   11
         Top             =   375
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4950
      TabIndex        =   12
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4950
      TabIndex        =   13
      Top             =   780
      Width           =   1100
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1590
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   14
         Top             =   1515
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1110
         Left            =   3450
         TabIndex        =   7
         Top             =   330
         Width           =   1110
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1215
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   2145
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���ݿ�(&D)"
         Height          =   180
         Index           =   5
         Left            =   300
         TabIndex        =   15
         Top             =   1575
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   5
         Top             =   1170
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmset����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mblnChangePassword As Boolean  '���뱻�޸Ĺ�
Private mlngIcdev As Long
Private st%
 
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Sub cmdTest_Click()
    If gcn����.State = adStateOpen Then gcn����.Close
    If Not IsNumeric(TxtEdit(4).Text) Then
        MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error Resume Next
'    gcn����.Open "Provider=SQLOLEDB.1;Initial Catalog=" & Trim(txtEdit(3).Tag) & ";Password=" & Trim(txtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(txtEdit(0).Tag) & ";Data Source=" & Trim(txtEdit(2).Tag)
    gcn����.Open "Provider=MSDASQL.1;Password=" & Trim(TxtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(TxtEdit(0).Text) & ";Data Source=" & Trim(TxtEdit(2).Text)
    
    If Err <> 0 Then
        MsgBox "ҽ��ǰ�÷���������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
'    If mblnChangePassword = True Then
'        '��������ɹ�
'        txtEdit(4).Enabled = True
'    End If

    MsgBox "ҽ��ǰ�÷��������ӳɹ�", vbInformation, gstrSysName
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    
    '���ж��ַ��ĺϷ���
    For lngCount = TxtEdit.LBound To TxtEdit.UBound
        If zlCommFun.StrIsValid(TxtEdit(lngCount).Text, TxtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll TxtEdit(lngCount)
            TxtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If Not IsNumeric(TxtEdit(4).Text) Then
        MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    '�����ӽ��в���
    If gcn����.State = adStateClosed Then
        On Error Resume Next
'        gcn����.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
            Trim(txtEdit(2).Text), Trim(txtEdit(0).Text), Trim(txtEdit(1).Tag)
'        gcn����.Open "Provider=SQLOLEDB.1;Initial Catalog=" & Trim(txtEdit(3).Tag) & ";Password=" & Trim(txtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(txtEdit(0).Tag) & ";Data Source=" & Trim(txtEdit(2).Tag)
        gcn����.Open "Provider=MSDASQL.1;Password=" & Trim(TxtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(TxtEdit(0).Text) & ";Data Source=" & Trim(TxtEdit(2).Text)
        
        If Err <> 0 Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    On Error Resume Next
    mlngIcdev = init_com(TxtEdit(4).Text - 1) 'Init COM2
    If mlngIcdev <> 0 Then
        If MsgBox("���ڳ�ʼ��ʧ�ܣ����鴮�ڡ��Ƿ�������棿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            TxtEdit(4).SetFocus
            Exit Function
        End If
    End If
    st = close_com()
    IsValid = True
End Function

Public Function ��������() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    
    mblnOK = False
    On Error GoTo errHandle
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����)
    
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "�����û���"
                TxtEdit(0).Text = str����ֵ
            Case "���ݷ�����"
                TxtEdit(2).Text = str����ֵ
            Case "�����û�����"
                TxtEdit(1).Text = "        "    '������
                TxtEdit(1).Tag = str����ֵ
            Case "�������ݿ�"
                TxtEdit(3).Text = str����ֵ
        End Select
        rsTemp.MoveNext
    Loop
    On Error Resume Next
    TxtEdit(4).Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���") + 1
    
    mblnChange = False
    mblnChangePassword = False
    frmset����.Show vbModal, frmҽ�����
    �������� = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'�����û���','" & TxtEdit(0).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'�����û�����','" & TxtEdit(1).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'���ݷ�����','" & TxtEdit(2).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'�������ݿ�','" & TxtEdit(3).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    '����ǰʹ�õĴ���д��ע���֮��
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", CStr(TxtEdit(4).Text - 1)
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 1 Then
        TxtEdit(1).Tag = TxtEdit(1).Text
        mblnChangePassword = True
    End If
    
    '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
    If gcn����.State = adStateOpen Then gcn����.Close
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
    If Index = 4 Then
        If CheckIsInclude(UCase(Chr(KeyAscii)), "������") = True Then KeyAscii = 0
    End If
End Sub

Private Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "����ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "�ɴ�ӡ�ַ�"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Index = 4 Then
        If Not IsNumeric(TxtEdit(4).Text) Then
            MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        End If
    End If
End Sub
