VERSION 5.00
Begin VB.Form frmset��ɽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmset��ɽ.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1440
      TabIndex        =   19
      Top             =   3510
      Width           =   3345
   End
   Begin VB.ComboBox cbo���õ��� 
      Height          =   300
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3090
      Width           =   3345
   End
   Begin VB.Frame fraIC 
      Caption         =   "IC������"
      Height          =   1245
      Left            =   120
      TabIndex        =   8
      Top             =   1740
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1320
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   780
         Width           =   990
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   10
         Text            =   "1"
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����֤��(&V)"
         Height          =   180
         Index           =   5
         Left            =   210
         TabIndex        =   12
         Top             =   840
         Width           =   990
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
      TabIndex        =   16
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4950
      TabIndex        =   17
      Top             =   780
      Width           =   1100
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1545
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4695
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����ļ�λ��"
      Height          =   180
      Left            =   315
      TabIndex        =   18
      Top             =   3585
      Width           =   1080
   End
   Begin VB.Label lbl���õ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���õ���(&Q)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   315
      TabIndex        =   14
      Top             =   3150
      Width           =   990
   End
End
Attribute VB_Name = "frmset��ɽ"
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
Private Declare Function IC_InitComm Lib "DCIC32.DLL" (ByVal Port%) As Long
Private Declare Function IC_ExitComm% Lib "DCIC32.DLL" (ByVal icdev As Long)
 
Private Sub cbo���õ���_Change()
    If cbo���õ���.ListIndex = 1 Then
        Text1.Enabled = True
    Else
        Text1.Enabled = False
    End If
End Sub

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
    If gcn��ɽ.State = adStateOpen Then gcn��ɽ.Close
    If Not IsNumeric(TxtEdit(3).Text) Then
        MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error Resume Next
    If cbo���õ���.ListIndex = 1 Then
        gcn��ɽ.Open "Provider=SQLOLEDB.1;Initial Catalog=hw_interface;Password=" & TxtEdit(1).Text & ";Persist Security Info=True;User ID=" & TxtEdit(0).Text & ";Data Source=" & TxtEdit(2).Text
    Else
        gcn��ɽ.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
                Trim(TxtEdit(2).Text), Trim(TxtEdit(0).Text), Trim(TxtEdit(1).Tag)
    End If
    
    If Err <> 0 Then
        MsgBox "ҽ��ǰ�÷���������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If mblnChangePassword = True Then
        '��������ɹ�
        TxtEdit(4).Enabled = True
    End If

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
    If TxtEdit(4).Enabled = True And TxtEdit(4).Text = "" Then
        MsgBox "����ѯ��IC����Ӧ�̺���д����֤�롣", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not IsNumeric(TxtEdit(3).Text) Then
        MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    '�����ӽ��в���
    If gcn��ɽ.State = adStateClosed Then
        On Error Resume Next
        gcn��ɽ.Open "Provider=SQLOLEDB.1;Initial Catalog=hw_interface;Password=" & TxtEdit(1).Text & ";Persist Security Info=True;User ID=" & TxtEdit(0).Text & ";Data Source=" & TxtEdit(2).Text
        
'          gcn��ɽ.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
              Trim(txtEdit(2).Text), Trim(txtEdit(0).Text), Trim(txtEdit(1).Tag)
        
        If Err <> 0 Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    On Error Resume Next
    mlngIcdev = IC_InitComm(TxtEdit(3).Text - 1) 'Init COM2
    If mlngIcdev <= 0 Then
        If MsgBox("���ڳ�ʼ��ʧ�ܣ����鴮�ڡ��Ƿ�������棿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            TxtEdit(3).SetFocus
            Exit Function
        End If
    End If
    st = IC_ExitComm(mlngIcdev)  'Close COM
    IsValid = True
End Function

Public Function ��������() As Boolean
'���ܣ���������ϣ��˾��ҽ���ӿ�
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    Dim int���õ��� As Integer
    
    mblnOK = False
    On Error GoTo errHandle
    
    
    gstrSQL = "select ������,����ֵ from ���ղ��� " & _
              " where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�����ɽ)
    
    int���õ��� = 0
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "��ɽ�û���"
                TxtEdit(0).Text = str����ֵ
            Case "��ɽ������"
                TxtEdit(2).Text = str����ֵ
            Case "��ɽ�û�����"
                TxtEdit(1).Text = "        "    '������
                TxtEdit(1).Tag = str����ֵ
            Case "����֤��"
                TxtEdit(4).Text = str����ֵ
            Case "���õ���"
                int���õ��� = Val(str����ֵ)
            Case "�����ļ�λ��"
                Text1.Text = str����ֵ
        End Select
        rsTemp.MoveNext
    Loop
    If TxtEdit(4).Text = "" Then TxtEdit(4).Enabled = True
    On Error Resume Next
'    If GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���") = "" Then
    TxtEdit(3).Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���") + 1
    
    'Modified By ���� ���� 06:07:34
    With cbo���õ���
        .Clear
        .AddItem "��ɽ"
        .AddItem "ǭ��"
        .AddItem "����"
        .AddItem "��ɽ"
        .ListIndex = int���õ���
    End With
    
    mblnChange = False
    mblnChangePassword = False
    frmset��ɽ.Show vbModal, frmҽ�����
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
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_�����ɽ & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����ɽ & ",null,'��ɽ�û���','" & TxtEdit(0).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����ɽ & ",null,'��ɽ�û�����','" & TxtEdit(1).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����ɽ & ",null,'��ɽ������','" & TxtEdit(2).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����ɽ & ",null,'����֤��','" & TxtEdit(4).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    'Modified By ���� ���� 06:07:51
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����ɽ & ",null,'���õ���','" & cbo���õ���.ListIndex & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����ɽ & ",null,'�����ļ�λ��','" & Text1.Text & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    gcnOracle.CommitTrans
    '����ǰʹ�õĴ���д��ע���֮��
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", CStr(TxtEdit(3).Text - 1)
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
    If gcn��ɽ.State = adStateOpen Then gcn��ɽ.Close
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
    If Index = 3 Then
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
    If Index = 3 Then
        If Not IsNumeric(TxtEdit(3).Text) Then
            MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        End If
    End If
End Sub
