VERSION 5.00
Begin VB.Form frmSet�㽭 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ������"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1545
      Left            =   80
      TabIndex        =   9
      Top             =   105
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1110
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1215
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   0
         Top             =   330
         Width           =   2145
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1110
         Left            =   3450
         TabIndex        =   3
         Top             =   330
         Width           =   1110
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4910
      TabIndex        =   7
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4910
      TabIndex        =   6
      Top             =   300
      Width           =   1100
   End
   Begin VB.Frame fraIC 
      Caption         =   "IC������"
      Height          =   735
      Left            =   80
      TabIndex        =   8
      Top             =   1740
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1290
         MaxLength       =   40
         TabIndex        =   4
         Text            =   "1"
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�Ŵ���"
         Height          =   180
         Index           =   4
         Left            =   1695
         TabIndex        =   15
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����(&D)"
         Height          =   180
         Index           =   3
         Left            =   195
         TabIndex        =   14
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.ComboBox cbo���õ��� 
      Height          =   300
      Left            =   1400
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2625
      Width           =   3345
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
      Left            =   275
      TabIndex        =   13
      Top             =   2685
      Width           =   990
   End
End
Attribute VB_Name = "frmSet�㽭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mblnChange As Boolean
Private mblnChangePassword As Boolean  '���뱻�޸Ĺ�
 
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
    If gcn�㽭.State = adStateOpen Then gcn�㽭.Close
    If Not IsNumeric(TxtEdit(3).Text) Then
        MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error Resume Next
    If cbo���õ���.ListIndex = 0 Then
        gcn�㽭.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
                Trim(TxtEdit(2).Text), Trim(TxtEdit(0).Text), Trim(TxtEdit(1).Tag)
    End If
    
    If Err <> 0 Then
        MsgBox "ҽ��ǰ�÷���������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
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
    
    If Not IsNumeric(TxtEdit(3).Text) Then
        MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    '�����ӽ��в���
    If gcn�㽭.State = adStateClosed Then
        On Error Resume Next
        If cbo���õ���.ListIndex = 0 Then
            gcn�㽭.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
                Trim(TxtEdit(2).Text), Trim(TxtEdit(0).Text), Trim(TxtEdit(1).Tag)
        End If
        If Err <> 0 Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
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
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�㽭)
    
    int���õ��� = 0
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "�㽭�û���"
                TxtEdit(0).Text = str����ֵ
            Case "�㽭������"
                TxtEdit(2).Text = str����ֵ
            Case "�㽭�û�����"
                TxtEdit(1).Text = "        "    '������
                TxtEdit(1).Tag = str����ֵ
            Case "���õ���"
                int���õ��� = Val(str����ֵ)
        End Select
        rsTemp.MoveNext
    Loop
    On Error Resume Next
    TxtEdit(3).Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���") + 1
    
    With cbo���õ���
        .Clear
        .AddItem "��Ϫҽ��"
        .ListIndex = int���õ���
    End With
    
    mblnChange = False
    mblnChangePassword = False
    frmSet�㽭.Show vbModal, frmҽ�����
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
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_�㽭 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�㽭 & ",null,'�㽭�û���','" & TxtEdit(0).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�㽭 & ",null,'�㽭�û�����','" & TxtEdit(1).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�㽭 & ",null,'�㽭������','" & TxtEdit(2).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    'Modified By ���� ���� 06:07:51
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�㽭 & ",null,'���õ���','" & cbo���õ���.ListIndex & "',5)"
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
    If gcn�㽭.State = adStateOpen Then gcn�㽭.Close
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

