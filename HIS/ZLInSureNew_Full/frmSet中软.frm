VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   Icon            =   "frmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽ��������"
      Height          =   1545
      Left            =   150
      TabIndex        =   12
      Top             =   3000
      Width           =   4155
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   19
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   18
         Top             =   1110
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   8
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   14
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   10
         Left            =   390
         TabIndex        =   17
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   9
         Left            =   570
         TabIndex        =   15
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   8
         Left            =   390
         TabIndex        =   13
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.Frame fra���Ĳ��� 
      Caption         =   "���Ĳ���"
      Height          =   4365
      Left            =   4440
      TabIndex        =   20
      Top             =   180
      Width           =   4605
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   24
         Top             =   718
         Width           =   1635
      End
      Begin VB.CheckBox chk 
         Caption         =   "�Ƿ񶨵����(&9)"
         Height          =   255
         Index           =   0
         Left            =   1590
         TabIndex        =   39
         Top             =   3434
         Width           =   1785
      End
      Begin VB.CheckBox chk 
         Caption         =   "�Ƿ�������(&0)"
         Height          =   225
         Index           =   1
         Left            =   1590
         TabIndex        =   40
         Top             =   3780
         Width           =   1665
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   22
         Top             =   330
         Width           =   2835
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1590
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   26
         Top             =   1106
         Width           =   1635
      End
      Begin VB.CommandButton cmdĿ¼ 
         Caption         =   "��"
         Height          =   240
         Index           =   1
         Left            =   4140
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3090
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1590
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   1494
         Width           =   1635
      End
      Begin VB.CommandButton cmdĿ¼ 
         Caption         =   "��"
         Height          =   240
         Index           =   0
         Left            =   4140
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2685
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   37
         Top             =   3046
         Width           =   2835
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   34
         Top             =   2658
         Width           =   2835
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   32
         Top             =   2270
         Width           =   2835
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   30
         Top             =   1882
         Width           =   2835
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "FTP��½�û�(&2)"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   23
         Top             =   778
         Width           =   1260
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "Զ������(&1)"
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   21
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "FTP�û�����(&3)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   25
         Top             =   1166
         Width           =   1260
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "Զ���ϴ�Ŀ¼(&5)"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   29
         Top             =   1942
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "Զ������Ŀ¼(&6)"
         Height          =   180
         Index           =   5
         Left            =   210
         TabIndex        =   31
         Top             =   2330
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����ϴ�Ŀ¼(&7)"
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   33
         Top             =   2718
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������Ŀ¼(&8)"
         Height          =   180
         Index           =   7
         Left            =   210
         TabIndex        =   36
         Top             =   3106
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "FTP����ȷ��(&4)"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   27
         Top             =   1554
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7440
      TabIndex        =   42
      Top             =   4710
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6120
      TabIndex        =   41
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Frame fraҽԺ 
      Caption         =   "ҽԺ��Ϣ"
      Height          =   2745
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4155
      Begin VB.CheckBox chk 
         Caption         =   "���޲���(&V)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   630
         TabIndex        =   11
         Top             =   2400
         Width           =   1365
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����Ը�����(&I)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   2370
         TabIndex        =   10
         Top             =   2070
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Caption         =   "ȫ�ԷѲ���(&L)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   630
         TabIndex        =   9
         Top             =   2070
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����Ը�����(&F)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   2370
         TabIndex        =   7
         Top             =   1380
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Caption         =   "ȫ�ԷѲ���(&A)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   630
         TabIndex        =   6
         Top             =   1410
         Width           =   1485
      End
      Begin VB.ComboBox cmbװǮ 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   690
         Width           =   1785
      End
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   1785
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ�����ʻ���ʹ�÷�Χ(&B)"
         Height          =   180
         Index           =   3
         Left            =   330
         TabIndex        =   8
         Top             =   1770
         Width           =   2430
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�շ�ʱ�����ʻ���ʹ�÷�Χ(&R)"
         Height          =   180
         Index           =   2
         Left            =   330
         TabIndex        =   5
         Top             =   1110
         Width           =   2430
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "װǮģʽ(&N)"
         Height          =   180
         Index           =   1
         Left            =   330
         TabIndex        =   3
         Top             =   765
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҽԺ����(&G)"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   390
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum�ı�
    TextԶ������ = 0
    Text��½�û� = 1
    Text�û����� = 2
    Textȷ������ = 3
    TextԶ���ϴ� = 4
    TextԶ������ = 5
    Text�����ϴ� = 6
    Text�������� = 7
    textҽ���û� = 8
    Textҽ������ = 9
    Textҽ�������� = 10
End Enum

Private Enum enumѡ��
    Check������� = 0
    Check�������� = 1
    Check�շ�ȫ�Է� = 2
    Check�շ������Ը� = 3
    Check����ȫ�Է� = 4
    Check���������Ը� = 5
    Check���㳬�� = 6
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Private mlng���� As Long, mlng���� As Long

Private Sub cmb����_Click()
    mblnChange = True
End Sub

Private Sub cmbװǮ_Change()
    mblnChange = True
End Sub

Private Sub cmdTest_Click()
    If OraDataOpen(gcn����, TxtEdit(Textҽ��������).Text, TxtEdit(textҽ���û�).Text, TxtEdit(Textҽ������).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
    
    If TxtEdit(Textҽ������).Tag = TxtEdit(Textҽ������).Text Then
        cmb����.Enabled = True
        cmbװǮ.Enabled = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    cmb����.Enabled = False
    cmbװǮ.Enabled = False
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

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    
    If cmb����.Text = "" Then
        MsgBox "������ҽԺ����", vbInformation, gstrSysName
        If cmb����.Enabled = True Then cmb����.SetFocus
        Exit Function
    End If
    If cmbװǮ.Text = "" Then
        MsgBox "������װǮģʽ��", vbInformation, gstrSysName
        If cmbװǮ.Enabled = True Then cmbװǮ.SetFocus
        Exit Function
    End If
    
    For lngCount = TxtEdit.LBound To TxtEdit.UBound
        If zlCommFun.StrIsValid(TxtEdit(lngCount).Text, TxtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll TxtEdit(lngCount)
            TxtEdit(lngCount).SetFocus
            Exit Function
        End If
        
        If lngCount < TextԶ���ϴ� Then
            If Len(TxtEdit(lngCount).Text) = 0 Then
                strTitle = Mid(lblEdit(lngCount).Caption, 1, InStr(lblEdit(lngCount).Caption, "(") - 1)
                If MsgBox("��" & strTitle & "�����Ϊ�տ���ʹ�ϴ������޷������������Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    zlControl.TxtSelAll TxtEdit(lngCount)
                    TxtEdit(lngCount).SetFocus
                    Exit Function
                End If
            End If
        End If
    Next
    
    '������ȷ��
    If TxtEdit(Text�û�����).Text <> TxtEdit(Textȷ������).Text Then
        MsgBox "��ȷ���������������һ�¡�", vbInformation, gstrSysName
        zlControl.TxtSelAll TxtEdit(Text�û�����)
        TxtEdit(Text�û�����).SetFocus
        Exit Function
    End If
    
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & "," & mlng���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'ҽԺ����','" & cmb����.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'װǮģʽ','" & cmbװǮ.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'�շѸ����ʻ�ʹ�÷�Χ','" & _
                chk(Check�շ�ȫ�Է�).Value & chk(Check�շ������Ը�).Value & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'��������ʻ�ʹ�÷�Χ','" & _
                chk(Check����ȫ�Է�).Value & chk(Check���������Ը�).Value & chk(Check���㳬��).Value & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'ҽ���û���','" & TxtEdit(textҽ���û�).Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'ҽ���û�����','" & _
            IIf(TxtEdit(Textҽ������).Tag = "", "", EncryptStr(TxtEdit(Textҽ������).Tag, 256, True)) & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'ҽ��������','" & TxtEdit(Textҽ��������).Text & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & mlng���� & ",'Զ������','" & TxtEdit(TextԶ������).Text & "',8)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & mlng���� & ",'FTP��½�û�','" & TxtEdit(Text��½�û�).Text & "',9)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & mlng���� & ",'FTP�û�����','" & _
            IIf(TxtEdit(Text�û�����).Tag = "", "", EncryptStr(TxtEdit(Text�û�����).Tag, 256, True)) & "',10)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & mlng���� & ",'Զ���ϴ�Ŀ¼','" & TxtEdit(TextԶ���ϴ�).Text & "',11)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & mlng���� & ",'Զ������Ŀ¼','" & TxtEdit(TextԶ������).Text & "',12)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & mlng���� & ",'�����ϴ�Ŀ¼','" & TxtEdit(Text�����ϴ�).Text & "',13)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & mlng���� & ",'��������Ŀ¼','" & TxtEdit(Text��������).Text & "',14)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & mlng���� & ",'����ҽ�ƻ���','" & chk(Check�������).Value & "',15)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & mlng���� & ",'��������','" & chk(Check��������).Value & "',16)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If gcn����.State = adStateClosed Then
        If OraDataOpen(gcn����, TxtEdit(Textҽ��������).Text, TxtEdit(textҽ���û�).Text, TxtEdit(Textҽ������).Tag, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Private Sub chk_Click(Index As Integer)
    mblnChange = True
    
    If Index = Check�շ�ȫ�Է� Or Index = Check����ȫ�Է� Then
        If chk(Index).Value = 1 Then
            chk(Index + 1).Value = 1
            chk(Index + 1).Enabled = False
        Else
            chk(Index + 1).Enabled = True
        End If
    End If
End Sub

Private Sub cmdĿ¼_Click(Index As Integer)
    Dim strTitle As String
    Dim strPath As String
    
    If Index = 0 Then
        strTitle = "��ѡ�񱣴��ϴ��ļ���Ŀ¼��"
    Else
        strTitle = "��ѡ�񱣴������ļ���Ŀ¼��"
    End If
    
    strPath = zlCommFun.OpenDir(Me.hwnd, strTitle)
    If strPath <> "" Then
        '����Ŀ¼��
        TxtEdit(Index + 6).Text = strPath
        TxtEdit(Index + 6).SetFocus
    End If
    
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text�û����� Or Index = Textҽ������ Then
        TxtEdit(Index).Tag = TxtEdit(Index).Text
    End If
    
    If Index = Textҽ�������� Or Index = Textҽ������ Or Index = textҽ���û� Then
        '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
        If gcn����.State = adStateOpen Then gcn����.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
End Sub

Public Function ��������(ByVal lng���� As Long, ByVal lng���� As Long) As Boolean
'���ܣ������붫�󰢶��ɵ�ҽ���ӿ�
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    
    mblnOK = False
    mlng���� = lng����
    mlng���� = lng����
    
    On Error GoTo errHandle
    
    cmbװǮ.AddItem "0.����װǮ"
    cmbװǮ.AddItem "1.����װǮ"
    cmbװǮ.AddItem "2.����װǮ"
    
    cmb����.AddItem "33.���ȼ׼�"
    cmb����.AddItem "32.�����Ҽ�"
    cmb����.AddItem "23.���ȼ׼�"
    cmb����.AddItem "22.�����Ҽ�"
    cmb����.AddItem "13.һ�ȼ׼�"
    cmb����.AddItem "12.һ���Ҽ�"
    cmb����.AddItem "0.�޼� "
    
    gstrSQL = "select ������,����ֵ from ���ղ��� " & _
              " where ����=[1] and (���� is null or ����=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����, lng����)
    
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽԺ����"
                SetComboByText cmb����, str����ֵ, False, " "
            Case "װǮģʽ"
                SetComboByText cmbװǮ, str����ֵ, False, " "
            Case "ҽ���û���"
                TxtEdit(textҽ���û�) = str����ֵ
            Case "ҽ��������"
                TxtEdit(Textҽ��������) = str����ֵ
            Case "ҽ���û�����"
                TxtEdit(Textҽ������).Text = "        "    '������
                If str����ֵ = "" Then
                    TxtEdit(Textҽ������).Tag = ""
                Else
                    TxtEdit(Textҽ������).Tag = EncryptStr(str����ֵ, 256, False)
                End If
            Case "Զ������"
                TxtEdit(TextԶ������) = str����ֵ
            Case "FTP��½�û�"
                TxtEdit(Text��½�û�) = str����ֵ
            Case "FTP�û�����"
                TxtEdit(Text�û�����).Text = "        "    '������
                TxtEdit(Textȷ������).Text = "        "    '������
                If str����ֵ = "" Then
                    TxtEdit(Text�û�����).Tag = ""
                Else
                    TxtEdit(Text�û�����).Tag = EncryptStr(str����ֵ, 256, False)
                End If
            Case "Զ���ϴ�Ŀ¼"
                TxtEdit(TextԶ���ϴ�) = str����ֵ
            Case "Զ������Ŀ¼"
                TxtEdit(TextԶ������) = str����ֵ
            Case "�����ϴ�Ŀ¼"
                TxtEdit(Text�����ϴ�) = str����ֵ
            Case "��������Ŀ¼"
                TxtEdit(Text��������) = str����ֵ
            Case "����ҽ�ƻ���"
                chk(Check�������).Value = IIf(str����ֵ = "1", 1, 0)
            Case "��������"
                chk(Check��������).Value = IIf(str����ֵ = "1", 1, 0)
'            Case "�շѸ����ʻ�ʹ�÷�Χ"
'                chk(Check�շ�ȫ�Է�).Value = IIf(Left(str����ֵ, 1) = "1", 1, 0)
'                chk(Check�շ������Ը�).Value = IIf(Mid(str����ֵ, 2, 1) = "1", 1, 0)
'                'ȫ�Է�����
'                If chk(Check�շ�ȫ�Է�).Value = 1 Then
'                    chk(Check�շ������Ը�).Value = 1
'                    chk(Check�շ������Ը�).Enabled = False
'                End If
'            Case "��������ʻ�ʹ�÷�Χ"
'                chk(Check����ȫ�Է�).Value = IIf(Left(str����ֵ, 1) = "1", 1, 0)
'                chk(Check���������Ը�).Value = IIf(Mid(str����ֵ, 2, 1) = "1", 1, 0)
'                chk(Check���㳬��).Value = IIf(Mid(str����ֵ, 3, 1) = "1", 1, 0)
'                'ȫ�Է�����
'                If chk(Check����ȫ�Է�).Value = 1 Then
'                    chk(Check���������Ը�).Value = 1
'                    chk(Check���������Ը�).Enabled = False
'                End If
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet����.Show vbModal, frmҽ�����
    
    �������� = mblnOK
    If mblnOK = False Then
        '����ʧ�ܣ��ر����ӡ��������û�����������û�������������
        If gcn����.State = adStateOpen Then gcn����.Close
    End If
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
