VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ղ�������"
   ClientHeight    =   3270
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5745
   Icon            =   "frmSet����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4515
      TabIndex        =   10
      Top             =   990
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4515
      TabIndex        =   11
      Top             =   1470
      Width           =   1100
   End
   Begin VB.Frame fraTop 
      Caption         =   "���в���"
      Height          =   2400
      Left            =   810
      TabIndex        =   1
      Top             =   690
      Width           =   3525
      Begin VB.CheckBox chk 
         Caption         =   "��������ҽ����Ŀ(&G)"
         Height          =   195
         Index           =   1
         Left            =   570
         TabIndex        =   9
         Top             =   1950
         Width           =   2265
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   3
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1485
         Width           =   645
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ҫ������֤(&V)"
         Height          =   195
         Index           =   0
         Left            =   570
         TabIndex        =   6
         Top             =   1170
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   5
         Top             =   735
         Width           =   645
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ҽ�����ų���(&R)"
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   2
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���볤��(&P)"
         Height          =   180
         Index           =   1
         Left            =   930
         TabIndex        =   7
         Top             =   1545
         Width           =   990
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����֤�ų���(&T)"
         Height          =   180
         Index           =   2
         Left            =   570
         TabIndex        =   4
         Top             =   795
         Width           =   1350
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmSet����.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ��ָ��������𣬳��򽫰�����Ҫ�����С�"
      Height          =   180
      Left            =   795
      TabIndex        =   0
      Top             =   240
      Width           =   3780
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum�༭
    Text���ų��� = 0
    Text���볤�� = 1
    Text����֤�� = 2
End Enum

Private Enum enumѡ��
    Check���� = 0
    Check��������ҽ����Ŀ = 1
End Enum

Dim mlng���� As Long, mlng���� As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim colPara As New Collection
    Dim lngCount As Long
    
    If Val(TxtEdit(Text���ų���).Text) > 26 Then
        MsgBox "���ų��Ȳ��ܳ���26λ��", vbInformation, gstrSysName
        TxtEdit(Text���ų���).SetFocus
        Exit Sub
    End If
    
    If Val(TxtEdit(Text���볤��).Text) > 8 Then
        MsgBox "���볤�Ȳ��ܳ���8λ��", vbInformation, gstrSysName
        TxtEdit(Text���볤��).SetFocus
        Exit Sub
    End If
    If Val(TxtEdit(Text����֤��).Text) > 26 Then
        MsgBox "����֤�ų��Ȳ��ܳ���26λ��", vbInformation, gstrSysName
        TxtEdit(Text����֤��).SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & "," & mlng���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    colPara.Add mlng���� & ",'���ų���','" & Int(Val(TxtEdit(Text���ų���).Text))
    colPara.Add mlng���� & ",'����֤����','" & Int(Val(TxtEdit(Text����֤��).Text))
    colPara.Add mlng���� & ",'������֤','" & chk(Check����).Value
    colPara.Add mlng���� & ",'���볤��','" & IIf(chk(Check����).Value = 1, Int(Val(TxtEdit(Text���볤��).Text)), 0)
    '��һ���ֲ�������������
    colPara.Add "null,'��������ҽ����Ŀ','" & chk(Check��������ҽ����Ŀ).Value
    
    For lngCount = 1 To colPara.Count
        gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & colPara(lngCount) & "'," & lngCount & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub chk_Click(Index As Integer)
    mblnChange = True
    If Index = Check���� Then
        TxtEdit(Text���볤��).Enabled = chk(Check����).Value
        lblEdit(Text���볤��).Enabled = chk(Check����).Value
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
    zlCommFun.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Public Function ��������(ByVal lng���� As Long, ByVal lng���� As Long) As Boolean
'���ܣ�������������ҽ������Ҫ�Ĳ���
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    
    mblnOK = False
    mlng���� = lng����
    mlng���� = lng����
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1] and (���� is null or ����=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����, lng����)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "���ų���"
                TxtEdit(Text���ų���).Text = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "����֤����"
                TxtEdit(Text����֤��).Text = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "������֤"
                chk(Check����).Value = IIf(rsTemp("����ֵ") = 1, 1, 0)
            Case "��������ҽ����Ŀ"
                chk(Check��������ҽ����Ŀ).Value = IIf(rsTemp("����ֵ") = 1, 1, 0)
            Case "���볤��"
                TxtEdit(Text���볤��).Text = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet����.Show vbModal, frmҽ�����
    �������� = mblnOK
End Function

