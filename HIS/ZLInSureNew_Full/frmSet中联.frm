VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ղ�������"
   ClientHeight    =   3870
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   Icon            =   "frmSet����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3285
      TabIndex        =   19
      Top             =   3390
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4575
      TabIndex        =   20
      Top             =   3390
      Width           =   1100
   End
   Begin VB.Frame fraButtom 
      Caption         =   "�����ʻ�֧����Χ"
      Height          =   2550
      Left            =   2940
      TabIndex        =   11
      Top             =   690
      Width           =   2925
      Begin VB.CheckBox chk 
         Caption         =   "ȫ�ԷѲ���(&A)"
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   13
         Top             =   590
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����Ը�����(&F)"
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   14
         Top             =   925
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Caption         =   "ȫ�ԷѲ���(&L)"
         Height          =   255
         Index           =   3
         Left            =   660
         TabIndex        =   16
         Top             =   1520
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����Ը�����(&I)"
         Height          =   255
         Index           =   4
         Left            =   660
         TabIndex        =   17
         Top             =   1855
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Caption         =   "���޲���(&V)"
         Height          =   255
         Index           =   5
         Left            =   660
         TabIndex        =   18
         Top             =   2190
         Width           =   1365
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�շ�ʱ��ʹ�÷�Χ"
         Height          =   180
         Index           =   2
         Left            =   270
         TabIndex        =   12
         Top             =   330
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��ʹ�÷�Χ"
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   15
         Top             =   1260
         Width           =   1440
      End
   End
   Begin VB.Frame fraTop 
      Caption         =   "���в���"
      Height          =   2550
      Left            =   180
      TabIndex        =   1
      Top             =   690
      Width           =   2565
      Begin VB.CheckBox chk97 
         Caption         =   "���������޶����"
         Height          =   225
         Left            =   345
         TabIndex        =   23
         Top             =   2130
         Width           =   1965
      End
      Begin VB.TextBox txt97 
         Height          =   270
         Left            =   1110
         TabIndex        =   21
         Top             =   1800
         Width           =   780
      End
      Begin VB.CheckBox chk 
         Caption         =   "�ȿ�����"
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   10
         Top             =   1500
         Width           =   2055
      End
      Begin VB.CheckBox chk 
         Caption         =   "�շ�ʹ��ҽ������(&G)"
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   9
         Top             =   1170
         Width           =   2025
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   3
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1485
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ҫ������֤(&V)"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   1170
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   5
         Top             =   735
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Ԫ"
         Height          =   180
         Left            =   2025
         TabIndex        =   24
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "�����޶�"
         Height          =   285
         Left            =   315
         TabIndex        =   22
         Top             =   1830
         Width           =   750
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ҽ�����ų���(&R)"
         Height          =   180
         Index           =   0
         Left            =   300
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
         Left            =   660
         TabIndex        =   7
         Top             =   1545
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����֤�ų���(&T)"
         Height          =   180
         Index           =   2
         Left            =   300
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
    Check�շ�ȫ�Է� = 1
    Check�շ������Ը� = 2
    Check����ȫ�Է� = 3
    Check���������Ը� = 4
    Check���㳬�� = 5
    Check�շ�ҽ������ = 6
    Check�ȿ����� = 7
End Enum

Dim mlng���� As Long, mlng���� As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub Form_Load()
    '20051220 �¶� ����ҽԺ����
    If mlng���� = TYPE_������Ժ Then
        txt97.Visible = True
        chk97.Visible = True
        Label1.Visible = True
        Label2.Visible = True
    Else
        txt97.Visible = False
        chk97.Visible = False
        Label1.Visible = False
        Label2.Visible = False
    End If
End Sub

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
    
    If Val(txtEdit(Text���ų���).Text) > 26 Then
        MsgBox "���ų��Ȳ��ܳ���26λ��", vbInformation, gstrSysName
        txtEdit(Text���ų���).SetFocus
        Exit Sub
    End If
    
    If Val(txtEdit(Text���볤��).Text) > 8 Then
        MsgBox "���볤�Ȳ��ܳ���8λ��", vbInformation, gstrSysName
        txtEdit(Text���볤��).SetFocus
        Exit Sub
    End If
    If Val(txtEdit(Text����֤��).Text) > 26 Then
        MsgBox "����֤�ų��Ȳ��ܳ���26λ��", vbInformation, gstrSysName
        txtEdit(Text����֤��).SetFocus
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
    colPara.Add mlng���� & ",'���ų���','" & Int(Val(txtEdit(Text���ų���).Text))
    colPara.Add mlng���� & ",'����֤����','" & Int(Val(txtEdit(Text����֤��).Text))
    colPara.Add mlng���� & ",'������֤','" & chk(Check����).Value
    colPara.Add mlng���� & ",'���볤��','" & IIf(chk(Check����).Value = 1, Int(Val(txtEdit(Text���볤��).Text)), 0)
    '��һ���ֲ�������������
    colPara.Add "null,'�շ�ʹ��ҽ������','" & chk(Check�շ�ҽ������).Value
    colPara.Add "null,'�շѸ����ʻ�ʹ�÷�Χ','" & _
                chk(Check�շ�ȫ�Է�).Value & chk(Check�շ������Ը�).Value
    colPara.Add "null,'��������ʻ�ʹ�÷�Χ','" & _
                chk(Check����ȫ�Է�).Value & chk(Check���������Ը�).Value & chk(Check���㳬��).Value
    colPara.Add "null,'�ȿ�����','" & chk(Check�ȿ�����).Value
    
    For lngCount = 1 To colPara.Count
        gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & colPara(lngCount) & "'," & lngCount & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    '20051220 �¶� ����ҽԺ����
    If mlng���� = TYPE_������Ժ Then
        If Val(txt97) > 0 Then
            gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",0,'�����޶�','" & Val(txt97) & "'," & lngCount + 1 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
        
        If chk97.Value = 1 Then
            gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",0,'�����ֹ','" & 1 & "'," & lngCount + 2 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End If
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
        txtEdit(Text���볤��).Enabled = chk(Check����).Value
        lblEdit(Text���볤��).Enabled = chk(Check����).Value
    End If
    If Index = Check�շ�ȫ�Է� Or Index = Check����ȫ�Է� Then
        If chk(Index).Value = 1 Then
            chk(Index + 1).Value = 1
            chk(Index + 1).Enabled = False
        Else
            chk(Index + 1).Enabled = True
        End If
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
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
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1] and (���� is null or ����=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����, lng����)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "���ų���"
                txtEdit(Text���ų���).Text = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "����֤����"
                txtEdit(Text����֤��).Text = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "������֤"
                chk(Check����).Value = IIf(rsTemp("����ֵ") = 1, 1, 0)
            Case "�շ�ʹ��ҽ������"
                chk(Check�շ�ҽ������).Value = IIf(rsTemp("����ֵ") = 1, 1, 0)
            Case "���볤��"
                txtEdit(Text���볤��).Text = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "�շѸ����ʻ�ʹ�÷�Χ"
                str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                chk(Check�շ�ȫ�Է�).Value = IIf(Left(str����ֵ, 1) = "1", 1, 0)
                chk(Check�շ������Ը�).Value = IIf(Mid(str����ֵ, 2, 1) = "1", 1, 0)
                'ȫ�Է�����
                If chk(Check�շ�ȫ�Է�).Value = 1 Then
                    chk(Check�շ������Ը�).Value = 1
                    chk(Check�շ������Ը�).Enabled = False
                End If
            Case "��������ʻ�ʹ�÷�Χ"
                str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                chk(Check����ȫ�Է�).Value = IIf(Left(str����ֵ, 1) = "1", 1, 0)
                chk(Check���������Ը�).Value = IIf(Mid(str����ֵ, 2, 1) = "1", 1, 0)
                chk(Check���㳬��).Value = IIf(Mid(str����ֵ, 3, 1) = "1", 1, 0)
                'ȫ�Է�����
                If chk(Check����ȫ�Է�).Value = 1 Then
                    chk(Check���������Ը�).Value = 1
                    chk(Check���������Ը�).Enabled = False
                End If
            Case "�ȿ�����"
                str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "0", rsTemp("����ֵ"))
                chk(Check�ȿ�����).Value = IIf(Left(str����ֵ, 1) = "1", 1, 0)
            '20051220 �¶� ��������
            Case "�����޶�"
                txt97 = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "�����ֹ"
                chk97.Value = IIf(rsTemp("����ֵ") = 1, 1, 0)
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet����.Show vbModal, frmҽ�����
    �������� = mblnOK
End Function
