VERSION 5.00
Begin VB.Form frmSet����ɽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ղ�������"
   ClientHeight    =   2295
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3660
   Icon            =   "frmSet����ɽ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1110
      TabIndex        =   5
      Top             =   1740
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2395
      TabIndex        =   6
      Top             =   1740
      Width           =   1100
   End
   Begin VB.Frame fraTop 
      Caption         =   "���в���"
      Height          =   1410
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   3345
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   2
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   4
         Top             =   735
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmSet����ɽ.frx":000C
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ҽ�����ų���(&R)"
         Height          =   180
         Index           =   0
         Left            =   990
         TabIndex        =   1
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����֤�ų���(&T)"
         Height          =   180
         Index           =   1
         Left            =   990
         TabIndex        =   3
         Top             =   795
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmSet����ɽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum�༭
    Text���ų��� = 0
    Text����֤�� = 1
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
    
    If Val(TxtEdit(Text���ų���).Text) > 25 Then
        MsgBox "���ų��Ȳ��ܳ���25λ��", vbInformation, gstrSysName
        TxtEdit(Text���ų���).SetFocus
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

Public Function ��������() As Boolean
'���ܣ�������������ҽ������Ҫ�Ĳ���
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    
    mblnOK = False
    mlng���� = TYPE_��������ɽ
    mlng���� = 0
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1] and (���� is null or ����=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����, mlng����)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "���ų���"
                TxtEdit(Text���ų���).Text = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "����֤����"
                TxtEdit(Text����֤��).Text = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet����ɽ.Show vbModal, frmҽ�����
    �������� = mblnOK
End Function
