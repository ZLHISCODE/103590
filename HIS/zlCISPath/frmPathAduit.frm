VERSION 5.00
Begin VB.Form frmPathAduit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���"
   ClientHeight    =   3810
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6135
   Icon            =   "frmPathAduit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   6
      Top             =   3285
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4815
      TabIndex        =   5
      Top             =   3285
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   6030
   End
   Begin VB.TextBox txtContent 
      Height          =   1140
      Left            =   345
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1815
      Width           =   5550
   End
   Begin VB.OptionButton optAduit 
      Caption         =   "��˲�ͨ��"
      Height          =   225
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   1005
      Width           =   1305
   End
   Begin VB.OptionButton optAduit 
      Caption         =   "���ͨ��"
      Height          =   225
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   1005
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   765
      Width           =   6030
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "ͨ����ͨ��������(&M):"
      Height          =   180
      Left            =   330
      TabIndex        =   8
      Top             =   1490
      Width           =   1980
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   180
      Picture         =   "frmPathAduit.frx":6852
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ϸ����ٴ�·���������Ƿ����Ҫ�󣬾���ͨ����ͨ����"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   810
      TabIndex        =   7
      Top             =   180
      Width           =   5115
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPathAduit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng·��ID As Long
Private mlng�汾�� As Long
                   
Private mintFunc  As Integer     '1=���, 2=ҩ�������
Private mblnOK As Boolean

Public Function ShowAudit(ByVal frmParent As Object, ByVal lng·��ID As Long, ByVal lng�汾�� As Long, ByVal intFunc As Integer) As Boolean
    On Error GoTo errHand
 
    mlng·��ID = lng·��ID
    mlng�汾�� = lng�汾��
    mintFunc = intFunc
    mblnOK = False
    Me.Show 1, frmParent
    
    ShowAudit = mblnOK
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim intType As Integer

    If zlCommFun.ActualLen(txtContent.Text) > 200 Then
        MsgBox "����������ֻ���� 100 �����ֻ� 200 ���ַ���", vbInformation, gstrSysName
        txtContent.SetFocus: Exit Sub
    End If
    If optAduit(1).Value And Trim(txtContent.Text) = "" Then
        MsgBox "��˲�ͨ��ʱ����¼��ԭ��", vbInformation, gstrSysName
        txtContent.SetFocus: Exit Sub
    End If
    
    If mintFunc = 1 Then 'ҽ������
        intType = IIf(optAduit(0).Value = True, 1, 2)
    ElseIf mintFunc = 2 Then
        intType = IIf(optAduit(0).Value = True, 3, 4)
    End If
    
    On Error GoTo errH
    
    strSql = "Select Nvl(���״̬, 0) As ���״̬ From �ٴ�·��Ŀ¼ Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯ·�����״̬", mlng·��ID)
    If rsTmp.RecordCount > 0 Then
       If (InStr(",1,2,", intType) > 0 And Not (Val(rsTmp!���״̬ & "") = 1 Or Val(rsTmp!���״̬ & "") = 2)) Or _
            (InStr(",3,4,", intType) > 0 And Val(rsTmp!���״̬ & "") <> 1) Then
           MsgBox "��ǰ·��״̬�Ѹı䲻�ܽ������,��ˢ�º�����!", vbInformation, gstrSysName
           Unload Me
           Exit Sub
       End If
    End If

    strSql = "Zl_�ٴ�·�����_Insert(" & intType & "," & mlng·��ID & "," & mlng�汾�� & ",'" & Trim(txtContent.Text) & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "�ٴ�·�����")
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub txtContent_GotFocus()
    zlControl.TxtSelAll txtContent
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtContent_LostFocus()
    Me.txtContent.Text = Replace(Me.txtContent, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub

