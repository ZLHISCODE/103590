VERSION 5.00
Begin VB.Form frmBillSuperviseParaSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3345
      TabIndex        =   4
      Top             =   2670
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4560
      TabIndex        =   3
      Top             =   2670
      Width           =   1100
   End
   Begin VB.Frame fraSplitDown 
      Height          =   120
      Left            =   -75
      TabIndex        =   1
      Top             =   2250
      Width           =   9600
   End
   Begin VB.Frame fraTopSplit 
      Height          =   120
      Left            =   -15
      TabIndex        =   0
      Top             =   885
      Width           =   9600
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   405
      Picture         =   "frmBillSuperviseParaSet.frx":0000
      Top             =   285
      Width           =   480
   End
   Begin VB.Label lblNotes 
      Caption         =   "�����ʵ�����,�����������ز���"
      Height          =   540
      Left            =   1110
      TabIndex        =   2
      Top             =   450
      Width           =   4320
   End
End
Attribute VB_Name = "frmBillSuperviseParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnOk As Boolean
Public Function zlSetPara(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:�������óɹ�,����true,����ķ���False
    '����:���˺�
    '����:2011-07-14 17:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnOk = False
    Me.Show 1, frmMain
    zlSetPara = mblnOk
End Function
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim blnHavePrivs As Boolean, intData As Integer
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    If isValied = False Then Exit Sub
    mblnOk = True: Unload Me
End Sub
Private Sub InitPara()
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
End Sub
Private Sub Form_Load()
    Call InitPara
End Sub
 
