VERSION 5.00
Begin VB.Form frmEarnRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����\ͣ��ԭ��"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmEarnRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4560
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtԭ�� 
      Height          =   1695
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lblԭ�� 
      Caption         =   "����ԭ��(�����¼��50�����֣������)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmEarnRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrԭ�� As String

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Me.txtԭ��.Text = "" Then
        Exit Sub
    End If
    
    If zlCommFun.ActualLen(Me.txtԭ��.Text) > 100 Then
        MsgBox "��ǰ�����ԭ�򳬳�50�����֣�����㣩��", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    mstrԭ�� = Me.txtԭ��.Text
    
    Unload Me
End Sub

Public Sub ShowMe(ByVal intType As Integer, ByRef strԭ�� As String)
    If intType = 1 Then
        Me.lblԭ��.Caption = "����ԭ��(�����¼��50�����֣������)"
    Else
        Me.lblԭ��.Caption = "ͣ��ԭ��(�����¼��50�����֣������)"
    End If
    
    mstrԭ�� = ""
    
    Me.Show 1
    
    strԭ�� = mstrԭ��
End Sub
