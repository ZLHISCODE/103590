VERSION 5.00
Begin VB.Form frmProReport 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "ר�Ʊ���"
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "����ͨ��"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
End
Attribute VB_Name = "frmProReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Call gobjParent.SendReport(glngAdviceId, "�ӿڵ��ò���", "���")
End Sub
