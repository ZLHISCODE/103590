VERSION 5.00
Begin VB.Form frmInSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�豸����"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4245
      TabIndex        =   1
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4245
      TabIndex        =   0
      Top             =   120
      Width           =   1100
   End
End
Attribute VB_Name = "frmInSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
'    SaveSetting "ZLSOFT", "����ȫ��", "�е�����Ϣ", chkBottom.Value
    Unload Me
End Sub

Private Sub Form_Load()
'    strBottom = GetSetting("ZLSOFT", "����ȫ��", "������Ϣ", "������Ϣ")
End Sub
