VERSION 5.00
Begin VB.Form frmPatientHistorySetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   Icon            =   "frmPatientHistorySetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   300
      Left            =   150
      TabIndex        =   4
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   300
      Left            =   1620
      TabIndex        =   2
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ҷ�Χ"
      Height          =   1065
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   3735
      Begin VB.OptionButton OptInPatient 
         Caption         =   "סԺ����"
         Height          =   300
         Left            =   2160
         TabIndex        =   1
         Top             =   420
         Width           =   1425
      End
      Begin VB.OptionButton OptOutPatient 
         Caption         =   "�������"
         Height          =   300
         Left            =   270
         TabIndex        =   0
         Top             =   420
         Value           =   -1  'True
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmPatientHistorySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    '��ʾ����
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    With frmPatientHistoryQuery
        .GetFilterDate IIf(Me.OptOutPatient.Value = True, True, False)
    End With
    Unload Me
End Sub
