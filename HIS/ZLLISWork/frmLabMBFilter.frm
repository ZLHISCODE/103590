VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLabMBFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "frmLabMBFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkDate 
      Caption         =   "����ʱ��"
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   285
      Left            =   4740
      TabIndex        =   8
      Top             =   1380
      Width           =   1155
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   285
      Left            =   3300
      TabIndex        =   7
      Top             =   1380
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   6495
   End
   Begin VB.TextBox txt�Լ����� 
      Height          =   300
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   2085
   End
   Begin VB.TextBox txt���԰�� 
      Height          =   300
      Left            =   1020
      TabIndex        =   0
      Top             =   120
      Width           =   2085
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1350
      TabIndex        =   3
      Top             =   630
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   99418115
      CurrentDate     =   39497
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   4020
      TabIndex        =   4
      Top             =   630
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   99418115
      CurrentDate     =   39497
   End
   Begin VB.Label Label1 
      Caption         =   "---"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   690
      Width           =   405
   End
   Begin VB.Label lbl���԰�� 
      AutoSize        =   -1  'True
      Caption         =   "���԰��"
      Height          =   180
      Left            =   210
      TabIndex        =   5
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lbl�Լ����� 
      AutoSize        =   -1  'True
      Caption         =   "�Լ�����"
      Height          =   180
      Left            =   3270
      TabIndex        =   1
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frmLabMBFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilter As String
Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    strFilter = Me.txt���԰�� & ";" & Me.txt�Լ����� & ";" & Me.chkDate.Value & "," & Me.dtpBegin & "," & Me.dtpEnd
    Unload Me
End Sub

Private Sub Form_Load()
    Me.dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    Me.dtpEnd.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
End Sub


Public Function ShowMe(objfrm As Object) As String
    Me.Show vbModal, objfrm
    ShowMe = strFilter
End Function
