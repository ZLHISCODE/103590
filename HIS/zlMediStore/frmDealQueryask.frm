VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmDealQueryask 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ѯ��������"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Caption         =   "��Χ����"
      Height          =   825
      Left            =   120
      TabIndex        =   11
      Top             =   1365
      Width           =   4305
      Begin VB.CheckBox chk���� 
         Caption         =   "��ĩӦ����������(&4)"
         Height          =   240
         Index           =   3
         Left            =   2190
         TabIndex        =   5
         Top             =   495
         Width           =   2010
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����֧����������(&3)"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   495
         Width           =   2010
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�����޹���������(&2)"
         Height          =   240
         Index           =   1
         Left            =   2190
         TabIndex        =   3
         Top             =   210
         Width           =   2010
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�ڳ�Ӧ����������(&1)"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   2010
      End
   End
   Begin VB.Frame fraRangeSelect 
      Caption         =   "ʱ��ѡ��"
      Height          =   1170
      Left            =   90
      TabIndex        =   6
      Top             =   75
      Width           =   2565
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   855
         TabIndex        =   1
         Top             =   630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   23658499
         CurrentDate     =   36257
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   855
         TabIndex        =   0
         Top             =   270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   23658499
         CurrentDate     =   36257
      End
      Begin VB.Label lblStartDate 
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����"
         Height          =   180
         Left            =   90
         TabIndex        =   8
         Top             =   330
         Width           =   735
      End
      Begin VB.Label lblEndDate 
         BackStyle       =   0  'Transparent
         Caption         =   "��ֹ����"
         Height          =   180
         Left            =   75
         TabIndex        =   7
         Top             =   690
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2940
      TabIndex        =   10
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2940
      TabIndex        =   9
      Top             =   210
      Width           =   1100
   End
End
Attribute VB_Name = "frmDealQueryask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnAskOk As Boolean
Private Sub CmdCancel_Click()
    blnAskOk = False
    Me.Hide
End Sub

Private Sub CmdOk_Click()
    blnAskOk = True
    SaveSetting "ZLHIS", "Ӧ����ҩ���ѯ", "Para", IIf(Me.chk����(0).Value = 1, "1", "0") & _
                                          IIf(Me.chk����(1).Value = 1, "1", "0") & _
                                          IIf(Me.chk����(2).Value = 1, "1", "0") & _
                                          IIf(Me.chk����(3).Value = 1, "1", "0")
    Me.Hide
End Sub

Private Sub dtpStartDate_Change()
    If Me.dtpStartDate.Value > Me.dtpEndDate.Value Then
        Me.dtpEndDate.Value = Me.dtpStartDate.Value
    End If
End Sub

Private Sub dtpEndDate_Change()
    If Me.dtpStartDate.Value > Me.dtpEndDate.Value Then
        Me.dtpStartDate.Value = Me.dtpEndDate.Value
    End If
End Sub

Private Sub Form_Load()
    Dim strSql As String

    Me.dtpEndDate.MaxDate = currentdate()
    Me.dtpEndDate.Value = Me.dtpEndDate.MaxDate
    Me.dtpStartDate.MaxDate = Me.dtpEndDate.Value
    Me.dtpStartDate.Value = DateAdd("m", -1, Me.dtpEndDate.Value)

End Sub
