VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTimeSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ӧ�����ѯ����"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frmTimeSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraTemp2 
      Caption         =   "��������"
      Height          =   1005
      Left            =   150
      TabIndex        =   5
      Top             =   1620
      Width           =   4575
      Begin VB.CheckBox chkData 
         Caption         =   "��ĩӦ����Ϊ��(&4)"
         Height          =   225
         Index           =   3
         Left            =   2370
         TabIndex        =   9
         Top             =   660
         Width           =   1845
      End
      Begin VB.CheckBox chkData 
         Caption         =   "����֧����Ϊ��(&3)"
         Height          =   225
         Index           =   2
         Left            =   300
         TabIndex        =   8
         Top             =   660
         Width           =   1845
      End
      Begin VB.CheckBox chkData 
         Caption         =   "�����޹���Ϊ��(&2)"
         Height          =   225
         Index           =   1
         Left            =   2370
         TabIndex        =   7
         Top             =   300
         Width           =   1845
      End
      Begin VB.CheckBox chkData 
         Caption         =   "�ڳ�Ӧ����Ϊ��(&1)"
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   300
         Width           =   1845
      End
   End
   Begin VB.Frame fraTemp1 
      Caption         =   "ʱ�䷶Χ"
      Height          =   1365
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   3165
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1260
         TabIndex        =   4
         Top             =   810
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   68419587
         CurrentDate     =   36279
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1260
         TabIndex        =   2
         Top             =   390
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   68419587
         CurrentDate     =   36279
         MinDate         =   2
      End
      Begin VB.Label lblTimeStart 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��(&B)"
         Height          =   180
         Left            =   180
         TabIndex        =   1
         Top             =   450
         Width           =   990
      End
      Begin VB.Label lblTimeStop 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��(&E)"
         Height          =   180
         Left            =   180
         TabIndex        =   3
         Top             =   870
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   11
      Top             =   570
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3630
      TabIndex        =   10
      Top             =   150
      Width           =   1100
   End
End
Attribute VB_Name = "frmTimeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mdatBegin As Date, mdatEnd As Date
Dim mstrData As String

Private Sub chkData_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub



Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpEnd.SetFocus
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then chkData(0).SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "��ʼʱ����ڽ���ʱ���ˡ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    mdatBegin = dtpBegin.Value
    mdatEnd = dtpEnd.Value
    mstrData = chkData(0).Value & chkData(1).Value & chkData(2).Value & chkData(3).Value
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function GetTimeScope(datBegin As Date, datEnd As Date, strData As String, ByVal frmOwner As Form) As Boolean
'--------------------------------------------------------------
'���ܣ���ȡӦ�����ѯ���ڷ�Χ����������
'������datBegin---------��ʼ����
'      datEnd-----------��������
'      strData----------���������ַ���
'      frmOwner---------���ô���
'���أ��Ƿ��ѯ
'˵����
'--------------------------------------------------------------
    Dim intTemp As Long
    
    dtpBegin.Value = datBegin
    dtpEnd.Value = datEnd
    dtpBegin.MaxDate = zlDatabase.Currentdate
    dtpEnd.MaxDate = dtpBegin.MaxDate
    
    For intTemp = 0 To 3
        chkData(intTemp).Value = Val(Mid(strData, intTemp + 1, 1))
    Next
    
    frmTimeSet.Show vbModal, frmOwner
    GetTimeScope = mblnOK
    If mblnOK = True Then
        datBegin = mdatBegin
        datEnd = mdatEnd
        strData = mstrData
    End If
End Function
