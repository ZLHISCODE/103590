VERSION 5.00
Begin VB.Form frm�嵥����Set 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frm�嵥����Set.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame fra��λ 
      Caption         =   "ҩƷ���ֵ�λ"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton opt��λ 
         Caption         =   "ҩ�ⵥλ(&4)"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton opt��λ 
         Caption         =   "סԺ��λ(&3)"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton opt��λ 
         Caption         =   "���ﵥλ(&2)"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton opt��λ 
         Caption         =   "�ۼ۵�λ(&1)"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm�嵥����Set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long
Private mstrPrivs As String
Private mblnHavePriv As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub
 
Public Sub ��������(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '-------------------------------------------------------------------------------------------
    '����:�ṩ���ϼ��������
    '����:frmMain-���������
    '     lngModule-ģ���
    '     strPrivs-Ȩ�޴�
    '����:
    '����:lesfeng
    '�޸�:2010/02/25
    '-------------------------------------------------------------------------------------------
    mlngModule = lngModule:    mstrPrivs = strPrivs
    mblnHavePriv = IsHavePrivs(mstrPrivs, "��������")
    
    Call InitDate
    Me.Show vbModal, frmMain
    
End Sub

Sub InitDate()
    ''''''''''''''''''''''''''''''''''
    '����               ��ʹ������
    ''''''''''''''''''''''''''''''''''
    Dim strTmp As String
    Dim i As Long
     
    'ѡ��Ĭ�ϵ�λ
    strTmp = Trim(zlDatabase.GetPara("��λ", glngSys, mlngModule, , Array(fra��λ, opt��λ(0), opt��λ(1), opt��λ(2), opt��λ(3)), mblnHavePriv))
    Select Case strTmp
    Case "0"
        opt��λ(0).Value = True
    Case "1"
        opt��λ(1).Value = True
    Case "2"
        opt��λ(2).Value = True
    Case "3"
        opt��λ(3).Value = True
    End Select
End Sub

Private Function SaveDate() As Boolean
    '------------------------------------------------------------------------------------------------
    '����       ��������
    '------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTmp As String
    
    For i = 0 To opt��λ.Count - 1
        If opt��λ(i).Value Then
            strTmp = i
        End If
    Next
    
    Err = 0: On Error GoTo ErrHand:
    Call zlDatabase.SetPara("��λ", strTmp, glngSys, mlngModule, IIf(opt��λ(0).Enabled = True, True, False))
    SaveDate = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Function

Private Sub cmdHelp_Click()
    '����:���ð���
    '�޸�:lesfeng
    '����:2010-02-25
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdOK_Click()
    If SaveDate = False Then Exit Sub
    Unload Me
End Sub

Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function

