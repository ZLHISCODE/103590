VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotifySet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frmNotifySet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4935
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdSetup 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   165
      TabIndex        =   4
      Top             =   2295
      Width           =   1100
   End
   Begin VB.CommandButton cmdPriv 
      Caption         =   "Ԥ��(&O)"
      Height          =   350
      Left            =   1335
      TabIndex        =   5
      Top             =   2295
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   2490
      TabIndex        =   6
      Top             =   2295
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3645
      TabIndex        =   7
      Top             =   2295
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "��������"
      Height          =   2010
      Left            =   195
      TabIndex        =   8
      Top             =   120
      Width           =   4470
      Begin VB.TextBox txt�߿��� 
         Height          =   300
         Left            =   1215
         TabIndex        =   3
         Top             =   915
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Left            =   1230
         TabIndex        =   1
         Top             =   390
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   126156803
         CurrentDate     =   36576
      End
      Begin VB.Label lblEdit 
         Caption         =   "�߿���"
         Height          =   180
         Left            =   375
         TabIndex        =   2
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "֪ͨ������ӡ������ָ����ֹ���������ڼ��ڵķ���Ƿ�������"
         ForeColor       =   &H00800000&
         Height          =   450
         Left            =   690
         TabIndex        =   9
         Top             =   1365
         Width           =   3465
      End
      Begin VB.Label lbl��ֹ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ֹ����"
         Height          =   180
         Left            =   420
         TabIndex        =   0
         Top             =   465
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmNotifySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mbytType As Byte '1��ʾ��ӡ,2��ʾԤ��
Private mblnFirst As Boolean
Private mstrPrivs As String
Private mblncmdPriv As Boolean
Private mblnOk As Boolean, mstr��ֹ���� As String, mdbl�߿��� As Double
Public Function ShowSet(ByVal frmMain As Form, strPrivs As String, ByVal blncmdPriv As Boolean, ByRef bytType As Byte, ByRef str��ֹ���� As String, ByRef dbl�߿��� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�߿������������
    '���:frmMain-���õĸ�����
    '     blncmdPriv-�Ƿ���ʾԤ����ť
    '����:bytType-0 ��ʾȡ�� 1��ʾ��ӡ,2��ʾԤ��
    '     str��ֹ����
    '     dbl�߿���
    '����:
    '����:���˺�
    '����:2010-01-20 11:58:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytType = 0: mblnOk = False: mdbl�߿��� = 0: mstr��ֹ���� = str��ֹ����: mstrPrivs = strPrivs: mblncmdPriv = blncmdPriv
    Me.Show 1, frmMain
    str��ֹ���� = mstr��ֹ����: dbl�߿��� = mdbl�߿���
    ShowSet = mblnOk: bytType = mbytType
End Function
Private Sub cmdCancel_Click()
    mblnOk = False:
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    mstr��ֹ���� = Format(dtp.Value, "yyyy-mm-dd"): mdbl�߿��� = Val(txt�߿���.Text)
    mbytType = 1: mblnOk = True:  Unload Me
End Sub

Private Sub cmdPriv_Click()
    mstr��ֹ���� = Format(dtp.Value, "yyyy-mm-dd"): mdbl�߿��� = Val(txt�߿���.Text)
    mbytType = 2: mblnOk = True: Unload Me:
End Sub

Private Sub cmdSetup_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1139_3", Me
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mstr��ֹ���� <> "" And IsDate(mstr��ֹ����) Then dtp.Value = CDate(mstr��ֹ����)
End Sub

Private Sub Form_Load()
    mblnFirst = True: mbytType = 0
    txt�߿���.Text = zlDatabase.GetPara("�߿���", glngSys, 1139, "", Array(txt�߿���, lblEdit), InStr(1, mstrPrivs, ";��������;") > 0)
    dtp.Value = DateAdd("d", -1, zlDatabase.Currentdate)
    cmdPriv.Visible = mblncmdPriv
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
  Call zlDatabase.SetPara("�߿���", Val(txt�߿���.Text), glngSys, 1139, InStr(1, mstrPrivs, ";��������;") > 0)
End Sub

Private Sub txt�߿���_GotFocus()
    zlControl.TxtSelAll txt�߿���
End Sub

Private Sub txt�߿���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txt�߿���_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt�߿���, KeyAscii, m�����ʽ
End Sub

Private Sub txt�߿���_LostFocus()
    txt�߿���.Text = Format(Val(txt�߿���.Text), "0.00")
End Sub
