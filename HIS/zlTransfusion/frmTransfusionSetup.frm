VERSION 5.00
Begin VB.Form frmTransfusionSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmTransfusionSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5351.164
   ScaleMode       =   0  'User
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkTimeCall 
      Caption         =   "�����ƶ����й���"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.CheckBox chk�ӵ����� 
      Caption         =   "�ӵ���ֱ�ӽ��봩��״̬"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3045
   End
   Begin VB.CheckBox chkAutoReady 
      Caption         =   "ͨ�����ҹ����ҵ����˺��Զ��ӵ�"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Frame frmCardSet 
      Caption         =   "�豸����"
      Height          =   675
      Left            =   270
      TabIndex        =   8
      Top             =   2040
      Width           =   4470
      Begin VB.CommandButton cmdCardSet 
         Caption         =   "����(&P)"
         Height          =   350
         Left            =   2985
         TabIndex        =   9
         Top             =   210
         Width           =   1100
      End
   End
   Begin VB.Frame fra 
      Caption         =   "��ѡ�񱾹���վ��ʾ�ĵ�������"
      Height          =   660
      Left            =   270
      TabIndex        =   3
      Top             =   1200
      Width           =   4485
      Begin VB.CheckBox chkType 
         Caption         =   "����"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "��Һ"
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   5
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "ע��"
         Height          =   195
         Index           =   2
         Left            =   2355
         TabIndex        =   6
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Ƥ��"
         Height          =   195
         Index           =   3
         Left            =   3435
         TabIndex        =   7
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   210
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2955
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   11
      Top             =   2955
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   12
      Top             =   2955
      Width           =   1100
   End
End
Attribute VB_Name = "frmTransfusionSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'2015-09-15�����Ρ�������Һ�Զ��ӵ�������

Public mstrPrivs As String
Public mlng����ID As Long 'IN:��ǰִ�п���ID
Public mblnOk As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCardSet_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strPar As String, i As Long
    Dim strType As String
    Dim blnModify As Boolean
    
    'ִ�м䷶Χ
    blnModify = False
    If InStr(mstrPrivs, "��������") > 0 Then blnModify = True
    
    '�ӵ���ֱ�ӽ��봩��״̬
    Call zlDatabase.SetPara("�ӵ�ֱ�Ӵ���", chk�ӵ�����.Value, glngSys, 1264)
    
    '�ƶ�����
    Call zlDatabase.SetPara("�ƶ�����", chkTimeCall.Value, glngSys, 1264)
    
    '2008-11-12
    strType = ""
    For i = 0 To chkType.Count - 1
        strType = strType & "," & chkType(i).Value
    Next
    Call zlDatabase.SetPara("��ʾ��������", Mid(strType, 2), glngSys, 1264, blnModify)
    
    '2012-05-14 10.30 sp ���
    Call zlDatabase.SetPara("������Һ�Զ��ӵ�", chkAutoReady.Value, glngSys, 1264, blnModify)
    
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
End Sub

Private Sub Form_Load()
    Dim strType As String, i As Integer
    Dim intType As Integer '������������
    Dim blnModify As Boolean
    
    mblnOk = False
    blnModify = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    
    cmdCardSet.Enabled = blnModify
    
    '�ӵ���ֱ�ӽ��봩��״̬
    chk�ӵ�����.Value = Val(zlDatabase.GetPara("�ӵ�ֱ�Ӵ���", glngSys, 1264, ""))
    
    '�ƶ���ʱ����
    chkTimeCall.Value = Val(zlDatabase.GetPara("�ƶ�����", glngSys, 1264))
        
    '2008-11-12
    'strType = zlDatabase.GetPara("��ʾ��������", glngSys, 1264, "1,1,1,1", Array(Me.chkType(0), Me.chkType(1), Me.chkType(2), Me.chkType(3)), blnModify, intType)
    strType = zlDatabase.GetPara("��ʾ��������", glngSys, 1264, "1,1,1,1")
    For i = 0 To chkType.Count - 1
        chkType(i).Value = Val(Split(strType, ",")(i))
    Next
    '2012-05-14
    chkAutoReady.Value = Val(zlDatabase.GetPara("������Һ�Զ��ӵ�", glngSys, 1264, "", Array(chkAutoReady), blnModify))
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
    mstrPrivs = ""
End Sub

