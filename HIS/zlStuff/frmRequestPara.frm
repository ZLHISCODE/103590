VERSION 5.00
Begin VB.Form frmRequestPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmRequestPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk����˲� 
      Caption         =   "������Ҫ�˲������ƿ�"
      Height          =   375
      Left            =   1020
      TabIndex        =   10
      Top             =   1440
      Width           =   3105
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "���ݴ�ӡ����(&S)"
      Height          =   350
      Left            =   1020
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2385
      Width           =   2400
   End
   Begin VB.Frame fra 
      Height          =   120
      Index           =   1
      Left            =   -30
      TabIndex        =   9
      Top             =   2910
      Width           =   5790
   End
   Begin VB.Frame fra 
      Height          =   120
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   645
      Width           =   5580
   End
   Begin VB.ComboBox Cboָ����λ 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1050
      Width           =   2415
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   30
      TabIndex        =   6
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2940
      TabIndex        =   4
      Top             =   3150
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4275
      TabIndex        =   5
      Top             =   3150
      Width           =   1100
   End
   Begin VB.CheckBox chkSavePrint 
      Caption         =   "���̴�ӡ"
      Height          =   375
      Left            =   1020
      TabIndex        =   2
      Top             =   1875
      Width           =   3105
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   45
      Picture         =   "frmRequestPara.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "���ѡ����̴�ӡ�����ڵ����У����ݴ��̺��Զ���ӡ�����򲻴�ӡ��"
      Height          =   615
      Left            =   630
      TabIndex        =   8
      Top             =   225
      Width           =   3180
   End
   Begin VB.Label lbl���ϵ�λ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ϵ�λ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   1110
      Width           =   720
   End
End
Attribute VB_Name = "frmRequestPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Private mstrPrivs As String
Private mlngModule As Long
Private mblnHavePriv As Boolean '�Ƿ��в�������Ȩ��

Private Sub Cboָ����λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chkSavePrint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '����:�����ݿⱣ���������
    '����:����ɹ�����True,���򷵻�False
    '����:���˺�
    '����:2007/12/24
    '------------------------------------------------------------------------------------------
 
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
    Call zlDatabase.SetPara("���̴�ӡ", IIf(chkSavePrint.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("���ĵ�λ", Cboָ����λ.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("������Ҫ�˲������ƿ�", IIf(chk����˲�.Value = 1, 1, 0), glngSys, mlngModule)
    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function
Private Sub cmdOk_Click()
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub
Private Sub initPara()
    '-----------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:
    '����:���˺�
    '�޸�:2007/12/24
    '-----------------------------------------------------------------------------------------------------------------------------------------
    Dim strReg As String
    mblnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "��������")
    chkSavePrint.Value = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0", Array(chkSavePrint), mblnHavePriv)) = 1, 1, 0)
    Me.CmdPrintSet.Enabled = InStr(1, mstrPrivs, ";���ݴ�ӡ;") <> 0
    chk����˲�.Value = IIf(Val(zlDatabase.GetPara("������Ҫ�˲������ƿ�", glngSys, mlngModule, "0", Array(chk����˲�), mblnHavePriv)) = 1, 1, 0)
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0", Array(Cboָ����λ, lbl���ϵ�λ), mblnHavePriv))
    With Cboָ����λ
        .Clear
        .AddItem "ɢװ��λ"
        .AddItem "��װ��λ"
        .ListIndex = Val(strReg)
    End With
End Sub
Public Sub ���ò���(ByVal lngModule As Long, frmMain As Form, Optional ByVal strFunction As String = "", Optional strPrivs As String = "")
    '-----------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ý���
    '����:
    '����:���˺�
    '�޸�:2007/12/24
    '-----------------------------------------------------------------------------------------------------------------------------------------
    mstrFunction = strFunction: mlngModule = lngModule:    mstrPrivs = IIf(strPrivs = "", gstrPrivs, strPrivs)
    Call initPara
    frmRequestPara.Show vbModal, frmMain
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_" & glngModul
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

