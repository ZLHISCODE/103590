VERSION 5.00
Begin VB.Form frm����������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   Icon            =   "frm�����������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      Begin VB.Frame fraUnit 
         Caption         =   "��ʾ��λ"
         Height          =   735
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   4335
         Begin VB.OptionButton optUnit 
            Caption         =   "���λ"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   13
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton optUnit 
            Caption         =   "��С��λ"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   12
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkCheck 
         Caption         =   "һ�㸶����Ҫ����Ԥ��(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtGetBefor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   200
         Left            =   2080
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "1"
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtGetMonth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   200
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "5"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox chkOldPrice 
         Caption         =   "���ø���ʱ��"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblCheck 
         Caption         =   "����Ԥ��׶�ֻ�й�ѡȫ����ϸ�󣬲��ܽ�����ˡ�"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   750
         Width           =   4335
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "˵������������ѡ��ʱ������˼���Ʊʱ�����������ʽ���á�"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2040
         X2              =   2520
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblOldPrice3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ڽ���ʱ����ǰ��     ���¡�"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   2790
      End
      Begin VB.Label lblOldPrice1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��������ҪĬ��Ϊ     ����ǰ���ų���ǰ������"
         Height          =   180
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   4410
      End
      Begin VB.Label lblOldPrice2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ǰ�����õ��£���Ȼ�����һ�죩��Ϊ����ʱ�䣬��ʼʱ"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   4500
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2325
         X2              =   2805
         Y1              =   1635
         Y2              =   1635
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   3840
      TabIndex        =   2
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   345
      Left            =   2640
      TabIndex        =   1
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   1100
   End
End
Attribute VB_Name = "frm�����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String '����ģ�������,���������Caption���磺�⹺�������
Private mblnSelect As Boolean   '�ԣңգű���ѡ��ȷ���˳�������Ϊȡ���˳���
Private mstrPrivs As String
Private mlngModule  As String
Private mblnHavePriv As Boolean

Private Sub chkOldPrice_Click()
    If chkOldPrice.Value = 0 Then
        txtGetMonth.Enabled = False
        txtGetBefor.Enabled = False
    Else
        If mblnHavePriv Then
            txtGetMonth.Enabled = True
            txtGetBefor.Enabled = True
        End If
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, 4)
End Sub

Private Sub cmdOK_Click()
    If SaveData = False Then Exit Sub
    mblnSelect = True
    Unload Me
End Sub

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------
    '����:�������ֵ
    '����:����ɹ�,����true,���򷵻�False
    '����:lesfeng
    '����:2010/03/25
    '---------------------------------------------------------------------------------------------
    Dim strKey As String
    Dim strOldPrice As String, strCheck As String, strUnit As String
    Err = 0: On Error GoTo ErrHand:
    
    gcnOracle.BeginTrans
    If chkOldPrice.Value = 0 Then
        strOldPrice = "0-" & Val(txtGetMonth.Text) & "-" & Val(txtGetBefor.Text)
    Else
        strOldPrice = "1-" & Val(txtGetMonth.Text) & "-" & Val(txtGetBefor.Text)
    End If
    
    strCheck = IIf(chkCheck.Value = 1, "1", "0")
    strUnit = IIf(optUnit(0).Value, "0", "1")
    
    Call zldatabase.SetPara("���ø���ʱ��", strOldPrice, glngSys, mlngModule, IIf(chkOldPrice.Enabled = True, True, False))
    Call zldatabase.SetPara("һ�㸶����Ҫ����Ԥ��", strCheck, glngSys, mlngModule, IIf(chkCheck.Enabled = True, True, False))
    Call zldatabase.SetPara("��ʾ��λѡ��", strUnit, glngSys, mlngModule, IIf(optUnit(0).Enabled = True, True, False))
    
    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Function

Private Sub InitPara()
    '------------------------------------------------------------------------------------------
    '����:�Ա��ز�����������
    '����:
    '����:���˺�
    '�޸�:2007/12/19
    '------------------------------------------------------------------------------------------
    Dim strOldPrice As String, strCheck As String, strUnit As String
    Dim arrHead As Variant
    
    mblnHavePriv = InStr(mstrPrivs, ";��������;") > 0
    
'    chkSavePrint.Value = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, , Array(chkSavePrint), mblnHavePriv)) = 1, 1, 0)
'    chkVerifyPrint.Value = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, , Array(chkVerifyPrint), mblnHavePriv)) = 1, 1, 0)
'
'    If mlngModule <> 316 Then
'        chk�Զ����.Value = IIf(Val(zlDatabase.GetPara("��Զ����", glngSys, mlngModule, , Array(chk�Զ����), mblnHavePriv)) = 1, 1, 0) ' And IsHavePrivs(mstrPrivs, "����")
'    End If
    
    strCheck = zldatabase.GetPara("һ�㸶����Ҫ����Ԥ��", glngSys, mlngModule, , Array(chkCheck), mblnHavePriv)
    strOldPrice = zldatabase.GetPara("���ø���ʱ��", glngSys, mlngModule, , Array(chkOldPrice, lblOldPrice1, lblOldPrice2, lblOldPrice3, txtGetMonth, txtGetBefor), mblnHavePriv)
    strUnit = zldatabase.GetPara("��ʾ��λѡ��", glngSys, mlngModule, , Array(optUnit), mblnHavePriv)
    
    If Val(strCheck) = 1 Then
        chkCheck.Value = 1
    Else
        chkCheck.Value = 0
    End If
    
    If Val(strUnit) = 1 Then
        optUnit(1).Value = True
    Else
        optUnit(0).Value = True
    End If
    
    If InStr(1, strOldPrice, "-") > 0 Then
        arrHead = Split(strOldPrice, "-")
        chkOldPrice.Value = Val(arrHead(0))
        txtGetMonth.Text = arrHead(1)
        txtGetBefor.Text = arrHead(2)
        If chkOldPrice.Value = 0 Then
            txtGetMonth.Enabled = False
            txtGetBefor.Enabled = False
        End If
    Else
        txtGetMonth.Enabled = False
        txtGetBefor.Enabled = False
    End If
End Sub

Public Function ���ò���(frmParent As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '------------------------------------------------------------------------------------------
    '����:�Ա��ز�����������
    '����:���óɹ�,����True,���򷵻�False
    '����:���˺�
    '�޸�:2007/12/19
    '------------------------------------------------------------------------------------------
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mstrFunction = frmParent.Caption
    mblnSelect = False
    
    Call InitPara
    
    frm�����������.Show vbModal, frmParent
    ���ò��� = mblnSelect
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtGetMonth_GotFocus()
    With txtGetMonth
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtGetMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtGetMonth_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtGetMonth, KeyAscii, m����ʽ
End Sub

Private Sub txtGetBefor_GotFocus()
    With txtGetBefor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtGetBefor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtGetBefor_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtGetBefor, KeyAscii, m����ʽ
End Sub

