VERSION 5.00
Begin VB.Form frmPayExitParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frmPayExitParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6000
   StartUpPosition =   1  '����������
   Begin VB.Frame fra�豸���� 
      Caption         =   " ���ܿ��������豸���� "
      Height          =   1000
      Left            =   3360
      TabIndex        =   16
      Top             =   3720
      Width           =   2535
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame fra 
      Caption         =   " �������� "
      Height          =   1000
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   3135
      Begin VB.CheckBox chkDetailPage 
         Caption         =   "������һ�δ���ر�ʱ��ҳǩ"
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   2745
      End
      Begin VB.CheckBox chkSendByNo 
         Caption         =   "�����ݺŷ���"
         Height          =   420
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2130
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " ҵ������ "
      Height          =   1272
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   3090
      Begin VB.ComboBox cbo�շѴ��� 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   2280
      End
      Begin VB.CheckBox chkҵ�� 
         Caption         =   "�շѵ�(&S)"
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   1150
      End
      Begin VB.CheckBox chkҵ�� 
         Caption         =   "���ʵ�(&J)"
         Height          =   285
         Index           =   1
         Left            =   1850
         TabIndex        =   1
         Top             =   240
         Width           =   1150
      End
      Begin VB.CheckBox chkҵ�� 
         Caption         =   "���ʱ�(&B)"
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   1150
      End
      Begin VB.Label lbl�շѴ��� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�շѴ���"
         Height          =   420
         Left            =   120
         TabIndex        =   20
         Top             =   825
         Width           =   465
      End
      Begin VB.Label lbl�������� 
         Caption         =   "��������"
         Height          =   420
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   4920
      Width           =   8775
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   705
      Width           =   8775
   End
   Begin VB.Frame fra 
      Caption         =   " ��ӡ��Ʊ������ "
      Height          =   1305
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   5775
      Begin VB.OptionButton opt��ӡ��ʽ 
         Caption         =   "����ӡ(&N)"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton opt��ӡ��ʽ 
         Caption         =   "�Զ���ӡ(&A)"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton opt��ӡ��ʽ 
         Caption         =   "��ʾ��ӡ(&M)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "Ʊ�ݴ�ӡ����"
         Height          =   360
         Left            =   3120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   690
         Width           =   1875
      End
      Begin VB.ComboBox cboƱ������ 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2070
      End
      Begin VB.Label lblƱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ��(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   " ȱʡ��λ "
      Height          =   1270
      Index           =   0
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Width           =   2364
      Begin VB.OptionButton opt��λ 
         Caption         =   "��װ��λ(&2)"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton opt��λ 
         Caption         =   "ɢװ��λ(&1)"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   8
      Top             =   5160
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4680
      TabIndex        =   7
      Top             =   5175
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   6
      Top             =   5175
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmPayExitParaSet.frx":030A
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "��������ѡ��Ŀ,������صĴ�ӡ�����ϵ�λ�����Ʊ�ݵ�����"
      Height          =   390
      Index           =   0
      Left            =   735
      TabIndex        =   9
      Top             =   390
      Width           =   5205
   End
End
Attribute VB_Name = "frmPayExitParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mblnExit As Boolean
Private mlngModule As Long
Private mstrPrivs As String
Private mblnHavePriv As Boolean

Private Sub cboƱ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chk��ӡ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub



Private Sub chk��λ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab

End Sub



 
 


Private Sub chkҵ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub
Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1723)
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
    Dim strҵ������ As String
    Dim n As Integer
    
    strҵ������ = IIf(chkҵ��(0).Value = 1, "24", "0")
    strҵ������ = strҵ������ & IIf(chkҵ��(1).Value = 1, ",25", ",0")
    strҵ������ = strҵ������ & IIf(chkҵ��(2).Value = 1, ",26", ",0")
    
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
   
    Call zlDatabase.SetPara("���ϴ�ӡ���ѷ�ʽ", IIf(opt��ӡ��ʽ(0).Value = True, 0, IIf(opt��ӡ��ʽ(1).Value = True, 1, 2)), glngSys, mlngModule)
    Call zlDatabase.SetPara("��ѯҵ������", strҵ������, glngSys, mlngModule)
    Call zlDatabase.SetPara("���ĵ�λ", IIf(opt��λ(1).Value = True, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("�����ݺŷ���", chkSendByNo.Value, glngSys, mlngModule)
    Call zlDatabase.SetPara("�շѴ�����ʾ��ʽ", cbo�շѴ���.ListIndex, glngSys, mlngModule)
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\���ķ��Ź���", "������һ�δ���ر�ʱ��ҳǩ", Me.chkDetailPage.Value)

    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdOk_Click()
    If SaveSet = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    
    If cboƱ������.ListIndex < 0 Then
        ShowMsgBox "�����ú�Ʊ��!"
        cboƱ������.SetFocus
    End If
    Select Case cboƱ������.ListIndex
    Case 0
        '���ݴ�ӡ
        strBill = "ZL1_BILL_1723"
    Case 1
        '�嵥��ӡ
        strBill = "ZL1_BILL_1723_1"
    Case 2
        '��������֪ͨ��
        strBill = "ZL1_BILL_1723_2"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim i As Long
    Dim strArr As Variant
    Dim str�������� As String
    Dim BlnSelect As Boolean
    Dim n As Integer
    Dim int�շѴ��� As Integer
    
    mblnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "��������")
    
    With cbo�շѴ���
        .Clear
        .AddItem "1-��ʾ���еĴ���"
        .AddItem "2-����ʾ���շѴ���"
        .AddItem "3-����ʾδ�շѴ���"
        .ListIndex = 0
    End With
    
    With cboƱ������
        .Clear
        .AddItem "1-���Ĵ�����"
        .AddItem "2-��ӡ�ѷ����嵥"
        .AddItem "3-����֪ͨ����ӡ"
        .ListIndex = 0
    End With
  
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0", Array(opt��λ(0), opt��λ(1)), mblnHavePriv))
    If Val(strReg) >= 0 And Val(strReg) <= 1 Then
        opt��λ(Val(strReg)).Value = True
    Else
        opt��λ(0).Value = True
    End If
      
    strReg = Trim(zlDatabase.GetPara("���ϴ�ӡ���ѷ�ʽ", glngSys, mlngModule, "0", Array(opt��ӡ��ʽ(0), opt��ӡ��ʽ(1), opt��ӡ��ʽ(2)), mblnHavePriv))
    
    If Val(strReg) >= 0 And Val(strReg) <= 2 Then
        opt��ӡ��ʽ(Val(strReg)).Value = True
    Else
        opt��ӡ��ʽ(0).Value = True
    End If
 
    strReg = Trim(zlDatabase.GetPara("��ѯҵ������", glngSys, mlngModule, "", Array(lbl��������, chkҵ��(0), chkҵ��(1), chkҵ��(2), Frame3), mblnHavePriv))
    If strReg = "" Then strReg = "24,25,26"
    strArr = Split(strReg & "," & "," & ",", ",")
    For i = 0 To UBound(strArr)
        If i > 2 Then Exit For
        chkҵ��(i).Value = IIf(Val(strArr(i)) > 0, 1, 0)
    Next
    
    chkSendByNo.Value = IIf(Val(zlDatabase.GetPara("�����ݺŷ���", glngSys, mlngModule, , Array(chkSendByNo), mblnHavePriv)) = 1, 1, 0)
    
    int�շѴ��� = Val(zlDatabase.GetPara("�շѴ�����ʾ��ʽ", glngSys, mlngModule, 0, Array(lbl�շѴ���, cbo�շѴ���), mblnHavePriv))
    If int�շѴ��� >= 0 And int�շѴ��� <= 2 Then
        cbo�շѴ���.ListIndex = int�շѴ���
    Else
        cbo�շѴ���.ListIndex = 0
    End If
    
    'ע������
    Me.chkDetailPage.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "���ķ��Ź���", "������һ�δ���ر�ʱ��ҳǩ", 0))
End Sub
 
Public Function ShowSetPara(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '����:���ò������
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '�޸�:2007/12/24
    '-----------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    '���ز�������
     Me.Show 1, frmMain
    ShowSetPara = mblnOk
End Function
