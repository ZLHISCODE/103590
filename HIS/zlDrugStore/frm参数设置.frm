VERSION 5.00
Begin VB.Form frm�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frm��������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra�Ƿ�������ʾ 
      Caption         =   "��ǰ�ⷿҩƷ�Ƿ�������ʾ(������������ʱ)"
      Height          =   735
      Left            =   180
      TabIndex        =   24
      Top             =   4440
      Width           =   6975
      Begin VB.OptionButton opt��������ʾ 
         Caption         =   "������������ʾ��ǰ�ⷿ�Ŀ������"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton opt��������ʾ 
         Caption         =   "����ǰ�ⷿ��ҩƷ��������ʾ"
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   25
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame frm��ǰ��� 
      Caption         =   "�ʱ��ǰ�ⷿ�����ʾ��ʽ"
      Height          =   735
      Left            =   180
      TabIndex        =   21
      Top             =   3480
      Width           =   3270
      Begin VB.OptionButton opt��ǰ��� 
         Caption         =   "��ʾʵ������"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton opt��ǰ��� 
         Caption         =   "��ʾ��������"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frm�Է���� 
      Caption         =   "�ʱ�Է��ⷿ�����ʾ��ʽ"
      Height          =   735
      Left            =   3570
      TabIndex        =   18
      Top             =   3480
      Width           =   3615
      Begin VB.OptionButton opt�Է���� 
         Caption         =   "��ʾʵ������"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton opt�Է���� 
         Caption         =   "��ʾ��������"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fra�����ʾ���� 
      Caption         =   "�����ʾ����"
      Height          =   855
      Left            =   180
      TabIndex        =   15
      Top             =   2520
      Width           =   6975
      Begin VB.CheckBox chkShow 
         Caption         =   "��ʾ�޿���ҩƷ"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbComment 
         Caption         =   "˵��������ʱҩƷѡ�������Ƿ���ʾ�޿���ҩƷ��¼"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   495
         Width           =   6420
      End
   End
   Begin VB.TextBox txt��ѯ���� 
      Height          =   300
      Left            =   4395
      TabIndex        =   12
      Text            =   "1"
      Top             =   2130
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Frame fraSort 
      Caption         =   "����ʽ"
      Height          =   1770
      Left            =   3510
      TabIndex        =   8
      Top             =   240
      Width           =   3675
      Begin VB.ComboBox Cbo���� 
         Height          =   300
         ItemData        =   "frm��������.frx":000C
         Left            =   120
         List            =   "frm��������.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   390
         Width           =   2415
      End
      Begin VB.ComboBox Cbo���� 
         Height          =   300
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   390
         Width           =   885
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "    �����������ã���Ӱ�����б༭�����е��ݵ���ʾ���ݵ�����ʽ��ȱʡ�����û������˳����ʾ�����ݵ�����"
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   180
         TabIndex        =   11
         Top             =   930
         Width           =   3345
      End
   End
   Begin VB.ComboBox Cboָ����λ 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   180
      TabIndex        =   2
      Top             =   480
      Width           =   3255
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   120
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmd��ӡ���� 
         Caption         =   "��ӡ����(&P)"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   2985
      End
      Begin VB.CheckBox chkPrintCode 
         Caption         =   "���̻���˺��ӡҩƷ����"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox chkVerifyPrint 
         Caption         =   "��˴�ӡ����"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CheckBox chkSavePrint 
         Caption         =   "���̴�ӡ����"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   90
      TabIndex        =   7
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4590
      TabIndex        =   5
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5940
      TabIndex        =   6
      Top             =   5400
      Width           =   1100
   End
   Begin VB.Label lbl��ѯ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ѯ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3600
      TabIndex        =   14
      Top             =   2190
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5340
      TabIndex        =   13
      Top             =   2190
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblҩƷ��λ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ��λ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Private mlngMode As Long
Dim mstrPrivs As String
Private mblnSetPara As Boolean                          '�Ƿ���в�������Ȩ��

Private Const M_LNG_FRMWIDTH_1 = 3800
Private Const M_LNG_FRMWIDTH_2 = 7500
Private Const M_LNG_FRMHEIGHT_1 = 3200
Private Const M_LNG_FRMHEIGHT_2 = 6315


Private Sub Cbo����_Click()
    If Cbo����.ListCount < 1 Then Exit Sub
    Cbo����.Enabled = Not (Cbo����.ListIndex = 0)
    If Not Cbo����.Enabled Then Cbo����.ListIndex = 0
End Sub


Private Sub chkSavePrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1
End Sub

Private Sub chkVerifyPrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOK_Click()
    If mlngMode = 1343 Then
        If Trim(txt��ѯ����.Text) = "" Then
            MsgBox "�������ѯ������1��-365�죩��", vbInformation, gstrSysName
            txt��ѯ����.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt��ѯ����.Text) Then
            MsgBox "��ѯ�����к��зǷ��ַ���", vbInformation, gstrSysName
            txt��ѯ����.SetFocus
            Exit Sub
        End If
        If Val(txt��ѯ����.Text) < 1 Or Val(txt��ѯ����.Text) > 365 Then
            MsgBox "��ѯ��������С��1������365�죡", vbInformation, gstrSysName
            txt��ѯ����.SetFocus
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    
    Select Case mlngMode
        Case 1343   'ҩƷ����
'            zldatabase.SetPara "�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "ҩƷ��λ", Cboָ����λ.ListIndex, glngSys, mlngMode
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngMode
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "��ӡҩƷ����", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngMode
            zldatabase.SetPara "��ʾ�޿��ҩƷ", chkShow.Value, glngSys, mlngMode
            zldatabase.SetPara "�ʱ��ǰ�ⷿ�����ʾ��ʽ", IIf(opt��ǰ���(0).Value = True, 0, 1), glngSys, mlngMode
            zldatabase.SetPara "�ʱ�Է��ⷿ�����ʾ��ʽ", IIf(opt�Է����(0).Value = True, 0, 1), glngSys, mlngMode
            zldatabase.SetPara "��ǰ�ⷿҩƷ�����Ƿ�������ʾ", IIf(opt��������ʾ(0).Value = True, 0, 1), glngSys, mlngMode
        Case 1344   'Э�����
'            zldatabase.SetPara "�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "ҩƷ��λ", Cboָ����λ.ListIndex, glngSys, mlngMode
    End Select
    
    Unload Me
End Sub

Public Sub ���ò���(frmParent As Object, ByVal strPrivs As String, Optional ByVal intMode As Integer = 1344, Optional ByVal strFunction As String = "")
    mstrFunction = strFunction
    mlngMode = intMode
    mstrPrivs = strPrivs
    
    Dim int�Ƿ�ѡ��ⷿ As Integer
    Dim intҩƷ��λ As Integer
    Dim str���� As String
    Dim int���̴�ӡ As Integer
    Dim int��˴�ӡ As Integer
    Dim int��ӡҩƷ���� As Integer
    Dim int��ѯ���� As Integer
    Dim int��ʾ�޿��ҩƷ As Integer
    Dim int��ǰ�����ʾ��ʽ As Integer
    Dim int�Է������ʾ��ʽ As Integer
    Dim int��ǰ��水������ʾ As Integer
    
    mblnSetPara = IsHavePrivs(mstrPrivs, "��������")
    
    'ȡ������˽�в���
    Select Case mlngMode
        Case 1343   'ҩƷ����
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngMode, 0, Array(lblҩƷ��λ, Cboָ����λ), mblnSetPara))
            str���� = zldatabase.GetPara("����", glngSys, mlngMode, "00", Array(fraSort, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngMode, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngMode, 0, Array(chkVerifyPrint), mblnSetPara))
            int��ӡҩƷ���� = Val(zldatabase.GetPara("��ӡҩƷ����", glngSys, mlngMode, 0, Array(chkPrintCode), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngMode, 7, Array(lbl��ѯ����, txt��ѯ����, lbl����), mblnSetPara))
            int��ʾ�޿��ҩƷ = Val(zldatabase.GetPara("��ʾ�޿��ҩƷ", glngSys, mlngMode, 0, Array(fra�����ʾ����, chkShow), mblnSetPara))
            int��ǰ�����ʾ��ʽ = Val(zldatabase.GetPara("�ʱ��ǰ�ⷿ�����ʾ��ʽ", glngSys, mlngMode, 0, Array(opt��ǰ���(0), opt��ǰ���(1)), mblnSetPara))
            int�Է������ʾ��ʽ = Val(zldatabase.GetPara("�ʱ�Է��ⷿ�����ʾ��ʽ", glngSys, mlngMode, 0, Array(opt�Է����(0), opt�Է����(1)), mblnSetPara))
            int��ǰ��水������ʾ = Val(zldatabase.GetPara("��ǰ�ⷿҩƷ�����Ƿ�������ʾ", glngSys, mlngMode, 0, Array(opt��������ʾ(0), opt��������ʾ(1)), mblnSetPara))
        Case 1344   'Э�����
'            int�Ƿ�ѡ��ⷿ = Val(zldatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngMode, 0, Array(chkStock, Label2), mblnSetPara))
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngMode, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngMode, 0, Array(chkVerifyPrint), mblnSetPara))
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngMode, 0, Array(lblҩƷ��λ, Cboָ����λ), mblnSetPara))
    End Select
    
    '���ݲ���ֵ����
'    If int�Ƿ�ѡ��ⷿ = 0 Then
'        chkStock.Value = 0
'    Else
'        chkStock.Value = 1
'    End If
    If int���̴�ӡ = 0 Then
        chkSavePrint.Value = 0
    Else
        chkSavePrint.Value = 1
    End If
    
    If int��˴�ӡ = 0 Then
        chkVerifyPrint.Value = 0
    Else
        chkVerifyPrint.Value = 1
    End If
    
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1
    
    If int��ӡҩƷ���� = 0 Then
        chkPrintCode.Value = 0
    Else
        chkPrintCode.Value = 1
    End If
    
    With Cboָ����λ
        .Clear
        .AddItem "ȱʡ����ǰ�ⷿ��Ӧ�ĵ�λ��"
        If glngSys \ 100 = 8 Then
            .AddItem "�ɹ���λ"
            .AddItem "�ۼ۵�λ"
        Else
            .AddItem "ҩ�ⵥλ"
            .AddItem "���ﵥλ"
            .AddItem "סԺ��λ"
            .AddItem "�ۼ۵�λ"
        End If
        .ListIndex = intҩƷ��λ
    End With
    
    fra�����ʾ����.Visible = False
    
    Select Case mlngMode
        Case 1343   '����
            fra�����ʾ����.Visible = True
'            Frame3.Top = Frame2.Top
'            Frame2.Visible = True
'            chkVerifyPrint.Visible = False
'            Label3.Caption = Replace(Label3.Caption, "��˴�ӡ���ͬ��", "")
            lblҩƷ��λ.Visible = True
            Cboָ����λ.Visible = True
            
            fraSort.Visible = True
            Me.Width = M_LNG_FRMWIDTH_2
            Me.Height = M_LNG_FRMHEIGHT_2
            
            CmdCancel.Top = Me.Height - CmdCancel.Height - 500
            CmdOK.Top = CmdCancel.Top
            CmdHelp.Top = CmdCancel.Top
            
            CmdCancel.Left = M_LNG_FRMWIDTH_2 - CmdCancel.Width - 400
            CmdOK.Left = CmdCancel.Left - CmdOK.Width - 200
            
            Dim strValue As String
            mstrFunction = strFunction
            
            'װ��ȱʡ����
            With Cbo����
                .Clear
                .AddItem "����˳��"
                .ItemData(.NewIndex) = 0
                .AddItem "����"
                .ItemData(.NewIndex) = 1
                .AddItem "ҩƷ����"
                .ItemData(.NewIndex) = 2
                .AddItem "�ⷿ��λ"
                .ItemData(.NewIndex) = 3
                .ListIndex = 0
            End With
            With Cbo����
                .Clear
                .AddItem "����"
                .ItemData(.NewIndex) = 0
                .AddItem "����"
                .ItemData(.NewIndex) = 1
                .ListIndex = 0
            End With
            
            'ȡ�����ֶμ��������Ϊȱʡ������cbo����.Enabled=False
            strValue = str����
            Cbo����.ListIndex = Mid(strValue, 1, 1)
            Cbo����.ListIndex = Right(strValue, 1)
            Cbo����.Enabled = Not (Cbo����.ListIndex = 0)
            
            lbl��ѯ����.Visible = True
            txt��ѯ����.Visible = True
            lbl����.Visible = True
            
            txt��ѯ����.Text = int��ѯ����
            
            chkShow.Value = IIf(int��ʾ�޿��ҩƷ = 1, 1, 0)
            
            If int��ǰ�����ʾ��ʽ = 1 Then
                opt��ǰ���(1).Value = True
            Else
                opt��ǰ���(0).Value = True
            End If
            
            If int�Է������ʾ��ʽ = 1 Then
                opt�Է����(1).Value = True
            Else
                opt�Է����(0).Value = True
            End If
            
            If int��ǰ��水������ʾ = 1 Then
                opt��������ʾ(1).Value = True
            Else
                opt��������ʾ(0).Value = True
            End If
            
        Case 1344   'Э��
'            Frame3.Top = Frame2.Top + Frame2.Height + cmd��ӡ����.Height + 200
            cmd��ӡ����.Top = Cboָ����λ.Top + Cboָ����λ.Height + 200
             Frame3.Top = cmd��ӡ����.Top + cmd��ӡ����.Height - 300
'            Me.Height = 4000

            fraSort.Visible = False
            Me.Width = M_LNG_FRMWIDTH_1
            Me.Height = 3800
            CmdCancel.Top = Frame3.Top + Frame3.Height + 300
            CmdCancel.Left = M_LNG_FRMWIDTH_1 - CmdCancel.Width - 200
            CmdOK.Top = CmdCancel.Top
            CmdOK.Left = CmdCancel.Left - CmdOK.Width - 50
            CmdHelp.Top = CmdCancel.Top
    End Select
'    cmd��ӡ����.Top = IIf(mlngMode = 1343, cmd��ӡ����.Top, Cboָ����λ.Top)
    
    frm��������.Show vbModal, frmParent
End Sub

Private Sub cmd��ӡ����_Click()
    Dim strBill As String
    Select Case mstrFunction
    Case "ҩƷ�������"
        strBill = Split(cbo����.Text, "(")(0)
    Case "Э��ҩƷ���"
        strBill = "ZL1_BILL_1344"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Load()
    Me.cmd��ӡ����.Caption = "Ʊ�ݡ�" & Replace(mstrFunction, "����", "") & "������ӡ����"
    chkPrintCode.Visible = mlngMode = 1343
    '�����б�����
    cbo����.Visible = mlngMode = 1343
    cbo����.AddItem "ZL1_BILL_1304(���ݴ�ӡ)"
    cbo����.AddItem "ZL1_INSIDE_1343_1(ҩƷ�����ӡ)"
    cbo����.ListIndex = 0
End Sub

