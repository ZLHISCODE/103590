VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParaset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmParaset.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   25
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4200
      TabIndex        =   24
      Top             =   5760
      Width           =   1100
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frmParaset.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��������"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra���ĵ�λ"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra��ӡ����"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra����"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fra���� 
         Caption         =   " ��������"
         Height          =   4920
         Left            =   4080
         TabIndex        =   18
         Top             =   480
         Width           =   2400
         Begin VB.CheckBox chk�ƿ����� 
            Caption         =   "�ƿ����ñ��ϡ����͡����ջ���"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.Frame fra��ѯ���� 
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1955
            Begin VB.TextBox txt��ѯ���� 
               Height          =   300
               Left            =   840
               TabIndex        =   20
               Text            =   "7"
               Top             =   60
               Width           =   300
            End
            Begin MSComCtl2.UpDown upd��ѯ���� 
               Height          =   300
               Left            =   1080
               TabIndex        =   21
               Top             =   60
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txt��ѯ����"
               BuddyDispid     =   196614
               OrigLeft        =   1800
               OrigTop         =   360
               OrigRight       =   2055
               OrigBottom      =   735
               Max             =   90
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label lbl���� 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   1440
               TabIndex        =   23
               Top             =   120
               Width           =   180
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
               Left            =   0
               TabIndex        =   22
               Top             =   120
               Width           =   720
            End
         End
         Begin VB.Label lbl�ƿ����� 
            Caption         =   "ע�⣺������򹴣���ô����д�ƿⵥ������һ����˲�������˺��Զ���ɱ��ϡ����͡�������һ����"
            ForeColor       =   &H00000080&
            Height          =   900
            Left            =   120
            TabIndex        =   27
            Top             =   1095
            Visible         =   0   'False
            Width           =   2865
         End
      End
      Begin VB.Frame fra��ӡ���� 
         Caption         =   " ��ӡ����"
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   3675
         Begin VB.CheckBox chkVerifyPrint 
            Caption         =   "��˺��ӡ"
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkSavePrint 
            Caption         =   "���̺��ӡ"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "���ݴ�ӡ����(&S)"
            Height          =   350
            Left            =   360
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   600
            Width           =   2925
         End
      End
      Begin VB.Frame fra���ĵ�λ 
         Caption         =   " ���ĵ�λ"
         Height          =   1665
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3675
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   390
            Width           =   2655
         End
         Begin VB.ComboBox CboUnit1 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   780
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "ע����ѡ��һ�����ĵ�λ���������Ľ�ʹ�øõ�λ���а�װ��ʾ�Ͱ�װ����"
            ForeColor       =   &H00000080&
            Height          =   405
            Left            =   120
            TabIndex        =   11
            Top             =   1170
            Width           =   3315
         End
         Begin VB.Label lbl�̵�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�̵��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   10
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl�̵㵥 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�̵㵥"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   " ����ʽ"
         Height          =   1785
         Left            =   120
         TabIndex        =   2
         Top             =   2250
         Width           =   3675
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   390
            Width           =   2415
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   2580
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   885
         End
         Begin VB.Label lbl����˵�� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ע�������������ã���Ӱ�����б༭�����е��ݵ���ʾ���ݵ�����ʽ��ȱʡ�����û������˳����ʾ�����ݵ�����"
            ForeColor       =   &H00000080&
            Height          =   600
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   3345
         End
      End
      Begin VB.Frame fra�������� 
         Caption         =   " ��������"
         Height          =   1785
         Left            =   120
         TabIndex        =   16
         Top             =   2250
         Width           =   3675
         Begin VB.CheckBox chk����˲� 
            Caption         =   "������Ҫ�˲������ƿ�"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   3105
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   360
      TabIndex        =   0
      Top             =   5760
      Width           =   1100
   End
End
Attribute VB_Name = "frmParaset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Private mlngModule As Long '
Private mstrPrivs As String '
Private mblnHavePriv As Boolean
Private mblnFirstLoad As Boolean    '��¼�Ƿ��һ�μ���
Private mfrmMain As Object '������


Private Sub Cbo����_Click()
    If cbo����.ListCount < 1 Then Exit Sub
    cbo����.Enabled = Not (cbo����.ListIndex = 0)
    If Not cbo����.Enabled Then cbo����.ListIndex = 0
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
    If txt��ѯ����.Text > 7 Then
        If MsgBox("��ѯ��������7���ˣ����ܽ�����ҳ���������Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    gcnOracle.BeginTrans
    
    Call zlDatabase.SetPara(IIf(mlngModule = 1719, "�̵��λ", "���ĵ�λ"), cboUnit.ListIndex, glngSys, mlngModule)
    If CboUnit1.Visible Then
        Call zlDatabase.SetPara("��¼����λ", CboUnit1.ListIndex, glngSys, mlngModule)
    End If
    
    '����û�е����������
    If mlngModule <> 1722 Then
        Call zlDatabase.SetPara("��������", CStr(cbo����.ListIndex) & CStr(cbo����.ListIndex), glngSys, mlngModule)
    End If
    
    Call zlDatabase.SetPara("���̴�ӡ", IIf(chkSavePrint.Value = 1, 1, 0), glngSys, mlngModule)
    
    '����û����˴�ӡ����
    If mlngModule <> 1722 Then
        Call zlDatabase.SetPara("��˴�ӡ", IIf(chkVerifyPrint.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    
    '��������������Ĳ���
    If mlngModule = 1722 Then
        Call zlDatabase.SetPara("������Ҫ�˲������ƿ�", IIf(chk����˲�.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModule
    
    '�����ƿ�������Ĳ���
    If mlngModule = 1716 Then
        Call zlDatabase.SetPara("�ƿ�����", IIf(chk�ƿ�����.Value = 1, 1, 0), glngSys, mlngModule, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    End If
    
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
    '-------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:
    '����:���˺�
    '�޸�:2007/12/24
    '-------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim strBidMess As String
    Dim int��ѯ���� As Integer
    
    'װ��ȱʡ����
    With cbo����
        .Clear
        .AddItem "����˳��"
        .ItemData(.NewIndex) = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 1
        .AddItem "��������"
        .ItemData(.NewIndex) = 2
        If mstrFunction = "�����̵����" Then
            .AddItem "�ⷿ��λ"
            .ItemData(.NewIndex) = 3
        End If
        .ListIndex = 0
    End With
    
    With cbo����
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    If mlngModule <> 1722 Then
        strValue = zlDatabase.GetPara("��������", glngSys, mlngModule, "00", Array(cbo����, cbo����, fra����, lbl����˵��), mblnHavePriv)
        strValue = IIf(strValue = "", "00", strValue)
        cbo����.ListIndex = Val(Mid(strValue, 1, 1))
        cbo����.ListIndex = Val(Right(strValue, 1))
        cbo����.Enabled = Not (cbo����.ListIndex = 0)
    End If
    
    chkSavePrint.Value = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0", Array(chkSavePrint), mblnHavePriv)) = 1, 1, 0)
    chkVerifyPrint.Value = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0", Array(chkVerifyPrint), mblnHavePriv)) = 1, 1, 0)
    
    With CboUnit1
        .Clear
        .AddItem "ɢװ��λ"
        .AddItem "��װ��λ"
    End With

    With cboUnit
        .Clear
        .AddItem "ɢװ��λ"
        .AddItem "��װ��λ"
    End With
    cboUnit.ListIndex = IIf(Val(zlDatabase.GetPara(IIf(mlngModule = 1719, "�̵��λ", "���ĵ�λ"), glngSys, mlngModule, "0", Array(cboUnit, lbl�̵��), mblnHavePriv)) = 1, 1, 0)
    If mstrFunction <> "�����̵����" Then
        CboUnit1.Visible = False
        lbl�̵��.Visible = False
        lbl�̵㵥.Visible = False
        cboUnit.Left = lbl�̵��.Left
        Label2.Top = lbl�̵㵥.Top
    Else
        CboUnit1.ListIndex = IIf(Val(zlDatabase.GetPara("��¼����λ", glngSys, mlngModule, "0", Array(CboUnit1, lbl�̵㵥), mblnHavePriv)) = 1, 1, 0)
    End If
    
    int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModule, 1))
    txt��ѯ����.Text = int��ѯ����
    
    fra��������.Visible = False
    Select Case mstrFunction
        Case "�����ƿ����"
            Me.Width = Me.Width + 700 '�ı���
            tabMain.Width = tabMain.Width + 700
            fra����.Width = fra����.Width + 700
            
            cmdOK.Left = cmdOK.Left + 700
            cmdCancel.Left = cmdCancel.Left + 700
            
            '���ÿɼ�
            chk�ƿ�����.Visible = True
            lbl�ƿ�����.Visible = True
            
            chk�ƿ�����.Value = Val(zlDatabase.GetPara("�ƿ�����", glngSys, mlngModule, "0", , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)))
            
        Case "�����̵����"

        Case "�����⹺������"

        Case "���ļƻ�����"
        
        Case "�����깺����"
            
        Case "�������ù���"
        
        Case "�����������"
            fra����.Visible = False
            fra��������.Visible = True
            fra��ӡ����.Top = fra����.Top
            fra��������.Top = fra��ӡ����.Top + fra��ӡ����.Height + 150
            chkVerifyPrint.Visible = False
            chk����˲�.Value = IIf((zlDatabase.GetPara("������Ҫ�˲������ƿ�", glngSys, mlngModule, "0")) = 0, 0, 1)
        Case "���ĵ��۹���"
            fra����.Visible = False
            fra��ӡ����.Visible = False
            
            fra����.Height = fra���ĵ�λ.Height
            tabMain.Height = fra����.Top + fra����.Height + 200
            
            Me.Height = tabMain.Top + tabMain.Height + cmdHelp.Height + 650
            cmdHelp.Top = tabMain.Top + tabMain.Height + 100
            cmdCancel.Top = cmdHelp.Top
            cmdCancel.Left = Me.Width - cmdCancel.Width - 200
            cmdOK.Top = cmdHelp.Top
            cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
        Case Else
    End Select

    Me.cmdPrintSet.Enabled = InStr(1, gstrPrivs, ";���ݴ�ӡ;") <> 0

End Sub

Public Sub ���ò���(ByVal lngModule As Long, ByVal strPrivs As String, ByVal frmMain As Form, Optional ByVal strFunction As String = "")
    '-------------------------------------------------------------------------------------------------------------
    '����:������ص��ݲ����Ŀ��Ʋ���
    '����:lngModule-ģ���
    '     strȨ�޴�-Ȩ�޴�
    '     frmMain-���õ�������
    '     strFunction-����˵��
    '����:
    '����:���˺�
    '�޸�:2007/12/24
    '-------------------------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs: mlngModule = lngModule: mstrFunction = strFunction
    mblnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "��������")
    Set mfrmMain = frmMain
    
    Call initPara
    frmParaset.Show vbModal, frmMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_" & glngModul
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFirstLoad = False
End Sub

