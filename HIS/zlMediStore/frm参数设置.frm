VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8700
   Icon            =   "frm��������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleMode       =   0  'User
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabMain 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frm��������.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra����ʽ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraҩƷ��λ"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra��ӡ����"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame fra��ӡ���� 
         Caption         =   " ��ӡ����"
         Height          =   1980
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   4000
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   480
            TabIndex        =   49
            Text            =   "Combo1"
            Top             =   1200
            Width           =   3135
         End
         Begin VB.CheckBox chkPrintCode 
            Caption         =   "���̻���˺��ӡҩƷ����"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   3015
         End
         Begin VB.CommandButton cmd��ӡ���� 
            Caption         =   "��ӡ����(&P)"
            Height          =   315
            Left            =   480
            TabIndex        =   24
            Top             =   1560
            Width           =   3135
         End
         Begin VB.CheckBox chkSendPrint 
            Caption         =   "���ͺ��ӡ����"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox chkSavePrint 
            Caption         =   "���̺��ӡ����"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1635
         End
         Begin VB.CheckBox chkVerifyPrint 
            Caption         =   "��˺��ӡ����"
            Height          =   255
            Left            =   2040
            TabIndex        =   21
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame fraҩƷ��λ 
         Caption         =   " ҩƷ��λ"
         Height          =   1785
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   4000
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   390
            Width           =   2655
         End
         Begin VB.ComboBox CboUnit1 
            Height          =   300
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   780
            Width           =   2655
         End
         Begin VB.Label lblUnitComment 
            Caption         =   "ע����ѡ��һ��ҩƷ��λ������ҩƷ��ʹ�øõ�λ���а�װ��ʾ�Ͱ�װ����"
            ForeColor       =   &H00000080&
            Height          =   405
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   3315
         End
         Begin VB.Label lbl�̵�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���װ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   17
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl�̵㵥 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "С��װ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   " ��������"
         Height          =   5355
         Left            =   4200
         TabIndex        =   8
         Top             =   480
         Width           =   4200
         Begin VB.Frame frm�Է���� 
            Caption         =   "�ʱ�Է��ⷿ�����ʾ��ʽ"
            Height          =   735
            Left            =   120
            TabIndex        =   45
            Top             =   1920
            Width           =   3975
            Begin VB.OptionButton opt�Է���� 
               Caption         =   "��ʾ��������"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   47
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton opt�Է���� 
               Caption         =   "��ʾʵ������"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   46
               Top             =   360
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.Frame frm��ǰ��� 
            Caption         =   "�ʱ��ǰ�ⷿ�����ʾ��ʽ"
            Height          =   735
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   3975
            Begin VB.OptionButton opt��ǰ��� 
               Caption         =   "��ʾ��������"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   44
               Top             =   360
               Width           =   1455
            End
            Begin VB.OptionButton opt��ǰ��� 
               Caption         =   "��ʾʵ������"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   43
               Top             =   360
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.CheckBox chkALLPlanPoint 
            Caption         =   "ȫԺ�ƻ�����վ��"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   4200
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Frame fraҩƷ�ƻ���Ӧ������ 
            Caption         =   " ҩƷ�ɹ��ƻ���Ӧ������"
            Height          =   1725
            Left            =   120
            TabIndex        =   35
            Top             =   2400
            Width           =   3960
            Begin VB.ComboBox cbo��Ӧ�̷�Χ 
               Height          =   300
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   660
               Width           =   2700
            End
            Begin VB.ComboBox cbo��Ӧ��ѡ�� 
               Height          =   300
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   300
               Width           =   2700
            End
            Begin VB.Label Label3 
               Caption         =   "ע��ҩƷ�ɹ��ƻ��༭������ҩƷ��Ӧ�̵�Ĭ�ϴ����Լ��ֹ�ѡ��Ӧ��ʱ�Ŀ�ѡ��Χ"
               ForeColor       =   &H00000080&
               Height          =   495
               Left            =   120
               TabIndex        =   40
               Top             =   1080
               Width           =   3765
            End
            Begin VB.Label lbl��Ӧ�̷�Χ 
               AutoSize        =   -1  'True
               Caption         =   "ѡ��Χ"
               Height          =   180
               Left            =   120
               TabIndex        =   39
               Top             =   720
               Width           =   720
            End
            Begin VB.Label lbl��Ӧ��ѡ�� 
               AutoSize        =   -1  'True
               Caption         =   "Ĭ��ѡ��"
               Height          =   180
               Left            =   120
               TabIndex        =   38
               Top             =   360
               Width           =   720
            End
         End
         Begin VB.Frame fra��ѯ���� 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2175
            Begin VB.ComboBox cboDay 
               Height          =   300
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   60
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txt��ѯ���� 
               Height          =   300
               Left            =   840
               TabIndex        =   30
               Text            =   "7"
               Top             =   60
               Width           =   300
            End
            Begin MSComCtl2.UpDown upd��ѯ���� 
               Height          =   300
               Left            =   1140
               TabIndex        =   31
               Top             =   60
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txt��ѯ����"
               BuddyDispid     =   196636
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
               TabIndex        =   33
               Top             =   120
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
               Left            =   1440
               TabIndex        =   32
               Top             =   120
               Width           =   180
            End
         End
         Begin VB.Frame fra�̵�ʱ�䷶Χ 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Visible         =   0   'False
            Width           =   1700
            Begin VB.TextBox txt�̵�ʱ�� 
               Height          =   300
               Left            =   840
               TabIndex        =   26
               Text            =   "3"
               Top             =   60
               Width           =   300
            End
            Begin MSComCtl2.UpDown UpD�̵�ʱ�� 
               Height          =   300
               Left            =   1140
               TabIndex        =   27
               Top             =   60
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               BuddyControl    =   "txt�̵�ʱ��"
               BuddyDispid     =   196641
               OrigLeft        =   1800
               OrigTop         =   360
               OrigRight       =   2055
               OrigBottom      =   735
               Max             =   90
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "�̵�ʱ��"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   0
               TabIndex        =   34
               Top             =   120
               Width           =   720
            End
            Begin VB.Label lblday 
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               Height          =   195
               Left            =   1440
               TabIndex        =   28
               Top             =   120
               Width           =   255
            End
         End
         Begin VB.CheckBox chk�������� 
            Caption         =   "������������"
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Frame fraҩƷ�ƻ��۸���ʾ��ʽ 
            Caption         =   " ҩƷ�ɹ��ƻ��۸���ʾ��ʽ"
            Height          =   735
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Visible         =   0   'False
            Width           =   3960
            Begin VB.OptionButton Opt��� 
               Caption         =   "�ɱ��ۺ��ۼ�"
               Height          =   180
               Left            =   2160
               TabIndex        =   11
               Top             =   375
               Width           =   1400
            End
            Begin VB.OptionButton Opt�ɱ��� 
               Caption         =   "�ɱ���"
               Height          =   180
               Left            =   120
               TabIndex        =   12
               Top             =   375
               Width           =   900
            End
            Begin VB.OptionButton Opt�ۼ� 
               Caption         =   "�ۼ�"
               Height          =   180
               Left            =   1200
               TabIndex        =   10
               Top             =   375
               Width           =   720
            End
         End
      End
      Begin VB.Frame fra����ʽ 
         Caption         =   " ����ʽ"
         Height          =   1515
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   4000
         Begin VB.ComboBox Cbo���� 
            Height          =   300
            ItemData        =   "frm��������.frx":0028
            Left            =   120
            List            =   "frm��������.frx":002A
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   390
            Width           =   2415
         End
         Begin VB.ComboBox Cbo���� 
            Height          =   300
            Left            =   2580
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   390
            Width           =   885
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ע�������������ã���Ӱ�����б༭�����е��ݵ���ʾ���ݵ�����ʽ��ȱʡ�����û������˳����ʾ�����ݵ�����"
            ForeColor       =   &H00000080&
            Height          =   675
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   3345
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   2
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6000
      TabIndex        =   1
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   240
      TabIndex        =   0
      Top             =   6240
      Width           =   1100
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Dim mstrPrivs As String
Dim mlngModul As Long
Dim mblnSetPara As Boolean      '�Ƿ���в�������Ȩ��
Private mint�̵�ʱ�� As Integer  '������¼���õ��̵�ʱ�䷶Χ

Private Sub Cbo����_Click()
    If Cbo����.ListCount < 1 Then Exit Sub
    Cbo����.Enabled = Not (Cbo����.ListIndex = 0)
    If Not Cbo����.Enabled Then Cbo����.ListIndex = 0
End Sub

Private Sub chkSavePrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1 Or chkSendPrint.Value = 1
End Sub

Private Sub chkSendPrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1 Or chkSendPrint.Value = 1
End Sub

Private Sub chkVerifyPrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1 Or chkSendPrint.Value = 1
End Sub

Private Sub chk��������_Click()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    
    If chk��������.Value = 0 Then
        gstrSQL = "Select �ڼ� From ҩƷ���� Where Length(�ڼ�) > 4"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rsTemp.RecordCount > 0 Then
            MsgBox "��������ģʽ���Ѿ��������ݣ������޸ģ�", vbInformation, gstrSysName
            chk��������.Value = 1
        End If
    End If
    Exit Sub
errH:
If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    If ISValid = False Then Exit Sub
    
    Select Case mlngModul
        Case 1300   'ҩƷ�⹺������
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��ӡҩƷ����", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1301   'ҩƷ����������
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1302   'ҩƷ����������
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��ӡҩƷ����", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1303   'ҩƷ����۵�������
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1304   'ҩƷ�ƿ����
            zldatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "���ʹ�ӡ", IIf(chkSendPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��ӡҩƷ����", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
            zldatabase.SetPara "�ʱ��ǰ�ⷿ�����ʾ��ʽ", IIf(opt��ǰ���(0).Value = True, 0, 1), glngSys, mlngModul
            zldatabase.SetPara "�ʱ�Է��ⷿ�����ʾ��ʽ", IIf(opt�Է����(0).Value = True, 0, 1), glngSys, mlngModul
        Case 1305   'ҩƷ���ù���
            zldatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��ӡҩƷ����", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "������������", IIf(chk��������.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1306   'ҩƷ�����������
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��ӡҩƷ����", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1307   'ҩƷ�̵����
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "С��װ��λ", CboUnit1.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��ѯ����", Val(cboDay.ItemData(cboDay.ListIndex)), glngSys, mlngModul

            zldatabase.SetPara "�̵�ʱ�䷶Χ����", txt�̵�ʱ��.Text, glngSys, mlngModul
        Case 1330   'ҩƷ�ƻ�����
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "�۸���ʾ��ʽ", IIf(Opt�ɱ���.Value = True, "0", IIf(Opt�ۼ�.Value = True, "1", "2")), glngSys, mlngModul
            zldatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "��Ӧ��Ĭ��ѡ��", cbo��Ӧ��ѡ��.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "��Ӧ��ѡ��Χ", cbo��Ӧ�̷�Χ.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
            zldatabase.SetPara "ȫԺ�ƻ�����վ��", IIf(chkALLPlanPoint.Value = 1, "1", "0"), glngSys, mlngModul
        Case 1331   'ҩƷ��������
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1333 'ҩƷ���۹���
            zldatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
    End Select
           
    Unload Me
End Sub

Private Function ISValid() As Boolean
    Dim i As Integer
    
    If Val(txt��ѯ����.Text) > 7 Then
        If MsgBox("��ѯʱ�����7����ܻᵼ�²�ѯ�������Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txt��ѯ����.SetFocus
            zlControl.TxtSelAll txt��ѯ����
            Exit Function
        End If
    End If
    If Val(txt��ѯ����.Text) = 0 Then
        MsgBox "��ѯʱ��������0�����������룡", vbInformation, gstrSysName
        txt��ѯ����.SetFocus
        zlControl.TxtSelAll txt��ѯ����
        Exit Function
    End If
    
    ISValid = True
End Function

Private Sub loadCboDay()
    With cboDay
        .AddItem "��ʾ����"
        .ItemData(.NewIndex) = 1
        .AddItem "��ʾ7��֮��"
        .ItemData(.NewIndex) = 7
        
        .Visible = True
        lbl��ѯ����.Caption = "��ѯ��Χ"
        txt��ѯ����.Visible = False
    End With
End Sub

Public Sub ���ò���(frmParent As Object, ByVal strPrivs As String, Optional ByVal strFunction As String = "")
    mstrFunction = strFunction
    mstrPrivs = strPrivs
    mlngModul = glngModul
    Dim str���ݴ�ӡ As String
    
    'ͨ�ã�˽��ģ�飩
    Dim str���� As String
    Dim int���̴�ӡ As Integer
    Dim int��˴�ӡ As Integer
    Dim int��ӡҩƷ���� As Integer
        
    '������Ҫ��ͨģ�飨˽��ģ�飩
    Dim intҩƷ��λ As Integer
    Dim int�ɱ�����Դ As Integer
    Dim int��ѯ���� As Integer
        
    '�����̵㣨˽��ģ�飩
    Dim intС��װ��λ As Integer
        
    '����ҩƷ�ƻ���˽��ģ�飩
    Dim int�۸���ʾ��ʽ As Integer
    Dim int��Ӧ��ѡ�� As Integer
    Dim int��Ӧ�̷�Χ As Integer
    Dim intPlanPoint As Integer
    
    '�����ƿ�(˽��)
    Dim int���ʹ�ӡ As Integer
    Dim int��ǰ��� As Integer
    Dim int�Է���� As Integer
    
    '��������
    Dim int�������� As Integer
    
    Dim i As Integer
    
    On Error Resume Next
    
    mblnSetPara = zlStr.IsHavePrivs(mstrPrivs, "��������")
    
    'ȡ����ֵ
    Select Case mlngModul
        Case 1300   'ҩƷ�⹺������
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int��ӡҩƷ���� = Val(zldatabase.GetPara("��ӡҩƷ����", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1301   'ҩƷ����������
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1302   'ҩƷ����������
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int��ӡҩƷ���� = Val(zldatabase.GetPara("��ӡҩƷ����", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1303   'ҩƷ����۵�������
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1304   'ҩƷ�ƿ����
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int���ʹ�ӡ = Val(zldatabase.GetPara("���ʹ�ӡ", glngSys, mlngModul, 0, Array(chkSendPrint), mblnSetPara))
            int��ӡҩƷ���� = Val(zldatabase.GetPara("��ӡҩƷ����", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
            int��ǰ��� = Val(zldatabase.GetPara("�ʱ��ǰ�ⷿ�����ʾ��ʽ", glngSys, mlngModul, 0, Array(opt��ǰ���(0), opt��ǰ���(1)), mblnSetPara))
            int�Է���� = Val(zldatabase.GetPara("�ʱ�Է��ⷿ�����ʾ��ʽ", glngSys, mlngModul, 0, Array(opt�Է����(0), opt�Է����(1)), mblnSetPara))
        Case 1305   'ҩƷ���ù���
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int��ӡҩƷ���� = Val(zldatabase.GetPara("��ӡҩƷ����", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            int�������� = Val(zldatabase.GetPara("������������", glngSys, mlngModul, 0, Array(chk��������), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1306   'ҩƷ�����������
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int��ӡҩƷ���� = Val(zldatabase.GetPara("��ӡҩƷ����", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1307   'ҩƷ�̵����
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            intС��װ��λ = Val(zldatabase.GetPara("С��װ��λ", glngSys, mlngModul, 0, Array(lbl�̵㵥, CboUnit1), mblnSetPara))
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
                        
            mint�̵�ʱ�� = Val(zldatabase.GetPara("�̵�ʱ�䷶Χ����", glngSys, mlngModul, 30))
            txt�̵�ʱ��.Text = mint�̵�ʱ��
            UpD�̵�ʱ��.Value = mint�̵�ʱ��
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1330   'ҩƷ�ƻ�����
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            int�۸���ʾ��ʽ = Val(zldatabase.GetPara("�۸���ʾ��ʽ", glngSys, mlngModul, 1, Array(fraҩƷ�ƻ��۸���ʾ��ʽ, Opt�ɱ���, Opt�ۼ�, Opt���), mblnSetPara))
            int���̴�ӡ = Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int��Ӧ��ѡ�� = Val(zldatabase.GetPara("��Ӧ��Ĭ��ѡ��", glngSys, mlngModul, 0, Array(cbo��Ӧ��ѡ��), mblnSetPara))
            int��Ӧ�̷�Χ = Val(zldatabase.GetPara("��Ӧ��ѡ��Χ", glngSys, mlngModul, 0, Array(cbo��Ӧ�̷�Χ), mblnSetPara))
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
            intPlanPoint = Val(zldatabase.GetPara("ȫԺ�ƻ�����վ��", glngSys, mlngModul, 0, Array(chkALLPlanPoint), mblnSetPara))
            chkALLPlanPoint.Value = intPlanPoint
        Case 1331  'ҩƷ��������
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1333 'ҩƷ���۹���
            str���� = zldatabase.GetPara("����", glngSys, mlngModul, "00", Array(fra����ʽ, Cbo����, Cbo����, Label5), mblnSetPara)
            intҩƷ��λ = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
    End Select
    
    If mlngModul <> 1307 Then
        txt��ѯ����.Text = int��ѯ����
    Else '�̵�
        loadCboDay
        int��ѯ���� = IIf(int��ѯ���� <> 1 And int��ѯ���� <> 7, 7, int��ѯ����)
        For i = 0 To cboDay.ListCount - 1
            If int��ѯ���� = cboDay.ItemData(i) Then cboDay.ListIndex = i
        Next
    End If
    
    If strFunction = "ҩƷ�ƻ�����" Then
        str���ݴ�ӡ = "�ɹ��ƻ���ӡ"
    Else
        str���ݴ�ӡ = "���ݴ�ӡ"
    End If
    
    'װ��ȱʡ����
    With Cbo����
        .Clear
        .AddItem "����˳��"
        .ItemData(.NewIndex) = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 1
        .AddItem "ҩƷ����"
        .ItemData(.NewIndex) = 2
        
        If InStr("ҩƷ�̵����/ҩƷ�ƿ����/ҩƷ���ù���/ҩƷ�����������", strFunction) > 0 Then
            .AddItem "�ⷿ��λ"
            .ItemData(.NewIndex) = 3
        End If
     
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
    Cbo����.ListIndex = Mid(str����, 1, 1)
    Cbo����.ListIndex = Right(str����, 1)
    Cbo����.Enabled = Not (Cbo����.ListIndex = 0)
    
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
    
    If int��ӡҩƷ���� = 0 Then
        chkPrintCode.Value = 0
    Else
        chkPrintCode.Value = 1
    End If
    
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1
    
    If int�������� = 0 Then
        chk��������.Value = 0
    Else
        chk��������.Value = 1
    End If

    If mstrFunction = "ҩƷ�̵����" Then
        If glngSys \ 100 = 8 Then
            With CboUnit1
                .AddItem "�ɹ���λ"
                .AddItem "�ۼ۵�λ"
            End With
        Else
            With CboUnit1
                .AddItem "�ʹ��װ��ͬ"
                .AddItem "ҩ�ⵥλ"
                .AddItem "���ﵥλ"
                .AddItem "סԺ��λ"
                .AddItem "�ۼ۵�λ"
            End With
        End If
        CboUnit1.ListIndex = intС��װ��λ
        lblUnitComment.Caption = "    ��ѡ���̵�ʱ�Ĵ�С��װ���̵㵥���̵��༭ʱ����ѡ��װ�����̵㡣"
    Else
        CboUnit1.Visible = False
        lbl�̵��.Visible = False
        lbl�̵㵥.Visible = False
        cboUnit.Left = lbl�̵��.Left
    End If
    
    With cboUnit
        .Clear
        If glngSys \ 100 = 8 Then
            .AddItem "ȱʡ����ǰ�ⷿ��Ӧ�ĵ�λ��"
            .AddItem "�ɹ���λ"
            .AddItem "�ۼ۵�λ"
        Else
            If mlngModul <> 1333 Then   '���۲���Ҫ�ⷿ
                .AddItem "ȱʡ����ǰ�ⷿ��Ӧ�ĵ�λ��"
            End If
            .AddItem "ҩ�ⵥλ"
            .AddItem "���ﵥλ"
            .AddItem "סԺ��λ"
            .AddItem "�ۼ۵�λ"
        End If
        .ListIndex = intҩƷ��λ
    End With
    
    '�������������ģ����ʾ�����ز�ͬ��ģ���������
    chkSendPrint.Visible = False
    If strFunction = "ҩƷ�ƿ����" Then
        chkSendPrint.Value = IIf(int���ʹ�ӡ = 1, 1, 0)
        chkSendPrint.Visible = True
        
        chkPrintCode.Enabled = chkPrintCode.Enabled Or chkSendPrint.Value = 1
        
        If int��ǰ��� = 1 Then
            opt��ǰ���(1).Value = True
        Else
            opt��ǰ���(0).Value = True
        End If
        
        If int�Է���� = 1 Then
            opt�Է����(1).Value = True
        Else
            opt�Է����(0).Value = True
        End If
    Else
        frm��ǰ���.Visible = False
        frm�Է����.Visible = False
    End If
    
    fraҩƷ�ƻ��۸���ʾ��ʽ.Visible = False
    fraҩƷ�ƻ���Ӧ������.Visible = False
    chkALLPlanPoint.Visible = False
    If strFunction = "ҩƷ�ƻ�����" Then
        If int�۸���ʾ��ʽ = 0 Then
            Opt�ɱ���.Value = True
        ElseIf int�۸���ʾ��ʽ = 1 Then
            Opt�ۼ�.Value = True
        Else
            Opt���.Value = True
        End If
        
        chkALLPlanPoint.Visible = True
        cbo��Ӧ��ѡ��.Clear
        cbo��Ӧ��ѡ��.AddItem "1-ȡ�ϴ���⹩Ӧ��"
        cbo��Ӧ��ѡ��.AddItem "2-ȡ��ͬ��λ"
        cbo��Ӧ��ѡ��.ListIndex = IIf(int��Ӧ��ѡ�� < 0 Or int��Ӧ��ѡ�� > 1, 0, int��Ӧ��ѡ��)
        
        cbo��Ӧ�̷�Χ.Clear
        cbo��Ӧ�̷�Χ.AddItem "1-���й�Ӧ��"
        cbo��Ӧ�̷�Χ.AddItem "2-�б굥λ"
        cbo��Ӧ�̷�Χ.ListIndex = IIf(int��Ӧ�̷�Χ < 0 Or int��Ӧ�̷�Χ > 1, 0, int��Ӧ�̷�Χ)
        
        fraҩƷ�ƻ��۸���ʾ��ʽ.Visible = True
        fraҩƷ�ƻ���Ӧ������.Visible = True
        
        fraҩƷ�ƻ��۸���ʾ��ʽ.Top = fra��ѯ����.Top + fra��ѯ����.Height + 100
        fraҩƷ�ƻ��۸���ʾ��ʽ.Left = fra��ѯ����.Left
        
        fraҩƷ�ƻ���Ӧ������.Top = fraҩƷ�ƻ��۸���ʾ��ʽ.Top + fraҩƷ�ƻ��۸���ʾ��ʽ.Height + 150
        fraҩƷ�ƻ���Ӧ������.Left = fra��ѯ����.Left
        
        chkALLPlanPoint.Top = fraҩƷ�ƻ���Ӧ������.Top + fraҩƷ�ƻ���Ӧ������.Height + 150
        chkALLPlanPoint.Left = fraҩƷ�ƻ���Ӧ������.Left
    End If

    fra�̵�ʱ�䷶Χ.Visible = False
    If strFunction = "ҩƷ�̵����" Then
        cboUnit.Enabled = False
        fra�̵�ʱ�䷶Χ.Visible = True
    End If
    
    If strFunction = "ҩƷ����������" Then

    End If
    
    If strFunction = "ҩƷ���۹���" Then
        fra����ʽ.Visible = False
        fra��ӡ����.Visible = False
        fraҩƷ�ƻ��۸���ʾ��ʽ.Visible = False
        fraҩƷ�ƻ���Ӧ������.Visible = False
        chk��������.Visible = False
        fra�̵�ʱ�䷶Χ.Visible = False
        chkALLPlanPoint.Visible = False
        
        fra����.Height = fraҩƷ��λ.Height
        
        tabMain.Height = fra����.Top + fra����.Height + 200
        tabMain.Width = fra����.Left + fra����.Width + 200
        
        Me.Height = tabMain.Top + tabMain.Height + cmdHelp.Height + 650
        Me.Width = tabMain.Left + tabMain.Width + 200
        
        cmdHelp.Top = tabMain.Top + tabMain.Height + 100
        CmdCancel.Top = cmdHelp.Top
        CmdCancel.Left = Me.Width - CmdCancel.Width - 200
        cmdOK.Top = cmdHelp.Top
        cmdOK.Left = CmdCancel.Left - cmdOK.Width - 50
    End If
    
    If strFunction = "ҩƷ��������" Then
        fraҩƷ��λ.Visible = False
        fra����ʽ.Visible = False
        fra��ӡ����.Visible = False
        fraҩƷ�ƻ��۸���ʾ��ʽ.Visible = False
        fraҩƷ�ƻ���Ӧ������.Visible = False
        chk��������.Visible = False
        fra�̵�ʱ�䷶Χ.Visible = False
        chkALLPlanPoint.Visible = False
        
        fra����.Move fraҩƷ��λ.Left, fraҩƷ��λ.Top, fraҩƷ��λ.Width, fraҩƷ��λ.Height
        
        tabMain.Height = fra����.Top + fra����.Height + 200
        tabMain.Width = fra����.Left + fra����.Width + 200
        
        Me.Height = tabMain.Top + tabMain.Height + cmdHelp.Height + 650
        Me.Width = tabMain.Left + tabMain.Width + 200
        
        cmdHelp.Top = tabMain.Top + tabMain.Height + 100
        CmdCancel.Top = cmdHelp.Top
        CmdCancel.Left = Me.Width - CmdCancel.Width - 200
        cmdOK.Top = cmdHelp.Top
        cmdOK.Left = CmdCancel.Left - cmdOK.Width - 50

    End If
    
    
    chk��������.Visible = False
    If strFunction = "ҩƷ���ù���" Then
        chk��������.Visible = True
    End If
    
    If mlngModul = 1302 Or mlngModul = 1303 Or mlngModul = 1306 Then
        '1302 :�������;1303:�����; 1306����������
    End If
    
    frm��������.Show vbModal, frmParent
End Sub
Private Sub cmd��ӡ����_Click()
    Dim strBill As String
    Select Case mstrFunction
    Case "ҩƷ�⹺������"
        strBill = Split(cbo����.Text, "(")(0)
    Case "ҩƷ����������"
        strBill = Split(cbo����.Text, "(")(0)
    Case "ҩƷ����������"
        strBill = "ZL1_BILL_1301"
    Case "����۵�������"
        strBill = "ZL1_BILL_1303"
    Case "ҩƷ�ƿ����"
        strBill = Split(cbo����.Text, "(")(0)
    Case "ҩƷ���ù���"
        strBill = Split(cbo����.Text, "(")(0)
    Case "ҩƷ�����������"
        strBill = Split(cbo����.Text, "(")(0)
    Case "ҩƷ�̵����"
        strBill = "ZL1_BILL_1307"
    Case "ҩƷ�ƻ�����"
        strBill = "zl1_bill_1330"
    Case "ҩƷ���۹���"
        strBill = "ZL1_BILL_1333"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.cmd��ӡ����.Caption = "Ʊ�ݡ�" & Mid(mstrFunction, 1, Len(mstrFunction) - 2) & "������ӡ����"
    
    '�������ʱ�Ĳ���״̬
    fra��ѯ����.BackColor = &H8000000F
    fra�̵�ʱ�䷶Χ.BackColor = &H8000000F
    
    chkPrintCode.Visible = True
    cbo����.Visible = True
    Select Case mlngModul
        Case 1300
            '�����б�����
            cbo����.AddItem "ZL1_BILL_1300(���ݴ�ӡ)"
            cbo����.AddItem "ZL1_INSIDE_1300_1(ҩƷ�����ӡ)"
            cbo����.ListIndex = 0
        Case 1302
            '�����б�����
            cbo����.AddItem "ZL1_BILL_1302(���ݴ�ӡ)"
            cbo����.AddItem "ZL1_INSIDE_1302_1(ҩƷ�����ӡ)"
            cbo����.ListIndex = 0
        Case 1304
            chkPrintCode.Caption = "���̻����(����)���ӡҩƷ����"
            '�����б�����
            cbo����.AddItem "ZL1_BILL_1304(���ݴ�ӡ)"
            cbo����.AddItem "ZL1_INSIDE_1304_1(ҩƷ�����ӡ)"
            cbo����.ListIndex = 0
        Case 1305
            '�����б�����
            cbo����.AddItem "ZL1_BILL_1305(���ݴ�ӡ)"
            cbo����.AddItem "ZL1_INSIDE_1305_2(ҩƷ�����ӡ)"
            cbo����.ListIndex = 0
        Case 1306
            '�����б�����
            cbo����.AddItem "ZL1_BILL_1306(���ݴ�ӡ)"
            cbo����.AddItem "ZL1_INSIDE_1306_1(ҩƷ�����ӡ)"
            cbo����.ListIndex = 0
        Case Else
            chkPrintCode.Visible = False
            cbo����.Visible = False
    End Select
End Sub

Private Sub txt��ѯ����_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = 0
End Sub


Private Sub txt��ѯ����_Validate(Cancel As Boolean)
    If Val(txt��ѯ����.Text) > 7 Then
        If MsgBox("��ѯʱ�����7����ܻᵼ�²�ѯ�������Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = False
            txt��ѯ����.SetFocus
            zlControl.TxtSelAll txt��ѯ����
        End If
    End If
    If Val(txt��ѯ����.Text) = 0 Then
        MsgBox "��ѯʱ��������0�����������룡", vbInformation, gstrSysName
        Cancel = False
        txt��ѯ����.SetFocus
        zlControl.TxtSelAll txt��ѯ����
    End If
End Sub


Private Sub txt�̵�ʱ��_Change()
    UpD�̵�ʱ��.Value = Val(txt�̵�ʱ��.Text)
End Sub

Private Sub txt�̵�ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
        End If
        If Val(txt�̵�ʱ��.Text & Chr(KeyAscii)) > 90 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt�̵�ʱ��_Validate(Cancel As Boolean)
    If Val(txt�̵�ʱ��.Text) > 90 Then
        MsgBox "�̵�ʱ�䷶Χ���ܴ���3���£�", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub UpD�̵�ʱ��_Change()
    txt�̵�ʱ��.Text = UpD�̵�ʱ��.Value
End Sub


