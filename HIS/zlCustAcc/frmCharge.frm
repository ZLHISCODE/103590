VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCharge 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ʴ���"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6975
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraCancel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   8460
      TabIndex        =   7
      Top             =   5730
      Width           =   2115
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame fraOK 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   6540
      TabIndex        =   6
      Top             =   5640
      Width           =   2025
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   510
         TabIndex        =   8
         ToolTipText     =   "�ȼ���F2"
         Top             =   330
         Width           =   1275
      End
   End
   Begin VB.Frame fraʱ�� 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7830
      TabIndex        =   5
      Top             =   4710
      Width           =   2265
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   240
         TabIndex        =   11
         Top             =   210
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame fra������ 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   5460
      TabIndex        =   4
      Top             =   4830
      Width           =   2265
      Begin VB.ComboBox cbo������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   60
         TabIndex        =   10
         Top             =   120
         Width           =   2085
      End
   End
   Begin VB.Frame fra�� 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   9600
      TabIndex        =   3
      Top             =   180
      Width           =   855
      Begin VB.CheckBox chk�� 
         Caption         =   "��"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F8"
         Top             =   150
         Width           =   405
      End
   End
   Begin VB.Frame fraNO 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   8070
      TabIndex        =   1
      Top             =   180
      Width           =   1605
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   150
         Width           =   1425
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6615
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCharge.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11800
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   88
            Key             =   "�������"
            Object.ToolTipText     =   "�������"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "ҽ������"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":115E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":1798
            Key             =   "WB"
            Object.ToolTipText     =   $"frmCharge.frx":1DD2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraForm 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   60
      TabIndex        =   13
      Top             =   90
      Width           =   11205
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   4755
         Width           =   1800
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   9045
         TabIndex        =   38
         Top             =   2498
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.ComboBox cboִ�п��� 
         Height          =   300
         Index           =   0
         Left            =   7710
         TabIndex        =   37
         Top             =   2460
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtʵ�ս�� 
         Height          =   300
         Index           =   0
         Left            =   6675
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtӦ�ս�� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txt��׼���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   4965
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txt���㵥λ 
         Height          =   300
         Index           =   0
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txt��ҳID 
         Height          =   300
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txt��ʶ�� 
         Height          =   300
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txt����ID 
         Height          =   300
         Left            =   390
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txt���˿��� 
         Height          =   300
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtʵ�� 
         Height          =   300
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   5670
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox chk�Ӱ� 
         Caption         =   "�Ӱ�(&A)"
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   4770
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   3570
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   300
         Left            =   4710
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton cmdϸĿѡ�� 
         Caption         =   "��"
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���Ctrl+Enter"
         Top             =   2468
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   300
         Left            =   2040
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   8490
         TabIndex        =   14
         Top             =   840
         Width           =   1830
      End
      Begin VB.TextBox txt�շ���Ŀ 
         Height          =   300
         Index           =   0
         Left            =   1635
         TabIndex        =   39
         Top             =   2460
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.ComboBox cbo�շ���� 
         Height          =   300
         Index           =   0
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2460
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   420
         TabIndex        =   18
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox txt���˲��� 
         Height          =   300
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   4230
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtӦ�� 
         Height          =   300
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   5670
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "���ʵ�"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   23
         Top             =   210
         Visible         =   0   'False
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'����������������������������������������������������������������������������������������������������������������������������������������
'��ڲ�����
Public Enum UseType
    UseסԺ = 0
    Use���ҷ�ɢ = 1
    Useҽ������ = 2
    Use���� = 3
End Enum
Public Enum InState
    staִ�� = 0
    sta���� = 1
    sta���� = 2
    sta���� = 3
End Enum

'2.����ʼ״̬������
Public mlng����ID  As Long '���ʵ�ID
Public mbytUseType As UseType   '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���,3-�������
Public mbytInState As InState   '0-ִ��,1-���,2-����,3-����
Public mstrInNO As String       '�������ĵ��ݺ�
Public mlngUnitID As Long '��ǰ���ʲ���,Ϊ0ʱ��ʾ���в���
Public mlngDeptID As Long '��ǰ���ʿ���,Ϊ0ʱ��ʾ���п���
Public mlng����ID As Long  '���ҷ�ɢ������
Public mstrPrivs As String
Public mblnViewCancel As Boolean '�Ƿ�鿴���˵ĵ���(mbytInState=1ʱ��Ч)

Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
'����������������������������������������������������������������������������������������������������������������������������������������
'���ݶ���
Private mrsMedAudit As ADODB.Recordset  '�����������ķ�����Ŀ
Private mrsMedPayMode As ADODB.Recordset '���п��õ�ҽ�Ƹ��ʽ
Private mrsClass As ADODB.Recordset '���ݲ�����ȡ�ĵ�ǰ���õ��շ����
Private mrsUnit As ADODB.Recordset '��ѡ���ִ�п���
Private mrsInfo As New ADODB.Recordset '������Ϣ
Private mrs�������� As ADODB.Recordset  '��ѡ�Ŀ�������
Private mrs������ As ADODB.Recordset    '��ѡҽ���ͻ�ʿ

'�������
Private mobjBill As ExpenseBill  '������õ��ݶ������
Private mcolBillDetails As BillDetails '���ݵ��շ�ϸĿ��
Private mobjBillDetail As BillDetail   '���ݵ��շ�ϸĿ����
Private mcolBillInComes As BillInComes '�շ�ϸĿ��������Ŀ��
Private mobjBillIncome As BillInCome   '�շ�ϸĿ��������Ŀ����
Private mobjDetail As Detail           '�������շ�ϸĿ����
Private mcolDetails As Details   '���ﵥ�����շ�ϸĿ���ϡ��

'�������
Private mstrWarn As String '�Ѿ���������ѡ����������
Private mrsWarn As ADODB.Recordset  '����������

Private mlngRows As Long            '��ǰ���ʵ����շ�����
Private mintCurrentRow As Integer   '��ǰ���ʵ����к�
Private mblnCard As Boolean         '�Ƿ�ˢ���￨

Private mblnNOMoved As Boolean '�����ĵ����Ƿ��ں����ݱ���,�������洫ֵ,���ж�
Private mcurModiMoney As Currency '�޸ĵ���ʱԭ���ݵĽ��
Private mstrUnitIDs As String   '��ǰ����Ա�����в���ID

Private mcurPreMoney As Currency       '��¼�޸�ʱԭ������,�Ա���ȷ��ȡʵ��ʣ���
Private mblnOne As Boolean             '�Ƿ�ֻ��һ�������շ����
Private mcur���ý�� As Currency       '��ǰ���˿��õ������
Private mblnDo As Boolean   '��combobox��Clickְ�����ж��Ƿ�ִ��,����index=**ʱ��ʽִ��
Private marrDr() As String '��¼ҽ����"ID|����ID|���|����|����"

Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
    ʵʱ��� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
Private mstrFreeTable As String
Private mstrTitle As String '���ڴ�����Ի�����Ĵ�����

Private mobjPublicExpense As Object  '���ù�������
Private mintPriceGradeStartType As Integer
Private mstrPriceGrade As String

Public Sub MainProc()
    Dim tmpBill As ExpenseBill
    Dim i As Long, lngPre As Long, strPre As String, strTmp As String

    If mbytUseType <> Use���� Then
        mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    End If
    gblnOK = False: mblnDo = True
    Load frmCharge
    Set mobjBill = New ExpenseBill
    
    '��ʼ�����ݵĽ���
    If InitFace = False Then
        Unload frmCharge
        Exit Sub
    End If
    
    If mbytUseType <> Use���� Then
        mstrUnitIDs = GetUserUnits
    Else
        mstrUnitIDs = ""
    End If
    
    '��ʼ����������
    '��������,�Ƿ���ת������ݱ���
    If mbytInState = sta���� Then
        mblnNOMoved = zlDatabase.NOMoved(mstrFreeTable, mstrInNO, , 2, Me.Caption)
    Else
        If Not (mbytInState = staִ�� And mstrInNO = "") Then  '�޸�,����,����
            If zlDatabase.NOMoved(mstrFreeTable, mstrInNO, , 2, Me.Caption) Then
                If Not ReturnMovedExes(mstrInNO, 2, Me.Caption) Then Exit Sub
            End If
            mblnNOMoved = False
        End If
    End If
    
    '����ִ�б�ʾ�������޸�
    If mbytInState = staִ�� Or mbytInState = sta���� Then
        '�Ա�����Ҫ�õ���һЩ�������ݽ���װ�����������ѱ𡢿������ҡ�ִ�п���
        If Not InitData Then
            Unload frmCharge
            Exit Sub
        End If
    End If
    If mbytInState <> staִ�� Then   '��ʾ�����������ʵ���(1,2,3)
        '��Щ�����ܼ򵥣��ò�����ȥ������
        Call NewBill
        If Not ReadBill(mstrInNO) Then
            Unload frmCharge
            Exit Sub
        End If
        cboNO.Text = mstrInNO
    Else '����
        '��ȡ�õ��ݵ�����
        If mstrInNO <> "" Then '�޸ĵ���  ������ں󱸱��У�����ִ�е������ǰ�����˳�
            Call ImportBill(mstrInNO, mlngRows, mstrPriceGrade)
            If mobjBill.NO = "" Then
                MsgBox "������ȷ��ȡ�������ݣ�", vbInformation, gstrSysName
                Unload frmCharge
                Exit Sub
            Else
                mcurModiMoney = GetBillMoney(IIf(mbytUseType = Use����, 1, 2), mobjBill.NO) 'Ҫ�ڶ�ȡ������Ϣǰ�ȶ�
                
                lngPre = mobjBill.��������ID
                strPre = mobjBill.������

                txtPatient.Text = "-" & mobjBill.����ID
                Call txtPatient_KeyPress(13)
                
                If mbytUseType <> Use���� Then
                    Call ReCalcInsure '���¼���ͳ����
                End If
                
                '��ʾ����ԭ���ݺ�,��������µ��ݺ�
                cboNO.Text = mobjBill.NO
                txtDate.Text = Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss")
                chk�Ӱ�.Value = mobjBill.�Ӱ��־

                mblnDo = False
                    cbo��������.ListIndex = cbo.FindIndex(cbo��������, lngPre)
                    If cbo��������.ListIndex = -1 And lngPre <> 0 Then
                        strTmp = GET��������(lngPre)
                        If strTmp <> "" Then
                            cbo��������.AddItem strTmp
                            cbo��������.ListIndex = cbo��������.NewIndex
                            cbo��������.ItemData(cbo��������.NewIndex) = lngPre
                        End If
                    End If
                    
                    i = 0
                    If cbo��������.ListIndex <> -1 Then i = cbo��������.ItemData(cbo��������.ListIndex)
                    Call FillDoctor(i)
                    Call cbo.SeekIndex(cbo������, strPre, , True)
                    If cbo������.ListIndex = -1 And strPre <> "" Then
                        cbo������.AddItem strPre
                        cbo������.ListIndex = cbo������.NewIndex
                    End If
                mblnDo = True
                mobjBill.��������ID = lngPre
                mobjBill.������ = strPre

                '�޸�ʱӦ���浱ǰ����Ա������
                mobjBill.����Ա��� = UserInfo.���
                mobjBill.����Ա���� = UserInfo.����
                
                Call zlControl.CboLocate(cboBaby, mobjBill.Ӥ����, True)

                Call ShowDetails
                Call ShowMoney
                
                'byZT200302
                For i = 0 To mlngRows - 1
                    If mobjBill.Details("R" & i).Detail.��� Then
                        txt����(i).TabStop = False
                        txt����(i).Locked = True
                        txt��׼����(i).TabStop = True
                        txt��׼����(i).Locked = False
                    Else
                        txt����(i).TabStop = True
                        txt����(i).Locked = False
                        txt��׼����(i).TabStop = False
                        txt��׼����(i).Locked = True
                    End If
                    chk����(i).Enabled = mobjBill.Details("R" & i).�շ���� = "F" '����
                    If chk����(i).Enabled = False Then chk����(i).Value = 0
                    
                    'ִ�п���!!!
                    If mobjBill.Details("R" & i).�շ�ϸĿID <> 0 Then Call Fillִ�п���(i)
                    
                    If cboִ�п���(i).ListCount = 1 Then
                        cboִ�п���(i).TabStop = False
                    Else
                        cboִ�п���(i).TabStop = True
                    End If
                Next

                mcurPreMoney = CalcGridToTal
            End If
        Else
            Call NewBill
            If mbytUseType = Use���ҷ�ɢ And mlng����ID <> 0 Then
                txtPatient.Text = "-" & mlng����ID
                Call txtPatient_KeyPress(13)
            End If
        End If
    End If
    
    '��ʼ���ɹ�
    If Not gfrmMain Is Nothing Then
        frmCharge.Show vbModal, gfrmMain
    ElseIf glngMain <> 0 Then
        zlCommFun.ShowChildWindow frmCharge.hwnd, glngMain
    End If
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    'ǿ��Ҫѡ��һ��(��һ��)
    If cbo��������.ListIndex = -1 And cbo��������.ListCount <> 0 Then cbo��������.ListIndex = 0
End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
    If cbo������.Text <> "" Then
        If cbo.FindIndex(cbo������, zlStr.NeedName(cbo������.Text), True) = -1 Then cbo������.ListIndex = -1: cbo������.Text = ""
    End If
    If cbo������.Text = "" Then Call cbo������_KeyPress(vbKeyReturn)
    '����������ȷ��������ʱ,���ܴ�ʱ��ѡ������,��ȥ�����������Һ�����ѡ
    If gbln������ And cbo������.ListIndex = -1 And txtPatient.Text <> "" And cbo������.ListCount > 0 Then Cancel = True
End Sub

Private Sub chk����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    If mbytInState = sta���� Then
        cmdCancel.SetFocus
    ElseIf mbytInState = sta���� Then
        txtDate.SetFocus
    ElseIf mbytInState = sta���� Then
        cmdOK.SetFocus
    Else
        If mbytUseType = Use���ҷ�ɢ And mobjBill.���� <> "" Then
            If cbo�շ����(0).ListIndex = -1 And cbo�շ����(0).Visible = True Then
                cbo�շ����(0).SetFocus
            Else
                If txt�շ���Ŀ(0).TabStop = True Then
                    txt�շ���Ŀ(0).SetFocus
                Else
                    SendKeys "{TAB}"
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    mstrTitle = "���ʴ���"
    Call CreatePublicExpenseObject
    Call RestoreWinState(Me, App.ProductName, mstrTitle)
End Sub

Public Sub CreatePublicExpenseObject()
    '����:�����������ò���
    Err = 0: On Error Resume Next
    If mobjPublicExpense Is Nothing Then
        Set mobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If mobjPublicExpense Is Nothing Then Exit Sub
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If mobjPublicExpense.zlInitCommon(glngSys, gcnOracle, gstrDbUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    mintPriceGradeStartType = mobjPublicExpense.zlGetPriceGradeStartType()
    If mintPriceGradeStartType = 0 Then Exit Sub
    '��ȡվ��۸�ȼ�
    Call mobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", , , mstrPriceGrade)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrInNO = 0
    mlngUnitID = 0
    mlngDeptID = 0
    mlng����ID = 0
    mintCurrentRow = 0
    mblnViewCancel = False
    Set mrs�������� = Nothing
    Set mrs������ = Nothing
    Set mrsInfo = Nothing
    Set mrsMedAudit = Nothing
    Set mrsMedPayMode = Nothing
    Set mrsWarn = Nothing
    Call SaveWinState(Me, App.ProductName, mstrTitle)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If ActiveControl Is cbo������ Then Call cbo������_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
        Case vbKeyF6 '�����ǰ��������,�����µ�״̬
            If mbytInState = 0 Then
                If fraForm.Enabled Then '�������뵥��״̬'(����������²��˵���)
                    mstrInNO = ""
                    txtPatient.Text = "": txt����.Text = "": txt����.Text = "": mcur���ý�� = 0
                    Call NewBill
                    txtPatient.SetFocus
                ElseIf chk��.Value = Checked Then '�˾ݵ�״̬
                    chk��.Value = Unchecked
                    Call NewBill
                    Call SetDisible(True)
                    txtPatient.SetFocus
                ElseIf Not fraForm.Enabled Then '��ȡ���۵�����״̬
                    Call NewBill
                    Call SetDisible(True)
                    txtPatient.SetFocus
                End If
            End If
        Case vbKeyF7 '�л����뷨
            If Not gbln�����л� Then Exit Sub   '35242
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyF8 '��(�Զ������¼�)
            If chk��.Visible And fra��.Enabled And chk��.Enabled Then chk��.Value = IIf(chk��.Value = Checked, Unchecked, Checked)
        Case vbKeyReturn
            If Shift And vbCtrlMask = vbCtrlMask Then
                If ActiveControl.Name = "txt�շ���Ŀ" Then
                    Call cmdϸĿѡ��_Click(ActiveControl.Index)
                End If
            End If
        Case vbKeyEscape, vbKeyX
            If KeyCode = vbKeyX And Shift <> 4 Then Exit Sub
            Call cmdCancel_Click
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If gbln�����л� = False Then Exit Sub
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        zlDatabase.SetPara "���뷽ʽ", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
        gbytCode = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))
    End If
End Sub

Private Sub txt��ʶ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt����ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt���˲���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt���˿���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt���㵥λ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt����_Change()
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    txt����.Text = mobjBill.����
End Sub

Private Sub txt����_Change()
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    txt����.Text = mobjBill.����
End Sub

Private Sub txt����ID_Change()
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    txt����ID.Text = Format(mobjBill.����ID, "#;;;")
End Sub

Private Sub txt��ʶ��_Change()
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    txt��ʶ��.Text = Format(mobjBill.��ʶ��, "#;;;")
End Sub

Private Sub txtʵ�ս��_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtӦ�ս��_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt��ҳID_Change()
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    txt��ҳID.Text = Format(mobjBill.��ҳID, "#;;;")
End Sub

Private Sub txt���˲���_Change()
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    txt���˲���.Text = mobjBill.����
End Sub

Private Sub txt���˿���_Change()
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    txt���˿���.Text = mobjBill.����
End Sub

Private Sub txt�շ���Ŀ_Change(Index As Integer)
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    If txt�շ���Ŀ(Index).Locked = True Then
        txt�շ���Ŀ(Index).Text = mobjBill.Details("R" & Index).�շ�����
    End If
End Sub

Private Sub txt���㵥λ_Change(Index As Integer)
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    If txt���㵥λ(Index).Locked = True Then
        txt���㵥λ(Index).Text = mobjBill.Details("R" & Index).���㵥λ
    End If
End Sub

Private Sub txt����_Change(Index As Integer)
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If Not fraForm.Enabled Then Exit Sub
    
    If txt����(Index).Locked Then
        If mobjBill.Details("R" & Index).���� = 0 Then
            txt����(Index).Text = ""
        Else
            txt����(Index).Text = mobjBill.Details("R" & Index).����
        End If
    End If
End Sub

Private Sub txt��׼����_Change(Index As Integer)
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    If txt��׼����(Index).Locked Then
        If mobjBill.Details("R" & Index).��׼���� <> 0 Then
            txt��׼����(Index).Text = Format(mobjBill.Details("R" & Index).��׼����, "0.0000")
        Else
            txt��׼����(Index).Text = ""
        End If
    End If
End Sub

Private Sub txtӦ�ս��_Change(Index As Integer)
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    If mobjBill.Details("R" & Index).Ӧ�ս�� <> 0 Then
        txtӦ�ս��(Index).Text = Format(mobjBill.Details("R" & Index).Ӧ�ս��, gstrDec)
    Else
        txtӦ�ս��(Index).Text = ""
    End If
End Sub

Private Sub txtʵ�ս��_Change(Index As Integer)
    '���fraForm�����ã��ǿ϶��ǳ����ڸı�
    If fraForm.Enabled = False Then Exit Sub
    
    If mobjBill.Details("R" & Index).ʵ�ս�� <> 0 Then
        txtʵ�ս��(Index).Text = Format(mobjBill.Details("R" & Index).ʵ�ս��, gstrDec)
    Else
        txtʵ�ս��(Index).Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
    mblnCard = False
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    
    If txtPatient.Text = mobjBill.���� Then Exit Sub
    If Trim(txtPatient.Text) = "" Then
        '�մ����⴦��
        txtPatient.Text = mobjBill.����
        Exit Sub
    End If
    
    If Input����() = False Then
        Cancel = True
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Integer
    
    mblnCard = False
    
    On Error Resume Next

    If txtPatient.Locked Then
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    
    mblnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, glngSys)

    If Trim(Me.txtPatient.Text) = "" And KeyAscii = 13 Then
        With frmPatiSelect
            If mbytUseType = UseסԺ Or mbytUseType = Use���ҷ�ɢ Then
                .mlngUnitID = mlngUnitID
            ElseIf mbytUseType = Useҽ������ Then
                .mlngUnitID = mlngDeptID
            Else
                KeyAscii = 0
                Exit Sub
            End If
            .mbytUseType = mbytUseType
            .mstrPrivs = mstrPrivs
            Set .mfrmParent = Me
            .Show 1, Me
            Me.Refresh
        End With
    End If

    If mblnCard And Len(txtPatient.Text) = gbytCardNOLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then '�������ûس�
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        Else
            If txtPatient.Text = mobjBill.���� Then
                If cbo��������.ListIndex = -1 Then
                    cbo��������.SetFocus
                Else
                    If cbo�շ����(0).ListIndex = -1 And cbo�շ����(0).Visible = True Then
                        cbo�շ����(0).SetFocus
                    Else
                        If txt�շ���Ŀ(0).TabStop = True Then
                            txt�շ���Ŀ(0).SetFocus
                        Else
                            SendKeys "{TAB}"
                        End If
                    End If
                End If
                Exit Sub
            End If
        End If
        KeyAscii = 0
        
        If Input����() = True Then
            '����õ�����
            If cbo��������.ListIndex = -1 Then
                cbo��������.SetFocus
            Else
                If cbo�շ����(0).ListIndex = -1 And cbo�շ����(0).Visible = True Then
                    cbo�շ����(0).SetFocus
                Else
                    If txt�շ���Ŀ(0).TabStop = True Then
                        txt�շ���Ŀ(0).SetFocus
                    Else
                        SendKeys "{TAB}"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Function Input����() As Boolean
'���ܣ����벡������
'��������ʾ��Դ�ڼ��̻����
    '��ȡ������Ϣ
    Dim blnReturn As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strTemp As String
    Dim strSQL As String
    Dim blnOutMsg As Boolean
    If Not (mbytInState = 0 And mbytUseType = 1 And sta.Panels(2) Like "��һ��*") Then
        sta.Panels(2) = ""
    End If
    If mbytUseType = Use���� Then
       blnReturn = GetPatientOut(txtPatient.Text)
    Else
       blnReturn = GetPatientIn(txtPatient.Text, mblnCard, blnOutMsg)
    End If
    If Not blnReturn Then
        If mblnCard Then
            txtPatient.Text = ""
           If Not blnOutMsg Then MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����", vbInformation, gstrSysName
            Call ClearPatient
        Else
            If Not blnOutMsg Then MsgBox "���ܶ�ȡ������Ϣ��", vbInformation, gstrSysName
            If mstrInNO = "" Then
                strTemp = txtPatient.Text
                Call ClearPatient
                txtPatient.Text = strTemp
            End If
            zlControl.TxtSelAll txtPatient
        End If
        Exit Function
    Else
        '���￨������
        If Mid(gstrCardPass, 6, 1) = "1" And mblnCard Then
            If Not zlCommFun.VerifyPassWord(Me, "" & mrsInfo!����֤��, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
                txtPatient.Text = ""
                Call ClearPatient
                Exit Function
            End If
        End If
    
        '�жϸò��˵ķѱ��Ƿ����
        Call cbo.SeekIndex(cbo�ѱ�, IIf(IsNull(mrsInfo("�ѱ�")), "", mrsInfo("�ѱ�")), , True)
        If cbo�ѱ�.ListIndex = -1 Then
            txtPatient.Text = ""
            MsgBox "����" & IIf(IsNull(mrsInfo("����")), "", mrsInfo("����")) & "�ķѱ���Ϣ������Ч�����ܼ��ʣ�", vbInformation, gstrSysName
            Call ClearPatient
            Exit Function
        End If
        
        '�������ϼ����򴫽����Ĳ���
        If mbytUseType = Use���ҷ�ɢ And mrsInfo!����ID <> mlng����ID Then mlng����ID = 0

        If mbytUseType = Use���� Then
            '�ɹҺŵ�����ʱ��ִ�в��Ų���Ϊ����Ŀ�������
            If Not IsNull(mrsInfo("����ID")) Then
                If IsNull(mrsInfo!����) Then
                    txtPatient.Text = ""
                    MsgBox "�ò��˹Һ�ʱû�еǼǵ���,��Ҫ���벡��������", vbInformation, gstrSysName
                    Call ClearPatient
                    Set mrsInfo = New ADODB.Recordset
                    Exit Function
                End If
                
                mobjBill.����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                Set�������� IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
            End If
        Else
             '�Զ����ÿ�������(ͬʱ���ü��ʱ�����Ϣ)
            mobjBill.����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
            Set�������� IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
        End If
        
        '����Ԥ������Ϣ
        Set rsTmp = GetMoneyInfo(mrsInfo!����ID, CDbl(mcurModiMoney), Val("" & mrsInfo!����) > 0)
        If rsTmp.State = adStateOpen Then
            sta.Panels(3).Text = "Ԥ��:" & Format(rsTmp!Ԥ�����, "0.00")
            sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(rsTmp!�������, gstrDec)
            sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(rsTmp!Ԥ����� - rsTmp!�������, "0.00")
            cmdOK.Tag = rsTmp!Ԥ�����
            cmdCancel.Tag = rsTmp!�������
            mcur���ý�� = rsTmp!Ԥ����� - rsTmp!�������
        Else
            sta.Panels(3).Text = "Ԥ��:0.00/����:" & gstrDec & "/ʣ��:0.00"
            cmdOK.Tag = 0
            cmdCancel.Tag = 0
            mcur���ý�� = 0
        End If
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------
        '���˺�:26952
        Dim cur��� As Currency, curItemMoney As Currency, cur���ն� As Currency, curTotal As Currency
        cur��� = mcur���ý��
        curItemMoney = 0
        '���ݷ���
        curTotal = CalcGridToTal
        
        '���¶�ȡ���ն�
        cur���ն� = GetPatiDayMoney(mrsInfo!����ID)
        If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(2, Val(NVL(mrsInfo!����ID)))
        
        
        gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!����, Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, cur���, cur���ն� - mcurModiMoney, curTotal, Val(NVL(mrsInfo!������)), "", "", mstrWarn, , , True)
        '����:0;û�б���,����
        '     1:������ʾ���û�ѡ�����
        '     2:������ʾ���û�ѡ���ж�
        '     3:������ʾ�����ж�
        '     4:ǿ�Ƽ��ʱ���,����
        '     5.������ʾ���û�ѡ�����,��ֻ�������Ϊ���۵�
        If gbytWarn = 2 Or gbytWarn = 3 Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "":
            mlng����ID = 0
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Call ClearPatient: Exit Function
        End If
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        
        
        Call LoadPatientBaby(cboBaby, mrsInfo!����ID, mrsInfo!��ҳID)
                                
        '������Ϣ
        With mobjBill
            .���� = IIf(IsNull(mrsInfo!����), 0, mrsInfo!����)
            .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
            .��ҳID = IIf(IsNull(mrsInfo!��ҳID), 0, mrsInfo!��ҳID)
            .��ʶ�� = IIf(IsNull(mrsInfo!��ʶ��), 0, mrsInfo!��ʶ��)
            .���� = "" & mrsInfo!����
            .�Ա� = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
            .���� = IIf(IsNull(mrsInfo!����), 0, mrsInfo!����)
            .�ѱ� = IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�)
            .������ = IIf(IsNull(mrsInfo!������), 0, mrsInfo!������)

            .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
            .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
            .���� = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
            .���� = IIf(IsNull(mrsInfo!����), 0, mrsInfo!����)

            If cbo��������.ListIndex <> -1 Then
                mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
            Else
                mobjBill.��������ID = 0
            End If
        End With
        
        If Not IsNull(mrsInfo!����) Then
            MCPAR.�������� = gclsInsure.GetCapability(support��������, mrsInfo!����ID, mrsInfo!����)
            MCPAR.�����ϴ� = gclsInsure.GetCapability(support�����ϴ�, mrsInfo!����ID, mrsInfo!����)
            MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, mrsInfo!����ID, mrsInfo!����)
            MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, mrsInfo!����ID, mrsInfo!����)
            MCPAR.ʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, mrsInfo!����ID, mrsInfo!����)
        End If
        
        Call ShowPatient
        txtPatient.PasswordChar = ""

        If Not IsNull(mrsInfo!��Ժ����) And mbytUseType <> Use���� Then
            MsgBox "��������" & vbCrLf & vbCrLf & "�ò������� " & Format(mrsInfo!��Ժ����, "yyyy-MM-dd") & " ��Ժ�����ڶԸò���ǿ�ƽ��м��ʣ�", vbInformation, gstrSysName
            txtDate.Text = Format(mrsInfo!��Ժ����, "yyyy-MM-dd HH:mm:ss")
        Else
            txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
        
        '��ȡ�۸�ȼ�
        If mintPriceGradeStartType >= 2 Then
            Call mobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(NVL(mrsInfo!����ID)), Val(NVL(mrsInfo!��ҳID)), _
                NVL(mrsInfo!ҽ�Ƹ��ʽ), , , mstrPriceGrade)
        End If
        
        If mbytInState = 0 And mobjBill.Details.Count > 0 Then
            '���¼���۸�
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If
    End If
    Input���� = True
End Function

Private Sub ClearPatient()
'���ܣ����������Ϣ����ʾ
    With mobjBill
        .����ID = 0
        .��ҳID = 0

        .����ID = 0
        .����ID = 0
        .���� = ""
        .���� = ""

        .���� = ""
        .��ʶ�� = 0
        .���� = ""
        .�Ա� = ""
        .���� = ""
        .�ѱ� = ""
        .������ = 0
    End With
    Call ShowPatient
End Sub

Private Sub ShowPatient()
'���ܣ���ʾ������Ϣ
    With mobjBill
        txtPatient.Text = .����
        Call cbo.SeekIndex(cbo�Ա�, .�Ա�, , True)
        txt����.Text = .����
        Call cbo.SeekIndex(cbo�ѱ�, .�ѱ�, , True)
        txt����.Text = Format(.����, "#;;;")
        
        txt����ID.Text = Format(.����ID, "#;;;")
        txt��ҳID.Text = Format(.��ҳID, "#;;;")
        txt��ʶ��.Text = Format(.��ʶ��, "#;;;")
        txt���˲���.Text = .����
        txt���˿���.Text = .����
    End With
End Sub

Private Sub Set��������(ByVal lngID As Long)
'���ܣ����ÿ�������
'ע�⣺����������ҵ�Tag���������õĻ�����Ҫ������Ӧ�Ĵ���
    If cbo��������.Tag <> "" Then
        Select Case cbo��������.Tag
            Case "C1" '�������п���
                cbo��������.ListIndex = cbo.FindIndex(cbo��������, IIf(mobjBill.����ID = 0, lngID, mobjBill.����ID))
            Case "C2" '����Ա���ڿ���
                cbo��������.ListIndex = cbo.FindIndex(cbo��������, IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID))
            Case Else 'ָ������
                cbo��������.ListIndex = cbo.FindIndex(cbo��������, Val(cbo��������.Tag))
                If cbo��������.ListIndex < 0 Then
                    cbo��������.AddItem GET��������(cbo��������.Tag, mrs��������), 0
                    cbo��������.ListIndex = 0
                End If
        End Select
        
        If cbo��������.ListCount > 0 And cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
    Else
        cbo��������.ListIndex = cbo.FindIndex(cbo��������, lngID)
    End If
    
    If cbo��������.ListIndex = -1 Then
        mobjBill.��������ID = 0
    Else
        mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    End If
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo�Ա�.ListIndex <> -1 Then mobjBill.�Ա� = Mid(cbo�Ա�.Text, InStr(cbo�Ա�.Text, "-") + 1)
        SendKeys "{TAB}"
    End If
    If cbo�Ա�.Locked Then Exit Sub
    If SendMessage(cbo�Ա�.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
End Sub

Private Sub txt����_Gotfocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mobjBill.���� = txt����.Text
        SendKeys "{TAB}"
    End If
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub cbo�ѱ�_Click()
    If cbo�ѱ�.ListIndex <> -1 And Not mobjBill Is Nothing Then
        mobjBill.�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)

        If mbytInState = staִ�� Then
            If mobjBill.Details.Count = 0 Then Exit Sub
            '���¼���۸�
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If
    End If
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo�ѱ�.Locked Then
        If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn And cbo�ѱ�.ListIndex <> -1 Then
        mobjBill.�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)

        If mbytInState = staִ�� And mstrInNO <> "" Then
            '���¼���۸�
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If

        SendKeys "{TAB}"
    End If
'    If SendMessage(cbo�ѱ�.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
    lngIdx = zlControl.CboMatchIndex(cbo�ѱ�.hwnd, KeyAscii)
'    If lngIdx <> -2 Then cbo�ѱ�.ListIndex = lngIdx
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    

   Dim lngIdx As Long, lngҽ��ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cbo��������.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo������.ListIndex >= 0 Then lngҽ��ID = cbo������.ItemData(cbo������.ListIndex)
    If mrs�������� Is Nothing Then Call FillDept(lngҽ��ID)
    
    If zlSelectDept(Me, 0, cbo��������, mrs��������, cbo��������.Text) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub

'
'
'
'
'    Dim lngIdx As Long
'    If KeyAscii = 13 And cbo��������.ListIndex <> -1 Then
'        mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
'        SendKeys "{TAB}"
'        Exit Sub
'    End If
'    If cbo��������.Locked Then Exit Sub
'
'    If SendMessage(cbo��������.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
'    lngIdx = MatchIndex(cbo��������.hwnd, KeyAscii)
'    If lngIdx <> -2 Then cbo��������.ListIndex = lngIdx

    'ǿ��Ҫѡ��һ��(��һ��)
    If cbo��������.ListIndex = -1 And cbo��������.ListCount <> 0 Then cbo��������.ListIndex = 0
End Sub

Private Sub cbo��������_Click()
    Dim i As Long, strDoctor As String
    If Not mblnDo Then Exit Sub
       
    '��λҽ��
    cbo������.Clear
    If cbo��������.ListIndex <> -1 Then
        FillDoctor cbo��������.ItemData(cbo��������.ListIndex)
    End If

    '���ݶ���
    If mbytInState = 0 Then
        If cbo��������.ListIndex = -1 Then
            mobjBill.��������ID = 0
        Else
            mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
    End If
    
    '�������������Ŀ��ִ�п���
    'byZT200302
    If mbytInState = 0 And cbo��������.ListIndex <> -1 And cbo��������.Visible Then
        For i = 0 To mobjBill.Details.Count - 1
            With mobjBill.Details("R" & i)
                If .Detail.ִ�п��� = 6 Then '6-�����˿���
                    cboִ�п���(i).Clear
                    'Call ShowDetail(i)
                    .ִ�в���ID = cbo��������.ItemData(cbo��������.ListIndex)
                End If
            End With
        Next
    End If
End Sub

Private Function isCheck������Exists(ByVal str���� As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo������.ListCount - 1
        If zlStr.NeedName(cbo������.List(i)) = str���� Then
            If blnLocateItem Then cbo������.ListIndex = i
            isCheck������Exists = True
            Exit Function
        End If
    Next
End Function


Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Integer, intIdx As Integer, strResult As String, iCount As Integer
    Dim strText As String, strFilter As String, rsTemp As ADODB.Recordset
        
    If KeyAscii = vbKeyReturn Then
        strText = UCase(cbo������.Text)
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo������.List(cbo������.ListIndex) Then
                Call zlControl.CboSetIndex(cbo������.hwnd, -1)
            Else
                zlCommFun.PressKey vbKeyTab: Exit Sub
            End If
        End If
        
        If strText = "" Then
            cbo������.ListIndex = -1
        Else
            intIdx = -1
          strFilter = IIf(gbln��ʿ, "��Ա����<>''", "��Ա����<>'��ʿ'")
            '���˺�:22383
            '�ȸ��Ƽ�¼��
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrs������)
            Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
            Dim strCompents As String 'ƥ�䴮
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrs������.Filter = strFilter: iCount = 0
            With mrs������
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrs������.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                        If NVL(!���) = strText Then strResult = NVL(!����): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(NVL(!���)) = Val(strText) Then
                            If iCount = 0 Then strResult = NVL(!����)
                            iCount = iCount + 1
                        End If
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(mrs������!���) Like strText & "*" Then
                            If isCheck������Exists(NVL(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(NVL(!����)) = strText Then
                            If iCount = 0 Then strResult = NVL(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(NVL(!����)) Like strCompents Then
                            If isCheck������Exists(NVL(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                            If iCount = 0 Then strResult = NVL(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!���) Like strText & "*" Or Trim(NVL(!����)) Like strCompents Or Trim(NVL(!����)) Like strCompents Then
                            If isCheck������Exists(NVL(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                        End If
                    End Select
                    mrs������.MoveNext
                Loop
            End With
             If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = NVL(rsTemp!����)
            '���˺�:ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheck������Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                Select Case intInputType
                Case 0 '����ȫ����
                    rsTemp.Sort = "���"
                Case 1 '����ȫƴ��
                    rsTemp.Sort = "����"
                Case Else
                    '����ѡ������
'                    If gbyt��������ʾ = 1 Then '����
'                        rsTemp.Sort = "����"
'                    Else
                        rsTemp.Sort = "���"
                  '  End If
                End Select
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cbo������, rsTemp, True, "", "ȱʡ,ְ��,���ȼ���", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheck������Exists(NVL(rsReturn!����), True) Then
                                rsTemp.Close: Set rsTemp = Nothing
                                zlCommFun.PressKey vbKeyTab
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cbo������: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
                         
            
'
'            For i = 0 To cbo������.ListCount - 1
'                If InStr(cbo������.List(i), UCase(strText)) > 0 Then
'                    If intIdx = -1 Then cbo������.ListIndex = i
'                    intIdx = i
'                End If
'                If IsNumeric(strText) Then
'                    If cbo������.ItemData(i) = CDbl(strText) Then
'                        If intIdx = -1 Then cbo������.ListIndex = i
'                        intIdx = i
'                    End If
'                End If
'            Next
        End If
        If cbo������.ListIndex = -1 Then
            cbo������.Text = ""
            mobjBill.������ = UserInfo.����
        Else
            mobjBill.������ = zlStr.NeedName(cbo������.Text)
            If intIdx <> cbo������.ListIndex Then SendKeys "{F4}": Exit Sub
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub cbo������_Click()
    If Not mblnDo Then Exit Sub
    
    If mbytInState = 0 Then
        '���ݶ���
        mobjBill.������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
    End If
End Sub

Private Sub cboBaby_Click()
    mobjBill.Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub chk��_Click()
    Dim i As Long
    
    mstrInNO = ""
    '�ı������
    If chk��.Value = 1 Then
        chk��.ForeColor = &HFF&
        cboNO.Locked = False
        
        fraForm.Enabled = False
        fra������.Enabled = False
        fraʱ��.Enabled = False
    Else
        chk��.ForeColor = 0
        
        fraForm.Enabled = True
        fra������.Enabled = True
        fraʱ��.Enabled = True
        
        cboNO.Locked = True
    End If
        
    'btZY200302
    For i = 0 To mlngRows - 1
        cboִ�п���(i).Clear
    Next
    
    '��ʼ��
    Call NewBill
    
    'ɨβ����
    If chk��.Value = 1 Then
        cboNO.SetFocus
    Else
        '��������
        Call cbo��������_Click
        If mbytUseType = 1 And mlng����ID > 0 Then
            txtPatient.Text = "-" & mlng����ID
            Call txtPatient_KeyPress(13)
        Else
            txtPatient.SetFocus
        End If
    End If
End Sub

Private Sub chk�Ӱ�_Click()
    If mbytInState = sta���� Or chk��.Value = Checked Then Exit Sub
    If mbytInState = sta���� Then Exit Sub
    If Not chk�Ӱ�.Visible Then Exit Sub

    Dim blnAdd As Boolean

    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk�Ӱ�.Value = Unchecked And blnAdd Then
        If MsgBox("��ǰ���ڼӰ�ʱ�䷶Χ��,Ҫȡ���Ӱ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Checked
        End If
    End If
    If chk�Ӱ�.Value = Checked And Not blnAdd Then
        If MsgBox("��ǰ�����ڼӰ�ʱ�䷶Χ��,Ҫִ�мӰ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Unchecked
        End If
    End If
    mobjBill.�Ӱ��־ = IIf(chk�Ӱ�.Value = Checked, 1, 0)
    
    '���¼���۸�
    If Not mobjBill.Details.Count = 0 Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk�Ӱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        mobjBill.�Ǽ�ʱ�� = CDate(txtDate.Text)
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.�Ǽ�ʱ�� = CDate(txtDate.Text)
End Sub

Private Sub cboNO_GotFocus()
    cboNO.SelStart = 0
    cboNO.SelLength = Len(cboNO.Text)
    If chk��.Value = Checked Then
        cboNO.Locked = False
    Else
        cboNO.Locked = True
    End If
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, strOper As String, vDate As Date
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    zlControl.TxtCheckKeyPress cboNO, KeyAscii, m�ı�ʽ
 
    If KeyAscii = 13 And cboNO.Locked Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 14)

        If chk��.Value = 1 Then
            '�Ƿ���ת������ݱ���
            If zlDatabase.NOMoved(mstrFreeTable, cboNO.Text, , 2, Me.Caption) Then
                If Not ReturnMovedExes(cboNO.Text, 2, Me.Caption) Then Exit Sub
                mblnNOMoved = False
            End If
            
             '����Ȩ��
            If Not ReadBillInfo(IIf(mbytUseType = Use����, 1, 2), cboNO.Text, 2, strOper, vDate) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If mbytUseType = 0 And InStr(mstrPrivs, "���в���Ա") <= 0 Then
                If UserInfo.���� <> strOper Then
                    MsgBox "��û��""���в���Ա""Ȩ��,���ܶ�" & strOper & "�ĵ��ݽ�������!", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            If Not BillOperCheck(5, strOper, vDate, "����", cboNO.Text) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            If CheckExecute(cboNO.Text, mlng����ID, IIf(mbytUseType = Use����, 1, 2)) Then
                MsgBox "�ü��ʵ������Ѿ�ȫ��ִ��" & vbCrLf & "�����ɱ����ʵ��Ǽǵģ��������ʣ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If

            '�Ƿ��ѽ���
            'int��Դ-1-����;2-סԺ
            If HaveBilling(IIf(mbytUseType = Use����, 1, 2), cboNO.Text, False) <> 0 Then  'mlng����ID
                If BillExistInsure(cboNO.Text) <> 0 Then
                    MsgBox "��ҽ�����ʵ��ݰ����Ѿ����ʵ�����,�������ʣ�", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                cboNO.Text = "": cboNO.SetFocus: Exit Sub
                            End If
                        Case 2
                            MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,�������ʣ�", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End Select
                End If
            End If
            
            '�Ƿ������������¼
            If CheckRecalcRecord(cboNO.Text) Then
                MsgBox "���ָü��ʵ��ݴ��ڰ��ѱ�����Ĵ��۳����¼!" & vbCrLf & _
                    "����ǰ�밴�ѱ�������ã������˽����ܵ�������ǰ�Ĵ����Żݽ�", vbInformation, Me.Caption
            End If
            
            blnRead = ReadBill(cboNO.Text)
        End If

        If blnRead Then
            mstrInNO = cboNO.Text 'ȷ��ʱ��mstrInNOΪ׼
            cmdOK.SetFocus
        Else
            mstrInNO = "": cboNO.Text = "": cboNO.SetFocus
        End If
    End If
End Sub

Private Function CheckRecalcRecord(ByVal strNO As String) As Boolean
'���ܣ��ж�ָ�����˵�ָ�������Ƿ���ڰ��ѱ�����ĳ����¼(����Ϊ0�ļ�¼)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim blnסԺ As Boolean
    blnסԺ = (mstrFreeTable = "סԺ���ü�¼")
    
    Err = 0: On Error GoTo errH:
    strSQL = "Select Count(A.ID) Num" & vbNewLine & _
            "From " & mstrFreeTable & " A," & vbNewLine & _
            "     ( Select ����id," & IIf(blnסԺ, " ��ҳid, ���˲���id,", "0 as ��ҳid,0 as ���˲���ID,") & " ���˿���id, �շ�ϸĿid, ������Ŀid, ��������id, ִ�в���id, ����ʱ��" & vbNewLine & _
            "       From " & mstrFreeTable & vbNewLine & _
            "       Where NO = [1] And ���ʷ��� = 1" & vbNewLine & _
            "       Group By ����id," & IIf(blnסԺ, " ��ҳid, ���˲���id,", "") & "���˿���id, �շ�ϸĿid, ������Ŀid, ��������id, ִ�в���id, ����ʱ��) B" & vbNewLine & _
            "Where A.��¼���� = 2 And A.���� = 0 And A.����id+0 = B.����id " & _
                   IIf(blnסԺ, " And A.��ҳid = B.��ҳid And A.���˲���id + 0 = B.���˲���id ", "") & _
            "       And A.���˿���id + 0 = B.���˿���id And A.�շ�ϸĿid + 0 = B.�շ�ϸĿid And" & vbNewLine & _
            "      A.������Ŀid + 0 = B.������Ŀid And A.��������id + 0 = B.��������id And A.ִ�в���id + 0 = B.ִ�в���id And" & vbNewLine & _
            "      A.����ʱ�� = B.����ʱ��"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO)
    If rsTmp.RecordCount > 0 Then CheckRecalcRecord = rsTmp!Num > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    If mbytInState = staִ�� Then
        If Not CheckBillisZero Then
            If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim curTotal As Currency, cur���ն� As Currency
    Dim intInsure As Integer, i As Long
    Dim strInfo As String

    If mbytInState = sta���� Then '%%%
        'ҽ�����������ϴ�(ע���ж�˳��)
        If mbytUseType <> Use���� Then
            intInsure = BillExistInsure(mstrInNO) '�ж��Ƿ�ҽ�����˼ǵ���
            'ȥ����ҽ������ƥ����
        End If
        
        If mbytUseType = Use���� Then
            strSQL = "zl_������ʼ�¼_DELETE('" & mstrInNO & "','','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            strSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "','','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        'ҽ�����������ϴ�
        If mbytUseType <> Use���� And intInsure <> 0 Then
            If MCPAR.���������ϴ� And Not MCPAR.������ɺ��ϴ� Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
        End If
        
        gcnOracle.CommitTrans
        
        'ҽ�����������ϴ�
        If mbytUseType <> Use���� And intInsure <> 0 Then
            If MCPAR.���������ϴ� And MCPAR.������ɺ��ϴ� Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                End If
            End If
        End If
        
        gblnOK = True
        Unload Me: Exit Sub
    ElseIf mbytInState = sta���� Then
        If Not IsDate(txtDate.Text) Then
            MsgBox "������Ϸ��ķ���ʱ�䣡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        strInfo = Check����ʱ��(CDate(txtDate.Text), cboNO.Text)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
            
        If Not SaveModi() Then Exit Sub
        gblnOK = True: Unload Me: Exit Sub
    ElseIf chk��.Value = 0 Then '�������뵥��״̬'%%%
        If mrsInfo.State = adStateClosed Then
            MsgBox "û�з��ֲ�����Ϣ,��ȷ��������Ϣ��", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Sub
        End If
        If cbo�ѱ�.ListIndex = -1 Or mobjBill.�ѱ� = "" Then
            MsgBox "��ѡ���˷ѱ�", vbInformation, gstrSysName
            If cbo�ѱ�.Visible = True Then cbo�ѱ�.SetFocus: Exit Sub
        End If
        If mobjBill.��������ID = 0 Then
            MsgBox "��ȷ���������ң�", vbInformation, gstrSysName
            cbo��������.SetFocus
            Exit Sub
        End If
        
        If mobjBill.������ = "" And gbln������ Then
            MsgBox "�����뿪���ˣ�", vbInformation, gstrSysName
            cbo������.SetFocus: Exit Sub
        End If
        
        strSQL = ""
        For i = 0 To mlngRows - 1
            If mobjBill.Details("R" & i).�շ�ϸĿID <> 0 Then
                strSQL = "������"
                
                If mobjBill.Details("R" & i).ִ�в���ID = 0 Then
                    MsgBox "�����շ���Ŀ��ִ�в���û���ã�", vbInformation, gstrSysName
                    If cboִ�п���(i).Visible = True Then
                        cboִ�п���(i).SetFocus
                    Else
                        txt�շ���Ŀ(i).SetFocus
                    End If
                    Exit Sub
                End If
            End If
        Next
        If strSQL = "" Then
            MsgBox "������û���κ�����,����ȷ���뵥�����ݣ�", vbInformation, gstrSysName
            txt�շ���Ŀ(0).SetFocus: Exit Sub
        End If
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������ȷ�ķ������ڣ�", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        strInfo = Check����ʱ��(CDate(txtDate.Text), mrsInfo!����ID)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        If Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) > Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "ǿ�ƶԳ�Ժ���˼���ʱ������ʱ�䲻�ܴ��ڲ��˳�Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        If Not IsNull(mrsInfo!����) And Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) < Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "���õķ���ʱ�䲻��С��ҽ�����˵���Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        'ҽ���������ʼ��    ��Ϊ����Ա�������䵥��,��ȷ������,����Ҫ�ټ��һ��
        If InStr(mstrPrivsOpt, "���Ƹ�������") > 0 And (mbytUseType = UseסԺ Or mbytUseType = Use���ҷ�ɢ) Then      '����������һ�ָ�������Ȩ��,�ſ����Ǹ���
            If Not IsNull(mrsInfo!����) Then
                If Not MCPAR.�������� Then
                    For i = 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).���� < 0 Then
                                MsgBox "�����е� " & i & " ���Ǹ���,����ҽ����֧�ָ������ʣ�", vbInformation, gstrSysName
                                txtDate.SetFocus: Exit Sub
                        End If
                    Next
                End If
            End If
        End If
        '������������Ȩ�޼��
        If mbytUseType <> Use���� Then
            If Not PatiCanBilling(mrsInfo!����ID, NVL(mrsInfo!��ҳID, 0), mstrPrivsOpt) Then Exit Sub
            If zlPatiIS�����ѱ�Ŀ(mrsInfo!����ID, NVL(mrsInfo!��ҳID, 0)) = True Then Exit Sub
            If zlIsAllowFeeChange(Val(NVL(mrsInfo!����ID)), Val(NVL(mrsInfo!��ҳID))) = False Then Exit Sub             '����:49501
        End If
        
        'ҽ��������Ŀ�������
        If mbytUseType <> Use���� Then
            If Not IsNull(mrsInfo!����) Then
                If Not mrsMedAudit Is Nothing Then
                    If Not CheckExamine(mobjBill.Details, mrsMedAudit, mrsInfo!����) Then Exit Sub
                End If
                
                If MCPAR.ʵʱ��� Then
                    If gclsInsure.CheckItem(mrsInfo!����, 1, 2, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text))) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If
                
        '���ʷ��౨��
        If mbytInState = staִ�� Then
            mrsWarn.Filter = ""
            If mrsWarn.RecordCount > 0 Then
                '���ݷ���
                curTotal = CalcGridToTal
                If curTotal > 0 Then
                    '����Ԥ������Ϣ
                    Set rsTmp = GetMoneyInfo(mrsInfo!����ID, CDbl(mcurModiMoney), Val("" & mrsInfo!����) > 0)
                    If Not rsTmp Is Nothing Then
                        sta.Panels(3).Text = "Ԥ��:" & Format(rsTmp!Ԥ�����, "0.00")
                        sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(rsTmp!�������, gstrDec)
                        sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(rsTmp!Ԥ����� - rsTmp!�������, "0.00")
                        cmdOK.Tag = rsTmp!Ԥ�����
                        cmdCancel.Tag = rsTmp!�������
                        mcur���ý�� = rsTmp!Ԥ����� - rsTmp!�������
                    Else
                        sta.Panels(3).Text = "Ԥ��:0.00/����:" & gstrDec & "/ʣ��:0.00"
                        cmdOK.Tag = 0
                        cmdCancel.Tag = 0
                        mcur���ý�� = 0
                    End If
                    
                    '���¶�ȡ���ն�
                    cur���ն� = GetPatiDayMoney(mrsInfo!����ID)
                                    
                    If gbln�����������۷��� Then mcur���ý�� = mcur���ý�� - GetPriceMoneyTotal(2, mrsInfo!����ID)
                    
                    For i = 1 To mobjBill.Details.Count
                        gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!����, Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, mcur���ý��, cur���ն� - mcurModiMoney, curTotal, mobjBill.������, mobjBill.Details(i).�շ����, mobjBill.Details(i).Detail.�������, mstrWarn)
                        If gbytWarn = 2 Or gbytWarn = 3 Then Exit Sub
                    Next
                End If
            End If
        End If
        
        '��Ŀ���������(��Ҫ��Ϊ�����������۲���)
        If mbytUseType <> Use���� Then
            If Check������� > 0 Then Exit Sub
        End If
        
        If Not SaveBill Then
            Exit Sub
        Else
            If mstrInNO = "" Then
                sta.Panels(2) = "��һ�ŵ���:" & mobjBill.NO
                Call NewBill
                mstrInNO = ""
                If mlng����ID <> 0 And mbytUseType = 1 Then
                    txtPatient.Text = "-" & mlng����ID
                    Call txtPatient_KeyPress(13)
                Else
                    txtPatient.SetFocus
                End If
            Else '�޸�
                gblnOK = True: Unload Me
            End If
        End If
    ElseIf chk��.Value = 1 Then '�˵���״̬
        If mstrInNO = "" Then
            MsgBox "û�ж�ȡ��������,�������ʣ�", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If

        'ҽ�����������ϴ�(ע���ж�˳��)
        If mbytUseType <> Use���� Then
            intInsure = BillExistInsure(mstrInNO) '�ж��Ƿ�ҽ�����˼ǵ���
            'ȥ����ҽ������ƥ����
        End If

        If mbytUseType = Use���� Then
            strSQL = "zl_������ʼ�¼_DELETE('" & mstrInNO & "','','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            strSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "','','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        End If

        On Error GoTo errH
        gcnOracle.BeginTrans
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        'ҽ�����������ϴ�
        If mbytUseType <> Use���� And intInsure <> 0 Then
            If MCPAR.���������ϴ� And Not MCPAR.������ɺ��ϴ� Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
        End If
        
        gcnOracle.CommitTrans
        
        'ҽ�����������ϴ�
        If mbytUseType <> Use���� And intInsure <> 0 Then
            If MCPAR.���������ϴ� And MCPAR.������ɺ��ϴ� Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                End If
            End If
        End If
        
        On Error GoTo 0

        mstrInNO = "": cboNO.Text = ""
        txtPatient.Text = "": txt����.Text = ""
        mcur���ý�� = 0
        Call NewBill
        chk��.Value = 0
        txtPatient.SetFocus
    End If
    gblnOK = True
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo�շ����_GotFocus(Index As Integer)
    LocateItem Index, "�շ����"
    
    If cbo�շ����(Index).ListIndex = -1 And Index > 0 Then
        cbo�շ����(Index).ListIndex = cbo�շ����(Index - 1).ListIndex
    End If
End Sub

Private Sub txt�շ���Ŀ_GotFocus(Index As Integer)
    LocateItem Index, "�շ���Ŀ"
    zlControl.TxtSelAll txt�շ���Ŀ(Index)
End Sub

Private Sub txt���㵥λ_GotFocus(Index As Integer)
    LocateItem Index, "���㵥λ"
    zlControl.TxtSelAll txt���㵥λ(Index)
End Sub

Private Sub txt����_GotFocus(Index As Integer)
    LocateItem Index, "����"
    zlControl.TxtSelAll txt����(Index)
End Sub

Private Sub txt��׼����_GotFocus(Index As Integer)
    LocateItem Index, "��׼����"
    zlControl.TxtSelAll txt��׼����(Index)
End Sub

Private Sub txtʵ�ս��_GotFocus(Index As Integer)
    LocateItem Index, "ʵ�ս��"
    zlControl.TxtSelAll txtʵ�ս��(Index)
End Sub

Private Sub txtӦ�ս��_GotFocus(Index As Integer)
    LocateItem Index, "Ӧ�ս��"
    zlControl.TxtSelAll txtӦ�ս��(Index)
End Sub

Private Sub cboִ�п���_GotFocus(Index As Integer)
    LocateItem Index, "ִ�п���"
End Sub

Private Sub chk����_GotFocus(Index As Integer)
    LocateItem Index, "���ӱ�־"
End Sub

Private Sub cbo�շ����_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngIdx As Long
    
    If cbo�շ����(Index).Locked Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        Call Input�շ����(Index)
        SendKeys "{TAB}"
        Exit Sub
    End If
    
'    If SendMessage(cbo�շ����(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
    lngIdx = zlControl.CboMatchIndex(cbo�շ����(Index).hwnd, KeyAscii)
'    If lngIdx <> -2 Then cbo�շ����(Index).ListIndex = lngIdx
End Sub

Private Sub cbo�շ����_Validate(Index As Integer, Cancel As Boolean)
    If cbo�շ����(Index).Locked Then Exit Sub
    
    Call Input�շ����(Index)
End Sub

Private Function Input�շ����(ByVal Index As Long) As Boolean
    
    If cbo�շ����(Index).ListIndex <> -1 Then
        If mobjBill.Details("R" & Index).�շ���� <> Chr(cbo�շ����(Index).ItemData(cbo�շ����(Index).ListIndex)) Then
            'һ�������շ����,�����(����)ԭ�и���Ŀ����
            ClearDetail Index
            'Call CalcMoneys
            Call ShowMoney
        End If
    Else
        ClearDetail Index
    End If
    
    Input�շ���� = True
End Function

Private Sub txt�շ���Ŀ_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������
    strText = txt�շ���Ŀ(Index).Text
    If strText <> "" Then
        If mobjBill.Details("R" & Index).�շ���� = "" Then
            sta.Panels(2) = "û��ȷ���������,�����������"
            If cbo�շ����(Index).Visible = True Then
                txt�շ���Ŀ(Index).Text = ""
                If cbo�շ����(Index).Enabled Then cbo�շ����(Index).SetFocus
            End If
            Call Beep: Exit Sub
        End If
        Call GetDetails(txt�շ���Ŀ(Index).hwnd, strText, mobjBill.Details("R" & Index).�շ����)
        If Set�շ�ϸĿ(Index) = False Then
            zlControl.TxtSelAll txt�շ���Ŀ(Index)
            Exit Sub
        End If
    Else
        If mobjBill.Details("R" & Index).�շ�ϸĿID <> 0 Then
            Set�շ�ϸĿEmpty Index
        End If
        If Index = mlngRows - 1 Then
            cmdOK.SetFocus
        Else
            txt�շ���Ŀ(Index + 1).SetFocus
        End If
        Exit Sub
    End If

    SendKeys "{TAB}"
End Sub

Private Sub txt�շ���Ŀ_Validate(Index As Integer, Cancel As Boolean)
    Dim strText As String
    
    strText = txt�շ���Ŀ(Index).Text
    
    If strText = mobjBill.Details("R" & Index).�շ����� Then Exit Sub
    If Trim(strText) = "" Then
        '�մ����⴦��
        txt�շ���Ŀ(Index).Text = mobjBill.Details("R" & Index).�շ�����
        Exit Sub
    End If
    
    If strText <> "" Then
        If mobjBill.Details("R" & Index).�շ���� = "" Then
            sta.Panels(2) = "û��ȷ���������,�����������"
            Set mcolDetails = New Details
            Set�շ�ϸĿEmpty Index
            Exit Sub
        End If
        '���������ж�
        Call GetDetails(txt�շ���Ŀ(Index).hwnd, strText, mobjBill.Details("R" & Index).�շ����)
        If Set�շ�ϸĿ(Index) = False Then
            zlControl.TxtSelAll txt�շ���Ŀ(Index)
            Cancel = True
        End If
    End If
End Sub

Private Function Set�շ�ϸĿ(ByVal Index As Integer) As Boolean
    Dim lngDoUnit As Long, curTotal As Currency
    Dim int������Դ As Integer, curItemMoney As Currency
    
    If mcolDetails.Count = 0 Then
        sta.Panels(2) = "�Ҳ�����Ӧ���շ���Ŀ,��ȷ�������Ƿ���ȷ��"
        Call Beep: Exit Function
    ElseIf mcolDetails.Count = 1 Then
        'ȷ�����շ�ϸĿ
        Set mobjDetail = mcolDetails(1)
        
        'һЩ������Ŀ�ĺϷ��Լ��
        '����֧����Ŀ��Ӧ���
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!����) Then
                If Not CheckMediCareItem(mobjDetail.ID, mrsInfo!����, mobjDetail.����, mobjDetail.��� = False, mstrPriceGrade) Then
                    txt�շ���Ŀ(Index).Text = mobjBill.Details("R" & Index).Detail.����
                    Exit Function
                End If
                
                'ҽ�����˷�����ĿҪ������
                If mbytUseType <> Use���� Then
                    If mobjDetail.Ҫ������ And Not mrsMedAudit Is Nothing Then
                        mrsMedAudit.Filter = "��ĿID=" & mobjDetail.ID
                        If mrsMedAudit.RecordCount = 0 Then
                            MsgBox "��ǰ����δ����׼ʹ�ø���Ŀ��", vbInformation, gstrSysName
                            txt�շ���Ŀ(Index).Text = "": Exit Function
                        ElseIf Not IsNull(mrsMedAudit!��������) Then
                            If mrsMedAudit!�������� <= 0 Then
                                MsgBox "��ǰ����ʹ��[" & mobjDetail.���� & "]�Ѵﵽ��׼��ʹ������" & mrsMedAudit!ʹ������ & "��", vbInformation, gstrSysName
                                txt�շ���Ŀ(Index).Text = "": Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If mobjDetail.ID = mobjBill.Details("R" & Index).Detail.ID Then
           '��Ȼ����ǰ���Ǹ��������ò����ٸı���
           txt�շ���Ŀ(Index).Text = mobjDetail.����
           Set�շ�ϸĿ = True
           Exit Function
        End If
        
        '������Դ
        If mbytUseType = Use���� Then
            int������Դ = 1
        Else
            If mrsInfo.State = 1 Then
                '��ȡ����ʱ�Ѹ���Ȩ�������Ƿ����۲���
                If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
                    int������Դ = 2
                ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
                    int������Դ = 1
                End If
            Else
                int������Դ = 2
            End If
        End If
        '���˿���
        lngDoUnit = mobjBill.����ID
        If lngDoUnit = 0 Then lngDoUnit = Get��������ID
        lngDoUnit = Get�շ�ִ�п���ID(mobjDetail.ID, mobjDetail.ִ�п���, lngDoUnit, Get��������ID, int������Դ, mobjBill.����ID)

        '������޸ĸ��շ�ϸĿ��
        With mobjBill.Details("R" & Index)
            Set .Detail = mobjDetail
            Set .InComes = New BillInComes
            .���ӱ�־ = 0
            .���㵥λ = mobjDetail.���㵥λ
            .�շ���� = mobjDetail.���
            .�շ�ϸĿID = mobjDetail.ID
            .�շ����� = mobjDetail.����
            
            If txt����(Index).Tag <> "" Then
                .���� = Val(txt����(Index).Tag)
            Else
                .���� = 1
            End If
            .ִ�в���ID = lngDoUnit
            '���㵥�ۺͽ��
            Call CalcMoney(Index)
            
            '��Ŀ��ֵ��Ԥ���
        End With
        
        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
        If mbytInState = staִ�� Then
            mrsWarn.Filter = ""
            If mrsInfo.State = 1 And mrsWarn.RecordCount > 0 Then
                curTotal = GetBillTotal(mobjBill)
                If curTotal > 0 Then
                    If gbln�����������۷��� Then mcur���ý�� = mcur���ý�� - GetPriceMoneyTotal(2, mrsInfo!����ID)
                    
                    '���˺�:24491
                    curItemMoney = GetBillRowTotal(mobjBill.Details("R" & Index).InComes)
                    gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!����, Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, mcur���ý��, mrsInfo!���ն� - mcurModiMoney, curTotal, mobjBill.������, mobjDetail.���, mobjDetail.�������, mstrWarn, , curItemMoney)
                    If gbytWarn = 2 Or gbytWarn = 3 Then
                        ClearDetail Index 'ɾ���ո���Ҫ����ķ�����
                        txt�շ���Ŀ(Index).Text = ""
                        Exit Function
                    End If
                End If
            End If
        End If
        If mrsInfo.State = 1 And mbytUseType <> Use���� Then
            If Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� Then
                If gclsInsure.CheckItem(mrsInfo!����, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), Index)) = False Then
                    ClearDetail Index 'ɾ���ո���Ҫ����ķ�����
                    txt�շ���Ŀ(Index).Text = ""
                    Exit Function
                End If
            End If
        End If
        

        Call ShowDetails(Index)
        Call ShowMoney
    End If

    If mobjBill.Details("R" & Index).Detail.��� Then
        txt����(Index).TabStop = gblnTime
        txt����(Index).Locked = Not gblnTime
        txt��׼����(Index).TabStop = True
        txt��׼����(Index).Locked = False
    Else
        txt����(Index).TabStop = True
        txt����(Index).Locked = False
        txt��׼����(Index).TabStop = False
        txt��׼����(Index).Locked = True
    End If
    chk����(Index).Enabled = mobjBill.Details("R" & Index).�շ���� = "F" '����
    If chk����(Index).Enabled = False Then chk����(Index).Value = 0
    
    'ִ�п���!!!
    Call Fillִ�п���(Index)
    
    If cboִ�п���(Index).ListCount = 1 Then
        cboִ�п���(Index).TabStop = False
    Else
        cboִ�п���(Index).TabStop = True
    End If
        
    'byZT200302
    'cboִ�п���(Index).SelLength = 0
    
    Set�շ�ϸĿ = True
End Function

Private Function Set�շ�ϸĿEmpty(Index As Integer) As Boolean
    Dim lngDoUnit As Long
    

    mobjBill.Details.Remove "R" & Index
    mobjBill.Details.AddEmpty Index + 1
    mobjBill.Details("R" & Index).�շ���� = cbo�շ����(Index).Tag
    
    If Val(txt�շ���Ŀ(Index).Tag) > 0 Then
        Call GetInputDetail(Val(txt�շ���Ŀ(Index).Tag))
        Call Set�շ�ϸĿ(Index)
    Else
        Call ShowDetails(Index)
    End If
    Call ShowMoney

    txt����(Index).TabStop = False
    txt����(Index).Locked = False
    txt��׼����(Index).TabStop = False
    txt��׼����(Index).Locked = False
    chk����(Index).Enabled = False
    chk����(Index).Value = 0
    
    Set�շ�ϸĿEmpty = True
End Function

Private Sub cmdϸĿѡ��_Click(Index As Integer)
    Dim strSQL As String, str��׼��Ŀ As String
    Dim str��� As String, lng��Ŀid As Long
    Dim int������Դ As Integer, int���� As Integer
    
    Call LocateItem(Index, "ϸĿѡ��")
    
    '�շ����
    str��� = mobjBill.Details("R" & Index).�շ����
    If str��� <> "" Then str��� = "'" & str��� & "'"
        
    '������Դ
    If mbytUseType = -1 Then '���
        int������Դ = 0
    ElseIf mbytUseType = Use���� Then
        int������Դ = 1
    Else
        If mrsInfo.State = 1 Then
            '��ȡ����ʱ�Ѹ���Ȩ�������Ƿ����۲���
            If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
                int������Դ = 2
            ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
                int������Դ = 1
            End If
        Else
            'δȷ������,������,�ڱ���ʱ���
            If (InStr(mstrPrivsOpt, "�������ۼ���") > 0 And gbln��������) Or mbytUseType = 2 Then
                int������Դ = 0
            Else
                int������Դ = 2
            End If
        End If
    End If
    If mbytUseType <> -1 Then
        'ҽ��������׼��Ŀ
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!����) Then
                int���� = mrsInfo!����
                '���˺�:24862
                If zl_Check��׼��Ŀ(gclsInsure, int����, Val(NVL(mrsInfo!����ID)), False) Then str��׼��Ŀ = Get������׼��Ŀ(Val(NVL(mrsInfo!����ID)), "A.ID")
            End If
        End If
    End If
    
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, str���, , , str��׼��Ŀ, mstrPriceGrade)
    Me.Refresh
    txt�շ���Ŀ(Index).SetFocus
    
    If lng��Ŀid <> 0 Then
        If lng��Ŀid = mobjBill.Details("R" & Index).�շ�ϸĿID Then
            SendKeys "{TAB}": Exit Sub
        End If
        Call GetInputDetail(lng��Ŀid)
        If Not Set�շ�ϸĿ(Index) Then Exit Sub
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txt����_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If txt����(Index).Locked = True Then Exit Sub
    If KeyAscii <> vbKeyReturn Then
        If InStr("-.0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    If txt����(Index).Text = "" Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    If IsNumeric(txt����(Index).Text) = False Then
        MsgBox "������Ϸ�������", vbExclamation, gstrSysName
        zlControl.TxtSelAll txt����(Index)
        Exit Sub
    End If
    If Input����(Index) = True Then
        SendKeys "{TAB}"
    Else
        zlControl.TxtSelAll txt����(Index)
    End If
End Sub

Private Sub txt����_Validate(Index As Integer, Cancel As Boolean)
    If txt����(Index).Locked Then Exit Sub
    
    If Not IsNumeric(txt����(Index).Text) Or Val(txt����(Index).Text) > 100000 Then
        If mobjBill.Details("R" & Index).���� = 0 Then
            txt����(Index).Text = ""
        Else
            txt����(Index).Text = mobjBill.Details("R" & Index).����
        End If
        Exit Sub
    Else
        If CSng(txt����(Index).Text) = mobjBill.Details("R" & Index).���� Then Exit Sub
    End If
    
    If Input����(Index) Then
        Cancel = True
    End If
End Sub

Private Function Input����(ByVal Index As Long) As Boolean
    Dim sngPreTime  As Single, sngItemNum As Single
    Dim sngInput As Single, curTotal As Currency, curItemMoney As Currency
    Dim dbl�������� As Double
    
    With mobjBill.Details("R" & Index)
        If Val(txt����(Index).Text) > 100000 Then
            MsgBox "����ֵ�������", vbInformation, gstrSysName
            txt����(Index).Text = .����
            Exit Function
        End If
        If .Detail.¼������ > 0 And Val(txt����(Index).Text) > .Detail.¼������ Then
            If MsgBox("��������γ�����¼������" & .Detail.¼������ & ",�Ƿ����?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                txt����(Index).Text = .����
                Exit Function
            End If
        End If
        '��������
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!����) And .Detail.Ҫ������ And Not mrsMedAudit Is Nothing Then
                mrsMedAudit.Filter = "��ĿID=" & .�շ�ϸĿID
                If mrsMedAudit.RecordCount > 0 Then
                    If Not IsNull(mrsMedAudit!��������) Then
                        If Val(txt����(Index).Text) > mrsMedAudit!�������� Then
                            MsgBox "��������γ�������׼��ʹ������" & mrsMedAudit!�������� & "��", vbInformation, gstrSysName
                            txt����(Index).Text = .����
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If

        '�������
        If gcurMaxMoney > 0 Then
            If CSng(txt����(Index).Text) * Val(txt��׼����(Index).Text) > gcurMaxMoney Then
                If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                    txt����(Index).Text = .����
                    Exit Function
                End If
            End If
        End If
    End With
    
    
    sngInput = Format(Val(txt����(Index).Text), "0.000")
    If sngInput < 0 Then
        '����Ȩ�޼��
        If (mbytUseType = UseסԺ Or mbytUseType = Use���ҷ�ɢ) Then
            If InStr(mstrPrivsOpt, "���Ƹ�������") = 0 Then
                MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                txt����(Index).Text = mobjBill.Details("R" & Index).����
                Exit Function
            Else
                If mrsInfo.State = 1 Then
                    If Not IsNull(mrsInfo!����) Then
                        If Not MCPAR.�������� Then
                            MsgBox "����ҽ����֧�ֶ�ҽ�����˽��и������ʣ�", vbInformation, gstrSysName
                            txt����(Index).Text = mobjBill.Details("R" & Index).����
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        '����:26951
         If InStr(1, mstrPrivsOpt, ";�������ʲ���鷢����Ŀ;") = 0 Then
             '���ڸ�������ʱ����鱾��סԺ��������Ŀ����,�д�Ȩ��,����¼�벡��δ�������ķ�����Ŀ���г���,�����鱾��סԺ��������Ŀ�������ܳ���
            '�����Ϸ��Լ��
            sngItemNum = GetDetailNum(Index, dbl��������)
            '32106
            If Abs(sngInput) > sngItemNum - dbl�������� Then
                Select Case gbytBillOpt '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
                Case 0  '����
                    If Abs(sngInput) > sngItemNum Then
                        MsgBox "����Ŀ��������������������[" & sngItemNum & "]��", vbInformation, gstrSysName
                        txt����(Index).Text = mobjBill.Details("R" & Index).����
                        Exit Function
                    End If
                Case 1   '����
                    If Abs(sngInput) > sngItemNum Then
                        MsgBox "����Ŀ��������������������[" & sngItemNum & "]��", vbInformation, gstrSysName
                        txt����(Index).Text = mobjBill.Details("R" & Index).����
                        Exit Function
                    End If
                    If MsgBox("����Ŀ���������а������ѽᲿ��(δ��:" & Round(sngItemNum - dbl��������, 5) & "; �ѽ�:" & Round(dbl��������, 5) & ") ��" & vbCrLf & _
                        " �Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        txt����(Index).Text = mobjBill.Details("R" & Index).����
                        Exit Function
                    End If
                Case 2   '��ֹ
                    MsgBox "����Ŀ��������������������[" & sngItemNum & "]��", vbInformation, gstrSysName
                    txt����(Index).Text = mobjBill.Details("R" & Index).����
                    Exit Function
                End Select
            End If
         End If
    End If

    '��¼����ǰ�������Ա�ȡ������
    sngPreTime = mobjBill.Details("R" & Index).����
    '���ĸ�������
    mobjBill.Details("R" & Index).���� = sngInput
    Call CalcMoneys(Index)

    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
    mrsWarn.Filter = ""
    If mrsInfo.State = 1 And mrsWarn.RecordCount > 0 Then
        curTotal = GetBillTotal(mobjBill)
        If curTotal > 0 Then
    
            If gbln�����������۷��� Then mcur���ý�� = mcur���ý�� - GetPriceMoneyTotal(2, mrsInfo!����ID)
            '���˺�:24491
            curItemMoney = GetBillRowTotal(mobjBill.Details("R" & Index).InComes)
            
            gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!����, Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, mcur���ý��, mrsInfo!���ն� - mcurModiMoney, curTotal, mobjBill.������, mobjBill.Details("R" & Index).�շ����, mobjBill.Details("R" & Index).Detail.�������, mstrWarn, , curItemMoney)
            If gbytWarn = 2 Or gbytWarn = 3 Then
                mobjBill.Details("R" & Index).���� = sngPreTime
                txt����(Index).Text = sngPreTime
                Call CalcMoneys(Index)
                Exit Function
            End If
        End If
    End If
    
    If mrsInfo.State = 1 And mbytUseType <> Use���� Then
        If Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� Then
            If gclsInsure.CheckItem(mrsInfo!����, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), Index)) = False Then
                mobjBill.Details("R" & Index).���� = sngPreTime
                txt����(Index).Text = sngPreTime
                Call CalcMoneys(Index)
                Exit Function
            End If
        End If
    End If

    Call ShowDetails(Index)
    Call ShowMoney
    Input���� = True
End Function

Private Sub txt��׼����_KeyPress(Index As Integer, KeyAscii As Integer)
    If txt��׼����(Index).Locked Then Exit Sub
    
    If KeyAscii <> 13 Then
        If InStr(".0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    
    If Not IsNumeric(txt��׼����(Index).Text) Then
        MsgBox "������Ϸ����ۡ�", vbExclamation, gstrSysName
        zlControl.TxtSelAll txt��׼����(Index)
        Exit Sub
    End If
    
    If Input��׼����(Index) Then
        SendKeys "{TAB}"
    Else
        zlControl.TxtSelAll txt��׼����(Index)
    End If
End Sub

Private Sub txt��׼����_Validate(Index As Integer, Cancel As Boolean)
    If txt��׼����(Index).Locked Then Exit Sub
    
    If Not IsNumeric(txt��׼����(Index).Text) Then
        txt��׼����(Index).Text = Format(mobjBill.Details("R" & Index).��׼����, "0,000")
        Exit Sub
    Else
        If CCur(txt��׼����(Index).Text) = mobjBill.Details("R" & Index).��׼���� Then Exit Sub
    End If
    
    If Not Input��׼����(Index) Then
        Cancel = True
    End If
End Sub

Private Function Input��׼����(ByVal Index As Long) As Boolean
    Dim strScope As String, curTotal As Currency
    Dim curPreMoney As Currency, curItemMoney As Currency
    
    '���û�ж�Ӧ��������Ŀ,���޷�����
    If mobjBill.Details("R" & Index).Detail.��� And mobjBill.Details("R" & Index).InComes.Count > 0 Then
        '���۲��������븺��
        If Val(txt��׼����(Index).Text) < 0 Then
            MsgBox "��Ŀ�۸�Ӧ��Ϊ������Ҫ�������ã������븺��������ʵ�֣�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '��������뷶Χ
        If Not (mobjBill.Details("R" & Index).InComes(1).�ּ� = 0 And mobjBill.Details("R" & Index).InComes(1).ԭ�� = 0) Then
            strScope = CheckScope(mobjBill.Details("R" & Index).InComes(1).ԭ��, mobjBill.Details("R" & Index).InComes(1).�ּ�, CCur(txt��׼����(Index).Text))
            If strScope <> "" Then
                sta.Panels(2) = strScope
                Exit Function
            End If
        End If
        '�������
        If gcurMaxMoney > 0 Then
            If Val(txt��׼����(Index).Text) * Val(mobjBill.Details("R" & Index).����) > gcurMaxMoney Then
                If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                    Exit Function
                End If
            End If
        End If

        curPreMoney = mobjBill.Details("R" & Index).InComes(1).��׼����

        mobjBill.Details("R" & Index).InComes(1).��׼���� = txt��׼����(Index).Text '�����շ�ϸĿֻ�ܶ�Ӧһ��������Ŀ
        Call CalcMoneys(Index)

        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
        mrsWarn.Filter = ""
        If mrsInfo.State = 1 And mrsWarn.RecordCount > 0 Then
            curTotal = GetBillTotal(mobjBill)
            If curTotal > 0 Then
        
                If gbln�����������۷��� Then mcur���ý�� = mcur���ý�� - GetPriceMoneyTotal(2, mrsInfo!����ID)
                '���˺�:24491
                curItemMoney = GetBillRowTotal(mobjBill.Details("R" & Index).InComes)
                gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!����, Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, mcur���ý��, mrsInfo!���ն� - mcurModiMoney, curTotal, mobjBill.������, mobjBill.Details("R" & Index).�շ����, mobjBill.Details("R" & Index).Detail.�������, mstrWarn)
                If gbytWarn = 2 Or gbytWarn = 3 Then
                    mobjBill.Details("R" & Index).InComes(1).��׼���� = curPreMoney
                    txt��׼����(Index).Text = Format(curPreMoney, "0.0000")
                    If Val(txt��׼����(Index).Text) = 0 Then txt��׼����(Index).Text = ""
                    
                    Call CalcMoneys(Index)
                    Exit Function
                End If
            End If
        End If

        Call ShowDetails(Index)
        Call ShowMoney
    Else
        txt��׼����(Index).Text = "0"
        sta.Panels(2) = "����Ŀ�������ö�Ӧ�ķ�Ŀ�������޷�������ã�"
        Beep
    End If
    Input��׼���� = True
End Function

Private Sub cboִ�п���_KeyPress(Index As Integer, KeyAscii As Integer)
        

    Dim lngIdx As Long, lngҽ��ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cboִ�п���(Index).ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    Fillִ�п��� Index, True
    If zlSelectDept(Me, 0, cboִ�п���(Index), mrsUnit, cboִ�п���(Index)) = False Then
        KeyAscii = 0: Exit Sub
    End If
    
'    Exit Sub
'    Dim lngIdx As Long
'    If KeyAscii = 13 Then
'        KeyAscii = 0 'byZT200302
'        If cboִ�п���(Index).ListIndex <> -1 Then
'            mobjBill.Details("R" & Index).ִ�в���ID = cboִ�п���(Index).ItemData(cboִ�п���(Index).ListIndex)
'        End If
'        SendKeys "{TAB}"
'        Exit Sub
'    End If
'
'    If cboִ�п���(Index).Locked Then Exit Sub
'    If SendMessage(cboִ�п���(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
'    lngIdx = MatchIndex(cboִ�п���(Index).hwnd, KeyAscii)
'    If lngIdx <> -2 Then cboִ�п���(Index).ListIndex = lngIdx
'
'    'byZT200302
'    If cboִ�п���(Index).ListIndex = -1 And cboִ�п���(Index).ListCount <> 0 Then cboִ�п���(Index).ListIndex = 0
End Sub

Private Sub cboִ�п���_Validate(Index As Integer, Cancel As Boolean)
    If cboִ�п���(Index).ListIndex = -1 And cboִ�п���(Index).ListCount <> 0 Then cboִ�п���(Index).ListIndex = 0
    
    If cboִ�п���(Index).ListIndex <> -1 Then
        mobjBill.Details("R" & Index).ִ�в���ID = cboִ�п���(Index).ItemData(cboִ�п���(Index).ListIndex)
    Else
        mobjBill.Details("R" & Index).ִ�в���ID = 0
    End If
    
End Sub

Private Sub chk����_Click(Index As Integer)
'˵��������ȫ��Ϊ��Ҫ����,������ȫ��Ϊ��������
    Dim i As Long, strCheck As String, bytTime As Byte

    '������δ��������Ч
    
    For i = 0 To mlngRows - 1
        If mobjBill.Details("R" & i).�շ���� = "F" And chk����(i).Value = 0 And i <> Index Then bytTime = bytTime + 1
    Next
    If bytTime > 0 Then
        mobjBill.Details("R" & Index).���ӱ�־ = chk����(Index).Value
        Call CalcMoneys(Index)
        Call ShowDetails(Index)
        Call ShowMoney
    ElseIf chk����(Index).Value = 1 Then
        chk����(Index) = 0
        MsgBox "�����б�Ȼ��һ���������Ǹ���������", vbInformation, gstrSysName
        Exit Sub
    End If
    mobjBill.Details("R" & Index).���ӱ�־ = chk����(Index).Value
End Sub

Private Sub Fillִ�п���(ByVal lngRow As Long, Optional blnNotLoad As Boolean = False)
'���ܣ����ݵ��������������б������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String, i As Long
    Dim lng����ID As Long, lng����ID As Long, int������Դ As Integer
    
    If mbytInState <> staִ�� Then Exit Sub

    '������Դ
    If mbytUseType = Use���� Then
        int������Դ = 1
    Else
        If mrsInfo.State = 1 Then
            '��ȡ����ʱ�Ѹ���Ȩ�������Ƿ����۲���
            If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
                int������Դ = 2
            ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
                int������Դ = 1
            End If
        Else
            int������Դ = 2
        End If
    End If
    
    '���˿���
    lng����ID = mobjBill.����ID
    If lng����ID = 0 Then lng����ID = Get��������ID
    
    If int������Դ = 1 Then
        lng����ID = lng����ID
    Else
        lng����ID = mobjBill.����ID
        If lng����ID = 0 Then lng����ID = Get����ID(lng����ID)
    End If
    
    '0-����ȷ,1-���˿���,2-���˲���,3-�����˿���,4-ָ������
    Select Case mobjBill.Details("R" & lngRow).Detail.ִ�п���
        Case 0 '����ȷ
            mrsUnit.Filter = 0
        Case 1 '���˿���
            mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & mobjBill.Details("R" & lngRow).ִ�в���ID
        Case 2 '���˲���
            mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & mobjBill.Details("R" & lngRow).ִ�в���ID
        Case 3 '����Ա���ڿ���
            mrsUnit.Filter = "ID=" & IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID) & " Or ID=" & mobjBill.Details("R" & lngRow).ִ�в���ID
        Case 4 'ָ������
            strSQL = "" & _
            "   Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
            "   From �շ�ִ�п��� A,���ű� C" & _
            "   Where A.�շ�ϸĿID=[1]��And A.ִ�п���ID+0=C.ID " & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And (A.������Դ is NULL Or A.������Դ=[3])" & _
            "       And (A.��������ID is NULL Or A.��������ID=[2])" & _
            " Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details("R" & lngRow).�շ�ϸĿID, lng����ID, int������Դ)
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    strTmp = strTmp & "ID=" & rsTmp!ִ�п���ID & " OR "
                    rsTmp.MoveNext
                Next
                strTmp = strTmp & "ID=" & mobjBill.Details("R" & lngRow).ִ�в���ID & " OR "
                strTmp = Left(strTmp, Len(strTmp) - 4)
                mrsUnit.Filter = strTmp
            Else
                mrsUnit.Filter = "ID=" & IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID) & " Or ID=" & mobjBill.Details("R" & lngRow).ִ�в���ID
            End If
        Case 5 'Ժ��ִ��(Ԥ��,������δ��)
        Case 6 '�����˿���
           mrsUnit.Filter = "ID=" & Get��������ID & " Or ID=" & mobjBill.Details("R" & lngRow).ִ�в���ID
    End Select
    If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID) & " Or ID=" & mobjBill.Details("R" & lngRow).ִ�в���ID
    If blnNotLoad = True Then Exit Sub
    
    With cboִ�п���(lngRow)
        .Clear
        For i = 1 To mrsUnit.RecordCount
            strTmp = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
            If Not (SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                .AddItem strTmp
                .ItemData(.NewIndex) = mrsUnit!ID
                
                'ȱʡΪ��������
                If lngRow = 1 Then
                    If mrsUnit!ID = mobjBill.��������ID Then .ListIndex = .NewIndex
                '�������һ��ִ�п���һ��
                ElseIf lngRow > 1 Then
                    If mrsUnit!ID = mobjBill.Details("R" & (lngRow - 1)).ִ�в���ID And mobjBill.Details("R" & lngRow).Detail.ִ�п��� = mobjBill.Details("R" & (lngRow - 1)).Detail.ִ�п��� Then
                        .ListIndex = .NewIndex
                    ElseIf mrsUnit!ID = mobjBill.��������ID And .ListIndex = -1 Then
                        .ListIndex = .NewIndex
                    End If
                End If
            End If
            mrsUnit.MoveNext
        Next
        
        If mobjBill.Details("R" & lngRow).Detail.ִ�п��� = 4 Then   'ִ�п���Ϊָ�����ҵ�,ȱʡΪ����Ա���ڿ���
            For i = 0 To .ListCount - 1
                If .ItemData(i) = UserInfo.����ID Then .ListIndex = i: Exit For
            Next
        End If
        
        If .ListIndex = -1 Then '���û����ȡ���е�ִ�п���
            For i = 0 To .ListCount - 1
                If .ItemData(i) = mobjBill.Details("R" & lngRow).ִ�в���ID Then .ListIndex = i: Exit For
            Next
        End If
        
        If .ListIndex = -1 And .ListCount <> 0 Then .ListIndex = 0
        If .ListIndex <> -1 Then
            mobjBill.Details("R" & lngRow).ִ�в���ID = .ItemData(.ListIndex)
        Else
            mobjBill.Details("R" & lngRow).ִ�в���ID = 0
        End If
        
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub
Private Sub LocateItem(ByVal Index As Integer, ByVal Item As String)
    mintCurrentRow = Index
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'��������Ϊ�����޸�״̬
    cboNO.Locked = Not bln
    txtPatient.Locked = Not bln
    cbo��������.Locked = Not bln
    cbo������.Locked = Not bln

    chk�Ӱ�.Enabled = bln
    txtDate.Enabled = bln
    fraForm.Enabled = bln
End Sub

Private Function GetPatientIn(ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
'���ܣ���ȡ������Ϣ
'������blnCard=�Ƿ���￨ˢ��
    Dim strSQL As String, strIF As String
    
    On Error GoTo errH
        
    'a.�Ƿ����ǿ�Ƽ���Ȩ��
    If InStr(mstrPrivsOpt, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(mstrPrivsOpt, "��Ժ����ǿ�Ƽ���") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, "��Ժδ��ǿ�Ƽ���") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, "��Ժ����ǿ�Ƽ���") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)=0)"
    Else
        strIF = " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3"
    End If
    
    'b.�Ƿ���Լ����в�������
     If (mbytUseType = UseסԺ Or mbytUseType = Use���ҷ�ɢ) And InStr(mstrPrivs, "���в���") <= 0 Then
        If InStr(1, mstrUnitIDs, ",") = 0 Then
            strIF = strIF & " And B.��ǰ����ID+0=[3]"
        Else
            strIF = strIF & " And B.��ǰ����ID+0 IN(Select * From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
       
    'c.�Ƿ����۲��˼���Ȩ��
    If (InStr(mstrPrivsOpt, "�������ۼ���") > 0 And gbln��������) And (InStr(mstrPrivsOpt, "סԺ���ۼ���") > 0 And gblnסԺ����) Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, "�������ۼ���") > 0 And gbln�������� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, "סԺ���ۼ���") > 0 And gblnסԺ���� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,2)"
    Else
        strIF = strIF & " And Nvl(B.��������,0)=0"
    End If
    
    '���ñ�־-->��˱�־:58629
    strSQL = _
            "Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����,C.���� as ����,D.���� as ����," & _
            "   A.���￨��,A.����֤��,A.סԺ�� as ��ʶ��,B.��Ժ���� as ����, " & _
            "   nvl(B.����,A.����) as ����,nvl(B.�Ա�,A.�Ա�) as �Ա�,A.����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
            "   A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,zl_PatiDayCharge(A.����ID) as ���ն�, " & _
            "   Zl_Patiwarnscheme(B.����id, B.��ҳid) As ���ò���,B.����,Nvl(B.��������,0) as ��������,b.��˱�־,B.��������" & _
            " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,������� X " & _
            " Where A.����ID=B.����ID And A.סԺ����=B.��ҳID And B.��ǰ����ID=C.ID(+) And B.��Ժ����ID=D.ID(+) " & _
            " And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) And A.ͣ��ʱ�� is NULL " & strIF
        If mbytUseType <> Use���� Then '����:49501
            strSQL = strSQL & " And X.����(+)=2 and X.����(+)=1"
        Else
            strSQL = strSQL & " And X.����(+)=1 and X.����(+)=1"
        End If
    If blnCard Then '���￨��
        strInput = UCase(strInput)
        strSQL = strSQL & " And A.���￨��=[2]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "/" And IsNumeric(Mid(strInput, 2)) Then '��λ��
        If mlngUnitID = 0 Then '������ȷ��������ͨ������ȷ������
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = _
            "Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����," & _
            "   A.���￨��,A.����֤��,A.סԺ�� as ��ʶ��,B.��Ժ���� as ����," & _
            "   nvl(B.����,A.����) as ����,nvl(B.�Ա�,A.�Ա�) as �Ա�,A.����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
            "   A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,zl_PatiDayCharge(A.����ID) as ���ն�," & _
            "   Zl_Patiwarnscheme(B.����id, B.��ҳid) As ���ò���,B.����,Nvl(B.��������,0) as ��������,b.��˱�־,B.��������" & _
            " From ������Ϣ A,������ҳ B,��λ״����¼ C,������� X" & _
            " Where A.����ID=B.����ID And A.סԺ����=B.��ҳID" & _
            "   And Nvl(B.��ҳID,0)<>0 And A.����ID=C.����ID And A.����ID=X.����ID(+) And A.ͣ��ʱ�� is NULL " & _
            "   And C.����ID=[3] And C.����=[1] " & strIF
            
        If mbytUseType <> Use���� Then  '����:49501
            strSQL = strSQL & " And X.����(+)=2 and X.����(+)=1"
        Else
            strSQL = strSQL & " And X.����(+)=1 and X.����(+)=1"
        End If
            
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strSQL = strSQL & " And A.סԺ��=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(ҽ������)
        strSQL = strSQL & " And A.�����=[1]"
    Else '��������
        strSQL = strSQL & " And A.����=[2]"
    End If
        
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput, mlngUnitID, mstrUnitIDs)
    
    txtPatient.ForeColor = Me.ForeColor
    If Not mrsInfo.EOF Then
        If zlPatiIS�����ѱ�Ŀ(Val(NVL(mrsInfo!����ID)), NVL(mrsInfo!��ҳID, 0)) Then
            Set mrsInfo = Nothing
            Set mrsMedAudit = Nothing
            blnOutMsg = True
            Exit Function
        End If
        If mbytUseType <> Use���� Then
            '����:49501
            If zlIsAllowFeeChange(Val(NVL(mrsInfo!����ID)), Val(NVL(mrsInfo!��ҳID)), Val(NVL(mrsInfo!��˱�־))) = False Then
                Set mrsInfo = Nothing
                Set mrsMedAudit = Nothing
                blnOutMsg = True
                Exit Function
            End If
        End If
        
        txtPatient.ForeColor = zlDatabase.GetPatiColor(NVL(mrsInfo!��������))
        Set mrsMedAudit = GetAuditRecord(mrsInfo!����ID, mrsInfo!��ҳID)
        GetPatientIn = True: Exit Function
    Else
        Set mrsMedAudit = Nothing   'ҽ�����˱�����Ժ�ż���������
    End If
    
        
    'ҽ�����Ҽ��ʣ�û�з���סԺ(��Ժ���Ժ)����,�����ﲡ�˶�
    If mbytUseType = 2 And InStr(mstrPrivsOpt, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(mstrPrivsOpt, "��Ժ����ǿ�Ƽ���") > 0 Then
        strSQL = _
            "Select A.����ID,Nvl(A.סԺ����,0) ��ҳID,A.��ǰ����ID as ����ID,A.��ǰ����ID as ����ID," & _
            " A.��Ժʱ�� as ��Ժ����,A.���￨��,A.����֤��,A.סԺ��,A.��ǰ���� as ����,A.����,A.�Ա�,A.����," & _
            " A.��Ժʱ�� as ��Ժ����,A.�ѱ�,A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,null)) ������" & _
            ", Zl_Patiwarnscheme(A.����id) As ���ò���,NULL as סԺҽʦ,A.ҽ�Ƹ��ʽ," & _
            " zl_PatiDayCharge(A.����ID) as ���ն�,A.����,-1 as ��������" & _
            " From ������Ϣ A Where A.ͣ��ʱ�� is NULL "
        If blnCard Then '���￨��
            strSQL = strSQL & " And A.���￨��=[2]"
        ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
            strSQL = strSQL & " And A.����ID=[1]"
        ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(ҽ������)
            strSQL = strSQL & " And A.�����=[1]"
        Else '��������
            strSQL = strSQL & " And A.����=[2]"
        End If
        
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
        If Not mrsInfo.EOF Then
                If zlPatiIS�����ѱ�Ŀ(Val(NVL(mrsInfo!����ID)), NVL(mrsInfo!��ҳID, 0)) Then
                    Set mrsInfo = Nothing
                    blnOutMsg = True
                    Exit Function
                End If
        
            GetPatientIn = True
        Else
            Set mrsInfo = New ADODB.Recordset
        End If
    Else
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function


Private Function GetPatientOut(ByVal strInput As String) As Boolean
'���ܣ���ȡ������Ϣ
'˵�������﹦����Ժ,סԺ���ܳ�Ժ,�����ܶ�ȡ������Ϣ
'�ֶ��б�����ID,��ҳID,����ID,����ID,����,����,��Ժ����,���￨��,��ʶ��,����,����,�Ա�,����,�ѱ�,������
    Dim strSQL As String, strNO As String
    On Error GoTo errH
    
    strSQL = _
    " Select A.����ID,Nvl(A.סԺ����,0) ��ҳID,0 as ����ID,0 as ����ID,'' as ����,'' as ����," & _
    "       A.��Ժʱ�� as ��Ժ����,A.��Ժʱ�� as ��Ժ����,A.���￨��,A.����֤��,A.����� as ��ʶ��,'' as ����,A.����,A.�Ա�,A.����," & _
    "       A.�ѱ�,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,null)) ������,A.����,A.ҽ�Ƹ��ʽ," & _
    "       zl_PatiDayCharge(A.����ID) as ���ն�, Zl_Patiwarnscheme(A.����id) As ���ò���,-1 as ��������" & _
    " From ������Ϣ A " & _
    " Where A.ͣ��ʱ�� is NULL And Nvl(��ǰ����ID,0)=0 "
    
    If mblnCard Then '���￨��
        strInput = UCase(strInput)
        strSQL = strSQL & " and A.���￨��=[2]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " and A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " and A.�����=[1]"
    ElseIf Left(strInput, 1) = "." And IsNumeric(Mid(strInput, 2)) Then '�Һŵ���(�����Ϊִ�в���ID)
        strNO = GetFullNO(Mid(strInput, 2), 12)
        strSQL = _
        "Select A.����ID,0 ��ҳID,0 as ����ID,ִ�в���ID as ����ID,'' as ����,'' as ����," & _
        "       B.��Ժʱ�� as ��Ժ����,B.��Ժʱ�� as ��Ժ����,B.���￨��,Nvl(A.��ʶ��,B.�����) as ��ʶ��,'' as ����," & _
        "       A.����,A.�Ա�,A.����,A.�ѱ�,B.������,B.����,B.ҽ�Ƹ��ʽ,zl_PatiDayCharge(B.����ID) as ���ն�," & _
        "       Zl_Patiwarnscheme(A.����id) As ���ò���,-1 as ��������" & _
        " From ������ü�¼ A,������Ϣ B" & _
        " Where A.����ID=B.����ID(+) And A.��¼����=4 And A.��¼״̬=1" & _
             zlGetRegEventsCons("�Ӱ��־", "A") & _
        "       And A.NO=[3]"
    Else
        strSQL = strSQL & " and A.����=[2]"
    End If
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput, strNO)
    
    If Not mrsInfo.EOF Then
        GetPatientOut = True
    Else
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub CalcMoneys(Optional ByVal lngRow As Long = -1)
'���ܣ���������¼���ָ���л������еĽ��
'������lngRow=ָ����,Ϊ0��ʾ����������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long
    If mobjBill.Details.Count = 0 Then Exit Sub
    If lngRow = -1 Then
        For i = 0 To mlngRows - 1
            CalcMoney i
        Next
    Else
        CalcMoney lngRow
    End If
End Sub

Private Sub CalcMoney(ByVal lngRow As Long)
'���ܣ���������¼���ָ���еĽ��
'������lngRow=ָ����
'˵����1.ExpenseBill���ϵ�������Ӧ���ݵ��к�
'      2.���ֻ�ܶ�Ӧһ��������Ŀ:mobjBill.Details("R" & lngRow).InComes(1)
'      3.������ϸĿδ�����������Ŀ(��һ�μ���),��ʹ��Ĭ���ּ�
'      4.������ϸĿ�Ѿ������������Ŀ(����2��),���ֶ�����(Ҳ����δ��)�˵���,�򰴸õ��ۼ��㡣
    Dim i As Long, strInfo As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim curMoney As Currency '�û�����ı�۽��
    Dim strWherePriceGrade As String

    On Error GoTo errH
    If mstrPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "            Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From �շѼ�Ŀ" & vbNewLine & _
            "                               Where b.�շ�ϸĿId = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                     And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    strSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ���,B.ȱʡ�۸� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID=A.ID And C.ID=B.������ĿID " & _
        "   And Sysdate Between B.ִ������ And Nvl(B.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
        "   And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details("R" & lngRow).�շ�ϸĿID, mstrPriceGrade)
    If rsTmp.EOF Then
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set mobjBill.Details("R" & lngRow).InComes = New BillInComes
        Exit Sub
    End If
    
    If mobjBill.Details("R" & lngRow).Detail.��� Then
        If mobjBill.Details("R" & lngRow).InComes.Count = 0 Then '��һ�μ�����ȡȱʡֵ
            curMoney = Val(NVL(rsTmp!ȱʡ�۸�))
        Else                        '��ȡ����Ա��ǰ����ı�۽��
            curMoney = mobjBill.Details("R" & lngRow).InComes(1).��׼����
            '����û�����ı�۲������۷�Χ����ȡȱʡֵ
            If CheckScope(Val(NVL(rsTmp!ԭ��)), Val(NVL(rsTmp!�ּ�)), curMoney) <> "" Then
                curMoney = Val(NVL(rsTmp!ȱʡ�۸�))
            End If
        End If
    End If

    '�����ԭ�м�¼
    Set mobjBill.Details("R" & lngRow).InComes = New BillInComes

    '��д���з��ü�¼
    For i = 1 To rsTmp.RecordCount
        Set mobjBillIncome = New BillInCome
        With mobjBillIncome
            .������ĿID = rsTmp!������ĿID
            .������Ŀ = rsTmp!����
            .�վݷ�Ŀ = NVL(rsTmp!�վݷ�Ŀ)
            .ԭ�� = Val(NVL(rsTmp!ԭ��))
            .�ּ� = Val(NVL(rsTmp!�ּ�))
            If mobjBill.Details("R" & lngRow).Detail.��� Then
                .��׼���� = Format(curMoney, "0.0000")
                'ǿ�ư����θĳ� 1
                'mobjBill.Details("R" & lngRow).���� = 1
            Else
                .��׼���� = Format(IIf(IsNull(rsTmp!�ּ�), 0, rsTmp!�ּ�), "0.0000")
            End If

            'Ӧ�ս��=���� *  ����
            .Ӧ�ս�� = .��׼���� * mobjBill.Details("R" & lngRow).����
            '�������������ü���(����������Ŀ)
            If mobjBill.Details("R" & lngRow).���ӱ�־ = 1 And mobjBill.Details("R" & lngRow).�շ���� = "F" Then
                .Ӧ�ս�� = .Ӧ�ս�� * IIf(IsNull(rsTmp!�����շ���), 1, rsTmp!�����շ��� / 100)
            End If
            '�Ӱ�����ʼ���
            If mobjBill.�Ӱ��־ = 1 And mobjBill.Details("R" & lngRow).Detail.�Ӱ�Ӽ� Then
                .Ӧ�ս�� = .Ӧ�ս�� + .Ӧ�ս�� * IIf(IsNull(rsTmp!�Ӱ�Ӽ���), 0, rsTmp!�Ӱ�Ӽ��� / 100)
            End If

            .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))

            If mobjBill.Details("R" & lngRow).Detail.���ηѱ� Or .Ӧ�ս�� = 0 Then
                .ʵ�ս�� = .Ӧ�ս��
            Else
                .ʵ�ս�� = CCur(Format(ActualMoney(mobjBill.�ѱ�, .������ĿID, .Ӧ�ս��), gstrDec))
            End If
            
            '��ȡ��Ŀ������Ϣ,ҽ�����˲Ŵ���,����Ҫ����ҽ��
            If mrsInfo.State = 1 And mbytUseType <> Use���� Then
                If Not IsNull(mrsInfo!����) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details("R" & lngRow).�շ�ϸĿID, .ʵ�ս��, False, mrsInfo!����, _
                        mobjBill.Details("R" & lngRow).ժҪ & "||" & mobjBill.Details("R" & lngRow).����)
                    If strInfo <> "" Then
                        mobjBill.Details("R" & lngRow).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details("R" & lngRow).���մ���ID = Val(Split(strInfo, ";")(1))
                        .ͳ���� = Val(Split(strInfo, ";")(2))
                        mobjBill.Details("R" & lngRow).���ձ��� = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details("R" & lngRow).ժҪ = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details("R" & lngRow).Detail.���� = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
            End If
            
            mobjBill.Details("R" & lngRow).InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, , .ͳ����
        End With
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ShowDetails(Optional ByVal lngRow As Long = -1)
'���ܣ�ˢ����ʾָ���л������е�����
'������lngRow=ָ����,Ϊ-1��ʾ��ʾ������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, curTotal As Currency

    If lngRow = -1 Then
        For i = 0 To mlngRows - 1
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If

    curTotal = GetBillTotal(mobjBill)

    If IsNumeric(cmdOK.Tag) Then
        sta.Panels(3).Text = "Ԥ��:" & Format(cmdOK.Tag, "0.00")
        sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(CCur(cmdCancel.Tag) + curTotal, gstrDec)
        sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(mcur���ý�� - curTotal, "0.00")
    End If
End Sub

Private Sub ShowDetail(ByVal lngRow As Long)
'���ܣ�ˢ����ʾָ���е�����
'������lngRow=ָ���У�Ҳ�������е����-1
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, j As Long, curMoney As Currency
    Dim objBillDetail As BillDetail
    Dim strTmp As String
    

    If lngRow > mlngRows - 1 Then Exit Sub
    Set objBillDetail = mobjBill.Details("R" & lngRow)
    
    With objBillDetail
        
        If .�շ���� <> "" Then
            For i = 0 To cbo�շ����(lngRow).ListCount - 1
                If cbo�շ����(lngRow).ItemData(i) = Asc(.�շ����) Then
                    cbo�շ����(lngRow).ListIndex = i
                    Exit For
                End If
            Next
        End If
        'ˢ�µ�����
        '��Ŀ"
        txt�շ���Ŀ(lngRow).Text = .Detail.����
        '��λ"
        txt���㵥λ(lngRow).Text = .Detail.���㵥λ
        '�����ڵ�һ����ʾʱ��Ĭ������Ϊ1
        If .���� = 0 Then
            txt����(lngRow).Text = ""
        Else
            txt����(lngRow).Text = .����
        End If
        '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
        '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
        curMoney = 0
        If .InComes.Count > 0 Then
            For j = 1 To .InComes.Count
                curMoney = curMoney + .InComes(j).��׼����
            Next
        End If
        .��׼���� = curMoney
        txt��׼����(lngRow).Text = Format(curMoney, "0.0000")
        If Val(txt��׼����(lngRow).Text) = 0 Then txt��׼����(lngRow).Text = ""
        
        'Ӧ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
        curMoney = 0
        If .InComes.Count > 0 Then
            For j = 1 To .InComes.Count
                curMoney = curMoney + .InComes(j).Ӧ�ս��
            Next
        End If
        .Ӧ�ս�� = curMoney
        txtӦ�ս��(lngRow).Text = Format(curMoney, gstrDec)
        If Val(txtӦ�ս��(lngRow).Text) = 0 Then txtӦ�ս��(lngRow).Text = ""
        
        'ʵ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
        curMoney = 0
        If .InComes.Count > 0 Then
            For j = 1 To .InComes.Count
                curMoney = curMoney + .InComes(j).ʵ�ս��
            Next
        End If
        .ʵ�ս�� = curMoney
        txtʵ�ս��(lngRow).Text = Format(curMoney, gstrDec)
        If Val(txtʵ�ս��(lngRow).Text) = 0 Then txtʵ�ս��(lngRow).Text = ""
        
        'ִ�п���"
        If mbytInState = staִ�� Then
            mrsUnit.Filter = "ID=" & .ִ�в���ID
            If mrsUnit.RecordCount <> 0 Then
                'byZT200302
                Call cbo.SeekIndex(cboִ�п���(lngRow), mrsUnit!����, , True)
                If cboִ�п���(lngRow).ListIndex = -1 Then
                    cboִ�п���(lngRow).AddItem IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                    cboִ�п���(lngRow).ItemData(cboִ�п���(lngRow).NewIndex) = .ִ�в���ID
                    cboִ�п���(lngRow).ListIndex = cboִ�п���(lngRow).NewIndex
                End If
                'cboִ�п���(lngRow).Text = mrsUnit!���� & "-" & mrsUnit!����
            Else
                'byZT200302
                strTmp = GET��������(.ִ�в���ID, mrsUnit)
                If strTmp <> "" Then
                    Call cbo.SeekIndex(cboִ�п���(lngRow), strTmp, , True)
                    If cboִ�п���(lngRow).ListIndex = -1 Then
                        cboִ�п���(lngRow).AddItem strTmp
                        cboִ�п���(lngRow).ListIndex = cboִ�п���(lngRow).NewIndex
                    End If
                End If
                'cboִ�п���(lngRow).Text = Get��������(.ִ�в���ID)
            End If
        Else
            '�������ֻ(��)��ʾ����
            'byZT200302
            strTmp = GET��������(.ִ�в���ID, mrsUnit)
            If strTmp <> "" Then
                cboִ�п���(lngRow).AddItem GET��������(.ִ�в���ID, mrsUnit)
            End If
            'cboִ�п���(lngRow).Text = Get��������(.ִ�в���ID)
        End If
            '��־"
        If .�շ���� = "F" Then
            chk����(lngRow).Value = .���ӱ�־
        Else
            chk����(lngRow).Enabled = False
            chk����(lngRow).Value = 0
        End If
    End With
End Sub

Public Sub ShowMoney()
'���ܣ�ˢ����ʾ������Ŀ������
    Dim i As Integer, j As Integer
    Dim curTotal As Currency, curӦ��Total As Currency


    '�������ܷ�Ŀ
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            curTotal = curTotal + mobjBill.Details(i).InComes(j).ʵ�ս��
            curӦ��Total = curӦ��Total + mobjBill.Details(i).InComes(j).Ӧ�ս��
        Next
    Next

    txtӦ��.Text = Format(curӦ��Total, gstrDec)
    txtʵ��.Text = Format(curTotal, gstrDec)
End Sub

Private Sub GetInputDetail(ByVal lng��Ŀid As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngMediCareNO As Long
            
    Set mcolDetails = New Details
    
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!����)
    If lngMediCareNO > 0 Then
        strSQL = _
            " Select" & _
            " A.ID,A.���,B.���� as �������,A.����,A.����,A.���,A.���㵥λ," & _
            " A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.�������,M.Ҫ������,A.¼������" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,����֧����Ŀ M " & _
            " Where A.���=B.���� And A.ID=[1] And A.ID=M.�շ�ϸĿID(+) And M.����(+)=[2]"

    Else
        strSQL = _
            " Select" & _
            " A.ID,A.���,B.���� as �������,A.����,A.����,A.���,A.���㵥λ," & _
            " A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.�������,0 as Ҫ������,A.¼������" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B" & _
            " Where A.���=B.���� And A.ID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, lngMediCareNO)
    
    Set mobjDetail = New Detail
    With mobjDetail
        .ID = rsTmp!ID
        .��� = rsTmp!���
        .������� = rsTmp!�������
        .���� = rsTmp!����
        .���� = rsTmp!����
        .��� = NVL(rsTmp!���)
        .���㵥λ = NVL(rsTmp!���㵥λ)
        .��� = NVL(rsTmp!�Ƿ���, 0) = 1
        .�Ӱ�Ӽ� = NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1
        .���ηѱ� = NVL(rsTmp!���ηѱ�, 0) = 1
        .ִ�п��� = NVL(rsTmp!ִ�п���, 0)
        .������� = NVL(rsTmp!�������, 0)
        .���� = NVL(rsTmp!��������)
        .Ҫ������ = NVL(rsTmp!Ҫ������, 0) = 1
        .¼������ = Val("" & rsTmp!¼������)
        
        mcolDetails.Add .ID, .���, .�������, .����, .����, .����, .����, .���, .���㵥λ, .˵��, .���ηѱ�, .���, .�Ӱ�Ӽ�, .ִ�п���, .�������, .����, , , .Ҫ������, .¼������
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetDetails(ByVal lngHwnd As Long, ByVal str���� As String, Optional ByVal str��� As String)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim int������Դ As Integer, str��׼��Ŀ As String
    Dim lng��Ŀid As Long, int���� As Integer

    Set mcolDetails = New Details
    
    str���� = UCase(str����)
    
    '������Դ
    If mbytUseType = Use���� Then
        int������Դ = 1
    Else
        If mrsInfo.State = 1 Then
            '��ȡ����ʱ�Ѹ���Ȩ�������Ƿ����۲���
            If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
                int������Դ = 2
            ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
                int������Դ = 1
            End If
        Else
            'δȷ������,������,�ڱ���ʱ���
            If (InStr(mstrPrivsOpt, "�������ۼ���") > 0 And gbln��������) Or mbytUseType = 2 Then
                int������Դ = 0
            Else
                int������Դ = 2
            End If
        End If
    End If
    If mbytUseType <> -1 Then
        'ҽ��������׼��Ŀ
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!����) Then
                int���� = mrsInfo!����
                '���˺�:24862
                If zl_Check��׼��Ŀ(gclsInsure, int����, Val(NVL(mrsInfo!����ID)), False) Then str��׼��Ŀ = Get������׼��Ŀ(Val(NVL(mrsInfo!����ID)), "A.ID")
                
            End If
        End If
    End If
    
    sta.Panels("MedicareType").Text = ""
    If str��� <> "" Then str��� = "'" & str��� & "'"
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, str���, str����, lngHwnd, str��׼��Ŀ, mstrPriceGrade)
    If lng��Ŀid <> 0 Then
        Call GetInputDetail(lng��Ŀid)
        If int���� <> 0 Then sta.Panels("MedicareType").Text = Getҽ������(lng��Ŀid, int����)
    Else
        zlControl.TxtSelAll txt�շ���Ŀ(mintCurrentRow)
    End If
End Function

Private Sub NewBill()
'���ܣ���ʼ��һ���µĵ���(�������)
    Dim lngRow As Long
    
    mcurModiMoney = 0
    
    '������ݵ���ʱ��Ϣ
    mcurPreMoney = 0: sta.Panels(3).Text = ""
    cmdOK.Tag = "": cmdCancel.Tag = "": txtʵ��.Tag = ""
    
    txtʵ��.Text = gstrDec: txtӦ��.Text = gstrDec

    '���ʷ��౨��
    mstrWarn = ""
        
    cboNO.Text = ""
    chk�Ӱ�.Value = IIf(OverTime(zlDatabase.Currentdate), Checked, Unchecked)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    Call LoadPatientBaby(cboBaby, 0, 0)
    Call cbo��������_Click
    
    
    '�Ժ����ÿ������ҵ�ȱʡֵ
    Set mrsMedAudit = Nothing
    Set mrsInfo = New ADODB.Recordset
    Call GetEmptyBill(mlngRows)
    '���������ʾ��Ϣ
    Call ShowPatient
    Call ShowMoney
    
End Sub

Private Sub GetEmptyBill(ByVal Rows As Integer)
    Dim i As Integer
    
    
    Set mobjBill = New ExpenseBill
    
    For i = 0 To Rows - 1
        mobjBill.Details.AddEmpty i + 1
        mobjBill.Details("R" & i).�շ���� = cbo�շ����(i).Tag
        '���ø����շ���Ŀ
        If Val(txt�շ���Ŀ(i).Tag) > 0 Then
            Call GetInputDetail(Val(txt�շ���Ŀ(i).Tag))
            Call Set�շ�ϸĿ(i)
        Else
            ShowDetail i
        End If
    Next
    
    With mobjBill
        .��¼���� = 2
        .��¼״̬ = 1
        .�����־ = 2
        .������ = UserInfo.����
        .������ = zlStr.NeedName(cbo������.Text)
        .����Ա��� = UserInfo.���
        .����Ա���� = UserInfo.����
        .����ʱ�� = CDate(txtDate.Text)
        .�Ӱ��־ = chk�Ӱ�.Value
        
        If cbo��������.ListIndex = -1 Then
            Set�������� 0
        Else
            Set�������� cbo��������.ItemData(cbo��������.ListIndex)
        End If
    End With
End Sub

Private Sub ImportBill(strNO As String, ByVal Rows As Integer, _
    Optional ByVal strPriceGrade As String)
'���ܣ���ȡ���õ��ݵ����ݶ�����(Ŀǰ���Դ�����Ŀ,����������Ŀ)
'������
'      strNO=���ݺ�
'���أ���ŵ�����Ϣ�ĵ��ݶ���
'˵������Ϊ������ʱ��Ŀ�۸���Ϣ��������,���Է�������������¼���
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim int��� As Integer, blnDo As Boolean, i As Integer
    Dim cur���� As Currency, curʵ�� As Currency, curӦ�� As Currency
    Dim rsTmp As ADODB.Recordset, strSQL As String, strWherePriceGrade As String
    
    On Error GoTo errH
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (d.�۸�ȼ� = [3]" & vbNewLine & _
            "            Or (d.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From �շѼ�Ŀ" & vbNewLine & _
            "                               Where d.�շ�ϸĿId = �շ�ϸĿid And �۸�ȼ� = [3]" & vbNewLine & _
            "                                     And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And d.�۸�ȼ� Is Null"
    End If
    '�۸񸸺� is NULL:ֻȡÿ��������շ�ϸĿID��
    '�շѼ�Ŀ����:�¼���۸�,����ж���۸�,��һ���շ�ϸĿID�оͻ��ж��������ͬ�ļ�¼
    strSQL = "Select A.ID, A.��¼����, A.NO AS ���ݺ�, A.ʵ��Ʊ��, A.��¼״̬, A.���, A.��������, A.�۸񸸺�, A.���ʵ�id, A.����id," & _
                    IIf(mstrFreeTable = "סԺ���ü�¼", " A.�ಡ�˵�,A.��ҳid, A.���˲���id, A.����,", " 0 as �ಡ�˵�,0 as ��ҳid,0 as ���˲���id, A.���ʽ as ����, ") & vbNewLine & _
            "       A.ҽ�����, A.�����־, A.���ʷ���, A.����, A.�Ա�, A.����, A.��ʶ��," & vbNewLine & _
            "       A.���˿���id, A.�ѱ�, A.�շ����, A.�շ�ϸĿid, A.���㵥λ, A.����, A.��ҩ����, A.����, A.�Ӱ��־, A.���ӱ�־," & vbNewLine & _
            "       A.Ӥ����, A.������Ŀid, A.�վݷ�Ŀ, A.��׼����, A.Ӧ�ս��, A.ʵ�ս��, A.������, A.��������id, A.������," & vbNewLine & _
            "       A.����ʱ��, A.�Ǽ�ʱ��, A.ִ�в���id, A.ִ����, A.ִ��״̬, A.ִ��ʱ��, A.����, A.����Ա���, A.����Ա����," & vbNewLine & _
            "       A.����id, A.���ʽ��, A.���մ���id, A.������Ŀ��, A.���ձ���, A.ͳ����, A.�Ƿ��ϴ�, A.ժҪ, A.�Ƿ���," & vbNewLine & _
            "       A.��������, B.����, B.���, B.���� �շ�����, B.���㵥λ, B.�Ӱ�Ӽ�, B.���, B.���ηѱ�, B.˵��, B.ִ�п���," & vbNewLine & _
            "       B.�������� ԭ��������, B.�Ƿ���, C.���� As �������, D.������Ŀid As ������id, D.ԭ��ID,D.�շ�ϸĿID,D.ԭ��,D.�ּ�,D.ȱʡ�۸�,D.������ĿID,D.�Ӱ�Ӽ���,D.�����շ���," & vbNewLine & _
            "       E.���� As ������Ŀ, E.�վݷ�Ŀ As �ַ�Ŀ" & vbNewLine & _
            "From " & mstrFreeTable & " A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, �շѼ�Ŀ D, ������Ŀ E" & vbNewLine & _
            "Where E.ID = D.������Ŀid And D.�շ�ϸĿid = A.�շ�ϸĿid And A.�շ���� = C.���� And A.�շ�ϸĿid = B.ID And" & vbNewLine & _
            "      A.�۸񸸺� Is Null And A.��¼���� = 2 And A.NO = [1] And A.��¼״̬ = 1 And A.ִ��״̬ <> 1 And" & vbNewLine & _
            "      A.���ʵ�id = [2] And Sysdate Between D.ִ������ And Nvl(D.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
                    strWherePriceGrade & vbNewLine & _
            "Order By A.���"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mlng����ID, strPriceGrade)
    
    'û�м�¼���ǿյ���
    Call GetEmptyBill(Rows)
    If rsTmp.RecordCount <> 0 Then
        With rsTmp
            i = 1
            Do While Not .EOF
                
                If i = 1 Then
                    '����������
                    mobjBill.NO = !���ݺ�
                    mobjBill.��¼���� = IIf(IsNull(!��¼����), 0, !��¼����)
                    mobjBill.��¼״̬ = IIf(IsNull(!��¼״̬), 0, !��¼״̬)
                    mobjBill.����ID = IIf(IsNull(!����ID), 0, !����ID)
                    mobjBill.��ҳID = IIf(IsNull(!��ҳID), 0, !��ҳID)
                    mobjBill.����ID = IIf(IsNull(!���˲���ID), 0, !���˲���ID)
                    mobjBill.����ID = IIf(IsNull(!���˿���ID), 0, !���˿���ID)
                    mobjBill.���� = IIf(IsNull(!����), "", !����)
                    mobjBill.�Ա� = IIf(IsNull(!�Ա�), "", !�Ա�)
                    mobjBill.���� = IIf(IsNull(!����), "", !����)
                    mobjBill.��ʶ�� = IIf(IsNull(!��ʶ��), 0, !��ʶ��)
                    mobjBill.���� = "" & !����
                    mobjBill.�ѱ� = IIf(IsNull(!�ѱ�), "", !�ѱ�)
                    mobjBill.�����־ = IIf(IsNull(!�����־), 0, !�����־)
                    mobjBill.�Ӱ��־ = IIf(IsNull(!�Ӱ��־), 0, !�Ӱ��־)
                    mobjBill.Ӥ���� = IIf(IsNull(!Ӥ����), 0, !Ӥ����)
                    mobjBill.��������ID = IIf(IsNull(!��������ID), 0, !��������ID)
                    mobjBill.������ = IIf(IsNull(!������), "", !������)
                    mobjBill.������ = IIf(IsNull(!������), "", !������)
                    mobjBill.����Ա��� = IIf(IsNull(!����Ա���), "", !����Ա���)
                    mobjBill.����Ա���� = IIf(IsNull(!����Ա����), "", !����Ա����)
                    mobjBill.����ʱ�� = !����ʱ��
                    mobjBill.�Ǽ�ʱ�� = !�Ǽ�ʱ��
                    mobjBill.�ಡ�˵� = (IIf(IsNull(!�ಡ�˵�), 0, !�ಡ�˵�) = 1)
                End If
                
                '�����շ�ϸĿ,�����๲ͬ������һ���շ�ϸĿ
                Set objBillDetail = New BillDetail
                Set objBillDetail.Detail = New Detail
                
                If !��� > mlngRows Then
                    MsgBox "�����ʵ����շѸ�������С���Ѿ�����������ʾ�������ݡ�", vbExclamation, gstrSysName
                    mobjBill.NO = ""
                    Exit Sub
                End If
                
                objBillDetail.��� = !���
                objBillDetail.�շ���� = IIf(IsNull(!�շ����), "", !�շ����)
                objBillDetail.�շ�ϸĿID = IIf(IsNull(!�շ�ϸĿID), 0, !�շ�ϸĿID)
                objBillDetail.�շ����� = IIf(IsNull(!�շ�����), "", !�շ�����)
                objBillDetail.���㵥λ = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                objBillDetail.���� = IIf(IsNull(!����), 1, !����)
                objBillDetail.���ӱ�־ = IIf(IsNull(!���ӱ�־), 0, !���ӱ�־)
                objBillDetail.ժҪ = IIf(IsNull(!ժҪ), "", !ժҪ)
                objBillDetail.ִ�в���ID = IIf(IsNull(!ִ�в���ID), 0, !ִ�в���ID)
                
                If cbo�շ����(!��� - 1).Tag <> "" And cbo�շ����(!��� - 1).Tag <> objBillDetail.�շ���� Then
                    MsgBox "���Ƽ��ʵ���" & !��� & "�еĹ̶��շ�����뵥��ԭ�����ݲ�ͬ��", vbExclamation, gstrSysName
                    mobjBill.NO = ""
                    Exit Sub
                End If
                
                If Val(txt�շ���Ŀ(!��� - 1).Tag) > 0 And Val(txt�շ���Ŀ(!��� - 1).Tag) <> objBillDetail.�շ�ϸĿID Then
                    MsgBox "���Ƽ��ʵ���" & !��� & "�еĹ̶��շ���Ŀ�뵥��ԭ�����ݲ�ͬ��", vbExclamation, gstrSysName
                    mobjBill.NO = ""
                    Exit Sub
                End If
                
                
                objBillDetail.Detail.ID = !�շ�ϸĿID
                objBillDetail.Detail.���� = !����
                objBillDetail.Detail.��� = (IIf(IsNull(!�Ƿ���), 0, !�Ƿ���) = 1)
                objBillDetail.Detail.��� = IIf(IsNull(!���), "", !���)
                objBillDetail.Detail.���㵥λ = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                objBillDetail.Detail.�Ӱ�Ӽ� = (IIf(IsNull(!�Ӱ�Ӽ�), 0, !�Ӱ�Ӽ�) = 1)
                objBillDetail.Detail.��� = IIf(IsNull(!���), "", !���)
                objBillDetail.Detail.������� = IIf(IsNull(!�������), "", !�������)
                objBillDetail.Detail.���� = IIf(IsNull(!�շ�����), "", !�շ�����)
                objBillDetail.Detail.���ηѱ� = (IIf(IsNull(!���ηѱ�), 0, !���ηѱ�) = 1)
                objBillDetail.Detail.˵�� = IIf(IsNull(!˵��), "", !˵��)
                objBillDetail.Detail.ִ�п��� = IIf(IsNull(!ִ�п���), 0, !ִ�п���)
                objBillDetail.Detail.���� = IIf(IsNull(!��������), "" & !ԭ��������, !��������)
                objBillDetail.Detail.Ҫ������ = 0
                    
                Set objBillDetail.InComes = New BillInComes
                cur���� = 0: curʵ�� = 0: curӦ�� = 0
                
                Do
                    '�����������еļ۸��������¼���
                    If IIf(IsNull(!�Ƿ���), 0, !�Ƿ���) = 1 Then
                        If Abs(!��׼����) > Abs(IIf(IsNull(!�ּ�), 0, !�ּ�)) Then
                            objBillIncome.��׼���� = IIf(IsNull(!ȱʡ�۸�), 0, !ȱʡ�۸�)
                        Else
                            objBillIncome.��׼���� = !��׼����
                        End If
                    Else
                        objBillIncome.��׼���� = !�ּ�
                    End If
                    objBillIncome.������ĿID = IIf(IsNull(!������ID), 0, !������ID)
                    objBillIncome.������Ŀ = IIf(IsNull(!������Ŀ), "", !������Ŀ)
                    objBillIncome.�վݷ�Ŀ = IIf(IsNull(!�ַ�Ŀ), "", !�ַ�Ŀ)
                    objBillIncome.�ּ� = IIf(IsNull(!�ּ�), 0, !�ּ�)
                    objBillIncome.ԭ�� = IIf(IsNull(!ԭ��), 0, !ԭ��)
                    
                    'Ӧ�ս��=����*����*����
                    objBillIncome.Ӧ�ս�� = objBillIncome.��׼���� * IIf(IsNull(!����), 1, !����)
                    
                    '�������������ü���(����������Ŀ)
                    If IIf(IsNull(!���ӱ�־), 0, !���ӱ�־) = 1 And IIf(IsNull(!�շ����), "", !�շ����) = "F" Then
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * IIf(IsNull(!�����շ���), 1, !�����շ��� / 100)
                    End If
                    
                    '�Ӱ�����ʼ���
                    If IIf(IsNull(!�Ӱ��־), 0, !�Ӱ��־) = 1 And IIf(IsNull(!�Ӱ�Ӽ�), 0, !�Ӱ�Ӽ�) = 1 Then
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� + objBillIncome.Ӧ�ս�� * IIf(IsNull(!�Ӱ�Ӽ���), 0, !�Ӱ�Ӽ��� / 100)
                    End If
                    
                    '����ʵ�ս��
                    If IIf(IsNull(!���ηѱ�), 0, !���ηѱ�) = 1 Then
                        objBillIncome.ʵ�ս�� = objBillIncome.Ӧ�ս��
                    Else
                        objBillIncome.ʵ�ս�� = ActualMoney(mobjBill.�ѱ�, !������ID, objBillIncome.Ӧ�ս��)
                    End If
                    
                    objBillIncome.ʵ��Ʊ�� = IIf(IsNull(!ʵ��Ʊ��), "", !ʵ��Ʊ��)
                    
                    With objBillIncome
                        objBillDetail.InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ʵ��Ʊ��
                        cur���� = cur���� + .��׼����
                        curʵ�� = curʵ�� + .ʵ�ս��
                        curӦ�� = curӦ�� + .Ӧ�ս��
                    End With
                    
                    
                    '�ж���һ����¼�Ƿ����ڵ�ǰ��
                    blnDo = False
                    int��� = !���
                    .MoveNext
                    If Not .EOF Then blnDo = (int��� = !���)
                    i = i + 1
                Loop While blnDo And Not .EOF
                
                '�����һ���շ�ϸĿ������
                With objBillDetail
                    mobjBill.Details.Remove "R" & .��� - 1 '����ǰ�Ȱ���ǰ�Ŀռ�¼ɾ��
                    mobjBill.Details.Add .Detail, .�շ�ϸĿID, .�շ�����, .���, .�շ����, .���㵥λ, .����, cur����, curʵ��, curӦ��, .���ӱ�־, .ִ�в���ID, .InComes
                End With
            Loop
        End With
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
End Sub

Private Sub ClearDetail(ByVal lngRow As Long)
'���ܣ�ˢ����ʾָ���е�����
'������lngRow=ָ���У�Ҳ�������е����-1
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, j As Long, curMoney As Currency
    Dim objBillDetail As BillDetail
    

    If lngRow > mlngRows - 1 Then Exit Sub
    mobjBill.Details.Remove "R" & lngRow
    mobjBill.Details.AddEmpty lngRow + 1
    
    If cbo�շ����(lngRow).ListIndex <> -1 Then
        mobjBill.Details("R" & lngRow).�շ���� = Chr(cbo�շ����(lngRow).ItemData(cbo�շ����(lngRow).ListIndex))
    End If
    Call ShowDetail(lngRow)
End Sub

Private Function SaveBill() As Boolean
'����:���浱ǰ����ļ��ʵ���
'����:�����Ƿ�ɹ�
    Dim i As Integer, j As Integer, arrSQL As Variant
    Dim lngCurID As Long, strNO As String, strTmp As String, strSQL As String
    Dim intInsure As Integer, str��Ϣ As String
    Dim lngNO As Long, lngParent As Long, lngParentNO As Long, lngChildNO As Long

    mobjBill.NO = zlDatabase.GetNextNo(14)
    mobjBill.����ʱ�� = CDate(txtDate.Text)
    mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate

    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    
    lngChildNO = mlngRows + 1 '�۸�ĸ���ֻ�ܴӴ˿�ʼ
    For lngParentNO = 1 To mlngRows
        Set mobjBillDetail = mobjBill.Details("R" & lngParentNO - 1)
        If mobjBillDetail.���� <> 0 Then
            lngParent = 0
            For Each mobjBillIncome In mobjBillDetail.InComes
                lngParent = lngParent + 1
                If lngParent = 1 Then
                    '��һ��������Ŀ��Ϊ����¼
                    lngNO = lngParentNO
                Else
                    lngNO = lngChildNO
                    '�����Ҫ�ֹ�����
                    lngChildNO = lngChildNO + 1
                End If
                
                If mbytUseType = Use���� Then
                    '��������
                    With mobjBill
                        strSQL = "zl_������ʼ�¼_INSERT('" & .NO & "'," & lngNO & "," & .����ID & "," & .��ʶ�� & "," & _
                            "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & .�ѱ� & "'," & .�Ӱ��־ & "," & .Ӥ���� & "," & _
                            IIf(.����ID = 0, .��������ID, .����ID) & "," & .��������ID & ",'" & .������ & "',"
                    End With
                
                    '�շ�ϸĿ����
                    With mobjBillDetail
                        strSQL = strSQL & "Null," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "'," & _
                             "1," & .���� & "," & .���ӱ�־ & "," & .ִ�в���ID & ","
                    End With
                
                    '������Ŀ����
                    With mobjBillIncome
                        strSQL = strSQL & IIf(lngParent = 1, "Null", lngParentNO) & "," & .������ĿID & "," & _
                            "'" & .�վݷ�Ŀ & "'," & .��׼���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & ","
                    End With
                                                
                    '��������
                    strSQL = strSQL & _
                        "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "'" & mstrInNO & "',0,'" & UserInfo.��� & "','" & UserInfo.���� & "',NULL," & mlng����ID & ",'" & _
                        mobjBillDetail.ժҪ & "',Null,Null,Null,Null,Null,Null)"
                
                Else 'סԺ����
                    '��������
                    With mobjBill
                        strSQL = "zl_סԺ���ʼ�¼_INSERT('" & .NO & "'," & lngNO & "," & .����ID & "," & IIf(.��ҳID = 0, "NULL", .��ҳID) & "," & _
                            IIf(.��ʶ�� = 0, "NULL", .��ʶ��) & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & .���� & "','" & .�ѱ� & "'," & _
                            IIf(.����ID = 0, .��������ID, .����ID) & "," & IIf(.����ID = 0, .��������ID, .����ID) & "," & .�Ӱ��־ & "," & .Ӥ���� & "," & .��������ID & ",'" & .������ & "',"
                    End With
    
                    '�շ�ϸĿ����
                    With mobjBillDetail
                        strSQL = strSQL & "Null," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                        strSQL = strSQL & IIf(.������Ŀ��, 1, 0) & "," & IIf(.���մ���ID = 0, "NULL", .���մ���ID) & ",'" & .���ձ��� & "',"
                        strSQL = strSQL & "1," & .���� & "," & .���ӱ�־ & "," & .ִ�в���ID & ","
                    End With
    
                    '������Ŀ����
                    With mobjBillIncome
                        strSQL = strSQL & IIf(lngParent = 1, "Null", lngParentNO) & "," & .������ĿID & "," & _
                            "'" & .�վݷ�Ŀ & "'," & .��׼���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & ","
                        strSQL = strSQL & .ͳ���� & ","
                    End With
    
                    '��������
                    strSQL = strSQL & _
                        "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "'" & mstrInNO & "',0,'" & UserInfo.��� & "','" & UserInfo.���� & "',0,NULL," & mlng����ID & ",'" & _
                        mobjBillDetail.ժҪ & "',0,Null,Null,Null,Null,Null,Null,0,'" & mobjBillDetail.Detail.���� & "')"
                End If
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & strSQL
            Next
        End If
    Next

    '�޸�ǰ�˳�ԭ����
    If mstrInNO <> "" Then
        If mbytUseType <> Use���� Then
            'ҽ�����������ϴ�(�޸�֮ǰ�Ѿ��ж�)
            intInsure = BillExistInsure(mstrInNO)
            'ȥ����ҽ������ƥ����
        End If
    
        If mbytUseType = Use���� Then
            strSQL = "zl_������ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            strSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        End If
        If strSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & strSQL
        End If
    End If

    If UBound(arrSQL) >= 0 Then
        '��SQL���а��շ�ϸĿID����
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Mid(arrSQL(j), 1, InStr(arrSQL(j), ";") - 1)) < CLng(Mid(arrSQL(i), 1, InStr(arrSQL(i), ";") - 1)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next

        'ִ��SQL���
        On Error GoTo errH
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
        Next
        
        'ҽ���ӿ�
        If mbytUseType <> Use���� Then
            '1.ҽ�����������ϴ�
            If mstrInNO <> "" And intInsure <> 0 Then
                If MCPAR.���������ϴ� And Not MCPAR.������ɺ��ϴ� Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
            
            '2.����ʵʱ�ϴ�
            If Not IsNull(mrsInfo!����) Then
                If MCPAR.�����ϴ� And Not MCPAR.������ɺ��ϴ� Then
                    str��Ϣ = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , mrsInfo!����) Then
                        gcnOracle.RollbackTrans
                        If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
        gcnOracle.CommitTrans

        'ҽ���ӿ�
        If mbytUseType <> Use���� Then
            '1.ҽ�����������ϴ�
            If mstrInNO <> "" And intInsure <> 0 Then
                If MCPAR.���������ϴ� And MCPAR.������ɺ��ϴ� Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        MsgBox "����""" & mstrInNO & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                    End If
                End If
            End If
            
            '2.����ʵʱ�ϴ�
            If Not IsNull(mrsInfo!����) Then
                'ҽ�����������ϸ
                If MCPAR.�����ϴ� And MCPAR.������ɺ��ϴ� Then
                    str��Ϣ = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , mrsInfo!����) Then
                        If str��Ϣ <> "" Then
                            MsgBox str��Ϣ, vbInformation, gstrSysName
                        Else
                            MsgBox "����""" & mobjBill.NO & """��������ҽ������ʧ�ܣ��õ����ѱ��档", vbInformation, gstrSysName
                        End If
                    End If
                End If
            End If
        End If

        '���뵥����ʷ��¼(�������͵���)
        For i = 0 To cboNO.ListCount - 1
            strNO = strNO & "," & cboNO.List(i)
        Next
        strNO = mobjBill.NO & strNO
        cboNO.Clear
        For i = 0 To UBound(Split(strNO, ","))
            cboNO.AddItem Split(strNO, ",")(i)
            If i = 9 Then Exit For 'ֻ��ʾ10��
        Next
        
        'ҽ���ӿ�
        If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
    End If
    SaveBill = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function

Private Function ReadBill(strNO As String) As Integer
'���ܣ����ݵ��ݺŶ�ȡһ�ŵ��ݲ�����������
'������strFullNo=���ݺ�
    Dim rsTmp As ADODB.Recordset, rsPatiMoney As ADODB.Recordset
    Dim i As Integer, blnDeal As Boolean, strSQL As String
    Dim curTotal As Currency, curӦ��Total As Currency
    Dim strFullNO As String, strTmp As String
    Dim blnסԺ As Boolean
    
    On Error GoTo errH
    blnסԺ = (mstrFreeTable <> "������ü�¼")
    
    strFullNO = GetFullNO(strNO, 14)
    If mbytInState = sta���� Then
        '�жϸ��ż��ʵ��ܷ����ʣ������һ����¼ִ�оͲ�����
        strSQL = "select nvl(count(ִ��״̬),0) as  ����,nvl(sum(decode(ִ��״̬,1,1,0)),0) as ִ�� " & _
                " From " & IIf(mblnNOMoved, zlGetFullFieldsTable(mstrFreeTable), mstrFreeTable & " A") & _
                " Where  " & IIf(blnסԺ, " Nvl(�ಡ�˵�,0)=0 And ", "") & " ��¼״̬=1 " & _
                "       And ��¼����=2 and ���ʵ�ID=[2] And NO=[1]"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFullNO, mlng����ID)
        
        If rsTmp("����") = 0 Then
            MsgBox "û�з��ָõ���,�õ��ݿ����Ѿ����ϣ�", vbInformation, gstrSysName
            Exit Function
        End If
        If rsTmp("ִ��") > 0 Then
            MsgBox "���ܶԸõ�������,���Ѿ�ִ���ˣ�", vbInformation, gstrSysName
            Exit Function
        End If
        rsTmp.Close
    End If
    
    '��ȡ��������   '" & IIf(mblnNOMoved, iif(blnNOOnline
    
    strSQL = _
    " Select A.����ID," & IIf(blnסԺ, " Nvl(A.��ҳID,0)  as  ��ҳID,A.���˲���ID,B.���� as ���˲���,A.����,", "0  as  ��ҳID, 0 as ���˲���ID,'' as ���˲���, '' as ����,") & _
    "       A.��ʶ��,A.����,A.�Ա�,A.����,A.�ѱ�," & _
    "       A.���˿���ID,C.���� as ���˿���,A.��������ID," & _
    "       Nvl(A.�Ӱ��־,0) as �Ӱ��־,Nvl(A.Ӥ����,0) as Ӥ����," & _
    "       A.������,A.������,A.����Ա����,A.����ʱ��,A.����ID" & _
    " From " & IIf(mblnNOMoved, zlGetFullFieldsTable(mstrFreeTable), mstrFreeTable & " A") & IIf(blnסԺ, ",���ű� B", "") & ",���ű� C,��Ա�� D" & _
    " Where Rownum=1  " & _
            IIf(blnסԺ, " And Nvl(A.�ಡ�˵�,0)=0 And  A.���˲���ID=B.ID(+) ", "") & _
    "       And A.���ʵ�ID=[2] And A.��¼״̬" & IIf(mblnViewCancel, "=2", " IN(1,3)") & _
    "       And A.��¼����=2 And A.NO=[1]  and A.���˿���ID=C.ID(+) " & _
    "       And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & vbNewLine & _
    "       And Nvl(A.����Ա����,A.������)=D.����"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFullNO, mlng����ID)
    
    If rsTmp.EOF Then
        MsgBox "û�з��ָõ���,�õ��ݿ����Ѿ����ϣ�", vbInformation, gstrSysName
        Exit Function
    ElseIf mbytUseType <> Use���� Then
        '������ʲ���Ҫ�Կ���Ȩ�޽����ж�
        If mbytUseType = 0 Or mbytUseType = 1 Then
            If InStr(mstrPrivs, "���в���") = 0 And mlngUnitID > 0 Then
                If InStr(1, "," & mstrUnitIDs & ",", "," & IIf(IsNull(rsTmp!���˲���ID), 0, rsTmp!���˲���ID) & ",") = 0 Then
                    MsgBox "��û��Ȩ�޶�ȡ���������ĵ��ݣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        ElseIf mbytUseType = 2 Then
            If InStr(mstrPrivs, "���п���") = 0 And mlngDeptID > 0 Then
                If IIf(IsNull(rsTmp!��������ID), 0, rsTmp!��������ID) <> mlngDeptID Then
                    MsgBox "��û��Ȩ�޶�ȡ�������ҿ����ĵ��ݣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If

    '����ͷ
    cboNO.Text = strFullNO                                   '���ݺ�
    txtPatient.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����) '����
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����) '����
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����) '����
    txt��ʶ��.Text = IIf(IsNull(rsTmp!��ʶ��), "", rsTmp!��ʶ��) '��ʶ��
    txt����ID.Text = IIf(IsNull(rsTmp!����ID), "", rsTmp!����ID) '����ID
    txt��ҳID.Text = IIf(IsNull(rsTmp!��ҳID), "", rsTmp!��ҳID) '��ҳID
    txt���˲���.Text = IIf(IsNull(rsTmp!���˲���), "", rsTmp!���˲���) '���˲���
    txt���˿���.Text = IIf(IsNull(rsTmp!���˿���), "", rsTmp!���˿���) '���˿���

    '�Ա�
    Call cbo.SeekIndex(cbo�Ա�, IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�), , True)
    If cbo�Ա�.ListIndex = -1 And Not IsNull(rsTmp!�Ա�) Then
        cbo�Ա�.AddItem rsTmp!�Ա�, 0
        cbo�Ա�.ListIndex = 0
    End If
    
    '�ѱ�
    Call cbo.SeekIndex(cbo�ѱ�, IIf(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�), , True)
    If cbo�ѱ�.ListIndex = -1 And Not IsNull(rsTmp!�ѱ�) Then
        cbo�ѱ�.AddItem rsTmp!�ѱ�, 0
        cbo�ѱ�.ListIndex = 0
    End If
    
    txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    chk�Ӱ�.Value = IIf(IsNull(rsTmp!�Ӱ��־), 0, rsTmp!�Ӱ��־)
    Call LoadPatientBaby(cboBaby, rsTmp!����ID, rsTmp!��ҳID)
    Call zlControl.CboLocate(cboBaby, rsTmp!Ӥ����, True)
    
    mblnDo = False
    
        '����ȷ��ҽ��
        cbo��������.ListIndex = cbo.FindIndex(cbo��������, NVL(rsTmp!��������ID, 0))
        If cbo��������.ListIndex = -1 And Not IsNull(rsTmp!��������ID) Then
            cbo��������.AddItem GET��������(rsTmp!��������ID, mrs��������), 0
            cbo��������.ItemData(cbo��������.NewIndex) = rsTmp!��������ID
            cbo��������.ListIndex = cbo��������.NewIndex
        End If
        
        cbo������.Clear
        If cbo��������.ListIndex <> -1 Then
            Call FillDoctor(cbo��������.ItemData(cbo��������.ListIndex))
        End If
        Call cbo.SeekIndex(cbo������, NVL(rsTmp!������), , True)
        If cbo������.ListIndex = -1 And Not IsNull(rsTmp!������) Then
            cbo������.AddItem rsTmp!������, 0
            cbo������.ListIndex = cbo������.NewIndex
        End If
    
    mblnDo = True
    
    
    '���˷�����Ϣ
    If Not IsNull(rsTmp!����ID) Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!����ID)
        If Not rsPatiMoney Is Nothing Then
            sta.Panels(3).Text = "Ԥ��:" & Format(rsPatiMoney!Ԥ�����, "0.00") & _
            "/����:" & Format(rsPatiMoney!�������, gstrDec) & _
            "/ʣ��:" & Format(rsPatiMoney!Ԥ����� - rsPatiMoney!�������, "0.00")
        End If
    End If
    
    '��ȡ�����շ�ϸĿ
    strSQL = _
    " Select Decode(A.�۸񸸺�,NULL,A.���,A.�۸񸸺�) as ���," & _
    "       C.����,C.���� as ���,B.����,B.���,Nvl(A.��������,B.��������) ��������,A.���㵥λ," & _
    "       Avg(A.����) as ����,Sum(A.��׼����) as ����,Sum(A.Ӧ�ս��) as Ӧ�ս��, " & _
    "       Sum(A.ʵ�ս��) as ʵ�ս��,A.���ӱ�־,A.ִ�в���ID,D.���� as ִ�в��� " & _
    " From " & IIf(mblnNOMoved, zlGetFullFieldsTable(mstrFreeTable), mstrFreeTable & " A") & ",�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D " & _
    " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID " & _
    "       And A.��¼״̬" & IIf(mblnViewCancel, "=2", " IN(1,3)") & " And A.NO=[1]" & _
    "       " & IIf(mstrFreeTable = "סԺ���ü�¼", "And Nvl(A.�ಡ�˵�,0)=0 ", "") & " And A.��¼����=2 And A.���ʵ�ID=[2]" & _
    " Group by Decode(A.�۸񸸺�,NULL,A.���,A.�۸񸸺�),C.����,C.����," & _
    "       B.����,B.���,Nvl(A.��������,B.��������),A.���㵥λ,A.���ӱ�־,A.ִ�в���ID,D.����" & _
    " Order by Decode(A.�۸񸸺�,NULL,A.���,A.�۸񸸺�)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFullNO, mlng����ID)
    If rsTmp.EOF Then Exit Function
    
    '������
    curTotal = 0
    curӦ��Total = 0
    
    For i = 0 To mlngRows - 1
        blnDeal = False
        If Not rsTmp.EOF Then
            If rsTmp("���") = i + 1 Then
                
                cbo�շ����(i).AddItem rsTmp!���
                cbo�շ����(i).ListIndex = cbo�շ����(i).NewIndex
                txt�շ���Ŀ(i).Text = rsTmp!����
                txt���㵥λ(i).Text = IIf(IsNull(rsTmp!���㵥λ), "", rsTmp!���㵥λ)
                txt����(i).Text = rsTmp!����
                txt��׼����(i).Text = Format(rsTmp!����, "0.0000")
                If Val(txt��׼����(i).Text) = 0 Then txt��׼����(i).Text = ""
                txtӦ�ս��(i).Text = Format(rsTmp!Ӧ�ս��, gstrDec)
                If Val(txtӦ�ս��(i).Text) = 0 Then txtӦ�ս��(i).Text = ""
                txtʵ�ս��(i).Text = Format(rsTmp!ʵ�ս��, gstrDec)
                If Val(txtʵ�ս��(i).Text) = 0 Then txtʵ�ս��(i).Text = ""
                
                'byZT200302
                strTmp = rsTmp("ִ�в���")
                If strTmp <> "" Then
                    Call cbo.SeekIndex(cboִ�п���(i), strTmp, , True)
                    If cboִ�п���(i).ListIndex = -1 Then
                        cboִ�п���(i).AddItem rsTmp("ִ�в���")
                        cboִ�п���(i).ListIndex = cboִ�п���(i).NewIndex
                    End If
                End If
                'cboִ�п���(i).Text = rsTmp("ִ�в���")
                
                chk����(i).Value = rsTmp!���ӱ�־
                
                curTotal = curTotal + rsTmp("ʵ�ս��")
                curӦ��Total = curӦ��Total + rsTmp("Ӧ�ս��")
                rsTmp.MoveNext
                blnDeal = True
            End If
        End If
        'û�ҵ����ʵ�ֵ
        If blnDeal = False And Val(txt�շ���Ŀ(i).Tag) <= 0 Then
            cbo�շ����(i).ListIndex = cbo.FindIndex(cbo�շ����(i), Val(cbo�շ����(i).Tag))
            txt�շ���Ŀ(i).Text = ""
            txt���㵥λ(i).Text = ""
            txt����(i).Text = ""
            txt��׼����(i).Text = ""
            txtӦ�ս��(i).Text = ""
            txtʵ�ս��(i).Text = ""
            cboִ�п���(i).ListIndex = -1
            chk����(i).Value = 0
        End If
    Next
    If rsTmp.EOF = False Then
        MsgBox "�����ʵ����շ�������С���Ѿ�����������ʾ�������ݡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    

    txtʵ��.Text = Format(curTotal, gstrDec)
    txtӦ��.Text = Format(curӦ��Total, gstrDec)

    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetDetailNum(ByVal lngRow As Long, Optional dbl�������� As Double) As Single
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ָ��ϸĿ���ܼ�������(����������)
    '���:lngRow=��ǰ������
    '����:dbl��������-���ص�ǰ�Ѿ����ʵ�����
    '����:
    '����:���˺�
    '����:2010-08-19 18:02:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim lngNum As Long, i As Long
    Dim strSQL As String

    If mrsInfo.State = 1 Then
        '��ǰ�����е�����
        For i = 0 To mlngRows - 1
            If i <> lngRow And mobjBill.Details("R" & i).�շ�ϸĿID = mobjBill.Details("R" & lngRow).�շ�ϸĿID Then
                lngNum = lngNum + mobjBill.Details("R" & i).����
            End If
        Next
        dbl�������� = 0
        '���ݿ��е�����
        strSQL = _
        " Select Sum(Nvl(����,1)*����) as NUM," & _
        "           Sum(decode(����ID,NULL,0,1)* Nvl(����,1)*����) as ��������  " & _
        " From " & mstrFreeTable & _
        " Where �۸񸸺� is Null And ���ʷ���=1 And ��¼״̬<>0" & _
        "       And ����ID=[1] " & IIf(mstrFreeTable = "������ü�¼", "", " And Nvl(��ҳID,0)=[2]") & " And �շ�ϸĿID+0=[3]"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!����ID), Val("" & mrsInfo!��ҳID), mobjBill.Details("R" & lngRow).�շ�ϸĿID)
        If Not rsTmp.EOF Then
            lngNum = lngNum + NVL(rsTmp!Num, 0)
            dbl�������� = Val(NVL(rsTmp!��������))
        End If
        GetDetailNum = lngNum
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FillDoctor(Optional lng����ID As Long, Optional strMask As String)
'���ܣ���ȡ����дҽ���б�
'������strMask=����ƥ��ļ�������
    Dim i As Integer, lngOldID As Long
    Dim str����ҽ�� As String, strPre As String
    
    If cbo������.ListIndex <> -1 Then strPre = cbo������.List(cbo������.ListIndex)
    cbo������.Clear
    
    Call GetDoctor(lng����ID, gbln��ʿ And (gstr�շ���� = "" _
        Or gstr�շ���� Like "*'E'*" Or gstr�շ���� Like "*'M'*" Or gstr�շ���� Like "*'4'*"), mrs������, IIf(mbytUseType = Use����, 1, 2))

    If Not mrs������ Is Nothing Then
        If mrsInfo.State = 1 And mbytUseType <> Use���� Then If Not IsNull(mrsInfo!סԺҽʦ) Then str����ҽ�� = mrsInfo!סԺҽʦ
        
        i = IIf(mrs������.RecordCount = 0, 0, mrs������.RecordCount - 1)
        ReDim marrDr(i)
        
        For i = 1 To mrs������.RecordCount
            If lngOldID <> mrs������!ID Then
                If strMask = "" Then
                    cbo������.AddItem IIf(IsNull(mrs������!����), "", mrs������!���� & "-") & mrs������!����
                    cbo������.ItemData(cbo������.NewIndex) = Val(mrs������!���)
                Else
                    If InStr(IIf(IsNull(mrs������!����), "", mrs������!���� & "-") & mrs������!����, UCase(strMask)) Then
                        cbo������.AddItem IIf(IsNull(mrs������!����), "", mrs������!���� & "-") & mrs������!����
                        cbo������.ItemData(cbo������.NewIndex) = Val(mrs������!���)
                    ElseIf IsNumeric(strMask) Then
                        If CDbl(strMask) = CDbl(mrs������!���) Then
                            cbo������.AddItem IIf(IsNull(mrs������!����), "", mrs������!���� & "-") & mrs������!����
                            cbo������.ItemData(cbo������.NewIndex) = Val(mrs������!���)
                        End If
                    End If
                End If
                
                marrDr(cbo������.NewIndex) = mrs������!ID & "|" & mrs������!����ID & "|" & IIf(IsNull(mrs������!���), "", mrs������!���) & "|" & mrs������!���� & "|" & IIf(IsNull(mrs������!����), "", mrs������!����) & "|" & mrs������!ְ�� & "|" & mrs������!��Ա����
                
                If cbo������.List(cbo������.NewIndex) = strPre And cbo������.ListIndex = -1 Then
                    cbo������.ListIndex = cbo������.NewIndex
                End If
                If str����ҽ�� = mrs������!���� And cbo������.ListIndex = -1 Then
                    cbo������.ListIndex = cbo������.NewIndex
                End If
                lngOldID = mrs������!ID
            End If
            mrs������.MoveNext
        Next
        
        If cbo������.ListCount > 0 Then ReDim Preserve marrDr(cbo������.ListCount - 1)
        If cbo������.ListCount = 1 And cbo������.ListIndex = -1 Then cbo������.ListIndex = 0
    End If
End Sub
Private Function FillDept(Optional lng��ԱID As Long) As Long
'���ܣ���ȡ����ʾ����
'������lng��ԱID=ֻ��ȡָ����Ա���ڿ���(������ȱʡ��)
'���أ����Ҹ���
    
    Dim strSQL As String, i As Long
    Dim lngDeptID As Long, lngOldDepID As Long
    Dim strDepts As String  'ָ����Ա�����Ķ������
    
    On Local Error GoTo errH
            
    '��¼ԭ����,�������¶�λ
    If cbo��������.ListIndex <> -1 Then
        lngDeptID = cbo��������.ItemData(cbo��������.ListIndex)
    End If
    cbo��������.Clear
    
    If mrs�������� Is Nothing Then  'һ��Ҫ��Form_Unload������nothing
    
         '��ѡ��������(�����ҽ������,����������סԺ��)
        If (InStr(mstrPrivsOpt, "�������ۼ���") > 0 And gbln��������) Or mbytUseType = 2 Then
            strSQL = "1,2,3"
        Else
            strSQL = "2,3"
        End If
        If mbytUseType = UseסԺ Or mbytUseType = Use���ҷ�ɢ Then
            strSQL = _
                "Select Distinct A.ID,A.����,A.����,A.����,B.�������� " & _
                " from ���ű� A,��������˵�� B " & _
                " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
                " and B.����ID=A.ID and (B.������� IN(" & strSQL & ") AND B.�������� IN('�ٴ�','����')  or B.��������='����')" & _
                " Order by A.����"
        ElseIf mbytUseType = Useҽ������ Then
            'ҽ�����Ҽ���
            If InStr(mstrPrivs, "���п���") > 0 Then
                strSQL = _
                    "Select Distinct A.ID,A.����,A.����,A.����,B.�������� " & _
                    " from ���ű� A,��������˵�� B " & _
                    " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
                    " and B.����ID=A.ID and (B.������� IN(" & strSQL & ") AND B.�������� IN('���','����','����','����','Ӫ��') Or b.��������='����')" & _
                    " Order by A.����"
            Else
                strSQL = _
                    "Select Distinct A.ID,A.����,A.����,A.����,B.�������� " & _
                    " from ���ű� A,��������˵�� B " & _
                    " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
                    " and B.����ID=A.ID and (B.������� IN(" & strSQL & ") AND B.�������� IN('���','����','����','����','Ӫ��') Or b.��������='����')" & _
                    " And A.ID=" & mlngDeptID & _
                    " Order by A.����"
            End If
        ElseIf mbytUseType = Use���� Then
            strSQL = _
                " Select Distinct A.ID,A.����,A.����,A.����,B.�������� " & _
                " from ���ű� A,��������˵�� B " & _
                " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
                " and B.����ID=A.ID and (B.������� IN(1,3) AND B.�������� IN('�ٴ�','����') Or b.��������='����')" & _
                " Order by A.����"
        End If
        Set mrs�������� = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrs��������, strSQL, Me.Caption)
    End If
   
    If lng��ԱID <> 0 Then
        If mrs������ Is Nothing Then Call FillDoctor
        mrs������.Filter = "ID=" & lng��ԱID
        For i = 1 To mrs������.RecordCount
            strDepts = strDepts & " OR ID=" & mrs������!����ID      'filter��֧��in
            mrs������.MoveNext
        Next
        If strDepts <> "" Then
            mrs��������.Filter = Mid(strDepts, 4)
        Else
            mrs��������.Filter = "ID=0" '��Աû�����ò���,����ʾ��������
        End If
    Else
        mrs��������.Filter = ""
    End If
    
    If Not mrs��������.EOF Then
        For i = 1 To mrs��������.RecordCount
            If lngOldDepID <> mrs��������!ID Then   'һ�����ſ���ͬʱ�����������ٴ�,��������ͬ��
                cbo��������.AddItem IIf(zlIsShowDeptCode, mrs��������!���� & "-", "") & mrs��������!����
                cbo��������.ItemData(cbo��������.ListCount - 1) = mrs��������!ID
                
                If mrs��������!ID = mlngDeptID Then cbo��������.ListIndex = cbo��������.NewIndex
                
                If mrs��������!ID = lngDeptID And cbo��������.ListIndex = -1 Then
                    cbo��������.ListIndex = cbo��������.NewIndex
                End If
                
                lngOldDepID = mrs��������!ID
            End If
            mrs��������.MoveNext
        Next
        If cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
    End If
    
    FillDept = mrs��������.RecordCount
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CalcGridToTal(Optional blnӦ�� As Boolean) As Currency
    Dim objTmpDetail As BillDetail
    Dim objTmpIncome As BillInCome
    Dim i As Integer, intCol As Integer
    If mobjBill.Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Details
            For Each objTmpIncome In objTmpDetail.InComes
                If blnӦ�� Then
                    CalcGridToTal = CalcGridToTal + objTmpIncome.Ӧ�ս��
                Else
                    CalcGridToTal = CalcGridToTal + objTmpIncome.ʵ�ս��
                End If
            Next
        Next
    Else
        For i = 0 To mlngRows - 1
            CalcGridToTal = CalcGridToTal + Val(IIf(blnӦ��, txtӦ�ս��(i).Text, txtʵ�ս��(i).Text))
        Next
    End If
End Function

Private Function CheckBillisZero() As Boolean
'���ܣ��жϵ������е����Ƿ�������Ϊ0
    Dim i As Integer, j As Integer
    
    For i = 0 To mlngRows - 1
        If mobjBill.Details("R" & i).���� = 0 Then j = j + 1
    Next
    
    CheckBillisZero = (mlngRows = j)
End Function

Private Function SaveModi() As Boolean
'���ܣ����浱ǰ�޸ĵķ��õ���
    Dim strSQL As String

    strSQL = "zl_���˷��ü�¼_Update('" & cboNO.Text & "',2,'" & zlStr.NeedName(cbo������.Text) & "'," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),NULL," & IIf(mbytUseType = Use����, 1, 2) & " )"
    On Error GoTo errH
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Private Function InitFace() As Boolean
'���ܣ���ɶԵ��ݽ���ĳ�ʼ��
    Dim arrHead() As String, i As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim sngTemp As Single, arrBaby As Variant
    Dim ctlTemp As Control, varTemp As Variant
    Dim lngIndex As Long
    Dim blnContainer As Boolean       '�ÿؼ��Ƿ����һ��������������
    
    On Error GoTo errHandle
    
    InitFace = False
    If mbytUseType = Use���� Then
        mstrFreeTable = "������ü�¼"
    Else
        mstrFreeTable = "סԺ���ü�¼"
    End If
    'һ���õ�����ͷ
    strSQL = "select ����,�շ���Ŀ��,���÷�Χ,���,�߶�,����ɫ from �շѼ��ʵ� where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    '�жϿ�����
    If rsTmp.EOF Then
        If mstrInNO <> "" Then
            strSQL = "zl_�շѼ��ʵ�_Normalize('" & mstrInNO & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            MsgBox "����ѡ��ļ��ʵ������Զ�����ʵ������Ѿ���ɾ����" & vbCrLf & _
                "�õ����Ѹ�Ϊ����ͨ���ʵ�����������ˢ���б�", vbExclamation, gstrSysName
        Else
            MsgBox "����ѡ��ļ��ʵ������Զ�����ʵ������Ѿ���ɾ����" & vbCrLf & _
                "�����½��뱾����", vbExclamation, gstrSysName
        End If
        Exit Function
    End If
    If mbytUseType = Use���� And Mid(rsTmp("���÷�Χ"), 1, 1) <> "1" Then
        MsgBox "�����ʵ�����֧��������ʣ�������ˢ���б�", vbExclamation, gstrSysName
        Exit Function
    End If
    If mbytUseType = UseסԺ And Mid(rsTmp("���÷�Χ"), 2, 1) <> "1" Then
        MsgBox "�����ʵ�����֧��סԺ���ʣ�������ˢ���б�", vbExclamation, gstrSysName
        Exit Function
    End If
    If mbytUseType = Use���ҷ�ɢ And Mid(rsTmp("���÷�Χ"), 3, 1) <> "1" Then
        MsgBox "�����ʵ�����֧�ֿ��ҷ�ɢ���ʣ�������ˢ���б�", vbExclamation, gstrSysName
        Exit Function
    End If
    If mbytUseType = Useҽ������ And Mid(rsTmp("���÷�Χ"), 4, 1) <> "1" Then
        MsgBox "�����ʵ�����֧��ҽ�����Ҽ��ʣ�������ˢ���б�", vbExclamation, gstrSysName
        Exit Function
    End If
    '�ı䴰�ڵĴ�С
    sngTemp = Me.Width - Me.ScaleWidth   '�õ����ڴ�С��ͻ�����С�Ĳ�ֵ
    Me.Width = rsTmp("���") + sngTemp
    sngTemp = Me.Height - Me.ScaleHeight '�õ����ڴ�С��ͻ�����С�Ĳ�ֵ
    Me.Height = rsTmp("�߶�") + sta.Height + sngTemp
    fraForm.Left = 0: fraForm.Top = 0
    fraForm.Width = rsTmp("���"): fraForm.Height = rsTmp("�߶�")
    fraForm.BackColor = rsTmp("����ɫ")
    
    Me.Caption = "���ʴ���" & " - " & rsTmp("����")
    '�õ��ؼ�����
    mlngRows = rsTmp("�շ���Ŀ��")
    For i = 1 To mlngRows - 1
        Load cbo�շ����(i): Set cbo�շ����(i).Container = fraForm
        Load txt�շ���Ŀ(i): Set txt�շ���Ŀ(i).Container = fraForm
        Load cmdϸĿѡ��(i): Set cmdϸĿѡ��(i).Container = fraForm
        Load txt���㵥λ(i): Set txt���㵥λ(i).Container = fraForm
        Load txt����(i):     Set txt����(i).Container = fraForm
        Load txt��׼����(i): Set txt��׼����(i).Container = fraForm
        Load txtӦ�ս��(i): Set txtӦ�ս��(i).Container = fraForm
        Load txtʵ�ս��(i): Set txtʵ�ս��(i).Container = fraForm
        Load cboִ�п���(i): Set cboִ�п���(i).Container = fraForm
        Load chk����(i):     Set chk����(i).Container = fraForm
    Next
    rsTmp.Close
    '�����õ�������
    strSQL = "select ��Ӧ�ֶ�,���,����,����ֵ,˳���,���,����,���,�߶�,����,ǰ��ɫ,����ɫ,�Ƿ���ʾ,����,�߿���,͸��" & _
        " from �շѼ��ʵ����� where ����ID=[1] order by ˳���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    
    Do Until rsTmp.EOF
        blnContainer = False
        lngIndex = IIf(IsNull(rsTmp("���")), 0, rsTmp("���")) - 1
        Select Case rsTmp("����")
            Case "CheckBox"
                Select Case rsTmp("��Ӧ�ֶ�")
                    Case "���ӱ�־"
                        Set ctlTemp = chk����(lngIndex)
                    Case "�Ӱ��־"
                        Set ctlTemp = chk�Ӱ�
                    Case "��"
                        Set ctlTemp = chk��
                        blnContainer = True
                End Select
                ctlTemp.Caption = IIf(IsNull(rsTmp("����ֵ")), "", rsTmp("����ֵ"))
                ctlTemp.Height = rsTmp("�߶�")
                ctlTemp.ForeColor = rsTmp("ǰ��ɫ")
                ctlTemp.BackColor = rsTmp("����ɫ")
                ctlTemp.Appearance = rsTmp("����")
            Case "ComboBox"
                Select Case rsTmp("��Ӧ�ֶ�")
                    Case "�ѱ�"
                        Set ctlTemp = cbo�ѱ�
                    Case "�Ա�"
                        Set ctlTemp = cbo�Ա�
                    Case "NO"
                        Set ctlTemp = cboNO
                        blnContainer = True
                    Case "������"
                        Set ctlTemp = cbo������
                        blnContainer = True
                    Case "��������"
                        Set ctlTemp = cbo��������
                        cbo��������.Tag = IIf(IsNull(rsTmp("����ֵ")), "", rsTmp("����ֵ"))
                        If cbo��������.Tag <> "" Then
                            cbo��������.Locked = True
                            cbo��������.TabStop = True
                        End If
                    Case "�շ����"
                        Set ctlTemp = cbo�շ����(lngIndex)
                        ctlTemp.Tag = IIf(IsNull(rsTmp("����ֵ")), "", rsTmp("����ֵ"))
                        If ctlTemp.Tag = "0" Then ctlTemp.Tag = ""
                        ctlTemp.Locked = ctlTemp.Tag <> ""
                    Case "ִ�в���"
                        Set ctlTemp = cboִ�п���(lngIndex)
                    Case "Ӥ����"
                        Set ctlTemp = cboBaby
                            'Ӥ����
                        Call LoadPatientBaby(ctlTemp, 0, 0)
                End Select
                ctlTemp.ForeColor = rsTmp("ǰ��ɫ")
                ctlTemp.BackColor = rsTmp("����ɫ")
            Case "CommandButton"
                Select Case rsTmp("��Ӧ�ֶ�")
                    Case "ȡ��"
                        Set ctlTemp = cmdCancel
                        blnContainer = True
                    Case "ȷ��"
                        Set ctlTemp = cmdOK
                        blnContainer = True
                    Case "ϸĿѡ��"
                        Set ctlTemp = cmdϸĿѡ��(lngIndex)
                End Select
                ctlTemp.Caption = IIf(IsNull(rsTmp("����ֵ")), "", rsTmp("����ֵ"))
                ctlTemp.Height = rsTmp("�߶�")
            Case "Label"
                Load lbl(lbl.UBound + 1)
                Set ctlTemp = lbl(lbl.UBound)
                ctlTemp.Caption = Replace(IIf(IsNull(rsTmp("����ֵ")), "", rsTmp("����ֵ")), "[��λ����]", gstr��λ����)
                ctlTemp.Appearance = rsTmp("����")
                ctlTemp.BorderStyle = rsTmp("�߿���")
                ctlTemp.BackStyle = rsTmp("͸��")
                ctlTemp.ForeColor = rsTmp("ǰ��ɫ")
                ctlTemp.BackColor = rsTmp("����ɫ")
                ctlTemp.Height = rsTmp("�߶�")
            Case "TextBox"
                Select Case rsTmp("��Ӧ�ֶ�")
                    Case "����"
                        Set ctlTemp = txtPatient
                    Case "��ʶ��"
                        Set ctlTemp = txt��ʶ��
                    Case "����ID"
                        Set ctlTemp = txt����ID
                    Case "����"
                        Set ctlTemp = txt����
                    Case "����"
                        Set ctlTemp = txt����
                    Case "���˲���"
                        Set ctlTemp = txt���˲���
                    Case "���˿���"
                        Set ctlTemp = txt���˿���
                    Case "��Ժ����"
                        Set ctlTemp = txt��ҳID
                    Case "ʵ�պϼ�"
                        Set ctlTemp = txtʵ��
                    Case "Ӧ�պϼ�"
                        Set ctlTemp = txtӦ��
                    Case "�շ�ϸĿ"
                        Set ctlTemp = txt�շ���Ŀ(lngIndex)
                        ctlTemp.Tag = IIf(IsNull(rsTmp("����ֵ")), "", rsTmp("����ֵ"))
                        
                        If Val(ctlTemp.Tag) > 0 Then
                            '����ȷ��ֵ
                            cbo�շ����(lngIndex).Locked = True
                            txt�շ���Ŀ(lngIndex).Locked = True
                            txt�շ���Ŀ(lngIndex).TabStop = False
                            cmdϸĿѡ��(lngIndex).Enabled = False
                        End If
                    Case "���㵥λ"
                        Set ctlTemp = txt���㵥λ(lngIndex)
                    Case "����"
                        Set ctlTemp = txt����(lngIndex)
                        ctlTemp.Tag = IIf(IsNull(rsTmp("����ֵ")), "", rsTmp("����ֵ"))
                        
                        If Val(ctlTemp.Tag) > 0 Then
                            '����ȷ��ֵ
                            ctlTemp.Locked = True
                            ctlTemp.TabStop = False
                        End If
                    Case "��׼����"
                        Set ctlTemp = txt��׼����(lngIndex)
                    Case "ʵ�ս��"
                        Set ctlTemp = txtʵ�ս��(lngIndex)
                    Case "Ӧ�ս��"
                        Set ctlTemp = txtӦ�ս��(lngIndex)
                    Case "����ʱ��"
                        Set ctlTemp = txtDate
                        blnContainer = True
                End Select
                ctlTemp.Height = rsTmp("�߶�")
                ctlTemp.ForeColor = CLng(rsTmp("ǰ��ɫ"))
                ctlTemp.BackColor = CLng(rsTmp("����ɫ"))
                ctlTemp.Appearance = rsTmp("����")
                ctlTemp.BorderStyle = rsTmp("�߿���")
        End Select
        If blnContainer = True Then
            ctlTemp.Left = 0
            ctlTemp.Top = 0
            ctlTemp.Container.Left = rsTmp("���")
            ctlTemp.Container.Top = rsTmp("����")
            ctlTemp.Container.Width = rsTmp("���")
            ctlTemp.Container.Height = rsTmp("�߶�")
        Else
            ctlTemp.Left = rsTmp("���")
            ctlTemp.Top = rsTmp("����")
        End If
        
        ctlTemp.Width = rsTmp("���")
        varTemp = Split(rsTmp("����"), "|")
        ctlTemp.Font.Name = varTemp(0)
        ctlTemp.Font.Size = varTemp(1)
        ctlTemp.Font.Bold = varTemp(2) = "1"
        ctlTemp.Font.Italic = varTemp(3) = "1"
        ctlTemp.Font.Underline = varTemp(4) = "1"
        ctlTemp.Visible = rsTmp("�Ƿ���ʾ") = 1
        ctlTemp.TabIndex = rsTmp("˳���")
        rsTmp.MoveNext
    Loop
    

    '�������ݱ�Ҫ��ɵĹ������ý��沼��
    cboBaby.Enabled = mbytInState = staִ��
    Select Case mbytInState
        Case staִ��  'ִ��
            If mbytUseType <> Use���� And (InStr(mstrPrivsOpt, "סԺ����") = 0 Or mstrInNO <> "") Then
                fra��.Visible = False
                chk��.Visible = False
            End If
        Case sta���� '����
            fraNO.Enabled = False
            fra������.Enabled = False
            fraʱ��.Enabled = False
            fraForm.Enabled = False
            'Ϊ��ʹ�ؼ����ɼ�
            fraOK.Visible = False
            
            If mblnViewCancel = False Then
                fra��.Visible = False
            Else
                fra��.Enabled = False
                chk��.ForeColor = &HFF&
            End If
            cmdCancel.Caption = "�˳�(&X)"
        Case sta���� '����
            fra��.Visible = False
            fraForm.Enabled = False
            fraNO.Enabled = False
        Case sta���� '����
            fra��.Visible = False
            fraForm.Enabled = False
            fra������.Enabled = False
            fraʱ��.Enabled = False
            fraNO.Enabled = False
    End Select
    
    '��ȡ����ƥ�䷽ʽ
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gbln�����л� '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln�����л�
    If mbytInState = 0 Then
        '����ƥ�䷽ʽ��0-ƴ��,1-���,2-����
        If gbytCode = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf gbytCode = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
    End If
    
    InitFace = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lngCount As Long, strSQL As String

    On Error GoTo errH

    '�Զ�ʶ��Ӱ�
    If mbytInState <> 2 And mstrInNO = "" Then
        If OverTime(zlDatabase.Currentdate) Then chk�Ӱ�.Value = Checked
    End If

    '��ѡ�Ա�
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo�Ա�.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            rsTmp.MoveNext
        Next
    End If

    '��ѡ�ѱ�
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ� Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 And cbo�ѱ�.ListIndex = -1 Then cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
            rsTmp.MoveNext
        Next
    Else
        MsgBox "û�г�ʼ���ѱ����ȵ��ѱ�����н������ã�", vbInformation, gstrSysName
        Exit Function
    End If

   
    If FillDept() = 0 Then  '������listindex=0ʱ����FillDoctor
        If mbytUseType = Use���� Then
            MsgBox "û�г�ʼ�������ٴ�����,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Else
            MsgBox "û�г�ʼ��סԺ�ٴ�����,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        End If
        Exit Function
    End If

    '�����շ����
    If gstr�շ���� = "" Then
        strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where ���� Not In ('1','4','5','6','7') Order by ���"
    Else
        strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where ���� In(" & gstr�շ���� & ")  And ���� Not In('4','5','6','7') Order by ���"
    End If
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsClass.EOF Then
        MsgBox "û�����ÿ��õ��շ���𣨱�����ģʽ��֧��ҩƷ���ʣ���" & vbCrLf & _
               "�����ڱ��ز��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    For i = 0 To mlngRows - 1
        cbo�շ����(i).Clear
    Next
    
    lngCount = 1
    Do Until mrsClass.EOF
        For i = 0 To mlngRows - 1
            cbo�շ����(i).AddItem lngCount & "-" & mrsClass("���")
            cbo�շ����(i).ItemData(cbo�շ����(i).NewIndex) = Asc(mrsClass("����"))
            
            '����Ԥ��ֵ
            If mrsClass("����") = cbo�շ����(i).Tag Then
                cbo�շ����(i).ListIndex = cbo�շ����(i).NewIndex
                cbo�շ����(i).TabStop = False
            End If
        Next
        lngCount = lngCount + 1
        mrsClass.MoveNext
    Loop

    mblnOne = (mrsClass.RecordCount = 1)

    'ִ�в���(���������סԺ)
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID and B.������� IN(" & IIf(mbytUseType = Use����, "1", "2") & ",3) " & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        " Order by B.�������,A.����"
    Set mrsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsUnit.EOF Then
        MsgBox "û�г�ʼ��������Ϣ,�����޷�����ִ�в��š����ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If

    '��������
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")

    If mbytInState = 0 Then Set mrsWarn = GetUnitWarn
    Set mrsInfo = New ADODB.Recordset

    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt��ҳID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub ReCalcInsure()
'���ܣ��޸ĵ���ʱ,���¼���ͳ������������Ϣ
    Dim i As Integer, j As Integer
    Dim strInfo As String
    
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!����) Then
            For i = 1 To mobjBill.Details.Count
                For j = 1 To mobjBill.Details(i).InComes.Count
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(i).�շ�ϸĿID, mobjBill.Details(i).InComes(j).ʵ�ս��, False, mrsInfo!����, _
                        mobjBill.Details(i).ժҪ & "||" & mobjBill.Details(i).����)
                    If strInfo <> "" Then
                        mobjBill.Details(i).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details(i).���մ���ID = Val(Split(strInfo, ";")(1))
                        mobjBill.Details(i).InComes(j).ͳ���� = Val(Split(strInfo, ";")(2))
                        mobjBill.Details(i).���ձ��� = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(i).ժҪ = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(i).Detail.���� = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                Next
            Next
        End If
    End If
End Sub

Private Function BillExistInsure(strNO As String) As Integer
'���ܣ��ж�ָ����סԺ���ʵ����Ƿ��ҽ�����˼ǵ���
'������strNO=���ʵ��ݺ�
'���أ�������򷵻ز�������
'˵����1.ֻ��סԺҽ������,�������ﲡ�˵�ҽ������
'      2.���ʱ�ֻ���ص�һ�����˵�����,������ҲӦ��ֻ��һ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.���� From סԺ���ü�¼ A,������ҳ B" & _
        " Where A.��¼����=2 And A.��¼״̬ IN(1,3) And B.���� is Not NULL" & _
        " And A.NO=[1] And A.����ID=B.����ID And A.��ҳID=B.��ҳID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    If Not rsTmp.EOF Then BillExistInsure = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckMediCareItem(ByVal lng�շ�ϸĿID As Long, ByVal int���� As Integer, ByVal str�շ���Ŀ���� As String, _
    ByVal bln���� As Boolean, Optional ByVal strPriceGrade As String) As Boolean
'���ܣ��ж��շ���Ŀ�Ƿ������˱���֧����Ŀ
    Dim rsTmp As ADODB.Recordset, strSQL As String, dbl�۸� As Double, rs�۸� As ADODB.Recordset
    Dim strWherePriceGrade As String
    
    CheckMediCareItem = True
    
    If gbytҽ�������� = 0 Then Exit Function
    If gclsInsure.GetCapability(support��������ҽ����Ŀ, , int����) Then
        Exit Function
    End If
    On Error GoTo errH

   '���˺� ����:27286 ���۵ļ۸�Ϊ��Ĳ����м����� ����:2010-01-07 15:13:45
    If bln���� Then
        If strPriceGrade <> "" Then
            strWherePriceGrade = _
                "      And (b.�۸�ȼ� = [2]" & vbNewLine & _
                "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
                "              And Not Exists(Select 1" & vbNewLine & _
                "                             From �շѼ�Ŀ" & vbNewLine & _
                "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
                "                                   And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
        Else
            strWherePriceGrade = " And b.�۸�ȼ� Is Null"
        End If
        strSQL = _
        " Select  B.�ּ� " & _
        " From �շѼ�Ŀ B " & _
        " Where   ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
        "       And B.�շ�ϸĿID=[1]" & vbNewLine & _
                strWherePriceGrade
        Set rs�۸� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ǰ�۸�", lng�շ�ϸĿID, strPriceGrade)
        If rs�۸�.EOF = False Then
            dbl�۸� = Val(NVL(rs�۸�!�ּ�))
        Else
            dbl�۸� = 0
        End If
        If dbl�۸� = 0 Then Exit Function
    End If
    
    strSQL = "Select �շ�ϸĿID From ����֧����Ŀ Where �շ�ϸĿID=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng�շ�ϸĿID, int����)
        
    If rsTmp.RecordCount = 0 Then
        If gbytҽ�������� = 1 Then
            If MsgBox("û������""" & str�շ���Ŀ���� & """��Ӧ�ı�����Ŀ,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckMediCareItem = False
            End If
        ElseIf gbytҽ�������� = 2 Then
            MsgBox "û������""" & str�շ���Ŀ���� & """��Ӧ�ı�����Ŀ!", vbInformation, gstrSysName
            CheckMediCareItem = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check�������() As Integer
'���ܣ���鵱ǰ���˵ļ��ʷ�����Ŀ�ķ�������Ƿ�һ��
'˵������Ϊ�������������۲���,�����д˼��
'���أ���һ�µķ�����,Ϊ0ʱ����
    Dim i As Integer
    
    If mrsInfo.State = 0 Then Exit Function
    For i = 1 To mobjBill.Details.Count
        If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
            'סԺ���˻�סԺ���۲���,������ֻ�������������Ŀ
            If mobjBill.Details(i).Detail.������� = 1 Then
                MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """������������,�ò��˲���ʹ��.", vbInformation, gstrSysName
                Check������� = i: Exit Function
            End If
        ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
            '������Ժ����(ҽ������)���������۲���,������ֻ������סԺ����Ŀ
            If mobjBill.Details(i).Detail.������� = 2 Then
                MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """��������סԺ,�ò��˲���ʹ��.", vbInformation, gstrSysName
                Check������� = i: Exit Function
            End If
        End If
    Next
End Function
Private Function Get��������ID() As Long
    If cbo��������.ListIndex <> -1 Then
        Get��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        Get��������ID = IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID)
    End If
End Function
