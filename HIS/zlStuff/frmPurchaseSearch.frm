VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPurchaseSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   5130
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmPurchaseSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2055
      Left            =   1920
      TabIndex        =   39
      Top             =   4680
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   4695
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Χ(&R)"
      TabPicture(0)   =   "frmPurchaseSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��������(&D)"
      TabPicture(1)   =   "frmPurchaseSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra��������"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra�������� 
         Height          =   3810
         Left            =   -74760
         TabIndex        =   48
         Top             =   600
         Width           =   5505
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   1380
            TabIndex        =   50
            Top             =   2640
            Width           =   3765
         End
         Begin VB.CheckBox Chk��Ӧ�� 
            Caption         =   "��Ӧ��"
            Height          =   300
            Left            =   480
            TabIndex        =   19
            Top             =   660
            Width           =   1110
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   20
            Top             =   660
            Width           =   3255
         End
         Begin VB.CommandButton Cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   21
            Top             =   660
            Width           =   255
         End
         Begin VB.CommandButton Cmd���� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   18
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox Txt���� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   17
            Top             =   240
            Width           =   3255
         End
         Begin VB.CheckBox Chk���� 
            Caption         =   "��������"
            Height          =   300
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox Txt������ 
            Height          =   300
            Left            =   1380
            MaxLength       =   8
            TabIndex        =   30
            Top             =   1860
            Width           =   1365
         End
         Begin VB.TextBox Txt����� 
            Height          =   300
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   32
            Top             =   1860
            Width           =   1365
         End
         Begin VB.TextBox Txt��ʼ��Ʊ�� 
            Height          =   300
            Left            =   1380
            TabIndex        =   34
            Top             =   2250
            Width           =   1365
         End
         Begin VB.TextBox Txt������Ʊ�� 
            Height          =   300
            Left            =   3780
            TabIndex        =   36
            Top             =   2250
            Width           =   1365
         End
         Begin VB.CheckBox Chk������ 
            Caption         =   "������"
            Height          =   300
            Left            =   480
            TabIndex        =   22
            Top             =   1050
            Width           =   1155
         End
         Begin VB.TextBox txt������ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            TabIndex        =   23
            Top             =   1050
            Width           =   3255
         End
         Begin VB.CommandButton Cmd������ 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   24
            Top             =   1050
            Width           =   255
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   2
            Left            =   1650
            TabIndex        =   26
            Top             =   1455
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   119996419
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   2
            Left            =   3540
            TabIndex        =   28
            Top             =   1455
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   119996419
            CurrentDate     =   36263
         End
         Begin VB.CheckBox chk�������� 
            Caption         =   "��������"
            Height          =   300
            Left            =   480
            TabIndex        =   25
            Top             =   1462
            Width           =   1095
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��  ��"
            Height          =   180
            Left            =   750
            TabIndex        =   49
            Top             =   2700
            Width           =   540
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   4
            Left            =   3300
            TabIndex        =   27
            Top             =   1522
            Width           =   180
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   750
            TabIndex        =   29
            Top             =   1920
            Width           =   540
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Left            =   3120
            TabIndex        =   31
            Top             =   1920
            Width           =   540
         End
         Begin VB.Label Lbl��Ʊ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ʊ��"
            Height          =   180
            Left            =   750
            TabIndex        =   33
            Top             =   2310
            Width           =   540
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   35
            Top             =   2310
            Width           =   180
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   3810
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chk�޷�Ʊ 
            Caption         =   "�޷�Ʊ"
            Height          =   180
            Left            =   2760
            TabIndex        =   15
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CheckBox chk�з�Ʊ 
            Caption         =   "�з�Ʊ"
            Height          =   180
            Left            =   720
            TabIndex        =   14
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CheckBox chkYesVerifyBack 
            Caption         =   "������˿�"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2760
            TabIndex        =   13
            Top             =   3080
            Width           =   1215
         End
         Begin VB.CheckBox chkNOVerifyBack 
            Caption         =   "δ����˿�"
            Height          =   180
            Left            =   720
            TabIndex        =   12
            Top             =   3080
            Width           =   1215
         End
         Begin VB.CheckBox chkNot��ֵ�Ĳ� 
            Caption         =   "�Ǹ�ֵ�Ĳĵ���"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2760
            TabIndex        =   9
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chk��ֵ�Ĳ� 
            Caption         =   "��ֵ�Ĳĵ���"
            Enabled         =   0   'False
            Height          =   180
            Left            =   720
            TabIndex        =   8
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chk����������� 
            Caption         =   "�����������"
            Height          =   180
            Left            =   2760
            TabIndex        =   11
            Top             =   2680
            Value           =   1  'Checked
            Width           =   1425
         End
         Begin VB.TextBox txt��ʼNo 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   0
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt����NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "δ��˵���"
            Height          =   300
            Left            =   480
            TabIndex        =   2
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "����˵���"
            Height          =   300
            Left            =   480
            TabIndex        =   5
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chkStrike 
            Caption         =   "��������"
            Enabled         =   0   'False
            Height          =   180
            Left            =   720
            TabIndex        =   10
            Top             =   2680
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   3
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   120913923
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   4
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   86900739
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   86900739
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   7
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   86507523
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   47
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   2640
            TabIndex        =   46
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�������"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   45
            Top             =   1905
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   3345
            TabIndex        =   44
            Top             =   1905
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   43
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   42
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   38
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   37
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmPurchaseSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '�����ַ���
Private BlnAdvance As Boolean '�Ƿ�չ��
Private mdatStart As Date   '��ʼʱ��
Private mdatEnd As Date     '����ʱ��
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mfrmMain As Form    '������
Private mstrSelectTag As String     '��ǰѡ��Ķ���
Private mstrOthers(0 To 13) As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��,13-������Ϣ
Public lng����ID As Long
Private mstr��ֵ�Ĳ� As String      '������¼��ֵ�Ĳ��Ƿ�ѡ��
Private mint�з�Ʊ As Integer
Private mint�޷�Ʊ As Integer

Public Function GetSearch(ByVal frmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strOthers() As String, ByRef str��ֵ�Ĳ� As String, ByRef intNo��Ʊ As Integer, _
        ByRef intYes��Ʊ As Integer) As String
    mstrFind = ""
    mstrSelectTag = ""
    Set mfrmMain = frmMain
    If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    strOthers = mstrOthers
    str��ֵ�Ĳ� = mstr��ֵ�Ĳ�
    intNo��Ʊ = mint�޷�Ʊ
    intYes��Ʊ = mint�з�Ʊ
End Function


Private Sub chkStrike_Click()
    chk�����������.Enabled = chkStrike.Value = 1
End Sub

Private Sub chkStrike_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub
Private Sub chk�����������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Chk��Ӧ��_Click()
    txt��Ӧ��.Enabled = IIf(Chk��Ӧ��.Value = 1, True, False)
    Cmd��Ӧ��.Enabled = IIf(Chk��Ӧ��.Value = 1, True, False)
    
End Sub

Private Sub Chk��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chk��Ӧ��.Value = 1 Then
        txt��Ӧ��.SetFocus
    Else
        Chk������.SetFocus
    End If
End Sub


Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If chk���.Value = 0 Then
            cmdȷ��.SetFocus
        Else
            SendKeys vbTab
        End If
    End If
    
End Sub

Private Sub chk��������_Click()
    dtp��ʼʱ��(2).Enabled = chk��������.Value = 1
    dtp����ʱ��(2).Enabled = dtp��ʼʱ��(2).Enabled
End Sub

Private Sub chk��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Chk������_Click()
    Me.txt������.Enabled = IIf(Chk������.Value = 1, True, False)
    Cmd������.Enabled = IIf(Chk������.Value = 1, True, False)
End Sub

Private Sub Chk������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        If Chk������.Value = 1 Then
            txt������.SetFocus
        
        Else
            Txt������.SetFocus
        End If
    End If
End Sub

Private Sub chk����_Click()
    dtp��ʼʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    dtp����ʱ��(0).Enabled = IIf(chk����.Value = 1, True, False)
    chkNOVerifyBack.Enabled = IIf(chk����.Value = 1, True, False)
    If chk����.Value = 0 Then chkNOVerifyBack.Value = 0
End Sub

Private Sub chk���_Click()
    dtp��ʼʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    dtp����ʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk���.Value = 1, True, False)
    chk�����������.Enabled = chkStrike.Value = 1
    chkYesVerifyBack.Enabled = IIf(chk���.Value = 1, True, False)
    If chk���.Value = 0 Then chkYesVerifyBack.Value = 0
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk����_Click()
    Txt����.Enabled = IIf(Chk����.Value = 1, True, False)
    Cmd����.Enabled = IIf(Chk����.Value = 1, True, False)
End Sub

Private Sub Chk����_GotFocus()
    sstFilter.Tab = 1
    Chk����.SetFocus
End Sub

Private Sub Chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chk����.Value = 1 Then
        Txt����.SetFocus
    ElseIf Chk��Ӧ��.Visible = True Then
        Chk��Ӧ��.SetFocus
    End If
End Sub



Private Sub Cmd��Ӧ��_Click()
    Dim rsTemp As New Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txt��Ӧ��.hwnd)
    
    gstrSQL = "" & _
        "   Select id,�ϼ�ID,����,����,����,ĩ�� " & _
        "   From ��Ӧ�� " & _
        "   where (substr(����,5,1)=1 And (վ��=[1] or վ�� is null) Or Nvl(ĩ��,0)=0) " & _
        "   Start with �ϼ�ID is null connect by prior ID =�ϼ�ID order by level,ID"
    
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 2, "��Ӧ��ѡ��", True, "", "��ѡ����ϲ��ϵĹ�Ӧ��", True, True, True, vRect.Left - 15, vRect.Top, txt��Ӧ��.Height, blnCancel, False, False, gstrNodeNo)
        
    If rsTemp Is Nothing Or blnCancel Then Exit Sub
    If rsTemp.State <> 1 Then Exit Sub
    
    With rsTemp
        txt��Ӧ��.Text = zlStr.NVL(!����)
        txt��Ӧ��.Tag = zlStr.NVL(!Id)
    End With
End Sub

Private Sub Cmdȡ��_Click()
    mstrFind = ""
    Dim i As Integer
    For i = 0 To 13
        mstrOthers(i) = ""
    Next
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim δ��������� As String
    Dim ����������� As String
    
    mint�з�Ʊ = 0
    mint�޷�Ʊ = 0
    '�������
    If Chk����.Value = 1 Then
        If Txt����.Tag = 0 Then
            MsgBox "��ѡ�����ѯ��������Ϣ��", vbInformation, gstrSysName
            Me.Txt����.SetFocus
            Exit Sub
        End If
    End If
    If Chk��Ӧ��.Value = 1 Then
        If txt��Ӧ��.Tag = 0 Then
            MsgBox "��ѡ�����ѯ�����Ĺ�Ӧ����Ϣ��", vbInformation, gstrSysName
            Me.txt��Ӧ��.SetFocus
            Exit Sub
        End If
    End If
    If Chk������.Value = 1 Then
        If txt������.Tag = 0 Then
            MsgBox "��ѡ�����ѯ��������������Ϣ��", vbInformation, gstrSysName
            Me.txt������.SetFocus
            Exit Sub
        End If
    End If
    
    If chk����.Value = 0 And chk���.Value = 0 Then
        MsgBox "�Բ��𣬱���ѡ��һ���������ڻ����������!", vbInformation, gstrSysName
        chk����.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    For i = 0 To 13
        mstrOthers(i) = ""
    Next
    
    mstrFind = ""
    '������ѯ����
    '������Χ:[1]-�ⷿid,[2]:��ʼ��������,[3]������������,[4]��ʼ�������,[5] �����������,[6]-��¼״̬,[7]��ʼ���ݺ�,[8]�������ݺ�,[9]����id,[10]�Է�����id,[11]������,[12]�����[13]-��Ӧ��ID,[14]-������,[15]-��ʼ��������,[16]-������������,[17]-��ʼ��Ʊ��,[18]-������Ʊ��
    mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
    mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    mstrOthers(0) = IIf(chkStrike.Value = 1, "0", "1")
    
    'δ����µ�������
    If chkNOVerifyBack.Value = 0 Then '����ѡδ����˿⣬ֻ��ʾ����
        δ��������� = δ��������� & " and nvl(a.��ҩ��ʽ,0)=0 "
    End If
    '������µ�������
    If chkStrike.Value = 1 Then '������
        ����������� = IIf(chk�����������, "", " And nvl(A.����ID,0)=0 ")
    Else
        ����������� = "and a.��¼״̬ =[6]"
    End If
    If chkYesVerifyBack.Value = 0 Then  '����ѡ������˿⣬ֻ��ʾ����
        ����������� = ����������� & " and nvl(a.��ҩ��ʽ,0)=0 "
    End If
    
    If chk����.Value = 1 And chk���.Value = 1 Then '��������˺�δ��˵���
    
        mstrFind = " And ((A.�������� Between [2] And [3] and A.������� is null " & δ��������� & " )  or (A.������� Between [4] And [5] " & ����������� & "))"
        
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
        
    ElseIf chk���.Value = 1 Then

        mstrFind = " And A.������� Between [4] And [5] " & �����������

        mdatVerifyStart = Format(dtp��ʼʱ��(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp����ʱ��(1), "yyyy-mm-dd")
    ElseIf chk����.Value = 1 Then
        mstrFind = " And (A.�������� Between [2] And [3] and A.������� is null " & δ��������� & ")  "
        mdatStart = Format(dtp��ʼʱ��(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp����ʱ��(0), "yyyy-mm-dd")
    End If
    
    '��Ʊ
    If chk�з�Ʊ.Value = 1 And chk�޷�Ʊ.Value = 0 Then
        mstrFind = mstrFind & " And e.��Ʊ�� is not null "
        mint�з�Ʊ = 1
        mint�޷�Ʊ = 0
    ElseIf chk�޷�Ʊ.Value = 1 And chk�з�Ʊ.Value = 0 Then
        mstrFind = mstrFind & " And e.��Ʊ�� is null "
        mint�з�Ʊ = 0
        mint�޷�Ʊ = 1
    End If
    
    If chk��ֵ�Ĳ�.Value = 1 And chkNot��ֵ�Ĳ�.Value = 0 Then '��ֵ�Ĳ�
        mstr��ֵ�Ĳ� = " and  (a.����id > 1 or d.��ֵ����=1) "
    ElseIf chk��ֵ�Ĳ�.Value = 0 And chkNot��ֵ�Ĳ�.Value = 1 Then '�Ǹ�ֵ�Ĳ�
        mstr��ֵ�Ĳ� = " and (d.��ֵ����=0 or d.��ֵ���� is null) " '����1���ǲ�����˵ĵ���
    Else
        mstr��ֵ�Ĳ� = ""
    End If
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
        Me.txt��ʼNo = UCase(LTrim(Me.txt��ʼNo))
        intYear = Format(sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt��ʼNo) < 8 Then Me.txt��ʼNo = strYear & String(7 - Len(txt��ʼNo), "0") & Me.txt��ʼNo
    End If
    If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
        Me.txt����NO = UCase(LTrim(Me.txt����NO))
        intYear = Format(sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt����NO) < 8 Then Me.txt����NO = strYear & String(7 - Len(txt����NO), "0") & Me.txt����NO
    End If
    
    mstrOthers(1) = Trim(Me.txt��ʼNo.Text)
    mstrOthers(2) = Trim(Me.txt����NO.Text)

    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No >= [7] And A.No <=[8] "
    If Me.txt��ʼNo <> "" And Me.txt����NO = "" Then mstrFind = mstrFind & " And A.No >= [7] "
    If Me.txt��ʼNo = "" And Me.txt����NO <> "" Then mstrFind = mstrFind & " And A.No <=[8] "
    
    '��չ��ѯ����
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id),5-������,
    ' 6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��
    
    If Chk����.Value = 1 Then
        lng����ID = Txt����.Tag
        mstrFind = mstrFind & " And A.ҩƷID+0=[9]"
        mstrOthers(3) = Txt����.Tag
    End If
    If Me.Txt������ <> "" Then
        mstrFind = mstrFind & " And A.������ like '" & Me.Txt������ & "%'"
        mstrOthers(5) = Trim(Me.Txt������) & "%"
    End If
    If Me.Txt����� <> "" Then
        mstrFind = mstrFind & " And A.����� like [12]"
        mstrOthers(6) = Trim(Me.Txt�����) & "%"
    End If
    
    If Chk��Ӧ��.Value = 1 Then
        mstrFind = mstrFind & " And A.��ҩ��λID+0=[13]"
        mstrOthers(7) = txt��Ӧ��.Tag
    End If
    If Chk������.Value = 1 Then
        mstrFind = mstrFind & " And A.����=[14]"
        mstrOthers(8) = txt������.Text
    End If
    If chk��������.Value = 1 Then
        mstrFind = mstrFind & " And A.�������� Between [15] And [16] "
        mstrOthers(9) = Format(dtp��ʼʱ��(2), "yyyy-mm-dd")
        mstrOthers(10) = Format(dtp����ʱ��(2), "yyyy-mm-dd")
    End If
    mstrOthers(11) = Trim(Txt��ʼ��Ʊ��.Text)
    mstrOthers(12) = Trim(Txt������Ʊ��.Text)
    If Trim(Txt��ʼ��Ʊ��.Text) <> "" Or Trim(Txt������Ʊ��.Text) <> "" Then
         mstrFind = mstrFind & "   And Exists(Select 1 From Ӧ����¼ D Where a.Id=d.�շ�ID And  D.ϵͳ��ʶ=5 And D.��¼����=0 "
        If Me.Txt��ʼ��Ʊ�� <> "" And Me.Txt������Ʊ�� <> "" Then mstrFind = mstrFind & " And d.��Ʊ�� >= [17] And d.��Ʊ�� <=[18] "
        If Me.Txt��ʼ��Ʊ�� <> "" And Me.Txt������Ʊ�� = "" Then mstrFind = mstrFind & " And d.��Ʊ�� >=[17] "
        If Me.Txt��ʼ��Ʊ�� = "" And Me.Txt������Ʊ�� <> "" Then mstrFind = mstrFind & " And d.��Ʊ�� <=[18] "
        mstrFind = mstrFind & ")"
    End If
    If gblnCode = True And Trim(txt����.Text) <> "" Then
        mstrOthers(13) = UCase(Trim(txt����.Text))
        mstrFind = mstrFind & " And (A.��Ʒ���� Like [19] Or A.�ڲ����� Like [19])"
    End If
    
    Unload Me
End Sub

Private Sub Cmd������_Click()
    Dim rsTemp As New Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txt������.hwnd)
    
    gstrSQL = "Select rownum as id,null as �ϼ�id,����,����,����,1 as ĩ�� From ���������� " & _
              "Where (վ��=[1] or վ�� is null) "
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, "����������ѡ��", True, "", "ѡ�����������̻���", False, False, True, vRect.Left - 15, vRect.Top, txt������.Height, blnCancel, False, False, gstrNodeNo)
    
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    
    If rsTemp Is Nothing Then
        txt������.SetFocus
        Exit Sub
    End If
    If rsTemp.State <> 1 Then
        txt������.SetFocus
        Exit Sub
    End If
    With rsTemp
        txt������.Tag = 1
        txt������.Text = zlStr.NVL(!����)
        chk��������.SetFocus
    End With
End Sub

Private Sub Cmd����_Click()
    Dim RecReturn As Recordset
    
    Set RecReturn = Frm����ѡ����.ShowMe(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt���� = "[" & RecReturn!���� & "]" & RecReturn!����
    Txt����.Tag = RecReturn!����ID
        
    If Chk��Ӧ��.Visible = True Then
        Chk��Ӧ��.SetFocus
    End If
End Sub

Private Sub dtp����ʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp��ʼʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.dtp����ʱ��(Index).SetFocus
End Sub


Private Sub Form_Load()
    Me.dtp����ʱ��(0) = sys.Currentdate
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    
    Me.dtp��ʼʱ��(0) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    Me.dtp��ʼʱ��(1) = Me.dtp��ʼʱ��(0)
    
    
    Me.dtp��ʼʱ��(2) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    Me.dtp����ʱ��(2) = Me.dtp����ʱ��(0)
    
    lbl����.Visible = gblnCode
    txt����.Visible = gblnCode
    
    Me.txt��Ӧ��.Tag = 0
    Me.Txt����.Tag = 0
    Me.txt������.Tag = 0
    lng����ID = 0
    chk�����������.Enabled = False
    chk��ֵ�Ĳ�.Enabled = True
    chkNot��ֵ�Ĳ�.Enabled = True
    '�򿪼�¼��
    sstFilter.Tab = 0
    BlnAdvance = False
    
End Sub

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo ErrHandle
    CheckCompete = False
    
    gstrSQL = "" & _
        "   Select id,�ϼ�ID,����,����,ĩ��,���� " & _
        "   From ��Ӧ�� " & _
        "   Where ���� is Not NULL " & _
        "       And  (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
        "       And (substr(����,5,1)=1 And (վ��=[1] or վ�� is null) Or Nvl(ĩ��,0)=0) " & _
        "   Start with �ϼ�ID is NULL Connect by prior id=�ϼ�id"
    Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrNodeNo)
    
    With rsCompete
        If .EOF Then
            .Close
            MsgBox "���Ĺ�Ӧ����Ϣ��ȫ�����ڹ�ҩ��λ�������������Ĺ�Ӧ����Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    gstrSQL = "Select ����,����,���� From ���������� where (վ��=[1] Or վ�� is null) "
    Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-����������", gstrNodeNo)
    With rsCompete
        If .EOF Then
            MsgBox "������������Ϣ��ȫ,�����ֵ����������������������Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckCompete = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            Case "Provider"
                txt��Ӧ��.SetFocus
                txt��Ӧ��.SelStart = 0
                txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
        
            Case "Maker"
                txt������.SetFocus
                txt������.SelStart = 0
                txt������.SelLength = Len(txt������.Text)
            
            Case "Booker"
                Txt������.SetFocus
                Txt������.SelStart = 0
                Txt������.SelLength = Len(Txt������.Text)
            Case "Verify"
                Txt�����.SetFocus
                Txt�����.SelStart = 0
                Txt�����.SelLength = Len(Txt�����.Text)
        End Select
        Cancel = True
    End If
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Provider"
                    txt��Ӧ��.Text = .TextMatrix(.Row, 3)
                    txt��Ӧ��.Tag = .TextMatrix(.Row, 0)
                    Chk������.SetFocus
                Case "Maker"
                    txt������.Text = .TextMatrix(.Row, 1)
                    txt������.Tag = 1
                    chk��������.SetFocus
                Case "Booker"
                    Txt������ = .TextMatrix(.Row, 2)
                    Txt�����.SetFocus
                Case "Verify"
                    Txt����� = .TextMatrix(.Row, 2)
                    Txt��ʼ��Ʊ��.SetFocus
            End Select
            .Visible = False
            
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
        End If
    End With
End Sub

Private Sub sstFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        If sstFilter.Tab = 0 Then
            txt��ʼNo.SetFocus
        Else
            Chk����.SetFocus
        End If
    End If
End Sub

Private Sub sstFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        If sstFilter.Tab = 0 Then
            txt��ʼNo.SetFocus
        Else
            Chk����.SetFocus
        End If
    End If
    
End Sub

Private Sub txt��Ӧ��_GotFocus()
'    Tvw.Visible = False
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim recTmp As New Recordset
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(txt��Ӧ��)) <> "" Then
        txt��Ӧ�� = UCase(txt��Ӧ��)
        
        gstrSQL = "" & _
            "   Select id,����,����,���� " & _
            "   From ��Ӧ�� " & _
            "   where  ĩ��=1 And (վ��=[2] or վ�� is null) " & _
            "           And (substr(����,5,1)=1 Or Nvl(ĩ��,0)=0) " & _
            "           And (���� like [1] or ���� like [1] or ���� like [1])"
        
        Set recTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���Ĺ�Ӧ��", IIf(gstrMatchMethod = "0", "%", "") & txt��Ӧ��.Text & "%", gstrNodeNo)
        With recTmp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                txt��Ӧ��.Tag = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Provider"
                Set mshSelect.Recordset = recTmp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + txt��Ӧ��.Top + txt��Ӧ��.Height
                    .Left = sstFilter.Left + fra��������.Left + txt��Ӧ��.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 800
                    .ColWidth(3) = .Width - .ColWidth(1) - .ColWidth(2)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                End With
            Else
                txt��Ӧ�� = !����
                txt��Ӧ��.Tag = !Id
            End If
        End With
    End If
    
    If Chk������.Value = 1 Then
        txt������.SetFocus
    Else
        Chk������.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, 68, lng�ⷿID)
        End If
        OS.PressKey (vbKeyTab)
    End If

End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt������Ʊ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Me.cmdȷ��.SetFocus
End Sub

Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    
    lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, 68, lng�ⷿID)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt��ʼNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt��ʼ��Ʊ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt������Ʊ��.SetFocus
End Sub

Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Txt��ʼ��Ʊ��.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt�����.Text) = "" Then
            Txt��ʼ��Ʊ��.SetFocus
            Exit Sub
        End If
        Txt�����.Text = UCase(Txt�����.Text)
        
        gstrSQL = "" & _
            "   Select ���,����,���� " & _
            "   From ��Ա�� " & _
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) " & _
            "       and (վ��=[2] or վ�� is null) " & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
            "   order by ���"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�����", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%", gstrNodeNo)
        
        With rsTemp
            
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + Txt�����.Top - .Height ' + Txt�����.Height
                    .Left = sstFilter.Left + fra��������.Left + Txt�����.Left + Txt�����.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt����� = IIf(IsNull(!����), "", !����)
                Txt��ʼ��Ʊ��.SetFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub Txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Me.txt������ = "" Then Exit Sub
        If Trim(txt������) = "" Then Exit Sub
        
        strKey = GetMatchingSting(txt������.Text, False)
        gstrSQL = "" & _
            "   Select ����,����,���� " & _
            "   From ���������� " & _
            "   Where (վ��=[2] or վ�� is null) " & _
            "         and (���� like [1] or ���� like upper([1]) or ���� like upper([1]))"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKey, gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Maker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + txt������.Top + txt������.Height
                    .Left = sstFilter.Left + fra��������.Left + txt������.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt������ = IIf(IsNull(!����), "", !����)
                
                txt������.Tag = 1
                chk��������.SetFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt������.Text) = "" Then
            Txt�����.SetFocus
            Exit Sub
        End If
        Txt������.Text = UCase(Txt������.Text)
        
        gstrSQL = "" & _
            "   Select ���,����,���� " & _
            "   From ��Ա�� " & _
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) " & _
            "       and (վ��=[2] or վ�� is null) " & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
            "   order by ���"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra��������.Top + Txt������.Top - .Height '+ Txt������.Height
                    .Left = sstFilter.Left + fra��������.Left + Txt������.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                End With
            Else
                Txt������ = IIf(IsNull(!����), "", !����)
                Me.Txt�����.SetFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt����.Text) = "" Then Exit Sub
    sngLeft = Me.Left + sstFilter.Left + fra��������.Left + Txt����.Left
    sngTop = Me.Top + sstFilter.Top + fra��������.Top + Txt����.Top + Txt����.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - Txt����.Height - 3630
    End If
    
    strKey = Trim(Txt����.Text)
    If Mid(strKey, 1, 1) = "[" Then
        If InStr(2, strKey, "]") <> 0 Then
            strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
        Else
            strKey = Mid(strKey, 2)
        End If
    End If
    
    Set RecReturn = FrmMulitSel.ShowSelect(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strKey, sngLeft, sngTop, Txt����.Width, Txt����.Height)
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt���� = "[" & RecReturn!���� & "]" & RecReturn!����
    Txt����.Tag = RecReturn!����ID
    
    If Chk��Ӧ��.Visible = True Then
        If Chk��Ӧ��.Value = 1 Then
            txt��Ӧ��.SetFocus
        Else
            Chk��Ӧ��.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

