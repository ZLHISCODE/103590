VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClientUpgradeConfigure 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   16020
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkAllBefUpgrade 
      Caption         =   "Ԥ����"
      Height          =   280
      Left            =   870
      TabIndex        =   38
      Top             =   600
      Width           =   870
   End
   Begin VB.CheckBox chkAllUpgrade 
      Caption         =   "����"
      Height          =   280
      Left            =   165
      TabIndex        =   37
      Top             =   600
      Width           =   660
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3630
      ScaleHeight     =   255
      ScaleWidth      =   2625
      TabIndex        =   32
      Top             =   195
      Width           =   2650
      Begin VB.TextBox txtFind 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   45
         TabIndex        =   33
         Text            =   "������ͻ��ˡ�IP�����š���;"
         Top             =   30
         Width           =   2650
      End
   End
   Begin VB.PictureBox picMonthSet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   28
      Top             =   4785
      Visible         =   0   'False
      Width           =   255
      Begin VB.CommandButton cmdMonthSet 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -30
         Picture         =   "frmClientUpgradeConfigure.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   -30
         Width           =   285
      End
   End
   Begin VB.PictureBox Picpgb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   15
      ScaleHeight     =   735
      ScaleWidth      =   5475
      TabIndex        =   24
      Top             =   5325
      Visible         =   0   'False
      Width           =   5500
      Begin MSComctlLib.ProgressBar pgbThis 
         Height          =   390
         Left            =   60
         TabIndex        =   25
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   688
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblPer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "100 %"
         ForeColor       =   &H80000001&
         Height          =   180
         Left            =   4920
         TabIndex        =   27
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "INFO"
         Height          =   180
         Left            =   105
         TabIndex        =   26
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.CheckBox chkClientRepair 
      Caption         =   "���ÿͻ����޸�"
      Height          =   270
      Left            =   12900
      TabIndex        =   3
      Top             =   195
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "ˢ��(&Q)"
      Height          =   300
      Left            =   6435
      TabIndex        =   1
      Top             =   180
      Width           =   1000
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      Height          =   800
      Index           =   2
      Left            =   5370
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3975
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      Height          =   800
      Index           =   5
      Left            =   9270
      ScrollBars      =   1  'Horizontal
      TabIndex        =   15
      Top             =   4035
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      Height          =   800
      Index           =   1
      Left            =   5355
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2805
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      Height          =   800
      Index           =   0
      Left            =   5325
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1650
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7560
      ScaleHeight     =   315
      ScaleWidth      =   3525
      TabIndex        =   18
      Top             =   165
      Width           =   3525
      Begin VB.OptionButton optStatus 
         Caption         =   "����ʧ��"
         Height          =   270
         Index           =   2
         Left            =   2535
         TabIndex        =   7
         Top             =   45
         Width           =   1065
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "δ����"
         Height          =   270
         Index           =   0
         Left            =   1605
         TabIndex        =   6
         Top             =   45
         Width           =   960
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "����"
         Height          =   270
         Index           =   4
         Left            =   930
         TabIndex        =   5
         Top             =   45
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.Label lbloptStatus 
         AutoSize        =   -1  'True
         Caption         =   "����״̬"
         Height          =   180
         Left            =   75
         TabIndex        =   19
         Top             =   75
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdAllCollect 
      Caption         =   "ȫ���ռ�(&R)"
      Height          =   300
      Left            =   12900
      TabIndex        =   4
      Top             =   555
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picBtn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   60
      ScaleHeight     =   345
      ScaleWidth      =   15735
      TabIndex        =   17
      Top             =   6210
      Width           =   15735
      Begin VB.CommandButton cmdkillProcess 
         Caption         =   "�ͻ��˽��̹���(&P)"
         Height          =   300
         Left            =   7065
         TabIndex        =   35
         Top             =   0
         Width           =   1800
      End
      Begin VB.CommandButton cmdClientModify 
         Caption         =   "�ͻ��˿����޸�(&M)"
         Height          =   300
         Left            =   4875
         TabIndex        =   34
         Top             =   0
         Width           =   1965
      End
      Begin VB.PictureBox picTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   11835
         ScaleHeight     =   240
         ScaleWidth      =   1920
         TabIndex        =   30
         Top             =   15
         Width           =   1945
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   315
            Left            =   -30
            TabIndex        =   31
            ToolTipText     =   "��ʱ���ǰ����Ԥ��������ʱ����Ժ������ʽ����"
            Top             =   -30
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   106299395
            CurrentDate     =   42691
         End
      End
      Begin VB.CommandButton cmdTimeSet 
         Caption         =   "Ԥ����ʱ������(&T)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   13905
         TabIndex        =   12
         Top             =   0
         Width           =   1800
      End
      Begin VB.OptionButton optUpgradeTime 
         Caption         =   "��ʱ����"
         Height          =   210
         Index           =   1
         Left            =   10695
         TabIndex        =   11
         ToolTipText     =   "�������ö�ʱ����ʱ��㣬��ʱ�����ǰ�����Ԥ��������ʱ����Ժ�������ʽ����"
         Top             =   60
         Width           =   1170
      End
      Begin VB.OptionButton optUpgradeTime 
         Caption         =   "��������"
         Height          =   210
         Index           =   0
         Left            =   9540
         TabIndex        =   10
         ToolTipText     =   "�Կͻ��˹�ѡ�����󣬿ͻ��˵�½�����Զ���ʽ����"
         Top             =   60
         Width           =   1050
      End
      Begin VB.CommandButton cmdClientAaminSet 
         Caption         =   "�ͻ���ͨ�ù���Ա����(&K)"
         Height          =   300
         Left            =   2475
         TabIndex        =   9
         Top             =   0
         Width           =   2400
      End
      Begin VB.CommandButton cmdFileSeverSet 
         Caption         =   "�����ļ�����������(&S)"
         Height          =   300
         Left            =   60
         TabIndex        =   8
         Top             =   0
         Width           =   2200
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   4185
      Left            =   120
      TabIndex        =   2
      Top             =   585
      Width           =   4905
      _cx             =   8652
      _cy             =   7382
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClientUpgradeConfigure.frx":00F6
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList imgEdit 
      Left            =   10170
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":01CD
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":0767
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":0D01
            Key             =   "ǩ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":1053
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":78B5
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":E117
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":E5DF
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblClientsList 
      AutoSize        =   -1  'True
      Caption         =   "�ͻ��������嵥"
      Height          =   180
      Left            =   135
      TabIndex        =   36
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "�����޸�˵��"
      Height          =   180
      Index           =   2
      Left            =   5370
      TabIndex        =   23
      Top             =   3645
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "�ռ�˵��"
      Height          =   180
      Index           =   5
      Left            =   9270
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "Ԥ����˵��"
      Height          =   180
      Index           =   1
      Left            =   5370
      TabIndex        =   21
      Top             =   2505
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "����˵��"
      Height          =   180
      Index           =   0
      Left            =   5340
      TabIndex        =   20
      Top             =   1350
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   3135
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmClientUpgradeConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrLocationClientsName As String '��λ�ϴ�ѡ���пͻ�������
Private mblnFilter As Boolean '��¼��ǰ�Ƿ���˹���� true - �ѹ��� false -Ϊ����
Private mblnCancel As Boolean
Private mlngClinetNum As Long           '�ͻ�������
Private mlngUpFailClinetNum As Long '����ʧ�ܿͻ�������
Private mlngNotUpClinetNum As Long 'δ�����ͻ�������
Public blnRefreshData As Boolean '�����л�ˢ���жϱ�־
Private mblnAllowEdit As Boolean '��ǵ�ǰ�����Ƿ�����༭
Private mblnAllUpdateClick As Boolean '��Ϊ����chkAllUpgrade.valueֵ��ʱ�����ʽ����chkAllUpgrade_Click������Ҫ��ֵ��������ʽ����
Private mblnAllBefUpgrade As Boolean  '��mblnAllUpdateClickһ��

Private Enum SeverData
    Col_���� = 0
    Col_Ԥ���� = 1
    Col_�ռ� = 2
    Col_�ͻ��� = 3
    Col_IP = 4
    Col_���� = 5
    Col_��; = 6
    Col_���������� = 7
    Col_Ԥ����ʱ�� = 8
    Col_���¼�� = 9
    Col_������� = 10
    Col_Ԥ������� = 11
'    Col_�ռ���� = 12
    Col_�����޸���� = 12
    Col_����˵�� = 13
    Col_Ԥ����˵�� = 14
'    Col_�ռ�˵�� = 15
    Col_�����޸�˵�� = 15
    Col_����Ա = 16
    Col_���� = 17
    Col_���� = 18
End Enum

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
End Sub

Private Sub chkAllBefUpgrade_Click()
    If mblnAllBefUpgrade Then Exit Sub
    If chkAllBefUpgrade.value = 1 Then
        If MsgBox("�Ƿ�Ҫ����ȫ���ͻ�������Ԥ������", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
            mblnAllBefUpgrade = True
            chkAllBefUpgrade.value = 0
            mblnAllBefUpgrade = False
            Exit Sub
        End If
    Else
        If MsgBox("�Ƿ�Ҫȡ��ȫ���ͻ�������Ԥ������", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
            mblnAllBefUpgrade = True
            chkAllBefUpgrade.value = 1
            mblnAllBefUpgrade = False
            Exit Sub
        End If
    End If
    mblnAllBefUpgrade = False
    On Error GoTo errH
    Call UpdateData(Col_Ԥ����, chkAllBefUpgrade.value)
    RefreshData
    '������Ҫ������־
    Call SaveAuditLog(2, "ȫ��Ԥ����/ȡ��ȫ��Ԥ����", "�����пͻ���ִ��" & IIf(chkAllBefUpgrade.value = 1, "", "ȡ��") & "Ԥ��������")
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub chkAllUpgrade_Click()
    If mblnAllUpdateClick Then Exit Sub
    If chkAllUpgrade.value = 1 Then
        If MsgBox("�Ƿ�Ҫ����ȫ���ͻ�������������", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
            mblnAllUpdateClick = True
            chkAllUpgrade.value = 0
            mblnAllUpdateClick = False
            Exit Sub
        End If
    Else
        If MsgBox("�Ƿ�Ҫȡ��ȫ���ͻ�������������", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
            mblnAllUpdateClick = True
            chkAllUpgrade.value = 1
            mblnAllUpdateClick = False
            Exit Sub
        End If
    End If
    mblnAllUpdateClick = False
    On Error GoTo errH
    Call UpdateData(Col_����, chkAllUpgrade.value)
    RefreshData
    '������Ҫ������־
    Call SaveAuditLog(2, "ȫ������/ȡ��ȫ������", "�����пͻ��˽���" & IIf(chkAllUpgrade.value = 1, "", "ȡ��") & "��������")
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub chkClientRepair_Click()
    Dim strSQL As String
    
    On Error Resume Next
    If chkClientRepair.value = 0 Then
        strSQL = "update zltools.ZLReginfo set ���� = '" & 0 & "'where ��Ŀ = '��ֹ�ͻ����޸�'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "update zltools.ZLReginfo set ���� = '" & 1 & "'where ��Ŀ = '��ֹ�ͻ����޸�'"
        gcnOracle.Execute strSQL
    End If
    
End Sub

Private Sub cmdAllCollect_Click()
    Dim i As Long
    Dim strSQL As String
    Dim strUpdateVal As String
    Dim strTemp As String
    
    strTemp = cmdAllCollect.Caption
    strUpdateVal = IIf(strTemp = "ȫ���ռ�(&R)", "1", "0")
    If MsgBox("�Ƿ�Ҫ" & IIf(strUpdateVal = "1", "����", "ȡ��") & "ȫ���ͻ��������ռ���", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub

    On Error GoTo errH
    Call UpdateData(Col_�ռ�, strUpdateVal)
    cmdAllCollect.Caption = IIf(strTemp = "ȫ���ռ�(&R)", "ȫ�����ռ�(&R)", "ȫ���ռ�(&R)")
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub cmdClientAaminSet_Click()
    Load frmClientUpgradeAdmin
    frmClientUpgradeAdmin.Show 1, frmMDIMain
    If frmClientUpgradeAdmin.mblnOk Then
    End If
    Exit Sub
End Sub

Private Sub cmdClientModify_Click()
    Dim blnReturn   As Boolean
    Dim strIp       As String
    Dim strName     As String
    Dim lngRow      As Long
    
    With vsfMain
        If .Row >= .FixedRows Then
            lngRow = .Row
            strIp = .TextMatrix(lngRow, Col_IP)
            strName = .TextMatrix(lngRow, Col_�ͻ���)
            frmClientsEdit.ShowEdit strIp, strName, 1, blnReturn
            If Not blnReturn Then Exit Sub
            Call LoadClientsData
            lngRow = .FindRow(strName, , Col_�ͻ���)
            If lngRow >= .FixedRows Then
                .SetFocus
                .Row = lngRow
                .ShowCell lngRow, Col_�ͻ���
            End If
        End If
    End With
End Sub

Private Sub cmdFileSeverSet_Click()
    Dim frmSeverSet As New frmClientUpgradeSever
    If frmSeverSet.ShowMe(frmMDIMain) = True Then
        cmdRef_Click
    End If
End Sub

'������Ҫ��ɱ�Ľ����б�
Private Sub cmdkillProcess_Click()
    frmKillProcessManage.ShowMe ("0307")
End Sub

Private Sub cmdRef_Click()
    Call RefreshData
End Sub

Private Sub cmdTimeSet_Click()
    Load frmClientUpgradeTime
    frmClientUpgradeTime.Show 1, frmMDIMain
    If frmClientUpgradeTime.mblnOk Then
        LoadClientsData
        FilterData (mstrLocationClientsName)
        InitCombolist
    End If
    Exit Sub
ErrHandle:
    MsgBox "����ʧ�ܡ�" & vbNewLine & err.Description, vbExclamation, gstrSysName
End Sub

Private Sub dtpTime_Change()
'    Dim strNow As String
'    strNow = Format(CurrentDate(), "yyyy-MM-dd") & " 23:00"
'    dtpTime.value = strNow
    Call SaveUpgradeDate
End Sub

Private Sub Form_Load()
'    Call LoadClientsData
'    Call FilterData
    mblnAllowEdit = True
    Call InitVsfMain
    Call LoadSetting
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngTxtHeight As Long
    vsfMain.Height = Me.ScaleHeight - vsfMain.Top - 600
    
    lngTxtHeight = (vsfMain.Height - 960) / 3
    
    If lngTxtHeight < 300 Or Me.ScaleWidth < 8000 Then
        lblExplain.Item(0).Visible = False
        lblExplain.Item(1).Visible = False
        lblExplain.Item(2).Visible = False
'        lblExplain.Item(3).Visible = False
        txtExplain.Item(0).Visible = False
        txtExplain.Item(1).Visible = False
        txtExplain.Item(2).Visible = False
'        txtExplain.Item(3).Visible = False
        vsfMain.Width = Me.ScaleWidth - 100
        picStatus.Visible = False
    Else
        lblExplain.Item(0).Visible = True
        lblExplain.Item(1).Visible = True
        lblExplain.Item(2).Visible = True
'        lblExplain.Item(3).Visible = True
        txtExplain.Item(0).Visible = True
        txtExplain.Item(1).Visible = True
        txtExplain.Item(2).Visible = True
'        txtExplain.Item(3).Visible = True
        vsfMain.Width = Me.ScaleWidth - 100 - 2600
        picStatus.Visible = True
        With lblExplain
            .Item(0).Move vsfMain.Left + vsfMain.Width + 90, vsfMain.Top
            .Item(1).Move .Item(0).Left, .Item(0).Top + lngTxtHeight + 330
            .Item(2).Move .Item(1).Left, .Item(1).Top + lngTxtHeight + 330
'            .Item(3).Move .Item(2).Left, .Item(2).Top + lngTxtHeight + 250
            txtExplain.Item(0).Move .Item(0).Left, .Item(0).Top + 290, 2500, lngTxtHeight
            txtExplain.Item(1).Move .Item(1).Left, .Item(1).Top + 290, 2500, lngTxtHeight
            txtExplain.Item(2).Move .Item(2).Left, .Item(2).Top + 290, 2500, lngTxtHeight
'            txtExplain.Item(3).Move .Item(3).Left, .Item(3).Top + 210, 2500, lngTxtHeight
    '        cmdRef.Move .Item(0).Left + 750, picBtn.Top
        End With
    End If
'    picBtn.Top = Me.ScaleHeight - picBtn.Top - 60
    Call Picpgb.Move((Me.Width - Picpgb.Width) / 2, (Me.Top - Picpgb.Height) / 2 + 2000)
    picBtn.Top = vsfMain.Top + vsfMain.Height + 150
    picStatus.Left = vsfMain.Left + vsfMain.Width - picStatus.Width
    cmdRef.Left = picStatus.Left - cmdRef.Width - 300
    PicFind.Left = cmdRef.Left - PicFind.Width - 200
    lblFind.Left = PicFind.Left - lblFind.Width - 100
End Sub

Private Sub LoadClientsData()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim strArr() As String
    Dim blnUpgrade As Boolean
    Dim blnBefUpgrade As Boolean
    Dim blnCollect As Boolean
    Dim intBatch As Integer
    Dim i As Long
    
    With vsfMain
        .Rows = .FixedRows
        strSQL = "Select Max(����) As �������� From zlRegInfo Where ��Ŀ = '���������ļ�����'"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        intBatch = Val(rsTemp!�������� & "")
        
'        strSQL = "select ����վ,IP,����,����������,������־,�Ƿ�Ԥ����,�ռ���־,Ԥ��ʱ��,�������,Ԥ�����,�ռ�״̬,�޸�״̬,����˵��,Ԥ����˵��,�ռ�˵��,�޸�˵�� from zlclients"
        strSQL = "select A.����վ,A.IP,A.����,A.��;,A.�����ļ�������,B.λ��, A.������־,A.�Ƿ�Ԥ����,A.�ռ���־,A.Ԥ��ʱ��,A.�������,A.Ԥ�����,A.�ռ�״̬,A.�޸�״̬,A.����˵��,A.Ԥ����˵��,A.�ռ�˵��,A.�޸�˵��,A.����,A.����Ա�û�,A.����Ա���� from zlclients A,zlupgradeserver B where A.�����ļ������� = B.���(+)"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        
        '��������
        .Rows = rsTemp.RecordCount + 1
        .Redraw = flexRDNone
        i = 1
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, Col_�ͻ���) = Nvl(rsTemp.Fields("����վ"))
            .Cell(flexcpText, i, Col_IP) = Nvl(rsTemp.Fields("IP"))
            .Cell(flexcpText, i, Col_����) = Nvl(rsTemp.Fields("����"))
            .Cell(flexcpText, i, Col_��;) = Nvl(rsTemp.Fields("��;"))
            
            strTemp = Trim(Nvl(rsTemp.Fields("�����ļ�������"), ""))
            If Trim(rsTemp.Fields("λ��")) & "" = "" And strTemp <> "" Then
'                strSQL = "update ZLClients set �����ļ������� = null where �����ļ������� = " & vsfMain.TextMatrix(vsfMain.Row, Col_���)
'                gcnOracle.Execute strSQL
                .Cell(flexcpText, i, Col_����������) = ""
            Else
                .Cell(flexcpText, i, Col_����������) = IIf(strTemp <> "" And Trim(rsTemp.Fields("λ��")) <> "", Nvl(rsTemp.Fields("�����ļ�������"), "") & ":" & Nvl(rsTemp.Fields("λ��"), ""), "")
            End If
            
'            If Val(rsTemp!���� & "") < intBatch Then
'                .Cell(flexcpText, i, Col_���¼��) = "��Ҫ����"
'            Else
'                .Cell(flexcpText, i, Col_���¼��) = "�������"
'            End If
            
'            .Cell(flexcpText, i, Col_����) = IIf(Nvl(rsTemp.Fields("������־"), "") = "1", "��", "")
'            If .TextMatrix(i, Col_����) = "" Then blnUpgrade = True
'            .Cell(flexcpText, i, Col_Ԥ����) = IIf(Nvl(rsTemp.Fields("�Ƿ�Ԥ����"), "") = "1", "��", "")
'            If .TextMatrix(i, Col_Ԥ����) = "" Then blnBefUpgrade = True
'            .Cell(flexcpText, i, Col_�ռ�) = IIf(Nvl(rsTemp.Fields("�ռ���־"), "") = "1", "��", "")
'            If .TextMatrix(i, Col_�ռ�) = "" Then blnCollect = True

            .Cell(flexcpText, i, Col_����) = IIf(Nvl(rsTemp.Fields("������־"), "") = "1", True, False)
            If .Cell(flexcpText, i, Col_����) = False Then blnUpgrade = True
            .Cell(flexcpText, i, Col_Ԥ����) = IIf(Nvl(rsTemp.Fields("�Ƿ�Ԥ����"), "") = "1", True, False)
            If .Cell(flexcpText, i, Col_Ԥ����) = False Then blnBefUpgrade = True
            
            .Cell(flexcpText, i, Col_Ԥ����ʱ��) = Format(Nvl(rsTemp.Fields("Ԥ��ʱ��")), "hh:mm")
            
            strTemp = Nvl(rsTemp.Fields("�������"), "0")
            .Cell(flexcpData, i, Col_�������) = strTemp
            .Cell(flexcpText, i, Col_�������) = Decode(strTemp, "0", "δ����", "1", "���", "2", "ʧ��", "3", "��������", "")

            strTemp = Nvl(rsTemp.Fields("Ԥ�����"), "0")
            .Cell(flexcpData, i, Col_Ԥ�������) = strTemp
            .Cell(flexcpText, i, Col_Ԥ�������) = Decode(strTemp, "0", "δ����", "1", "���", "2", "ʧ��", "3", "��������", "")

'            strTemp = Nvl(rsTemp.Fields("�ռ�״̬"), "0")
'            .Cell(flexcpData, i, Col_�ռ����) = strTemp
'            .Cell(flexcpText, i, Col_�ռ����) = Decode(strTemp, "0", "δ�ռ�", "1", "���", "2", "ʧ��", "3", "�����ռ�", "")
            
            strTemp = Nvl(rsTemp.Fields("�޸�״̬"), "0")
            .Cell(flexcpData, i, Col_�����޸����) = strTemp
            .Cell(flexcpText, i, Col_�����޸����) = Decode(strTemp, "0", "δ�޸�", "1", "���", "2", "ʧ��", "3", "�����޸�", "")
            .Cell(flexcpText, i, Col_����˵��) = Nvl(rsTemp.Fields("����˵��"))
            .Cell(flexcpText, i, Col_Ԥ����˵��) = Nvl(rsTemp.Fields("Ԥ����˵��"))
'            .Cell(flexcpText, i, Col_�ռ�˵��) = Nvl(rsTemp.Fields("�ռ�˵��"))
            .Cell(flexcpText, i, Col_�����޸�˵��) = Nvl(rsTemp.Fields("�޸�˵��"))
            .Cell(flexcpText, i, Col_����Ա) = Nvl(rsTemp.Fields("����Ա�û�"))
            .Cell(flexcpText, i, Col_����) = Decipher(Nvl(rsTemp.Fields("����Ա����")))
            rsTemp.MoveNext
            i = i + 1
        Loop
        '�ı��������
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, Col_�ͻ���, .Rows - 1, Col_�ͻ���) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_IP, .Rows - 1, Col_IP) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_����, .Rows - 1, Col_����) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_����������, .Rows - 1, Col_����������) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_���¼��, .Rows - 1, Col_���¼��) = flexAlignCenterCenter
'            .Cell(flexcpAlignment, .FixedRows, Col_����, .Rows - 1, Col_����) = flexAlignCenterCenter
'            .Cell(flexcpAlignment, .FixedRows, Col_Ԥ����, .Rows - 1, Col_Ԥ����) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_�ռ�, .Rows - 1, Col_�ռ�) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_Ԥ����ʱ��, .Rows - 1, Col_Ԥ����ʱ��) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_�������, .Rows - 1, Col_�������) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_Ԥ�������, .Rows - 1, Col_Ԥ�������) = flexAlignCenterCenter
'            .Cell(flexcpAlignment, .FixedRows, Col_�ռ����, .Rows - 1, Col_�ռ����) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_�����޸����, .Rows - 1, Col_�����޸����) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_����˵��, .Rows - 1, Col_����˵��) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_Ԥ����˵��, .Rows - 1, Col_Ԥ����˵��) = flexAlignCenterCenter
'            .Cell(flexcpAlignment, .FixedRows, Col_�ռ�˵��) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_�����޸�˵��) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_����Ա, .Rows - 1, Col_����Ա) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_����, .Rows - 1, Col_����) = flexAlignLeftCenter
            .Cell(flexcpBackColor, .FixedRows, Col_��;, .Rows - 1, Col_��;) = RGB(210, 240, 255)  ' RGB(247, 247, 247)
            .Cell(flexcpBackColor, .FixedRows, Col_����������, .Rows - 1, Col_����������) = RGB(210, 240, 255)   'RGB(247, 247, 247)
            .Cell(flexcpBackColor, .FixedRows, Col_����, .Rows - 1, Col_�ռ�) = RGB(210, 240, 255)   'RGB(247, 247, 247)
            .Cell(flexcpBackColor, .FixedRows, Col_Ԥ����ʱ��, .Rows - 1, Col_Ԥ����ʱ��) = RGB(210, 240, 255)   'RGB(247, 247, 247)
        End If
        .Redraw = flexRDDirect
    End With

    
    '���ð���״̬
    mblnAllUpdateClick = True
    If blnUpgrade Then
        chkAllUpgrade.value = 0
    Else
        chkAllUpgrade.value = 1
    End If
    mblnAllUpdateClick = False
    mblnAllBefUpgrade = True
    If blnBefUpgrade Then
        chkAllBefUpgrade.value = 0
    Else
        chkAllBefUpgrade.value = 1
    End If
    mblnAllBefUpgrade = False
    If blnCollect Then
        cmdAllCollect.Caption = "ȫ���ռ�(&R)"
    Else
        cmdAllCollect.Caption = "ȫ�����ռ�(&R)"
    End If
    '���ر�������б�����
    InitCombolist
End Sub

Public Sub SetMenu(Optional lngrows As Long = -1)
    If lngrows = -1 Then lngrows = vsfMain.Rows - 1
'    frmMDIMain.stbThis.Panels(2).Text = "�б��й���ʾ��" & lngrows & "�����ݡ�"
    frmMDIMain.stbThis.Panels(2).Text = "�б��й���ʾ��" & mlngClinetNum & "���ͻ��ˣ�δ�����Ŀͻ�����" & mlngNotUpClinetNum & "��������ʧ�ܵĿͻ�����" & mlngUpFailClinetNum & "����"
End Sub

Private Function CheckIP(strIp As String) As Boolean
'���IP��ʽ�Ƿ���ȷ
    Dim sTmp() As String
    Dim i As Integer
    
    If strIp = "" Then CheckIP = False: Exit Function
    
    sTmp = Split(strIp, ".")
    If UBound(sTmp) <> 3 Then CheckIP = False: Exit Function
    
    For i = 0 To UBound(sTmp)
        If sTmp(i) = "" Then CheckIP = False: Exit Function
        
        If CLng(sTmp(i)) > 255 Or CLng(sTmp(i)) < 0 Or i > 3 Then CheckIP = False: Exit Function
    Next i
    
    CheckIP = True
End Function

Private Sub LoadSetting()
    '����ؼ����á�״̬��ȡ����
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errH
    
    lbloptStatus.Tag = "4" '����Ĭ��Ϊ��ʾȫ��
    
    txtFind.Tag = "������ͻ��ˡ�IP�����š���;"
    txtFind.Text = txtFind.Tag
    txtFind.ForeColor = vbGrayText

'    vsfMain_RowColChange
    
    '��ֹ�ͻ����޸� ���ö�ȡ��ɾ����
'    strSQL = "select ���� from ZLReginfo where ��Ŀ = '��ֹ�ͻ����޸�'"
'    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
'
'    If rsTemp.EOF Then
'        strSQL = "insert into zltools.ZLReginfo(��Ŀ,����) select '��ֹ�ͻ����޸�','0' from dual where not Exists (select 1 from zltools.ZLReginfo where ��Ŀ ='��ֹ�ͻ����޸�')"
'        gcnOracle.Execute strSQL
'        chkClientRepair.value = 0
'    Else
'        chkClientRepair.value = CInt(Nvl(rsTemp.Fields("����"), "0"))
'    End If
    
    '��ʱ�������ö�ȡ
    strSQL = "select ���� from ZLReginfo where ��Ŀ = '�ͻ�����������'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    If rsTemp.EOF = False Then
        dtpTime.value = Format(Nvl(rsTemp.Fields("����"), CurrentDate()), "yyyy-MM-dd hh:mm")
        optUpgradeTime.Item(1).value = True
    Else
        dtpTime.value = Format(CurrentDate(), "yyyy-MM-dd") & " 23:00"
        optUpgradeTime.Item(0).value = True
    End If

    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 1 = 0 Then
        Resume
    End If
End Sub

Private Sub ShowPercent(Optional blnVisible As Boolean, Optional strInfo As String, Optional sngPer As Single = -1, Optional blnPer As Boolean = False)
'��ʾ�������
    If blnVisible = False Then Picpgb.Visible = False: Exit Sub
    
    If sngPer = -1 Then
        pgbThis.value = 0
    Else
        If sngPer >= 1 Then
            pgbThis.value = CInt(sngPer)
        Else
            pgbThis.value = CInt(sngPer * 100)
        End If
    End If
    
    pgbThis.Max = 100

    lblInfo.Caption = strInfo
    lblPer.Caption = CInt(pgbThis.value) & " %"
    
    If blnPer = True Then
        lblPer.Visible = True
    Else
        lblPer.Visible = False
    End If
    Picpgb.Visible = True
    
End Sub

Private Sub InitCombolist()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim blnTemp As Boolean
    
    On Error GoTo errH:
    With vsfMain
        .Editable = flexEDKbdMouse
        '���������������б����������
        strSQL = "select ���,����,λ��,�Ƿ�����,�Ƿ�ȱʡ,�Ƿ��ռ� from zltools.zlupgradeserver"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        .ColComboList(Col_����������) = "#10*2;" & "" & vbTab & "" & vbTab & " "
        i = 2
        Do Until rsTemp.EOF
            If Nvl(rsTemp.Fields("�Ƿ�����"), "0") = "1" Or Nvl(rsTemp.Fields("�Ƿ�ȱʡ"), "0") = "1" Or Nvl(rsTemp.Fields("�Ƿ��ռ�"), "0") = "1" Then
                blnTemp = True
            Else
                blnTemp = False
            End If
            If blnTemp Then
                .ColComboList(Col_����������) = .ColComboList(Col_����������) & _
                "|#" & i * 10 & ";" & rsTemp.Fields("���") & "��" & vbTab & IIf(rsTemp.Fields("����") = 0, "����", "FTP") & vbTab & rsTemp.Fields("���") & ":" & rsTemp.Fields("λ��")
            End If
            rsTemp.MoveNext
            i = i + 1
        Loop
        '��;�����б����������
        strSQL = "select distinct ��; from zlclients order by ��;"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        Do Until rsTemp.EOF
            If Nvl(rsTemp!��;, "") <> "" Then
                .ColComboList(Col_��;) = .ColComboList(Col_��;) & "|" & rsTemp!��;
            End If
            rsTemp.MoveNext
        Loop
        .ColComboList(Col_��;) = " " & .ColComboList(Col_��;)
        
        'Ԥ����ʱ��������б�
        strSQL = "select ��Ŀ,���� from zltools.ZLReginfo where ��Ŀ = '�ͻ���Ԥ����ʱ���'"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        If Not rsTemp.EOF Then
            .ColComboList(Col_Ԥ����ʱ��) = Replace(Nvl(rsTemp!����), ",", "|")
            .ColComboList(Col_Ԥ����ʱ��) = " |" & .ColComboList(Col_Ԥ����ʱ��)
        End If
'        ��ϸ����
        .ColComboList(Col_�������) = "..."
        .ColComboList(Col_Ԥ�������) = "..."
        .ColComboList(Col_�����޸����) = "..."
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub optStatus_Click(Index As Integer)
    Dim lngRowCount As Long
    Dim i As Long
    If optStatus(Index).Visible = False Then Exit Sub
    lbloptStatus.Tag = Index
    Call FilterData(mstrLocationClientsName)
End Sub

Private Sub optUpgradeTime_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpTime.Enabled = False
            cmdTimeSet.Enabled = False
            chkAllBefUpgrade.Visible = False
            vsfMain.ColHidden(Col_Ԥ����) = True
            vsfMain.ColHidden(Col_Ԥ����ʱ��) = True
            vsfMain.ColHidden(Col_Ԥ�������) = True
            txtExplain(1).Enabled = False
            lblExplain(1).Enabled = False
            If optUpgradeTime(Index).Visible Then
                If chkAllBefUpgrade.value = 0 Then
                    Call chkAllBefUpgrade_Click
                Else
                    chkAllBefUpgrade.value = 0
                End If
                SaveUpgradeDate True
            End If
        Case 1
            dtpTime.Enabled = True
            cmdTimeSet.Enabled = True
            chkAllBefUpgrade.Visible = True
            vsfMain.ColHidden(Col_Ԥ����) = False
            vsfMain.ColHidden(Col_Ԥ����ʱ��) = False
            vsfMain.ColHidden(Col_Ԥ�������) = False
            txtExplain(1).Enabled = True
            lblExplain(1).Enabled = True
            If optUpgradeTime(Index).Visible Then
                mblnAllBefUpgrade = True
                chkAllBefUpgrade.value = 0
                mblnAllBefUpgrade = False
                SaveUpgradeDate
            End If
    End Select
End Sub

Private Sub txtExplain_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = txtFind.Tag Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    FilterData mstrLocationClientsName
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = txtFind.Tag
        txtFind.ForeColor = vbGrayText
    End If
End Sub

Private Sub vsfMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strTemp As String
    Dim strSave() As String
    Dim strSQL As String
    
    With vsfMain
        Select Case Col
        Case Col_��;
            strTemp = Trim(.Cell(flexcpTextDisplay, .Row, Col_��;))
            If strTemp <> "" Then
                strSQL = "update zltools.zlclients set ��; = '" & strTemp & "' where ����վ = '" & .TextMatrix(.Row, Col_�ͻ���) & "'"
                gcnOracle.Execute strSQL
            Else
                strSQL = "update zltools.zlclients set ��; = null where ����վ = '" & .TextMatrix(.Row, Col_�ͻ���) & "'"
                gcnOracle.Execute strSQL
            End If
        Case Col_����������
            strTemp = Trim(.Cell(flexcpTextDisplay, .Row, Col_����������))
            If strTemp <> "" Then
                strSave = Split(strTemp, ":")
                If IsNumeric(strSave(0)) = False Then Exit Sub
                strSQL = "update zltools.zlclients set �����ļ������� = " & Trim(strSave(0)) & " where ����վ = '" & .TextMatrix(.Row, Col_�ͻ���) & "'"
                gcnOracle.Execute strSQL
            Else
                strSQL = "update zltools.zlclients set �����ļ������� = null where ����վ = '" & .TextMatrix(.Row, Col_�ͻ���) & "'"
                gcnOracle.Execute strSQL
            End If
        Case Col_Ԥ����ʱ��
            strTemp = Trim(.Cell(flexcpTextDisplay, .Row, Col_Ԥ����ʱ��))
            If strTemp <> "" Then
                strTemp = Format(Now(), "yyyy/MM/dd") & " " & Format(strTemp, "hh:mm:00")
                strTemp = "to_date('" & strTemp & "','YYYY/MM/DD HH24:MI:SS')"
            Else
                strTemp = "NULL"
            End If
            strSQL = "update zltools.zlclients set Ԥ��ʱ�� = " & strTemp & " where ����վ = '" & .TextMatrix(.Row, Col_�ͻ���) & "'"
            gcnOracle.Execute strSQL
        End Select
    End With
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    txtExplain(0).Text = ""
    txtExplain(1).Text = ""
    txtExplain(2).Text = ""
'    txtExplain(3).Text = ""
    If NewRow < vsfMain.FixedRows Or NewCol < vsfMain.FixedCols Then Exit Sub
    
    With vsfMain
        mstrLocationClientsName = .TextMatrix(NewRow, Col_�ͻ���)
        txtExplain(0).Text = .TextMatrix(NewRow, Col_����˵��)
        txtExplain(1).Text = .TextMatrix(NewRow, Col_Ԥ����˵��)
'        txtExplain(2).Text = .TextMatrix(NewRow, Col_�ռ�˵��)
        txtExplain(2).Text = .TextMatrix(NewRow, Col_�����޸�˵��)
    End With
End Sub

Private Sub vsfMain_AfterSort(ByVal Col As Long, Order As Integer)
    vsfMain.Row = vsfMain.FindRow(mstrLocationClientsName, , Col_�ͻ���)
    vsfMain.ShowCell vsfMain.Row, 0
End Sub

Private Sub vsfMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Col_���������� And Col <> Col_Ԥ����ʱ�� And Col <> Col_��; And Col <> Col_������� And Col <> Col_Ԥ������� And Col <> Col_�����޸���� Then Cancel = True
End Sub

Private Sub vsfMain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = Col_���� Or Col = Col_Ԥ���� Then
        Cancel = True
    End If
End Sub

Private Sub vsfMain_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Row < vsfMain.FixedRows Then Exit Sub
    frmClientUpgradeLogView.mstrName = vsfMain.TextMatrix(Row, Col_�ͻ���)
    Load frmClientUpgradeLogView
    frmClientUpgradeLogView.Show 1, frmMDIMain
End Sub

Private Sub vsfMain_DblClick()
    Dim strSQL As String
    
    If mblnAllowEdit = False Then Exit Sub
    With vsfMain
        If .MouseRow <> .Row Then Exit Sub '��ѡ����˫����Ч�����ι̶���˫��
        Select Case .ColSel
        Case Col_����
            .TextMatrix(.RowSel, .ColSel) = IIf(.TextMatrix(.RowSel, .ColSel) = True, False, True)
            strSQL = "Zl_Zlclients_Update('" & .TextMatrix(.RowSel, Col_�ͻ���) & "'," & 0 & "," & IIf(.TextMatrix(.RowSel, .ColSel) = True, "1", "0") & ")"
            Call ExecuteProcedure(strSQL, Me.Caption)
            If .TextMatrix(.RowSel, .ColSel) = True Then
                If .TextMatrix(.Row, Col_�������) <> "δ����" Then
                    mlngNotUpClinetNum = mlngNotUpClinetNum + 1
                    If .TextMatrix(.Row, Col_�������) = "ʧ��" Then
                        mlngUpFailClinetNum = mlngUpFailClinetNum - 1
                    End If
                    SetMenu
                End If
                .TextMatrix(.Row, Col_�������) = "δ����"
                .TextMatrix(.Row, Col_����˵��) = ""
                .TextMatrix(.Row, Col_�����޸����) = "δ�޸�"
                .TextMatrix(.Row, Col_�����޸�˵��) = ""
                txtExplain(0).Text = ""
                txtExplain(2).Text = ""
            End If
            '������Ҫ������־
            Call SaveAuditLog(2, "����/ȡ������", "�Կͻ��ˡ�" & .TextMatrix(.Row, Col_�ͻ���) & "������" & IIf(.TextMatrix(.RowSel, .ColSel) = True, "", "ȡ��") & "��������")
        Case Col_Ԥ����
            .TextMatrix(.RowSel, .ColSel) = IIf(.TextMatrix(.RowSel, .ColSel) = True, False, True)
            strSQL = "Zl_Zlclients_Update('" & .TextMatrix(.RowSel, Col_�ͻ���) & "'," & 1 & "," & IIf(.TextMatrix(.RowSel, .ColSel) = True, "1", "0") & ")"
            Call ExecuteProcedure(strSQL, Me.Caption)
            If .TextMatrix(.RowSel, .ColSel) = True Then
                .TextMatrix(.Row, Col_Ԥ�������) = "δ����"
                .TextMatrix(.Row, Col_Ԥ����˵��) = ""
                txtExplain(1).Text = ""
            End If
            '������Ҫ������־
            Call SaveAuditLog(2, "Ԥ����/ȡ��Ԥ����", "�Կͻ��ˡ�" & .TextMatrix(.Row, Col_�ͻ���) & "������" & IIf(.TextMatrix(.RowSel, .ColSel) = True, "", "ȡ��") & "Ԥ��������")
        Case Col_�ռ�
            .TextMatrix(.RowSel, .ColSel) = IIf(.TextMatrix(.RowSel, .ColSel) = "��", "", "��")
            strSQL = "update zltools.ZlClients set �ռ���־ = " & IIf(.TextMatrix(.RowSel, .ColSel) = "��", "1", "0") & " where ����վ = '" & .TextMatrix(.RowSel, Col_�ͻ���) & "'"
            gcnOracle.Execute strSQL
        End Select
    End With
End Sub

Public Function CurrentDate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ������
    '-------------------------------------------------------------
    Dim rsOut As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    Set rsOut = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Current_date")
    If rsOut.RecordCount > 0 Then
        CurrentDate = IIf(IsNull(rsOut.Fields(0)), 0, rsOut.Fields(0))
    Else
        CurrentDate = 0
    End If
    Exit Function
ErrHandle:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If
End Function

Private Function UpdateData(strUpdateCol As String, strUpdateVal As String, Optional blnRef As Boolean = False) As Boolean
    Dim arrSQL() As Variant
    Dim i As Long, strName As String, blnTrans As Boolean
    Dim strUpdateField As String
    Dim strTemp As String
    Dim strSQL  As String
    Select Case strUpdateCol
        Case Col_����
            strUpdateField = 0
        Case Col_Ԥ����
            strUpdateField = 1
        Case Col_�ռ�
            strUpdateField = 2
        Case Else
            UpdateData = False: Exit Function
    End Select
    
    On Error GoTo errH:
    arrSQL() = Array()
    With vsfMain
        If .Rows < 1 Then Exit Function
        Me.Enabled = False
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            If .RowHidden(i) = False Then
                If ActualLen(strName & "," & .TextMatrix(i, Col_�ͻ���)) > 3900 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_Zlclients_Update('" & strName & "'," & strUpdateField & "," & strUpdateVal & ")"
                    
                    strName = Trim(.TextMatrix(i, Col_�ͻ���))
                Else
                    strName = IIf(strName = "", "", strName & ",") & (.TextMatrix(i, Col_�ͻ���))
                End If
                .TextMatrix(i, strUpdateCol) = IIf(strUpdateVal = "1", True, False)
            End If
        Next
        .Redraw = flexRDBuffered
        
        If strName <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_Zlclients_Update('" & strName & "'," & strUpdateField & "," & strUpdateVal & ")"
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
        
            strSQL = arrSQL(i)
            Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        Next
        gcnOracle.CommitTrans: blnTrans = False
                
        Me.Enabled = True
    End With
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    MsgBox err.Description, vbExclamation, gstrSysName
    If 1 = 0 Then
        Resume
    End If
End Function

Private Function SaveUpgradeDate(Optional blnDelete As Boolean = False) As Boolean
'    �洢��ʱ����ʱ������
'    blnDelete ɾ��ʱ�䣬true-ɾ����false-��ɾ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim strTime As String
    Dim intTemp As Integer
    
    On Error GoTo errH
    
    strTime = Trim(dtpTime.value)
    
     'ɾ���ͻ�������������Ŀ
    If blnDelete = True Then
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ = '�ͻ�����������'"
        Call OpenRecordset(rsTmp, strSQL, Me.Caption)
        
        If rsTmp.EOF = False Then
            strSQL = "Delete from zltools.zlRegInfo Where ��Ŀ='�ͻ�����������'"
            gcnOracle.Execute strSQL
        End If
        SaveUpgradeDate = True
        optUpgradeTime.Item(0).SetFocus
        Exit Function
    End If
    
    '�����洢
    Set rsTmp = New ADODB.Recordset
    strSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ = '�ͻ�����������'"
    Call OpenRecordset(rsTmp, strSQL, Me.Caption)
        
    If rsTmp.EOF = False Then
        strSQL = "Update zlRegInfo Set ����='" & strTime & "' Where ��Ŀ='�ͻ�����������'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�ͻ�����������',Null,'" & strTime & "')"
        gcnOracle.Execute strSQL
    End If
    
'    MsgBox "����ָ������ʱ�����!", vbInformation, gstrSysName

    If optUpgradeTime.Item(1).Visible = True Then optUpgradeTime.Item(1).SetFocus
    
    SaveUpgradeDate = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub FilterData(Optional strLocationClientsName As String)
    Dim strFind As String
    Dim strCondition As String
    Dim lngSelectRow As Long
    Dim i As Long
    
    On Error GoTo errH:
    strFind = IIf(txtFind.Tag <> txtFind.Text, txtFind.Text, "")
    strCondition = Decode(lbloptStatus.Tag, "4", "", "0", "δ����", "2", "ʧ��", "")
    
    mlngClinetNum = 0
    mlngNotUpClinetNum = 0
    mlngUpFailClinetNum = 0
    
    With vsfMain
        If .Rows < 1 Then Exit Sub
        lngSelectRow = .Row
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, Col_�������) = strCondition Or strCondition = "") And (InStr(Trim(.TextMatrix(i, Col_IP)), strFind) > 0 Or InStr(Trim(.TextMatrix(i, Col_�ͻ���)), strFind) > 0 Or InStr(Trim(.TextMatrix(i, Col_�ͻ���)), UCase(strFind)) > 0 Or InStr(Trim(.TextMatrix(i, Col_����)), strFind) > 0 Or InStr(Trim(.TextMatrix(i, Col_��;)), strFind) > 0) Then
                .RowHidden(i) = False
                mlngClinetNum = mlngClinetNum + 1
                Select Case .TextMatrix(i, Col_�������)
                    Case "δ����"
                        mlngNotUpClinetNum = mlngNotUpClinetNum + 1
                    Case "ʧ��"
                        mlngUpFailClinetNum = mlngUpFailClinetNum + 1
                End Select
            Else
                .RowHidden(i) = True
            End If
        Next
        
        lngSelectRow = .FindRow(strLocationClientsName, , Col_�ͻ���)
        If lngSelectRow > 0 Then
            If .RowHidden(lngSelectRow) = False Then
                .Row = lngSelectRow
            Else
                For i = 1 To .Rows - 1
                    If .RowHidden(i) = False Then .Row = i: Exit For
                Next
                If i = .Rows Then .Row = 0
            End If
        Else
            For i = 1 To .Rows - 1
                If .RowHidden(i) = False Then .Row = i: Exit For
            Next
            If i = .Rows Then .Row = 0
        End If
        vsfMain_AfterRowColChange -1, -1, .Row, Col_�ͻ���
        .ShowCell .Row, 0
        .Redraw = flexRDBuffered
    End With
    mblnFilter = True
    SetMenu
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub
Private Sub InitVsfMain()
    With vsfMain
        .Rows = .FixedRows
        .Cols = Col_����

        .Cell(flexcpText, 0, Col_�ͻ���) = "�ͻ���"
        .Cell(flexcpAlignment, 0, Col_�ͻ���) = flexAlignCenterCenter
        .ColWidth(Col_�ͻ���) = 1800
        
        .Cell(flexcpText, 0, Col_IP) = "IP"
        .Cell(flexcpAlignment, 0, Col_IP) = flexAlignCenterCenter
        .ColWidth(Col_IP) = 1500
        
        .Cell(flexcpText, 0, Col_����) = "����"
        .Cell(flexcpAlignment, 0, Col_����) = flexAlignCenterCenter
        .ColWidth(Col_����) = 1000
        
        .Cell(flexcpText, 0, Col_��;) = "��;"
        .Cell(flexcpAlignment, 0, Col_��;) = flexAlignCenterCenter
        .ColWidth(Col_��;) = 1000
        
        .Cell(flexcpText, 0, Col_����������) = "����������"
        .Cell(flexcpAlignment, 0, Col_����������) = flexAlignCenterCenter
        .ColWidth(Col_����������) = 3200
        
        .Cell(flexcpText, 0, Col_���¼��) = "���¼��"
        .Cell(flexcpAlignment, 0, Col_���¼��) = flexAlignCenterCenter
        .ColWidth(Col_���¼��) = 900
        .ColHidden(Col_���¼��) = True

        .Cell(flexcpText, 0, Col_����) = "����"
        .Cell(flexcpAlignment, 0, Col_����) = flexAlignCenterCenter
        .ColWidth(Col_����) = 700
        
        .Cell(flexcpText, 0, Col_Ԥ����) = "Ԥ����"
        .Cell(flexcpAlignment, 0, Col_Ԥ����) = flexAlignCenterCenter
        .ColWidth(Col_Ԥ����) = 900
        
        .Cell(flexcpText, 0, Col_�ռ�) = "�ռ�"
        .Cell(flexcpAlignment, 0, Col_�ռ�) = flexAlignCenterCenter
        .ColWidth(Col_�ռ�) = 700
        .ColHidden(Col_�ռ�) = True
        
        .Cell(flexcpText, 0, Col_Ԥ����ʱ��) = "Ԥ����ʱ��"
        .Cell(flexcpAlignment, 0, Col_Ԥ����ʱ��) = flexAlignCenterCenter
        .ColWidth(Col_Ԥ����ʱ��) = 1000

        .Cell(flexcpText, 0, Col_�������) = "�������"
        .Cell(flexcpAlignment, 0, Col_�������) = flexAlignCenterCenter
        .ColWidth(Col_�������) = 1800

        .Cell(flexcpText, 0, Col_Ԥ�������) = "Ԥ�������"
        .Cell(flexcpAlignment, 0, Col_Ԥ�������) = flexAlignCenterCenter
        .ColWidth(Col_Ԥ�������) = 1200

'        .Cell(flexcpText, 0, Col_�ռ����) = "�ռ����"
'        .Cell(flexcpAlignment, 0, Col_�ռ����) = flexAlignCenterCenter
'        .ColWidth(Col_�ռ����) = 1300
        
        .Cell(flexcpText, 0, Col_�����޸����) = "�����޸����"
        .Cell(flexcpAlignment, 0, Col_�����޸����) = flexAlignCenterCenter
        .ColWidth(Col_�����޸����) = 1200

        .Cell(flexcpText, 0, Col_����˵��) = "����˵��"
        .ColWidth(Col_����˵��) = 10
        .ColHidden(Col_����˵��) = True
        
        .Cell(flexcpText, 0, Col_Ԥ����˵��) = "Ԥ����˵��"
        .ColWidth(Col_Ԥ����˵��) = 10
        .ColHidden(Col_Ԥ����˵��) = True
        
'        .Cell(flexcpText, 0, Col_�ռ�˵��) = "�ռ�˵��"
'        .ColWidth(Col_�ռ�˵��) = 10
'        .ColHidden(Col_�ռ�˵��) = True
        
        .Cell(flexcpText, 0, Col_�����޸�˵��) = "�����޸�˵��"
        .ColWidth(Col_�����޸�˵��) = 10
        .ColHidden(Col_�����޸�˵��) = True
        
        .Cell(flexcpText, 0, Col_����Ա) = "����Ա"
        .ColWidth(Col_����Ա) = 1025
        .Cell(flexcpAlignment, 0, Col_����Ա) = flexAlignCenterCenter
        
        
        .Cell(flexcpText, 0, Col_����) = "����"
        .ColWidth(Col_����) = 1025
        .Cell(flexcpAlignment, 0, Col_����) = flexAlignCenterCenter
        'ѡ�п���
        .FocusRect = flexFocusSolid
        '���һ���Զ��п�
'        .ExtendLastCol = True
        '�����������
        .ScrollTrack = True
        '�Զ�����
        .WordWrap = True
        '�и�����
        .RowHeightMin = 300
        .RowHeightMax = 300
        '���������
        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
'        Call SetMenu
    End With
End Sub

Public Sub RefreshData()
    Call LoadClientsData
    Call FilterData(mstrLocationClientsName)
End Sub

Public Sub SetControlEnable(ByVal strProgFunc As String)
'����Ȩ���ַ������ÿؼ�״̬
'strProgFunc:Ȩ���ַ���
    Dim arrFunc() As String
    Dim i As Long
    
    mblnAllowEdit = False
    arrFunc = Split(strProgFunc, "|")
    For i = 0 To UBound(arrFunc)
        If arrFunc(i) = "�ͻ�����������" Then
            mblnAllowEdit = True
        End If
    Next
    '��û��Ȩ�ޣ���һЩ�ؼ���Ϊ������
    If mblnAllowEdit = False Then
        chkClientRepair.Enabled = False
        chkAllUpgrade.Enabled = False
        chkAllBefUpgrade.Enabled = False
        cmdAllCollect.Enabled = False
        txtExplain(0).Enabled = False
        txtExplain(1).Enabled = False
        txtExplain(2).Enabled = False
        txtExplain(5).Enabled = False
        cmdFileSeverSet.Enabled = False
        cmdClientAaminSet.Enabled = False
        optUpgradeTime(0).Enabled = False
        optUpgradeTime(1).Enabled = False
        cmdTimeSet.Enabled = False
        cmdClientModify.Enabled = False
        vsfMain.Editable = flexEDNone
        cmdkillProcess.Enabled = False
    End If
End Sub
