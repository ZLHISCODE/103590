VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.1#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiFeeQuery 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "���˷��ò�ѯ"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   -780
   ClientWidth     =   15045
   Icon            =   "frmPatiFeeQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptPati 
      Height          =   2265
      Left            =   120
      TabIndex        =   16
      Top             =   4980
      Width           =   3015
      _Version        =   589884
      _ExtentX        =   5318
      _ExtentY        =   3995
      _StockProps     =   0
      ShowGroupBox    =   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7365
      ScaleHeight     =   345
      ScaleWidth      =   3045
      TabIndex        =   26
      Top             =   210
      Width           =   3045
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   270
         Left            =   15
         TabIndex        =   28
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   476
         ShowSortName    =   0   'False
         IDKindStr       =   "��|����|1|1|0|0|0|0;ס|סԺ��|0|2|0|0|0|0;��|����|1|3|0|0|0|0;ҽ|ҽ����|0|4|0|0|0|0"
         CaptionAlignment=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12;CTRL++;CTRL+-;CTRL+P;CTRL+F;CTRL+F3;CTRL+F5;CTRL+A;CTRL+9;CTRL+U"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1800
         TabIndex        =   27
         ToolTipText     =   "���Ҳ���(Ctrl+F)"
         Top             =   30
         Width           =   1155
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrint 
      Height          =   780
      Left            =   5925
      TabIndex        =   22
      Top             =   420
      Visible         =   0   'False
      Width           =   2760
      _cx             =   4868
      _cy             =   1376
      Appearance      =   2
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
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      ExplorerBar     =   0
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
   Begin VB.PictureBox picCondition 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3840
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   21
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox PicRptPati 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   1335
      TabIndex        =   19
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame fraCondition 
      Height          =   4170
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   3270
      Begin zlIDKind.IDKindNew IDKindPati 
         Height          =   255
         Left            =   135
         TabIndex        =   29
         Top             =   2340
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   450
         ShowSortName    =   0   'False
         IDKindStr       =   "��|����|0|0|0|0|0|0;��|����|0|0|0|0|0|0;ס|סԺ��|1|0|0|0|0|0;ҽ|ҽ����|1|0|0|0|0|0"
         CaptionAlignment=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12;CTRL++;CTRL+-;CTRL+P;CTRL+F;CTRL+F3;CTRL+F5;CTRL+A;CTRL+9;CTRL+U"
         MustSelectItems =   "���￨"
         BackColor       =   -2147483633
      End
      Begin VB.Frame fraվ�� 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   90
         TabIndex        =   23
         Top             =   195
         Width           =   3120
         Begin VB.ComboBox cboNode 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   0
            Width           =   2085
         End
         Begin VB.Label lblվ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ժ��(&0)"
            Height          =   180
            Left            =   345
            TabIndex        =   25
            Top             =   60
            Width           =   630
         End
      End
      Begin VB.TextBox txtԤ�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1455
         TabIndex        =   12
         Top             =   2745
         Width           =   570
      End
      Begin VB.CheckBox chk����δ��˲��� 
         Caption         =   "����δ��˲���(&U)"
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   3405
         Width           =   1890
      End
      Begin VB.CheckBox chk����δ���岡�� 
         Caption         =   "����δ���岡��(&M)"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   3105
         Width           =   1890
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   1905
         TabIndex        =   15
         Top             =   3660
         Width           =   1100
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   10
         Top             =   2355
         Width           =   2085
      End
      Begin VB.ComboBox cboState 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1110
         TabIndex        =   7
         Top             =   1605
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   255328259
         CurrentDate     =   36257.9999884259
         MinDate         =   30682
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1110
         TabIndex        =   5
         Top             =   1260
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   255328259
         CurrentDate     =   36257
         MinDate         =   30682
      End
      Begin VB.CheckBox chkԤ���� 
         Caption         =   "Ԥ�����С��       Ԫ�Ĳ���"
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   2790
         Width           =   2745
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2010
         Width           =   2085
      End
      Begin VB.ComboBox cboUnit 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1110
         TabIndex        =   1
         Text            =   "cboUnit"
         Top             =   540
         Width           =   2085
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&3)"
         Height          =   180
         Left            =   90
         TabIndex        =   4
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         Caption         =   "����״̬(&2)"
         Height          =   180
         Left            =   90
         TabIndex        =   2
         Top             =   945
         Width           =   990
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "���˲���(&1)"
         Height          =   180
         Left            =   90
         TabIndex        =   0
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&4)"
         Height          =   180
         Left            =   90
         TabIndex        =   6
         Top             =   1665
         Width           =   990
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��(&5)"
         Height          =   180
         Left            =   255
         TabIndex        =   8
         Top             =   2070
         Width           =   930
      End
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   1440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeQuery.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeQuery.frx":0464
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   9255
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiFeeQuery.frx":0D3E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17348
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Key             =   "���"
            Object.ToolTipText     =   "�����Ƿ������"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   6180
      Left            =   3360
      TabIndex        =   20
      Top             =   780
      Width           =   9330
      _Version        =   589884
      _ExtentX        =   16457
      _ExtentY        =   10901
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatiFeeQuery.frx":15D2
      Left            =   840
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Menu mnuPop 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopAudit 
         Caption         =   "���(&A)"
      End
      Begin VB.Menu mnuPopUnAudit 
         Caption         =   "ȡ�����(&U)"
      End
      Begin VB.Menu mnuPopBilling 
         Caption         =   "����(&B)"
      End
      Begin VB.Menu mnuPopAudit_Line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDisp 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPopCard 
         Caption         =   "������Ϣ��Ƭ(&K)"
      End
      Begin VB.Menu mnuPopNotify 
         Caption         =   "��ӡ���Ŵ߿(&N)"
      End
      Begin VB.Menu mnuPopCurr 
         Caption         =   "��ӡ���Ŵ߿(&C)"
      End
   End
End
Attribute VB_Name = "frmPatiFeeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrPrivs As String
Private mlngModul As Long
Private mblnUnload As Boolean
Private mblnHavePara As Boolean
Private WithEvents mclsFeeQuery As clsFeeQuery
Attribute mclsFeeQuery.VB_VarHelpID = -1
Private mfrmPatiFeeVerfy As frmPatiFeeVerfy

Private mclsAdvices As Object
Private mcolSubForm As Collection
Private mfrmActive As Form

Private mintFindType As Integer
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstrPrePati As String
Private mrsDept As ADODB.Recordset
Private Enum mPan
    Condition = 1
    Pati = 2
End Enum
Private mblnSelPatiList As Boolean '�Ƿ�ѡ���˲����б�
Private Type t_ViewState
    OnePati As Boolean
End Type
Private mlngUnitID As Long
Private mvs As t_ViewState
Private mblnԤ�� As Boolean '����Ա�Ƿ���Ԥ��ģ�鹦��(��Ԥ��,�������ڷ��ò�ѯ�н�Ԥ��)
Private mstr��ֹ���� As String

'�ֶ���,���,�Ƿ��������;(���,�Ƿ�������鲻д��ʾ����������)
Private mstrPatiHead As String
'Private Const mstrPatiHead = "����ID;��ҳID;�Ǽ�ʱ��;״̬;��������;����ת��;��ǰ����ID;����;����;��ǰ����ID;���￨��;" & _
                           "���,30,0;����,60,0;סԺ��,60,0;����,50,0;�ѱ�,60,1;�Ա�,40,1;����,60,1;��Ժʱ��,100,0;��Ժʱ��,100,0;" & _
                           "��ǰ����,80,1;����,40,1;����,40,1;ҽ����,120,0;��ϵ�绰,80,0;ҽ�Ƹ��ʽ,120,1;" & _
                           "���,45,1;��������,100,1;��ǰ����,80,1"
Private mobjPatient As Object
Private mobjPlugIn As Object
Private mblnNotClick As Boolean
Private mblnȱʡ���� As Boolean
Private mbln����վ�� As Boolean
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrCaption As String
Private mbytFontSize As Byte
Private mintInsure As Integer '����:31883

Private Sub ReMoveCtrol()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ƶ��ؼ�λ��
    '����:���˺�
    '����:2012-06-19 11:29:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    cboUnit.Left = lblUnit.Left + lblUnit.Width + 20
    
    lblվ��.Top = cboNode.Top + (cboNode.Height - lblվ��.Height) \ 2
    cboNode.Left = cboUnit.Left
    cboNode.Width = dtpBegin.Width
    lblվ��.Left = cboNode.Left - 20 - lblվ��.Width
    fraվ��.Width = fraCondition.Width - 60
    fraվ��.Left = 15
    
    If mbln����վ�� Then
        cboUnit.Top = fraվ��.Top + fraվ��.Height + IIf(mbytFontSize = 9, 15, 60)
    Else
        cboUnit.Top = fraվ��.Top
    End If
    cboUnit.Width = dtpBegin.Width
    lblUnit.Top = cboUnit.Top + (cboUnit.Height - lblUnit.Height) \ 2
    cboUnit.Left = cboNode.Left
    
    cboState.Top = cboUnit.Top + cboUnit.Height + 50
    cboState.Left = cboNode.Left
    cboState.Width = dtpBegin.Width
    lblState.Top = cboState.Top + (cboState.Height - lblState.Height) \ 2
    
    dtpBegin.Top = cboState.Top + cboState.Height + 50
    dtpBegin.Height = cboState.Height: dtpEnd.Height = dtpBegin.Height
    dtpBegin.Left = cboNode.Left
    lblStartDate.Top = dtpBegin.Top + (dtpBegin.Height - lblStartDate.Height) \ 2
    
    dtpEnd.Top = dtpBegin.Top + dtpBegin.Height + 50
    dtpEnd.Left = cboNode.Left
    lblEndDate.Top = dtpEnd.Top + (dtpEnd.Height - lblEndDate.Height) \ 2
    
    dtpEnd.Top = dtpBegin.Top + dtpBegin.Height + 50
    dtpEnd.Left = cboNode.Left
    lblEndDate.Top = dtpEnd.Top + (dtpEnd.Height - lblEndDate.Height) \ 2
        
    txtסԺ��.Top = dtpEnd.Top + dtpEnd.Height + 50
    txtסԺ��.Height = cboNode.Height
    txtסԺ��.Left = cboNode.Left
    txtסԺ��.Width = dtpBegin.Width
    lblסԺ��.Top = txtסԺ��.Top + (txtסԺ��.Height - lblסԺ��.Height) \ 2
    
    txt����.Top = txtסԺ��.Top + txtסԺ��.Height + 50
    txt����.Height = cboNode.Height
    txt����.Left = cboNode.Left
    IDKindPati.Top = txt����.Top + (txt����.Height - IDKindPati.Height) \ 2
    txt����.Width = cboState.Width
    IDKindPati.Left = lblUnit.Left + lblUnit.Width - IDKindPati.Width
    txtԤ��.Top = txt����.Top + txt����.Height + 50: txtԤ��.Height = cboNode.Height
    chkԤ����.Top = txtԤ��.Top + (txtԤ��.Height - chkԤ����.Height) \ 2
    txtԤ��.Left = chkԤ����.Left + TextWidth("<Ԥ�����С�� >")
    
    chk����δ���岡��.Top = txtԤ��.Top + txtԤ��.Height + 50
    chk����δ��˲���.Top = chk����δ���岡��.Top + chk����δ���岡��.Height + 50
    cmdSearch.Top = chk����δ��˲���.Top + chk����δ��˲���.Height + 50
    
    dkpMain.Panes(1).MinTrackSize.Height = IIf(mbytFontSize = 9, 270, 300)
    dkpMain.Panes(1).MinTrackSize.Width = IIf(mbytFontSize = 9, 225, 295)
    dkpMain.RedrawPanes
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
 
 
Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:���˺�
    '����:2012-06-18 16:50:35
    '����:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
    Call ReMoveCtrol
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������С
    '����:���˺�
    '����:2012-06-18 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Call mclsFeeQuery.SetFontSize(bytSize)
    Call mfrmPatiFeeVerfy.SetFontSize(bytSize)
    Call mfrmActive.SetFontSize(bytSize)
    If Not mclsAdvices Is Nothing Then Call mclsAdvices.SetFontSize(bytSize)
    
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") 'ҳ��ؼ�
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Height = TextHeight("��") + 20
        Case UCase("VsFlexGrid")
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("��") * 1.5
        Case UCase("textBox")
          objCtrl.FontSize = mbytFontSize
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = mbytFontSize
        End Select
    Next
    IDKindPati.FontSize = mbytFontSize
    IDKindPati.Refrash
    Call Form_Resize
End Sub

Private Sub cboNode_Click()
    If mblnNotClick Then Exit Sub
    Call LoadUnits
End Sub

Private Sub cboState_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub cboUnit_Click()
    mlngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
End Sub

Private Sub cboUnit_Validate(Cancel As Boolean)
    Dim lngID As Long
    
    If cboUnit.ListIndex >= 0 Then Exit Sub
    lngID = mlngUnitID
   zlControl.CboLocate cboUnit, lngID, True
   If cboUnit.ListIndex < 0 And cboUnit.ListCount <> 0 Then cboUnit.ListIndex = 0
End Sub

Private Sub chk����δ���岡��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub chk����δ��˲���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
   Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    Dim strKind As String
    '����:42946
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        'If mobjICCard Is Nothing Then Exit Sub
        'txtFind.Text = mobjICCard.Read_Card()
        txtFind.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtFind.IMEMode = 0
        
        If txtFind.Text = "" Then Exit Sub
        ExecFindPati objCard, True
        Exit Sub
    End If
   txtFind.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtFind.IMEMode = 0
    If lng�����ID <= 0 Then
         txtFind.PasswordChar = ""
         '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
         txtFind.IMEMode = 0
         Exit Sub
    End If
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtFind.Text = strOutCardNO
    If txtFind.Text = "" Then
        txtFind.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtFind.IMEMode = 0
        Exit Sub
    End If
    
    ExecFindPati objCard, True
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtFind.IMEMode = 0
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtFind.Locked Then Exit Sub
    txtFind.Text = objPatiInfor.����
    If txtFind.Text = "" Then
        txtFind.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtFind.IMEMode = 0
        Exit Sub
    End If
    ExecFindPati objCard, True
End Sub

Private Sub IDKindPati_Click(objCard As zlIDKind.Card)
   Dim strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    Dim strKind As String, intFindType As Integer
    
    '����:42946
    strKind = objCard.����
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        txt����.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txt����.IMEMode = 0
        If txt����.Text = "" Then Exit Sub
        ExecFindPati objCard, , True
        Exit Sub
    End If
    txt����.PasswordChar = IIf(IDKindPati.ShowPassText, "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txt����.IMEMode = 0
    
    If objCard.�ӿ���� <= 0 Then
        txt����.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txt����.IMEMode = 0
        Exit Sub
    End If
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, objCard.�ӿ����, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txt����.Text = strOutCardNO
    If txt����.Text = "" Then
        txt����.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txt����.IMEMode = 0
        Exit Sub
    End If
    
    Call LoadPatients(objCard, True)
    Call tbcSub_SelectedChanged(tbcSub.Selected)
End Sub

Private Sub IDKindPati_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txt����.IMEMode = 0
End Sub
Private Sub IDKindPati_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txt����.Locked Then Exit Sub
    txt����.Text = objPatiInfor.����
    If txt����.Text = "" Then
        txt����.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txt����.IMEMode = 0
        Exit Sub
    End If
    
    Call LoadPatients(objCard, True)
    Call tbcSub_SelectedChanged(tbcSub.Selected)
End Sub

Private Sub mclsFeeQuery_RequestRefresh()
'���ܣ��Ӵ���Ҫ��ˢ��
    Call LoadPatients(IDKindPati.GetCurCard)
End Sub

Private Sub mclsFeeQuery_StatusTextUpdate(ByVal Text As String)
'���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    Me.sta.Panels(2).Text = Text
End Sub

Private Sub cboState_Click()
    Dim objControl As CommandBarButton
    
    dtpBegin.Enabled = cboState.Text = "��Ժ����" Or cboState.Text = "���в���"
    dtpEnd.Enabled = dtpBegin.Enabled
        
    rptPati.Columns(GetRptColumn(rptPati, "��Ժʱ��")).Visible = cboState.Text <> "��Ժ����"
    
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    If KeyAscii <> 13 Then Exit Sub
    
    If cboUnit.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Call InitUnits
    Dim strRootCaption As String
    strRootCaption = ""
    If InStr(mstrPrivs, ";���в���;") > 0 Then strRootCaption = "���в���"
    If cboNode.ListCount > 0 Then
        mrsDept.Filter = "վ��=" & cboNode.ItemData(cboNode.ListIndex)
    End If
    
    If zlSelectDept(Me, mlngModul, cboUnit, mrsDept, cboUnit.Text, True, strRootCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
     
End Sub


Private Sub ExecPrintMultiBill()
    Dim i As Long
    Dim rptr As ReportRecord
    '--27894
    If cboUnit.ItemData(cboUnit.ListIndex) = 0 Then
        MsgBox "������һ�δ�ӡ���в������˵�֪ͨ����" & vbCr & "��ѡ�����Ĳ��������ԣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    '����:��������:34770
    If frmPatiPressMoney.zlPatiPressMoney(Me, mlngModul, mstrPrivs, cboUnit.ItemData(cboUnit.ListIndex), cboUnit.Text) = False Then Exit Sub
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objRow As ReportRow, objControl As CommandBarControl
    Dim i As Long, blnSelect As Boolean
     
    Select Case Control.ID
        Case conMenu_File_PrintMultiBill
            ExecPrintMultiBill
        Case conMenu_File_Preview_Pati  '��ӡԤ�������б�
            Call zlRptPrint(2)
        Case conMenu_File_Print_Pati   '��ӡ�����б�
            Call zlRptPrint(1)
        Case conMenu_File_Excel_Pati   '�����б������Excel
            Call zlRptPrint(3)
        Case conMenu_Edit_PreBalanceAll
            Call ExecPreBalanceAll
        Case conMenu_Edit_Balance   '����
            Call ExecBalance
        Case conMenu_Edit_PrePayMoney '��Ԥ��
            Call ExecPrePayMoney
        Case conMenu_Manage_Change_InsureSel
            Call ModeInsurePatiDisease  '31883
        Case conMenu_Edit_FeeAudit  '���
            Call ExecAuditingAndCancelAudit(1)
        Case conMenu_Edit_OverFeeAudit '������
            Call ExecAuditingAndCancelAudit(2)
        Case conMenu_Edit_FeeUnAudit   'ȡ�����
            Call ExecAuditingAndCancelAudit(0)
            
        Case conMenu_View_OnePati   '���סԺֻ��ʾһ������
            Control.Checked = Not Control.Checked: mvs.OnePati = Control.Checked
            Call LoadPatients(IDKindPati.GetCurCard)
        Case conMenu_View_GroupCol * 10 + 1 To conMenu_View_GroupCol * 10 + UBound(Split(mstrPatiHead, ";"))
            Control.Checked = Not Control.Checked
            If Control.Checked Then
                i = GetRptColumn(rptPati, Control.Caption)
                rptPati.GroupsOrder.Add rptPati.Columns(i)
                rptPati.Columns(i).Visible = False
                rptPati.Populate
            Else
                Call rptRemoveGroupsItem(rptPati, Control.Caption)
            End If
'
'        Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + mcllBrushCard.Count  '���ҷ�ʽ
'            mintFindType = Val(Control.Parameter)
'            cbsMain.RecalcLayout
'            txtFind.Text = ""
'            txtFind.SetFocus
'            Call InitCardType
'        Case conMenu_View_Filter * 100# + 1 To conMenu_View_Filter * 100# + mcllBrushCard.Count
'            '�����˵�����ʾ
'            '���˺�:24913
'            IDKindPati.Tag = Val(Control.Parameter)
'            IDKindPati.Caption = Replace(Split(Control.Caption, "(")(0) & "��(&6)", " ", "")
'            If txt����.Enabled And txt����.Visible Then txt����.SetFocus
'            zlDatabase.setPara "���˹������", Val(IDKindPati.Tag), glngSys, mlngModul, mblnHavePara
'            Call InitSearchType
        Case conMenu_View_Find '����
            If Me.ActiveControl Is txtFind Then
                txtFind.SetFocus '��ʱ��Ҫ��λһ��
                If txtFind.Text <> "" Then
                    Call ExecFindPati(IDKIND.GetCurCard)
                End If
            Else
                txtFind.SetFocus
            End If
        Case conMenu_View_FindNext '������һ��
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call ExecFindPati(IDKIND.GetCurCard, True)
            End If
        Case conMenu_View_ToolBar_Button '������
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                cbsMain(i).Visible = Not cbsMain(i).Visible
            Next
            cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '��ť����
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                For Each objControl In cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '��ͼ��
            cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
            Control.Checked = Not Control.Checked
            cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '״̬��
            sta.Visible = Not sta.Visible
            Control.Checked = Not Control.Checked
            cbsMain.RecalcLayout
        Case conMenu_View_FontSize_S    'С����
            Call SetFontSize(0)
        Case conMenu_View_FontSize_L    '������
            Call SetFontSize(1)
        Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
            If rptPati.SelectedRows.Count > 0 Then
                If rptPati.SelectedRows(0).GroupRow Then
                    rptPati.SelectedRows(0).Expanded = False
                ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                    If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                        rptPati.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '���۵���λ��������,�����Զ�������¼�
            Call rptPati_SelectionChanged
        Case conMenu_View_Expend_CurExpend 'չ����ǰ��
            If rptPati.SelectedRows.Count > 0 Then
                rptPati.SelectedRows(0).Expanded = True
            End If
        Case conMenu_View_Expend_AllCollapse '�۵�������
            For Each objRow In rptPati.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '���۵���λ��������,�����Զ�������¼�
            Call rptPati_SelectionChanged
        Case conMenu_View_Expend_AllExpend 'չ��������
            For Each objRow In rptPati.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call zlHomePage(hWnd)
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(hWnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call zlMailTo(hWnd)
        Case conMenu_Help_About '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, hWnd, Name, Int((glngSys) / 100))
        Case conMenu_File_SchemeSet '������������
             Call zlSchemeSet
        Case conMenu_File_Exit '�˳�
            Unload Me
        Case Else
            Select Case Me.tbcSub.Selected.Tag
            Case "����", "ҽ��"
                 Call mclsFeeQuery.zlExecuteCommandBars(Control)
            End Select
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If sta.Visible Then Bottom = sta.Height
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean, strPrivsTemp As String
    
    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" Then Exit Sub
    blnVisible = True
    
    Select Case Control.ID
        '76511:������,2014-08-12,Ԥ�����в��˵�Ȩ���ж�ȱʧ
        Case conMenu_Edit_PreBalanceAll
            blnVisible = InStr(";" & mstrPrivs, ";Ԥ�����в���;") > 0
            Control.Category = "���ж�"
        Case conMenu_Edit_Balance
            strPrivsTemp = GetInsidePrivs(Enum_Inside_Program.p���˽���)
            blnVisible = InStr(strPrivsTemp, "סԺ���ý���") > 0
            Control.Category = "���ж�"
        Case conMenu_Edit_PrePayMoney 'Ԥ��
            strPrivsTemp = ";" & GetInsidePrivs(Enum_Inside_Program.pԤ����) & ";"
            blnVisible = InStr(strPrivsTemp, ";Ԥ���տ�;") > 0
            Control.Category = "���ж�"
        Case conMenu_Manage_Change_InsureSel
            blnVisible = InStr(";" & mstrPrivs, ";����ѡ��;") > 0
            Control.Category = "���ж�"
        Case conMenu_Edit_FeeAudit
            blnVisible = InStr(";" & mstrPrivs, ";��˲���;") > 0
            Control.Category = "���ж�"
        Case conMenu_Edit_OverFeeAudit  '������
            blnVisible = InStr(";" & mstrPrivs, ";��˲���;") > 0
            Control.Category = "���ж�"
        Case conMenu_Edit_FeeUnAudit
            blnVisible = InStr(";" & mstrPrivs, ";ȡ����˲���;") > 0
            Control.Category = "���ж�"
    End Select
    Control.Visible = blnVisible
    Control.Enabled = blnVisible    '51135 :��������Enabled����,��Ȼ���������
End Sub


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.tbcSub
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub


Private Sub ExecPreBalanceAll()
    Dim rptr As ReportRecord, rsTmp As ADODB.Recordset
    Dim arrInfo() As Variant, str������� As String, i As Integer
    Dim lng����ID As Long, int���� As Integer, blnDateMoved As Boolean
    Dim strҽ���� As String, str���� As String, str���� As String, dat�Ǽ�ʱ�� As Date
            
    arrInfo = Array()
    For Each rptr In rptPati.Records
        If Trim(rptr(GetRptRsColumn("ҽ����")).Value) <> "" And Trim(rptr(GetRptRsColumn("��Ժʱ��")).Value) = "" Then
            ReDim Preserve arrInfo(UBound(arrInfo) + 1)
             '����,����ID,����,ҽ����,����
            arrInfo(UBound(arrInfo)) = rptr(GetRptRsColumn("����")).Value & "|" & Val(rptr(GetRptRsColumn("����ID")).Value) & "|" & _
                rptr(GetRptRsColumn("����")).Value & "|" & rptr(GetRptRsColumn("ҽ����")).Value & "|" & rptr(GetRptRsColumn("����")).Value & "|" & rptr(GetRptRsColumn("�Ǽ�ʱ��")).Value
        End If
    Next
    
    If UBound(arrInfo) = -1 Then
        MsgBox "��ǰ�����嵥��û�з��ֵ�ǰ�������Ժҽ�����ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("�ò������Ե�ǰ�����嵥�е�������Ժҽ�����˽���Ԥ����," & _
        vbCrLf & "����ܻỨ�ѽϳ���ʱ��,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    For i = 0 To UBound(arrInfo)
        str���� = Split(arrInfo(i), "|")(0)
        lng����ID = Val(Split(arrInfo(i), "|")(1))
        int���� = Val(Split(arrInfo(i), "|")(2))
        strҽ���� = Split(arrInfo(i), "|")(3)
        str���� = Split(arrInfo(i), "|")(4)
        dat�Ǽ�ʱ�� = CDate(Split(arrInfo(i), "|")(5))
        
        If Not gclsInsure.GetCapability(support����_�������ú���ýӿ�, lng����ID, int����) Then
            blnDateMoved = zlDatabase.DateMoved(dat�Ǽ�ʱ��, , , Caption)
            
            Call zlCommFun.ShowFlash("���ڴ���ҽ������""" & str���� & """ ...", Me)
            Refresh
            
            Set rsTmp = GetVBalance(1, "סԺ���ý���", int����, lng����ID, , , , , blnDateMoved)
            If Not rsTmp Is Nothing Then
                If Not rsTmp.RecordCount = 0 Then
                    str������� = gclsInsure.WipeoffMoney(rsTmp, lng����ID, strҽ����, "0", int����, "|0") '������;����
                End If
            End If
        End If
    Next
    sta.Panels(3).Text = "Ԥ����ɹ�!"
    Call zlCommFun.StopFlash
    
    mstrPrePati = ""
    Call rptPati_SelectionChanged
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnSelect As Boolean, lngColTmp As Long, blnEnabled As Boolean, blnQueryFee As Boolean
    Dim strTemp As String
    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    If rptPati.SelectedRows.Count > 0 Then blnSelect = Not rptPati.SelectedRows(0).GroupRow
    
    Select Case Control.ID
        '�ļ�
        Case conMenu_File_PrintMultiBill, conMenu_File_PrintDayDetail
            Control.Enabled = rptPati.Records.Count > 0
    
        '�༭
        Case conMenu_Edit_PreBalanceAll
            Control.Enabled = rptPati.Records.Count > 0 And cboState.Text <> "��Ժ����"
        Case conMenu_Edit_Balance
            Control.Enabled = blnSelect
        Case conMenu_Edit_PrePayMoney 'Ԥ����
            Control.Enabled = blnSelect
        Case conMenu_Manage_Change_InsureSel
            '31883
             Control.Enabled = blnSelect And mintInsure <> 0
        Case conMenu_Edit_FeeAudit
            '�������
            Control.Enabled = blnSelect
            If blnSelect Then
                strTemp = Trim(rptPati.SelectedRows(0).Record(GetRptRsColumn("���״̬")).Value)
                Control.Enabled = (strTemp = "")
            End If
        Case conMenu_Edit_OverFeeAudit '������
            Control.Enabled = blnSelect
            If blnSelect Then
                strTemp = Trim(rptPati.SelectedRows(0).Record(GetRptRsColumn("���״̬")).Value)
                Control.Enabled = strTemp = "��ʼ"
            End If
        Case conMenu_Edit_FeeUnAudit
            'ȡ�����
            Control.Enabled = blnSelect
            If blnSelect Then
                strTemp = Trim(rptPati.SelectedRows(0).Record(GetRptRsColumn("���״̬")).Value)
                Control.Enabled = strTemp <> ""
            End If
       '�鿴
        Case conMenu_View_OnePati
            Control.Enabled = cboState.Text = "��Ժ����" Or cboState.Text = "���в���"
            Control.Checked = mvs.OnePati
        
        Case conMenu_View_GroupCol * 10 + 1 To conMenu_View_GroupCol * 10 + UBound(Split(mstrPatiHead, ";")) '������
            Control.Enabled = rptPati.Rows.Count > 1
            
            Control.Checked = False
            For lngColTmp = 0 To rptPati.GroupsOrder.Count - 1
                If Control.Caption = rptPati.GroupsOrder(lngColTmp).Caption Then
                    Control.Checked = True
                    Exit Sub
                End If
            Next
        
        Case conMenu_View_ToolBar_Button '������
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar '״̬��
            Control.Checked = Me.sta.Visible
        Case conMenu_View_Expend_CurExpend 'չ����ǰ��
            blnEnabled = False
            If rptPati.SelectedRows.Count > 0 Then
                If rptPati.SelectedRows(0).GroupRow Then
                    blnEnabled = Not rptPati.SelectedRows(0).Expanded
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
            blnEnabled = False
            If rptPati.SelectedRows.Count > 0 Then
                If rptPati.SelectedRows(0).GroupRow Then
                    blnEnabled = rptPati.SelectedRows(0).Expanded
                ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                    If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                        blnEnabled = rptPati.SelectedRows(0).ParentRow.Expanded
                    End If
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend '�۵�/չ����
            Control.Enabled = rptPati.GroupsOrder.Count > 0 And rptPati.Rows.Count > 0
        Case conMenu_View_FindType '���ҷ�ʽ
        Case conMenu_View_FindNext
            Control.Enabled = rptPati.Records.Count > 1
        Case conMenu_File_SchemeSet  '������������:35386
             Control.Visible = InStr(1, mstrPrivs, ";������������;") > 0
        Case conMenu_View_FontSize_S         'С����
             Control.Checked = mbytFontSize = 9
        Case conMenu_View_FontSize_L    '������
             Control.Checked = mbytFontSize <> 9
        Case Else
            Select Case tbcSub.Selected.Tag
            Case "����", "ҽ��"
                Call mclsFeeQuery.zlUpdateCommandBars(Control)
            End Select
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim i As Long, strKey As String, objControl As CommandBarControl
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
    Case Else
       Select Case tbcSub.Selected.Tag
       Case "����"
           Call mclsFeeQuery.zlPopupCommandBars(CommandBar)
       End Select
    End Select
End Sub

Private Sub cmdSearch_Click()
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "��ʼʱ�䲻�ܴ��ڽ���ʱ��!", vbInformation, gstrSysName
        If dtpBegin.Visible And dtpBegin.Enabled Then dtpBegin.SetFocus
        Exit Sub
    End If
    mlng����ID = 0
    Call LoadPatients(IDKindPati.GetCurCard)
    Call tbcSub_SelectedChanged(tbcSub.Selected)
 End Sub


Private Sub ExecBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�н��ʲ���
    '����:���˺�
    '����:2015-02-05 12:00:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    Dim bln�������۲��� As Boolean
    
    bln�������۲��� = ZlIsOutpatientObserve(mlng����ID, mlng��ҳID)
    strPrivs = ";" & GetInsidePrivs(Enum_Inside_Program.p���˽���) & ";"
    If Val(zlDatabase.GetPara("���ʽ�����", glngSys, 1137, "1")) = 0 Then
        If frmPatiBalanceTraditional.ShowMe(Me, _
            IIf(bln�������۲���, g_Ed_�������, g_Ed_סԺ����), strPrivs, mlng����ID, CStr(mlng��ҳID)) = False Then Exit Sub
    Else
        If frmPatiBalanceSplit.ShowMe(Me, _
            IIf(bln�������۲���, g_Ed_�������, g_Ed_סԺ����), strPrivs, mlng����ID, CStr(mlng��ҳID)) = False Then Exit Sub
    End If
    If MsgBox("��ǰ���ݿ����ѱ仯,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        mstrPrePati = ""
        Call RefreshData
    End If
End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPan.Pati
        Item.Handle = PicRptPati.hWnd
    Case mPan.Condition
        Item.Handle = picCondition.hWnd
    End Select
End Sub


'111515:���ϴ���2017/8/14�������С�����ͺ��Դ���
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    Call cbsMain_Resize
    txtסԺ��.Width = dtpBegin.Width
    txt����.Width = dtpBegin.Width
    cboNode.Width = dtpBegin.Width
    cboUnit.Width = dtpBegin.Width
    cboState.Width = dtpBegin.Width
End Sub

Private Sub picCondition_Resize()
    Err = 0: On Error Resume Next
    With fraCondition
        .Top = 0
        .Left = 0
        .Width = picCondition.ScaleWidth
        .Height = picCondition.ScaleHeight
        fraվ��.Width = .Width - 60
    End With
End Sub

 

Private Sub picFind_Resize()
    Err = 0: On Error Resume Next
    With picFind
        IDKIND.Left = .ScaleLeft
        txtFind.Left = IDKIND.Left + IDKIND.Width
        txtFind.Width = .ScaleWidth - txtFind.Left
        'txtFind.Width = .ScaleWidth - txtFind.Left - IIf(cmdReadCard.Visible, cmdReadCard.Width + 50, 0)
    End With
End Sub

Private Sub PicRptPati_Resize()
    Err = 0: On Error Resume Next
    With rptPati
        .Top = PicRptPati.ScaleTop
        .Left = PicRptPati.ScaleLeft
        .Width = PicRptPati.ScaleWidth
        .Height = PicRptPati.ScaleHeight
    End With
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBarPopup
    Dim objCommandBar As CommandBar
    Dim objControl As CommandBarControl
    Dim i As Long, j As Long, arrTmp As Variant
    If Button = 2 Then
        Set objHitTest = rptPati.HitTest(X, Y)
        If objHitTest.ht = xtpHitTestHeader Then
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(xtpControlButtonPopup, conMenu_View_GroupCol, True, True)
            Set objCommandBar = objPopup.CommandBar
        ElseIf objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = cbsMain.ActiveMenuBar.FindControl(xtpControlButtonPopup, conMenu_View_Expend, True, True)
                Set objCommandBar = objPopup.CommandBar
            Else
                Set objCommandBar = cbsMain.Add("PopupPati", xtpBarPopup)
                With objCommandBar.Controls
                    Set objControl = .Add(xtpControlButton, conMenu_File_Preview_Pati, "��ӡԤ�������б�(&T)")
                    objControl.BeginGroup = True:        objControl.IconId = conMenu_File_Preview
                    Set objControl = .Add(xtpControlButton, conMenu_File_Print_Pati, "��ӡ�����б�(&O)")
                    objControl.IconId = conMenu_File_Print
                    Set objControl = .Add(xtpControlButton, conMenu_File_Excel_Pati, "�����б������Excel(&E)")
                    objControl.IconId = conMenu_File_Excel
                    
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InsureSel, "����ѡ��(&Z)")
                    objControl.BeginGroup = True
                    
                    .Add(xtpControlButton, conMenu_Edit_PreBalance, "Ԥ�ᵱǰ����(&W)").BeginGroup = True
                    
                    .Add xtpControlButton, conMenu_Edit_Balance, "����(&B)"
                    .Add xtpControlButton, conMenu_Edit_Billing, "����(&C)"
                    
                    .Add(xtpControlButton, conMenu_Edit_FeeAudit, IIf(gTy_System_Para.byt������˷�ʽ = 1, "��ʼ���(&A)", "���(&A)")).BeginGroup = True
                    .Add(xtpControlButton, conMenu_Edit_OverFeeAudit, "������(&O)").IconId = 252
                    .Add xtpControlButton, conMenu_Edit_FeeUnAudit, "ȡ�����(&U)"
                    .Add(xtpControlButton, conMenu_Edit_PrePayMoney, "��Ԥ��(&P)").IconId = 3816
                
                    '���벡����Ϣ,һ���嵥,�߿
                    .Add(xtpControlButton, conMenu_View_PatInfor, "������ϸ��Ϣ(&K)").BeginGroup = True
                    .Add xtpControlButton, conMenu_File_PrintSingleBill, "��ӡ���Ŵ߿(&C)��"
                    .Add xtpControlButton, conMenu_File_PrintDayDetail, "��ӡһ���嵥(&D)��"
                End With
            End If
        End If
        If Not objCommandBar Is Nothing Then objCommandBar.ShowPopup
    End If
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Not rptPati.SelectedRows(0).GroupRow Then Call ShowPatiCard
End Sub


Private Sub ShowPatiCard()
    If mlng����ID <> 0 Then
        frmDegreeCard.mlng����ID = mlng����ID
        frmDegreeCard.mlng��ҳID = mlng��ҳID
        frmDegreeCard.Show 1, Me
    End If
End Sub

Private Sub rptPati_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    mblnSelPatiList = True
End Sub

Private Sub rptPati_SelectionChanged()
    Dim strTmp As String, lngסԺ���� As Long
    
    If rptPati.SelectedRows.Count = 0 Then Exit Sub  '���������
    
    With rptPati.SelectedRows(0)
        If Not .GroupRow Then
            mlng����ID = Val(.Record(GetRptRsColumn("����ID")).Value)
            mlng��ҳID = Val(.Record(GetRptRsColumn("��ҳID")).Value)
            mintInsure = Val(.Record(GetRptRsColumn("����")).Value)
            If .Record(GetRptRsColumn("��Ժʱ��")).Value = "" Then
                lngסԺ���� = DateDiff("d", CDate(Format(.Record(GetRptRsColumn("��Ժʱ��")).Value, "yyyy-mm-dd")), CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd")))
            Else
                lngסԺ���� = DateDiff("d", CDate(Format(.Record(GetRptRsColumn("��Ժʱ��")).Value, "yyyy-mm-dd")), CDate(Format(.Record(GetRptRsColumn("��Ժʱ��")).Value, "yyyy-mm-dd")))
            End If
            If lngסԺ���� = 0 Then lngסԺ���� = 1
            strTmp = "����ID:" & mlng����ID & ",����:" & .Record(GetRptRsColumn("����")).Value & _
                ",�� " & mlng��ҳID & " ��סԺ,סԺ����:" & lngסԺ���� & "��,��Ժʱ��:" & .Record(GetRptRsColumn("��Ժʱ��")).Value & ",��Ժʱ��:" & _
                .Record(GetRptRsColumn("��Ժʱ��")).Value
        Else
            mlng����ID = 0
            mlng��ҳID = 0
            mintInsure = 0
            strTmp = ""
        End If
    End With
    
    If mstrPrePati = mlng����ID & ":" & mlng��ҳID Then Exit Sub
    mstrPrePati = mlng����ID & ":" & mlng��ҳID
        
    sta.Panels(2).Text = strTmp
    Call tbcSub_SelectedChanged(tbcSub.Selected)
    If rptPati.Visible Then rptPati.SetFocus
End Sub


Private Sub RefreshData()
    If rptPati.Records.Count = 0 Then
        Call LoadPatients(IDKindPati.GetCurCard)
    Else
        Call rptPati_SelectionChanged
    End If
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Text = "������ɫ" Then Call zlDatabase.ShowPatiColorTip(Me)
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Visible Then Exit Sub
    
    Call SubWinDefCommandBar(Item)
    Call SubWinRefreshData(Item)
End Sub

Private Sub txtFind_Change()
    txtFind.Tag = ""
End Sub
Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, intLen As Integer
    Select Case IDKIND.GetCurCard.����
    Case "����"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        blnCard = zlCommFun.InputIsCard(txtFind, KeyAscii, IDKIND.ShowPassText)
        intLen = IDKIND.GetCardNoLen
    Case "����"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case "סԺ��"
        '63494:������,2013-10-25 ,סԺ�Ų��ܶ�λ�����б������
        If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "ҽ����"
    Case Else
            If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
           If IDKIND.GetCurCard.�ӿ���� > 0 Then
                blnCard = zlCommFun.InputIsCard(txtFind, KeyAscii, IDKIND.ShowPassText)
                intLen = IDKIND.GetCardNoLen
            End If
     End Select
     
    'ˢ����ϻ���������س�
    If (blnCard And Len(txtFind.Text) = intLen - 1 Or KeyAscii = 13) And KeyAscii <> 8 Then
        If KeyAscii <> 13 Then
            txtFind.Text = txtFind.Text & Chr(KeyAscii)
            txtFind.SelStart = Len(txtFind.Text)
        End If
        KeyAscii = 0:
        Call ExecFindPati(IDKIND.GetCurCard, , blnCard)
        zlControl.TxtSelAll txtFind
   End If
End Sub

Private Sub txt����_Change()
    txt����.Tag = ""
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call OpenIme(gstrIme)
End Sub
 
Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, intTYPE As Integer
    Dim strKind As String, intLen As Integer
    'Switch(intTmp = 0, "������(&6)", intTmp = 1, "���￨��(&6)", intTmp = 2, "���š�(&6)", intTmp = 3, "ҽ���š�(&6)", True, "������(&6)")
    intTYPE = Val(IDKindPati.Tag)
    
    txt����.PasswordChar = IIf(IDKindPati.ShowPassText, "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txt����.IMEMode = 0
    strKind = IDKindPati.GetCurCard.����    '56866
    'ȡȱʡ��ˢ����ʽ
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
    Select Case strKind
    Case "����"
        KeyAscii = Asc(Chr(KeyAscii))
        blnCard = zlCommFun.InputIsCard(txt����, KeyAscii, IDKindPati.ShowPassText)
        intLen = IDKindPati.GetCardNoLen
    Case "����"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case "סԺ��"
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "ҽ����"
    Case Else
            If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
           If IDKindPati.GetCurCard.�ӿ���� > 0 Then
                blnCard = zlCommFun.InputIsCard(txt����, KeyAscii, IDKindPati.ShowPassText)
                intLen = IDKindPati.GetCardNoLen
            End If
     End Select
     
    'ˢ����ϻ���������س�
    If blnCard And Len(txt����.Text) = intLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txt����.Text = txt����.Text & Chr(KeyAscii)
            txt����.SelStart = Len(txt����.Text)
        End If
        KeyAscii = 0
        If Trim(txt����.Text) <> "" Then
            If blnCard Then
                Call LoadPatients(IDKindPati.GetCurCard, blnCard)
                Call tbcSub_SelectedChanged(tbcSub.Selected)
            Else
                Call cmdSearch_Click
                zlControl.TxtSelAll txt����
            End If
        Else
            zlCommFun.PressKey vbKeyTab
        End If
   End If
     
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    txt����.Text = Trim(txt����.Text)
    Call OpenIme
End Sub

Private Sub txtסԺ��_GotFocus()
    zlControl.TxtSelAll txtסԺ��
End Sub

Private Sub txtסԺ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    '24547
    If Trim(txtסԺ��.Text) <> "" Then
        Call cmdSearch_Click
        zlControl.TxtSelAll txtסԺ��
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtסԺ��_Validate(Cancel As Boolean)
    txtסԺ��.Text = Trim(txtסԺ��.Text)
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2012-05-21 14:32:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    '�ֶ���,���,�Ƿ��������;(���,�Ƿ�������鲻д��ʾ����������)
      mstrPatiHead = "" & _
    "����ID;��ҳID;�Ǽ�ʱ��;״̬;��������;����ת��;��ǰ����ID;����;����;��ǰ����ID;���￨��;" & _
    "���,30,0;����,60,0;סԺ��,60,0;����,50,0;�ѱ�,60,1;�Ա�,40,1;����,60,1;��Ժʱ��,100,0;��Ժʱ��,100,0;" & _
    "��ǰ����,80,1;����,40,1;����,40,1;ҽ����,120,0;��ϵ�绰,80,0;ҽ�Ƹ��ʽ,120,1"
    If gTy_System_Para.byt������˷�ʽ = 0 Then
        mstrPatiHead = mstrPatiHead & ";���״̬;�����,45,1;��������,100,1;��ǰ����,80,1"
    Else
        mstrPatiHead = mstrPatiHead & ";���״̬,100,1;�����,45,1;��������,100,1;��ǰ����,80,1"
    End If
End Sub
Private Sub Form_Load()
    Dim lngTmp As Long, strTmp As String, DatTmp As Date, blnAdviceQuery As Boolean
    Dim objPan As Pane, strValue As String, objCondition As Pane
    mbytFontSize = IIf(Val(zlDatabase.GetPara("��ʾ�����С", glngSys, glngModul)) = 0, 9, 12)
    Call InitData    ' ��ʼ����Ҫ����
    mblnSelPatiList = False
    mstrPrivs = gstrPrivs: mlngModul = glngModul
    mblnHavePara = InStr(1, mstrPrivs, ";��������;") > 0
    Call InitMenus
    
    mstr��ֹ���� = ""
    blnAdviceQuery = GetInsidePrivs(pסԺҽ���´�) <> ""
    If InStr(GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "סԺ����") > 0 Then Call InitLocPar(Enum_Inside_Program.pסԺ����)
    If InStr(GetInsidePrivs(Enum_Inside_Program.p���˽���), "סԺ���ý���") > 0 Then Call InitLocPar(Enum_Inside_Program.p���˽���)
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����", , , True)) = 1 Then
        IDKindPati.IDKIND = IIf(Val(zlDatabase.GetPara("���˹������", glngSys, mlngModul, "1")) = 0, 1, Val(zlDatabase.GetPara("���˹������", glngSys, mlngModul, "1")))
        GetRegInFor g˽��ģ��, Me.Name, "IDKind", strValue
        IDKIND.IDKIND = IIf(Val(strValue) = 0, 1, Val(strValue))
    End If
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Set fraCondition.Container = picCondition
    fraCondition.Top = 0
    fraCondition.Left = 0
    '�˵���ʼ
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.VisualTheme = xtpThemeOffice2003
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    '���岼������ʼ
    '-----------------------------------------------------
    dkpMain.SetCommandBars Me.cbsMain
    dkpMain.VisualTheme = ThemeOffice2003
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = False
    
    Set objCondition = dkpMain.CreatePane(mPan.Condition, 220, 200, DockLeftOf, Nothing)
    objCondition.Title = "��ѯ����": objCondition.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPan = dkpMain.CreatePane(mPan.Pati, 200, 500, DockBottomOf, objCondition)
    objPan.Title = "�����б�": objPan.Options = PaneNoCloseable Or PaneNoFloatable
        
     'TabControl
    '-----------------------------------------------------
    Set mcolSubForm = New Collection
    Set mclsFeeQuery = New clsFeeQuery
    If blnAdviceQuery Then Set mclsAdvices = CreateObject("zlCISKernel.clsDockInAdvices")
    
    mcolSubForm.Add mclsFeeQuery.zlGetForm, "_����"
    If blnAdviceQuery Then mcolSubForm.Add mclsAdvices.zlGetForm, "_ҽ��"
    
    Set mfrmPatiFeeVerfy = New frmPatiFeeVerfy
    Load mfrmPatiFeeVerfy
    mcolSubForm.Add mfrmPatiFeeVerfy, "_ҽ�������"
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(0, "���ò�ѯ", mcolSubForm("_����").hWnd, 0).Tag = "����"
        If blnAdviceQuery Then .InsertItem(1, "ҽ����ѯ", mcolSubForm("_ҽ��").hWnd, 0).Tag = "ҽ��"
        .InsertItem(2, "ҽ�������", mcolSubForm("_ҽ�������").hWnd, 0).Tag = "ҽ�������"
        
        Call SubWinDefCommandBar(.Selected)   '��ʼˢ�¶���һ�β˵�����ť
        Call SubWinRefreshData(.Selected)
    End With
                    
    Call InitRPTPati
    
    PicRptPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    
    '���˲���
    If Not InitUnits Then mblnUnload = True: Exit Sub

    If InStr(";" & mstrPrivs, ";��Ժ���˲�ѯ;") = 0 Then
        strTmp = "��Ժ����,Ԥ��Ժ����"
    Else
        strTmp = "��Ժ����,Ԥ��Ժ����,��Ժ����,���в���"
    End If
    Call CboAddByStrings(cboState, strTmp, False)
    
    strTmp = zlDatabase.GetPara("����״̬", glngSys, mlngModul, "��Ժ����")
    Call zlControl.CboLocate(cboState, strTmp)
    
    
    lngTmp = Val(zlDatabase.GetPara("�������", glngSys, mlngModul, -1))
    If lngTmp > 100 Then lngTmp = 7 'ֻ���7��
    DatTmp = zlDatabase.Currentdate()
    '42849
    dtpEnd.Value = Format(DatTmp, "yyyy-mm-dd 23:59:59")
    If lngTmp = -1 Then
        dtpBegin.Value = CDate(Format(DateAdd("m", -1, DatTmp), "yyyy-mm-dd") & " 00:00:00")
    Else
        dtpBegin.Value = CDate(Format(DateAdd("d", -lngTmp, DatTmp), "yyyy-mm-dd") & " 00:00:00")
    End If
    
    chk����δ���岡��.Value = IIf(zlDatabase.GetPara("����δ���岡��", glngSys, mlngModul, "0") = "1", 1, 0)
    chk����δ��˲���.Value = IIf(zlDatabase.GetPara("����δ��˲���", glngSys, mlngModul, "0") = "1", 1, 0)
    
    mvs.OnePati = zlDatabase.GetPara("������ʾ", glngSys, mlngModul, "0") = "1"
        
         
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����", , , True)) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    End If

    Call RestoreWinState(Me, App.ProductName)
    Call tbcSub_SelectedChanged(tbcSub.Selected)
    '50793
    Call SetFontSize(mbytFontSize)
    Call picFind_Resize
    dkpMain.Panes(1).MinTrackSize.Height = IIf(mbytFontSize = 9, 270, 300)
    dkpMain.Panes(1).MinTrackSize.Width = IIf(mbytFontSize = 9, 225, 295)
    dkpMain.RedrawPanes
End Sub

Private Sub InitRPTPati()
    Dim arrTmp As Variant, arrItem As Variant, i As Long
    Dim rptCol As ReportColumn
    
    With rptPati
        Set .Container = PicRptPati
        arrTmp = Split(mstrPatiHead, ";")
        For i = 0 To UBound(arrTmp)
            arrItem = Split(arrTmp(i), ",")
            If UBound(arrItem) > 0 Then
                Set rptCol = .Columns.Add(i, CStr(arrItem(0)), Val(arrItem(1)), True)
                rptCol.Visible = True
                rptCol.Alignment = xtpAlignmentCenter
                
                rptCol.Groupable = Val(arrItem(2)) = 1
            Else
                Set rptCol = .Columns.Add(i, CStr(arrItem(0)), 0, False)
                rptCol.Visible = False
            End If
        Next
        
        .SetImageList imgPati
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .AutoColumnSizing = False
        .ShowGroupBox = True
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û���ҵ����������Ĳ���..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With

End Sub


Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ���˵���������
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long
    
    '��¼���в˵���ʽ
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        blnShowBar = cbsMain(2).Visible
        bytStyle = cbsMain(2).Controls(1).Style
    End If
    
    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hWnd)
        
    Me.Caption = objItem.Caption
        
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '���������¼���
    Call MainDefCommandBar
    
    '�Ӵ������¼���
    Select Case objItem.Tag
        Case "����"
            Call mclsFeeQuery.zlDefCommandBars(Me, Me.cbsMain, 0)
        Case "ҽ��"
            Call mclsAdvices.zlDefCommandBars(Me, Nothing, 1)
    End Select
    
    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = bytStyle
        Next
        cbsMain(lngCount).Visible = blnShowBar
        
    Next
    
    '�������RecalcLayout����������
    Call LockWindowUpdate(0)
    Set mfrmActive = mcolSubForm("_" & objItem.Tag)
    Call IDKIND.Refrash
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ������ݼ�״̬
    Dim blnDateMoved As Boolean
    Dim lngDeptID As Long
    
    Select Case objItem.Tag
        Case "����"
            If mlng����ID = 0 Then
                '����:25850
                If cboUnit.ListIndex >= 0 Then lngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
                Call mclsFeeQuery.zlRefresh(0, 0, 0, lngDeptID, 0, False, False, False)
            Else
                With rptPati.SelectedRows(0)
                    If Val(.Record(GetRptRsColumn("����ת��")).Value) = 1 Then
                        blnDateMoved = True
                    Else
                        blnDateMoved = zlDatabase.DateMoved(Format(.Record(GetRptRsColumn("��Ժʱ��")).Value, "yyyy-MM-dd 00:00:00"), , , Caption)
                    End If
                    '����:25850
                    lngDeptID = Val(.Record(GetRptRsColumn("��ǰ����ID")).Value)
                    If cboUnit.ListIndex >= 0 And lngDeptID = 0 Then lngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
                     
                    Call mclsFeeQuery.zlRefresh(mlng����ID, mlng��ҳID, Val(.Record(GetRptRsColumn("סԺ��")).Value), Val(.Record(GetRptRsColumn("��ǰ����ID")).Value), _
                        Val(.Record(GetRptRsColumn("����")).Value), blnDateMoved, Trim(.Record(GetRptRsColumn("��Ժʱ��")).Value) <> "", Trim(.Record(GetRptRsColumn("����")).Value) <> "", , Trim(.Record(GetRptRsColumn("��Ժʱ��")).Value) <> "")
                End With
            End If
        Case "ҽ�������"
                If mlng����ID <> 0 Then
                    With rptPati.SelectedRows(0)
                        If Val(.Record(GetRptRsColumn("����ת��")).Value) = 1 Then
                            blnDateMoved = True
                        Else
                            blnDateMoved = zlDatabase.DateMoved(Format(.Record(GetRptRsColumn("��Ժʱ��")).Value, "yyyy-MM-dd 00:00:00"), , , Caption)
                        End If
                    End With
                End If
                Call mfrmPatiFeeVerfy.ShowData(mlng����ID, mlng��ҳID, blnDateMoved)
        Case "ҽ��"
            If mlng����ID = 0 Then
                Call mclsAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
            Else
                With rptPati.SelectedRows(0)
                    '�ȸ��²�����صı������ٵ�ҽ����ˢ��
                    If Val(.Record(GetRptRsColumn("����ת��")).Value) = 1 Then
                        blnDateMoved = True
                    Else
                        blnDateMoved = zlDatabase.DateMoved(Format(.Record(GetRptRsColumn("��Ժʱ��")).Value, "yyyy-MM-dd 00:00:00"), , , Caption)
                    End If
                    
                    Call mclsFeeQuery.zlRefresh(mlng����ID, mlng��ҳID, Val(.Record(GetRptRsColumn("סԺ��")).Value), Val(.Record(GetRptRsColumn("��ǰ����ID")).Value), _
                        Val(.Record(GetRptRsColumn("����")).Value), blnDateMoved, Trim(.Record(GetRptRsColumn("��Ժʱ��")).Value) <> "", Trim(.Record(GetRptRsColumn("����")).Value) <> "", True)

                    Call mclsAdvices.zlRefresh(mlng����ID, mlng��ҳID, Val(.Record(GetRptRsColumn("��ǰ����ID")).Value), Val(.Record(GetRptRsColumn("��ǰ����ID")).Value), _
                        IIf(.Record(GetRptRsColumn("��Ժʱ��")).Value = "", IIf(Val(.Record(GetRptRsColumn("״̬")).Value) = 3, 1, 0), 2), Val(.Record(GetRptRsColumn("����ת��")).Value) = 1)
                End With
            End If
    End Select
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objBar As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim arrTmp As Variant, i As Long, j As Long
        
    '-----------------------------------------------------
    '�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        .Add xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��"
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview_Pati, "��ӡԤ�������б�(&T)")
        objControl.BeginGroup = True:        objControl.IconId = conMenu_File_Preview
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_Pati, "��ӡ�����б�(&O)")
         objControl.IconId = conMenu_File_Print
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel_Pati, "�����б������Excel(&E)")
         objControl.IconId = conMenu_File_Excel
        
        .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)").BeginGroup = True
        .Add xtpControlButton, conMenu_File_Print, "��ӡ(&P)"
        .Add xtpControlButton, conMenu_File_Excel, "�����Excel(&L)"
        .Add xtpControlButton, conMenu_File_PrintMultiBill, "��ӡ���Ŵ߿(&N)��"
        .Add(xtpControlButton, conMenu_File_SchemeSet, "������������(&F)��").BeginGroup = True
        .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)").BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        .Add xtpControlButton, conMenu_Edit_PreBalanceAll, "Ԥ�����в���(&I)"
        .Add xtpControlButton, conMenu_Edit_PreBalance, "Ԥ�ᵱǰ����(&W)"
        .Add xtpControlButton, conMenu_Edit_Balance, "����(&B)"
        .Add(xtpControlButton, conMenu_Edit_FeeAudit, IIf(gTy_System_Para.byt������˷�ʽ = 1, "��ʼ���(&A)", "���(&A)")).BeginGroup = True
        .Add(xtpControlButton, conMenu_Edit_OverFeeAudit, "������(&O)").IconId = 252
        .Add xtpControlButton, conMenu_Edit_FeeUnAudit, "ȡ�����(&U)"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InsureSel, "����ѡ��(&Z)")
        objControl.BeginGroup = True
       Set objControl = .Add(xtpControlButton, conMenu_Edit_PrePayMoney, "��Ԥ��(&P)")
       objControl.IconId = 3816: objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
       Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����
  
        Set objControl = .Add(xtpControlButton, conMenu_View_FontSize_S, "С����(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_FontSize_L, "������(&U)")

        Set objControl = .Add(xtpControlButton, conMenu_View_OnePati, "���סԺֻ��һ�β���(&O)")
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)")
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False)
        End With
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_GroupCol, "���˷�������(&G)")
        arrTmp = Split(mstrPatiHead, ";")
        For i = 0 To UBound(arrTmp)
            If UBound(Split(arrTmp(i), ",")) > 1 Then
                If Val(Split(arrTmp(i), ",")(2)) = 1 Then
                    j = j + 1
                    objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_View_GroupCol * 10 + j, Split(arrTmp(i), ",")(0)
                End If
            End If
        Next
        
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&N)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With
    
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With


    '���������⴦��
    '-----------------------------------------------------
    '���˵��Ҳ�Ĳ���
    With cbsMain.ActiveMenuBar.Controls
'        Set objPopup = .Add(xtpControlPopup, conMenu_View_FindType, "����")
'        objPopup.ID = conMenu_View_FindType
'        objPopup.flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = picFind.hWnd
        objCustom.flags = xtpFlagRightAlign
        IDKIND.BackColor = picFind.BackColor
    End With

    '����������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop) '����
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FeeAudit, IIf(gTy_System_Para.byt������˷�ʽ = 1, "��ʼ���", "���"))
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OverFeeAudit, "������")
        objControl.IconId = 252
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Balance, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PrePayMoney, "��Ԥ��")
        objControl.IconId = 3816
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
    End With
    
    
    
    '�����
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("F"), conMenu_View_Find
        .Add 0, VK_F3, conMenu_View_FindNext
        .Add 0, VK_F5, conMenu_View_Refresh
        
        .Add FCONTROL, Asc("A"), conMenu_Edit_FeeAudit
        If gTy_System_Para.byt������˷�ʽ = 1 Then
            .Add FCONTROL, Asc("O"), conMenu_Edit_OverFeeAudit
        End If
        .Add FCONTROL, Asc("U"), conMenu_Edit_FeeUnAudit
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With cbsMain.Options     '��������ˣ��ؼ��ڲ˵���һ����ʾʱû�е���update�¼�
'        .AddHiddenCommand conMenu_View_Owe
'        .AddHiddenCommand conMenu_View_UnAudit
        '.AddHiddenCommand conMenu_View_OnePati
    End With
        
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1139_3")   '��ӡ�߿
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, lngTmp As Long
    
    SaveWinState Me, App.ProductName
    lngTmp = Val(dtpEnd.Value - dtpBegin.Value)
    If lngTmp > 100 Then lngTmp = 7
    zlDatabase.SetPara "��ʾ�����С", IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
      
    zlDatabase.SetPara "����״̬", cboState.Text, glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "�������", lngTmp, glngSys, mlngModul, mblnHavePara
     zlDatabase.SetPara "����δ���岡��", IIf(chk����δ���岡��.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "����δ��˲���", IIf(chk����δ��˲���.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "������ʾ", IIf(mvs.OnePati, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "���˹������", IDKindPati.IDKIND, glngSys, mlngModul, mblnHavePara
    SaveRegInFor g˽��ģ��, Me.Name, "IDKind", IDKIND.IDKIND
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����", , , True)) = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    End If
    Call SaveWinState(Me, App.ProductName)
    mlng����ID = 0
    mlng��ҳID = 0
    mstrPrePati = ""
    mblnUnload = False
    
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mfrmActive = Nothing
    Set mclsFeeQuery = Nothing
    Set mclsAdvices = Nothing
    Set mobjPatient = Nothing
     
End Sub
Private Function CheckIsAllowAudit(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ��������
    '����:���˺�
    '����:�������,����true,���򷵻�False
    '����:2012-06-19 14:04:21
    '����:50778
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select NO,��¼���� From סԺ���ü�¼ Where ����ID=[1] and ��ҳID=[2] and ���ʷ���=1 And ��¼״̬=0 and rownum<=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
      MsgBox "�ò��˻���δ��Ч�ķ���,���ܽ������!", vbInformation + vbOKOnly, gstrSysName
      Exit Function
    End If
    CheckIsAllowAudit = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ExecAuditingAndCancelAudit(ByVal bytType As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ����˻�ȡ����˲���
    '����:bytType-0-ȡ�����;1-��ʼ��˻����;2-������;
    '����:���˺�
    '����:2012-05-21 11:57:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytAudit As Byte, strSQL As String, strTemp As String, strExpend As String
    Dim blnCheck As Boolean
    If gTy_System_Para.byt������˷�ʽ = 1 Then
        If CheckIsAllowAudit(mlng����ID, mlng��ҳID) = False Then Exit Sub
    End If
    
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear
        On Error GoTo errH:
        If Not mobjPlugIn Is Nothing Then
            Call mobjPlugIn.Initialize(gcnOracle, glngSys, mlngModul)
        End If
    End If
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        blnCheck = mobjPlugIn.PatiFeeAuditingAndCancelCheck(mlngModul, mlng����ID, mlng��ҳID, bytType = 0, strExpend)
        If Err = 0 Then
            '���ڼ��ӿ�
            If blnCheck = False Then
                MsgBox "�޷��Բ��˽���" & IIf(bytType = 0, "ȡ�����!", "���!"), vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            '�����ڼ��ӿڵĲ����
            Err.Clear
        End If
        On Error GoTo errH:
    End If
    
    With rptPati.SelectedRows(0)
        'bytAudit-0-δ���;1-��ʼ��˻������;2-������
        strTemp = Trim(.Record(GetRptRsColumn("���״̬")).Value)
        bytAudit = Switch(strTemp = "��ʼ" Or strTemp = "����", 1, strTemp = "���", 2, True, 0)
        If bytType = 0 Then
            bytAudit = Switch(bytAudit = 2, 1, bytAudit = 1, 0, True, 0)
        Else
            bytAudit = bytType
        End If
        
        On Error GoTo errH
        ' Zl_�������_Execute
        strSQL = "Zl_�������_Execute("
        '  ����id_In   ������ҳ.����id%Type,
        strSQL = strSQL & "" & mlng����ID & ","
        '  ��ҳid_In   ������ҳ.��ҳid%Type,
        strSQL = strSQL & "" & mlng��ҳID & ","
        '  ��˱�־_In ������ҳ.��˱�־%Type,
        strSQL = strSQL & "" & bytAudit & ","
        '  �����_In   ������ҳ.�����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ������ʽ_In Integer:=0 --������ʽ_In:0-���;1-ȡ�����
        strSQL = strSQL & IIf(bytType = 0, 1, 0) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        'Ϊ�ٶ��Ż�,������Call LoadPatients,ֱ�Ӹ�д״̬
        Select Case bytAudit
        Case 1
            .Record(GetRptRsColumn("���״̬")).Value = IIf(gTy_System_Para.byt������˷�ʽ = 1, "��ʼ", "����")
            .Record(GetRptRsColumn("�����")).Value = UserInfo.����
        Case 2
            .Record(GetRptRsColumn("���״̬")).Value = "���"
            .Record(GetRptRsColumn("�����")).Value = UserInfo.����
        Case Else
            .Record(GetRptRsColumn("���״̬")).Value = ""
            .Record(GetRptRsColumn("�����")).Value = ""
        End Select
        If bytAudit = "1" Or bytAudit = "2" Then
            .Record(GetRptRsColumn("����")).ForeColor = &H33AA22
        Else
            .Record(GetRptRsColumn("����")).ForeColor = .Record(GetRptRsColumn("סԺ��")).ForeColor   '��ԭ����������ɫ
        End If
    End With
    rptPati.Populate
    Select Case bytType
    Case 0
        sta.Panels(3).Text = "ȡ��" & IIf(gTy_System_Para.byt������˷�ʽ = 1, IIf(bytAudit = 0, "��ʼ", "���"), "") & "��˳ɹ�!"
    Case 1
        sta.Panels(3).Text = IIf(gTy_System_Para.byt������˷�ʽ = 1, "��ʼ", "") & "��˳ɹ�!"
    Case 2
        sta.Panels(3).Text = "�����˳ɹ�!"
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ����������
    Dim i As Long, strNodes As String, strDefaultNode As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set mrsDept = GetUnit(InStr(mstrPrivs, ";���в���;") = 0, "1,2,3", "����", True, True, True)
    strNodes = ""
    mblnNotClick = True
    With mrsDept
        Do While Not .EOF
            If Nvl(!վ��) <> "" And InStr(1, "," & strNodes & ",", "," & Nvl(!վ��) & ",") = 0 Then
                strNodes = strNodes & "," & Nvl(!վ��)
            End If
            If mrsDept!ID = UserInfo.����ID Then strDefaultNode = Nvl(!վ��)
             .MoveNext
        Loop
    End With
     
    cboNode.Clear
    If strNodes <> "" Then
        strNodes = Mid(strNodes, 2)
        gstrSQL = "" & _
        "   Select /*+ RULE */A.���,A.���� " & _
        "   From zlNodeList A,Table(Cast(f_num2list([1]) As Zltools.t_numlist)) J" & _
        "   where A.���=j.Column_Value " & _
        "   Order by ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNodes)
        With cboNode
             Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!���) & "-" & Nvl(rsTemp!����)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!���))
                If .ItemData(.NewIndex) = Val(strDefaultNode) Then
                    .ListIndex = .NewIndex
                End If
                rsTemp.MoveNext
             Loop
             If .ListIndex < 0 And .ListCount >= 1 Then .ListIndex = 0
        End With
    End If
    fraվ��.Visible = cboNode.ListCount > 0
    mbln����վ�� = cboNode.ListCount > 0
    If mrsDept.RecordCount <> 0 Then mrsDept.MoveFirst
        
    '����:50743
    If mrsDept.EOF Then
        MsgBox "û�з��ֻ�������Ϣ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    If Not mrsDept.EOF Then
        If LoadUnits = False Then Exit Function
    ElseIf InStr(";" & mstrPrivs, ";���в���;") > 0 Then
        MsgBox "û�з��ֻ�������Ϣ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    Call SizeWinCons
    mblnNotClick = False
    InitUnits = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SizeWinCons()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '����:���˺�
    '����:2011-02-28 18:06:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim lngTop As Long
     lngTop = IIf(cboNode.ListCount > 0, fraվ��.Top + fraվ��.Height, fraվ��.Top + 10)
     cboUnit.Top = lngTop
     lblUnit.Top = cboUnit.Top + (cboUnit.Height - lblUnit.Height) \ 2
     lngTop = cboUnit.Top + cboUnit.Height + 30
     cboState.Top = lngTop
     lblState.Top = cboState.Top + (cboState.Height - lblState.Height) \ 2
     dtpBegin.Top = cboState.Top + cboState.Height + 30
     lblStartDate.Top = dtpBegin.Top + (dtpBegin.Height - lblStartDate.Height) \ 2
     dtpEnd.Top = dtpBegin.Top + dtpBegin.Height + 30
     lblEndDate.Top = dtpEnd.Top + (dtpEnd.Height - lblEndDate.Height) \ 2
     txtסԺ��.Top = dtpEnd.Top + dtpEnd.Height + 30
     lblסԺ��.Top = txtסԺ��.Top + (txtסԺ��.Height - lblסԺ��.Height) \ 2
     txt����.Top = txtסԺ��.Top + txtסԺ��.Height + 30
     IDKindPati.Top = txt����.Top + (txt����.Height - IDKindPati.Height) \ 2
     txtԤ��.Top = txt����.Top + txt����.Height + 30
     chkԤ����.Top = txtԤ��.Top + (txtԤ��.Height - chkԤ����.Height) \ 2
     chk����δ���岡��.Top = txtԤ��.Top + txtԤ��.Height + 30
     chk����δ��˲���.Top = chk����δ���岡��.Top + chk����δ���岡��.Height + 30
     cmdSearch.Top = chk����δ��˲���.Top + chk����δ��˲���.Height + 30
End Sub
 
 
Private Function LoadUnits() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '���:strվ��-ָ��վ��,strվ��="",��ʾ������վ��
    '����:
    '����:���سɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2011-02-28 17:51:27
    '����:36048
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    mrsDept.Filter = 0
    If cboNode.ListIndex >= 0 Then
         mrsDept.Filter = "վ��=" & cboNode.ItemData(cboNode.ListIndex) & " or վ��=NULL"
    End If
    cboUnit.Clear
    If InStr(";" & mstrPrivs, ";���в���;") > 0 Then cboUnit.AddItem "���в���"
    With mrsDept
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            cboUnit.AddItem !���� & "-" & !����
            cboUnit.ItemData(cboUnit.ListCount - 1) = !ID
            If mrsDept!ID = UserInfo.����ID Then cboUnit.ListIndex = cboUnit.NewIndex
            .MoveNext
        Loop
    End With
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    LoadUnits = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadPatients(ByVal objCard As Card, Optional blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����Χ�ڵĲ����б�
    '���:blnCard-�Ƿ�ˢ��
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-12-01 14:13:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intBedLen As Integer
    Dim rsPati As ADODB.Recordset, dblԤ����� As Double
    Dim i As Long, j As Long, blnUnIndex As Boolean, strCount As String, strסԺ�� As String, str���� As String
    Dim DateBegin As Date, DateEnd As Date, lngUnitID As Long, lngPatiRow As Long, strPatiRow As String
    Dim objRecord As ReportRecord, objRow As ReportRow
    Dim objItem As ReportRecordItem, strKind As String, intTYPE As Integer
    Dim strWhere As String, strNodeNo As String   'վ��
    Dim lng�����ID As Long, lng����ID As Long
    
    Dim strPassWord As String, strErrMsg As String
    On Error GoTo errH
    Screen.MousePointer = 11
    
    strNodeNo = gstrNodeNo  '��ǰվ���
    
    Call zlCommFun.ShowFlash("����ͳ������,���Ժ� ...", Me)
    DoEvents
    Refresh
    mstrPrePati = ""
    DateBegin = CDate(Format(dtpBegin.Value, "yyyy-MM-dd HH:MM:SS"))
    DateEnd = CDate(Format(dtpEnd.Value, "yyyy-MM-dd HH:MM:SS"))
    '����:50743
    If cboUnit.ListIndex < 0 Then
        lngUnitID = -1
    Else
        lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    End If
    
    
    strסԺ�� = Trim(txtסԺ��.Text)
    str���� = Trim(txt����.Text)
    
    If cboState.Text = "��Ժ����" Then
        '��ǰ��Ժ�Ĳ���
        strSQL = " And A.��Ժ=1 And Nvl(B.״̬,0)<>3 And A.��ҳID=B.��ҳID "
    ElseIf cboState.Text = "��Ժ����" Then
        '���ڼ��ڳ�Ժ
        strSQL = " And B.��Ժ����<=[2] And B.��Ժ���� Between [1] And [2]"
        If mvs.OnePati Then strSQL = strSQL & " And A.��ҳID=B.��ҳID"
    ElseIf cboState.Text = "Ԥ��Ժ����" Then
        'Ԥ��Ժ����
        strSQL = " And A.��Ժ=1 And B.״̬=3 "
    Else '���в���
        If (strסԺ�� = "" And str���� = "") Or (strסԺ�� = "" And Len(str����) = 1 And IDKindPati.IDKIND = 1) Then
            strSQL = " And ((A.��Ժ=1  And A.��ҳID=B.��ҳID) Or (B.��Ժ���� Between [1] And [2]))"
        Else
            strSQL = ""
        End If
        If mvs.OnePati Then strSQL = strSQL & " And A.��ҳID=B.��ҳID"
    End If
    
    If strסԺ�� <> "" Then
        strSQL = strSQL & " And A.����ID = (Select distinct ����ID From ������ҳ Where סԺ��=[4])": blnUnIndex = True
    End If
    
    If str���� <> "" Then
         strKind = objCard.����
        Select Case strKind
            Case "����"
                lng����ID = Val(txt����.Tag)
                If blnCard Then
                    If IDKIND.Cards.��ȱʡ������ And Not IDKIND.GetfaultCard Is Nothing Then
                        lng�����ID = IDKIND.GetfaultCard.�ӿ����
                    Else
                        lng�����ID = "-1"
                    End If
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, str����, True, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then lng����ID = 0
                    txt����.Tag = lng����ID
                    lng����ID = Val(txt����.Tag)
                    strSQL = strSQL & " And A.����ID=[10]"
                ElseIf lng����ID <> 0 Then
                    strSQL = strSQL & " And A.����ID=[10]"
                Else
                    strSQL = strSQL & " And A.���� like [5]": If Len(str����) > 1 Then blnUnIndex = True    '�����ֻ��һ����ĸ,����Ӱ������,���Բ�������
                End If
            Case "����"
                strSQL = strSQL & " And B.��Ժ����=[6]":   blnUnIndex = True
            Case "ҽ����"
                strSQL = strSQL & " And (F.ҽ����=[6] or F.ҽ���� IS NULL and  D.��Ϣֵ=[6])":   blnUnIndex = True
            Case "סԺ��"
                If gblnÿ��סԺ��סԺ�� Or True Then
                    strSQL = strSQL & " And A.����ID = (Select distinct ����ID From ������ҳ Where סԺ��=[4])": blnUnIndex = True
                Else
                    strSQL = strSQL & " And A.סԺ��=[4]": blnUnIndex = True
                End If
                '����:50788
                strסԺ�� = str����
            Case Else
               lng�����ID = objCard.�ӿ����
                If lng�����ID > 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, str����, True, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(strKind, str����, True, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                strSQL = strSQL & " And A.����ID=[10]"
        End Select
    End If
    
    If cboState.Text = "��Ժ����" Then blnUnIndex = True
    If lngUnitID > 0 Then strSQL = strSQL & " And B.��ǰ����ID" & IIf(blnUnIndex, "+0", "") & "=[3]"
     If cboNode.ListCount > 0 And cboNode.ListIndex >= 0 Then
        'ֻ�ܲ��վ������в����Ĳ���
        strNodeNo = cboNode.ItemData(cboNode.ListIndex)
    End If
    
    If chk����δ���岡��.Value = 1 Then strSQL = strSQL & " And Nvl(X.�������,0) <> 0 And Exists (Select 1 From ����δ����� Where ����id = a.����id And Nvl(��ҳid,0) = Nvl(b.��ҳid,0) And Rownum < 2)"
    If chk����δ��˲���.Value = 1 Then strSQL = strSQL & " And ����� is Null"
     
    
    dblԤ����� = Val(txtԤ��.Text)
    intBedLen = GetMaxBedLen(lngUnitID)
    
    If chkԤ����.Value = 1 Then
        '����:27546
        '��ʾ����Ԥ�������С�ڣ����ٵĲ���
        '��ΪԤ��Ժ���˵ĳ�Ժʱ���ڱ䶯��¼��,����Ҫȡ���˱䶯��¼���ű�
        If cboState.Text = "��Ժ����" Or cboState.Text = "��Ժ����" Then
            strSQL = "Select A.����ID, B.��ҳID, A.�Ǽ�ʱ��, B.״̬, B.��������, B.����ת��, B.��Ժ����id As ��ǰ����id, A.���￨��, B.����, E.����, B.��ǰ����ID, Nvl(b.����, a.����) As ����, B.סԺ��, B.��Ժ���� As ����," & vbNewLine & _
                "       B.�ѱ�, Nvl(b.�Ա�, a.�Ա�) As �Ա�, Nvl(b.����, a.����) as ����, B.��Ժ����, B.��Ժ����, C.���� As ��ǰ����, Decode(Nvl(X.�������, 0), 0, '��', '') As ����," & vbNewLine & _
                "       Nvl(E.ҽ����, D.��Ϣֵ) ҽ����, A.��ͥ�绰, B.ҽ�Ƹ��ʽ,B.��˱�־, B.�����, B.��������, H.���� ��ǰ����" & vbNewLine & _
                "From ������Ϣ A, ������ҳ B, ������ҳ�ӱ� D, ҽ�����˵��� E, ҽ�����˹����� F, ������� X,����ģ����� X1, ���ű� C, ���ű� H" & vbNewLine & _
                "Where A.����ID = B.����ID And B.��Ժ����ID = C.ID And Nvl(B.��ҳID, 0) <> 0 And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) And" & vbNewLine & _
                "      D.��Ϣ��(+) = 'ҽ����' And A.����ID = X.����ID(+) And X.����(+) = 1 And X.����(+)=2 And B.����ID=X1.����ID(+) and B.��ҳid=X1.��ҳID(+) And A.����ID = F.����ID(+) And A.���� = F.����(+) And F.��־(+) = 1 And" & vbNewLine & _
                "      F.ҽ���� = E.ҽ����(+) And F.���� = E.����(+) And F.���� = E.����(+) And B.��ǰ����ID + 0 = H.ID" & vbNewLine & _
                "    And (H.վ��=[8] Or H.վ�� is Null)" & vbNewLine & strSQL & vbNewLine & _
                "Group by A.����ID, B.��ҳID, A.�Ǽ�ʱ��, B.״̬, B.��������, B.����ת��, B.��Ժ����id, A.���￨��, B.����, E.����, B.��ǰ����ID,  Nvl(b.����, a.����), B.סԺ��, B.��Ժ����,B.�ѱ�, Nvl(b.�Ա�, a.�Ա�), Nvl(b.����, a.����), B.��Ժ����, B.��Ժ����, C.����, Decode(Nvl(X.�������, 0), 0, '��', ''), " & vbNewLine & _
                "          Nvl(E.ҽ����, D.��Ϣֵ), A.��ͥ�绰, B.ҽ�Ƹ��ʽ,B.��˱�־, B.�����, B.��������, H.���� " & _
                "having (Max(nvl(X.Ԥ�����,0))-Max(nvl(x.�������,0))+Sum(nvl(X1.���,0)))<[7] " & vbNewLine & _
                IIf(lngUnitID = 0, " Order by סԺ�� Desc", " Order by LPAD(����,10,' ')")
                
        Else
            strSQL = "Select A.����ID, B.��ҳID, A.�Ǽ�ʱ��, B.״̬, B.��������, B.����ת��,B.��Ժ����id As ��ǰ����id, A.���￨��, B.����, E.����, B.��ǰ����ID, Nvl(b.����, a.����) As ����, B.סԺ��, B.��Ժ���� As ����," & vbNewLine & _
                "       B.�ѱ�, Nvl(b.�Ա�, a.�Ա�) As �Ա�, Nvl(b.����, a.����) as ����, B.��Ժ����, Decode(B.��Ժ����, Null, Z.��ʼʱ��, B.��Ժ����) ��Ժ����, C.���� As ��ǰ����," & vbNewLine & _
                "       Decode(Nvl(X.�������, 0), 0, '��', '') As ����, Nvl(E.ҽ����, D.��Ϣֵ) ҽ����, A.��ͥ�绰, B.ҽ�Ƹ��ʽ,B.��˱�־," & vbNewLine & _
                "       B.�����, B.��������, H.���� ��ǰ����" & vbNewLine & _
                "From ������Ϣ A, ������ҳ B, ������ҳ�ӱ� D, ���˱䶯��¼ Z, ҽ�����˵��� E, ҽ�����˹����� F, ������� X,����ģ����� X1, ���ű� C, ���ű� H" & vbNewLine & _
                "Where A.����ID = B.����ID And B.��Ժ����ID = C.ID And Nvl(B.��ҳID, 0) <> 0 And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) And" & vbNewLine & _
                "      D.��Ϣ��(+) = 'ҽ����' And A.����ID = X.����ID(+) And X.����(+) = 1 And X.����(+)=2 And A.����ID = F.����ID(+) And A.���� = F.����(+) And F.��־(+) = 1 And" & vbNewLine & _
                "      F.ҽ���� = E.ҽ����(+) And F.���� = E.����(+) And F.���� = E.����(+) And B.��ǰ����ID + 0 = H.ID" & vbNewLine & _
                "    And (H.վ��=[8] Or H.վ�� is Null)" & vbNewLine & _
                "   And B.����ID = Z.����ID(+) And B.��ҳID = Z.��ҳID(+) And Z.��ʼԭ��(+) = 10 And Z.���Ӵ�λ(+) = 0" & vbNewLine & _
                "   And B.����ID=X1.����ID(+) and B.��ҳid=X1.��ҳID(+)  " & vbNewLine & strSQL & vbNewLine & _
                "Group by A.����ID, B.��ҳID, A.�Ǽ�ʱ��, B.״̬, B.��������, B.����ת��,B.��Ժ����id, A.���￨��, B.����, E.����, B.��ǰ����ID, Nvl(b.����, a.����), B.סԺ��, B.��Ժ����," & _
                "         B.�ѱ�, Nvl(b.�Ա�, a.�Ա�) , Nvl(b.����, a.����), B.��Ժ����, Decode(B.��Ժ����, Null, Z.��ʼʱ��, B.��Ժ����) , C.����,Decode(Nvl(X.�������, 0), 0, '��', '')  , Nvl(E.ҽ����, D.��Ϣֵ) , A.��ͥ�绰, B.ҽ�Ƹ��ʽ," & _
                "         B.��˱�־,B.�����, B.��������, H.���� " & vbNewLine & _
                "having (Max(nvl(X.Ԥ�����,0))-Max(nvl(x.�������,0))+Sum(nvl(X1.���,0)))<[7] " & vbNewLine & _
                 IIf(lngUnitID = 0, " Order by סԺ�� Desc", " Order by LPAD(����,10,' ')")
        End If
        
    Else
            '��ΪԤ��Ժ���˵ĳ�Ժʱ���ڱ䶯��¼��,����Ҫȡ���˱䶯��¼���ű�
            If cboState.Text = "��Ժ����" Or cboState.Text = "��Ժ����" Then
                strSQL = "Select A.����ID, B.��ҳID, A.�Ǽ�ʱ��, B.״̬, B.��������, B.����ת��, B.��Ժ����id As ��ǰ����id, A.���￨��, B.����, E.����, B.��ǰ����ID, Nvl(b.����, a.����) As ����, B.סԺ��, B.��Ժ���� As ����," & vbNewLine & _
                    "       B.�ѱ�, Nvl(b.�Ա�, a.�Ա�) As �Ա� , Nvl(b.����, a.����) as ����, B.��Ժ����, B.��Ժ����, C.���� As ��ǰ����, Decode(Nvl(X.�������, 0), 0, '��', '') As ����," & vbNewLine & _
                    "       Nvl(E.ҽ����, D.��Ϣֵ) ҽ����, A.��ͥ�绰, B.ҽ�Ƹ��ʽ,B.��˱�־, B.�����, B.��������, H.���� ��ǰ����" & vbNewLine & _
                    "From ������Ϣ A, ������ҳ B, ������ҳ�ӱ� D, ҽ�����˵��� E, ҽ�����˹����� F, ������� X, ���ű� C, ���ű� H" & vbNewLine & _
                    "Where A.����ID = B.����ID And B.��Ժ����ID = C.ID And Nvl(B.��ҳID, 0) <> 0 And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) And" & vbNewLine & _
                    "      D.��Ϣ��(+) = 'ҽ����' And A.����ID = X.����ID(+) And X.����(+) = 1 And X.����(+)=2 And A.����ID = F.����ID(+) And A.���� = F.����(+) And F.��־(+) = 1 And" & vbNewLine & _
                    "      F.ҽ���� = E.ҽ����(+) And F.���� = E.����(+) And F.���� = E.����(+) And B.��ǰ����ID + 0 = H.ID" & vbNewLine & _
                    "    And (H.վ��=[8] Or H.վ�� is Null)" & vbNewLine & _
                    strSQL & IIf(lngUnitID = 0, " Order by סԺ�� Desc", " Order by LPAD(����,10,' ')")
            Else
                strSQL = "Select A.����ID, B.��ҳID, A.�Ǽ�ʱ��, B.״̬, B.��������, B.����ת��,B.��Ժ����id As ��ǰ����id, A.���￨��, B.����, E.����, B.��ǰ����ID,Nvl(b.����, a.����) As ����, B.סԺ��, B.��Ժ���� As ����," & vbNewLine & _
                    "       B.�ѱ�, Nvl(b.�Ա�, a.�Ա�) As �Ա�, Nvl(b.����, a.����) as ����, B.��Ժ����, Decode(B.��Ժ����, Null, Z.��ʼʱ��, B.��Ժ����) ��Ժ����, C.���� As ��ǰ����," & vbNewLine & _
                    "       Decode(Nvl(X.�������, 0), 0, '��', '') As ����, Nvl(E.ҽ����, D.��Ϣֵ) ҽ����, A.��ͥ�绰, B.ҽ�Ƹ��ʽ," & vbNewLine & _
                    "       B.��˱�־,B.�����, B.��������, H.���� ��ǰ����" & vbNewLine & _
                    "From ������Ϣ A, ������ҳ B, ������ҳ�ӱ� D, ���˱䶯��¼ Z, ҽ�����˵��� E, ҽ�����˹����� F, ������� X, ���ű� C, ���ű� H" & vbNewLine & _
                    "Where A.����ID = B.����ID And B.��Ժ����ID = C.ID And Nvl(B.��ҳID, 0) <> 0 And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) And" & vbNewLine & _
                    "      D.��Ϣ��(+) = 'ҽ����' And A.����ID = X.����ID(+) And X.����(+) = 1 And X.����(+)=2 And A.����ID = F.����ID(+) And A.���� = F.����(+) And F.��־(+) = 1 And" & vbNewLine & _
                    "      F.ҽ���� = E.ҽ����(+) And F.���� = E.����(+) And F.���� = E.����(+) And B.��ǰ����ID + 0 = H.ID" & vbNewLine & _
                    "    And (H.վ��=[8] Or H.վ�� is Null)" & vbNewLine & _
                    " And B.����ID = Z.����ID(+) And B.��ҳID = Z.��ҳID(+) And Z.��ʼԭ��(+) = 10 And Z.���Ӵ�λ(+) = 0" & vbNewLine & _
                    strSQL & IIf(lngUnitID = 0, " Order by סԺ�� Desc", " Order by LPAD(����,10,' ')")
            End If
    End If
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Caption, DateBegin, DateEnd, lngUnitID, Val(strסԺ��), str���� & "%", str����, dblԤ�����, strNodeNo, strסԺ��, lng����ID)
    '��¼����ѡ�еĲ���
    If rptPati.SelectedRows.Count > 0 And mlng����ID <> 0 Then
        If Not rptPati.SelectedRows(0).GroupRow And rptPati.SelectedRows(0).Childs.Count = 0 Then
            lngPatiRow = rptPati.SelectedRows(0).Index '���ڿ������¶�λ
            strPatiRow = rptPati.SelectedRows(0).Record.Tag
        End If
    End If
    rptPati.Records.DeleteAll
    
    If rsPati.RecordCount > 0 Then
        With rsPati
            
            For i = 1 To .RecordCount
                
                Set objRecord = rptPati.Records.Add()
                objRecord.Tag = !����ID & "," & Val("" & !��ҳID)
                '������
                objRecord.AddItem Val(!����ID)
                objRecord.AddItem Val("" & !��ҳID)
                objRecord.AddItem Format(!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss")
                objRecord.AddItem Val("" & !״̬)
                objRecord.AddItem Val("" & !��������)
                objRecord.AddItem Val("" & !����ת��)
                objRecord.AddItem Val("" & !��ǰ����id)
                objRecord.AddItem CStr("" & !����)
                objRecord.AddItem CStr("" & !����)
                objRecord.AddItem Val("" & !��ǰ����ID)
                objRecord.AddItem ("" & !���￨��)
                
                '��ʾ��
                Set objItem = objRecord.AddItem("")
                objItem.Icon = IIf(Val("" & !��������) = 0, 0, 1)
                
                Set objItem = objRecord.AddItem(CStr("" & !����))
                If IsNull(!�����) = False Then objItem.ForeColor = &H33AA22
                
                objRecord.AddItem "" & !סԺ��
                
                If intBedLen > Len("" & !����) Then
                    objRecord.AddItem Space(intBedLen - Len("" & !����)) & !����
                Else
                    objRecord.AddItem "" & !����
                End If
                objRecord.AddItem CStr("" & !�ѱ�)
                objRecord.AddItem CStr("" & !�Ա�)
                objRecord.AddItem CStr("" & !����)
                objRecord.AddItem Format(Nvl(!��Ժ����, ""), "yyyy-MM-dd HH:mm:ss")
                objRecord.AddItem Format(Nvl(!��Ժ����, ""), "yyyy-MM-dd HH:mm:ss")
                objRecord.AddItem CStr("" & !��ǰ����)
                objRecord.AddItem Val("" & !��ҳID)
                objRecord.AddItem CStr("" & !����)
                objRecord.AddItem CStr("" & !ҽ����)
                objRecord.AddItem CStr("" & !��ͥ�绰)
                objRecord.AddItem CStr("" & !ҽ�Ƹ��ʽ)
                If Val(Nvl(!��˱�־)) = 1 Then
                    objRecord.AddItem IIf(gTy_System_Para.byt������˷�ʽ = 1, "��ʼ", "����")
                ElseIf Val(Nvl(!��˱�־)) = 2 Then
                    objRecord.AddItem CStr("���")
                Else
                    objRecord.AddItem " "
                End If
                objRecord.AddItem CStr("" & !�����)
                objRecord.AddItem CStr("" & !��������)
                objRecord.AddItem CStr("" & !��ǰ����)
                
                For j = 0 To rptPati.Columns.Count - 1
                    If Not (IsNull(!�����) = False And rptPati.Columns(j).Caption = "����") Then
                        objRecord.Item(j).ForeColor = zlDatabase.GetPatiColor(Nvl(!��������))
                    End If
                Next
                                
                If Not mvs.OnePati Then
                    If InStr(strCount & ",", "," & rsPati!����ID & ",") = 0 Then strCount = strCount & "," & rsPati!����ID    '���ܶ��סԺ��ʾ�˶�����¼�����Բ���ֱ�����¼��
                End If
                rsPati.MoveNext
            Next
        End With
        rptPati.Populate
        
        'ȡָ��������
        If strPatiRow <> "" Then
            '�ȿ��ٶ�λ
            If lngPatiRow <= rptPati.Rows.Count - 1 Then
                If Not rptPati.Rows(lngPatiRow).GroupRow Then
                    If rptPati.Rows(lngPatiRow).Record.Tag = strPatiRow Then
                        Set objRow = rptPati.Rows(lngPatiRow)
                    End If
                End If
            End If
            '�ٽ��в���
            If objRow Is Nothing Then
                For i = 0 To rptPati.Rows.Count - 1
                    If Not rptPati.Rows(i).GroupRow Then
                        If rptPati.Rows(i).Record.Tag = strPatiRow Then
                            Set objRow = rptPati.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        End If
        'ȡ��һ���Ƿ�����
        If objRow Is Nothing Then
            For i = 0 To rptPati.Rows.Count - 1
                If Not rptPati.Rows(i).GroupRow Then Set objRow = rptPati.Rows(i): Exit For
            Next
        End If
        
        '��ѯ�����Ψһʱ,ȱʡ����ʾ��ǰ���˷��������Ϣ
        If Not (rsPati.RecordCount > 1 And lngPatiRow = 0) Then
            Set rptPati.FocusedRow = objRow '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        End If
                
        If mvs.OnePati Then
            i = rptPati.Records.Count
        Else
            i = UBound(Split(Mid(strCount, 2), ",")) + 1
        End If
                
        If cboState.Text = "��Ժ����" Then
            sta.Panels(3).Text = " �ò�����Ժ��������:" & i
        ElseIf cboState.Text = "��Ժ����" Then
            sta.Panels(3).Text = " ʱ��: " & Format(dtpBegin.Value, "yyyy-MM-dd") & " �� " & Format(dtpEnd.Value, "yyyy-MM-dd") & ",����:" & i
        ElseIf cboState.Text = "ԤԺ����" Then
            sta.Panels(3).Text = " Ԥ��Ժ��������:" & i
        Else
            sta.Panels(3).Text = " ���в���,����:" & i
        End If
    Else
        rptPati.Populate
        mlng����ID = 0: mlng��ҳID = 0
        Call tbcSub_SelectedChanged(tbcSub.Selected)
        sta.Panels(2).Text = "ָ��������û��ɸѡ���κβ���."
        sta.Panels(3).Text = ""
                  
    End If
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Refresh
    Exit Function
errH:
    Call zlCommFun.StopFlash
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetRptRsColumn(ByVal strColumn As String) As Long
'���ܣ������������ؼ�¼���������(����ı���˳���������)��û���ҵ�ʱ����-1
    Dim arrTmp As Variant, i As Long, strTmp As String
    
    arrTmp = Split(mstrPatiHead, ";")
    
    GetRptRsColumn = -1
    For i = 0 To UBound(arrTmp)
        strTmp = arrTmp(i)
        If InStr(1, strTmp, ",") > 0 Then strTmp = Mid(strTmp, 1, InStr(1, strTmp, ",") - 1)
        
        If strTmp = strColumn Then GetRptRsColumn = i: Exit For
    Next
End Function

Private Sub ExecFindPati(ByVal objCard As Card, Optional blnNext As Boolean, Optional blnCard As Boolean = True)
    Dim strKind As String, strValue As String, i As Long, lngPoint As Long
    Dim lngסԺ�� As Long, lng���� As Long, lng���￨ As Long, lng���� As Long, lngҽ���� As Long
    Dim lng�����ID As Long, lng����ID As Long, lng����IDCol As Long
    Dim strErrMsg  As String
    
    If rptPati.Records.Count = 0 Then Exit Sub
    strValue = Trim(txtFind.Text)
    If strValue = "" Then Call txtFind.SetFocus: Exit Sub
    strKind = objCard.����
    If Not IsNumeric(strValue) And strKind = "סԺ��" Then
        MsgBox "סԺ��Ҫ����������ֵ!", vbInformation, gstrSysName
        Call txtFind.SetFocus: Call zlControl.TxtSelAll(txtFind)
        Exit Sub
    End If
    
    If Not blnNext Then
        lngPoint = 0
    Else
        lngPoint = rptPati.SelectedRows(0).Index + 1
    End If
    
    Select Case strKind
        Case "����"
            lng���� = GetRptRsColumn("����")
        Case "סԺ��"
            lngסԺ�� = GetRptRsColumn("סԺ��")
        Case "����"
            lng���� = GetRptRsColumn("����")
            '������ˢ��
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
            strErrMsg = ""
            If blnNext = False Then
                If blnCard Then
                    If IDKIND.Cards.��ȱʡ������ And Not IDKIND.GetfaultCard Is Nothing Then
                        lng�����ID = IDKIND.GetfaultCard.�ӿ����
                    Else
                        lng�����ID = "-1"
                    End If
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strValue, True, lng����ID, "", strErrMsg, lng�����ID) = False Then lng����ID = 0
                    txtFind.Tag = lng����ID
                    lng����ID = Val(txtFind.Tag)
                    strKind = "����ID"
                    lng����IDCol = GetRptRsColumn("����ID")
                End If
            Else
                 lng����ID = Val(txtFind.Tag)
                 If lng����ID <> 0 Then strKind = "����ID": lng����IDCol = GetRptRsColumn("����ID")
           End If
        Case "ҽ����"
            lngҽ���� = GetRptRsColumn("ҽ����")
        Case Else ' "���￨"
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
            strErrMsg = ""
            If blnNext = False Then
                    lng�����ID = objCard.�ӿ����
                    If lng�����ID > 0 Then
                        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strValue, True, lng����ID, "", strErrMsg) = False Then GoTo NotFoundPati:
                        If lng����ID = 0 Then GoTo NotFoundPati:
                    Else
                        If gobjSquare.objSquareCard.zlGetPatiID(strKind, strValue, True, lng����ID, _
                            "", strErrMsg) = False Then GoTo NotFoundPati:
                    End If
                    If lng����ID <= 0 Then GoTo NotFoundPati:
                    txtFind.Tag = lng����ID
            End If
            lng����ID = Val(txtFind.Tag)
            strKind = "����ID"
            lng����IDCol = GetRptRsColumn("����ID")
    End Select
    '���Ҳ���
    For i = lngPoint To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            Select Case strKind
                Case "����"
                    If Trim(rptPati.Rows(i).Record(lng����).Value) = strValue Then Exit For
                Case "סԺ��"
                    If rptPati.Rows(i).Record(lngסԺ��).Value = strValue Then Exit For
                Case "����"
                    If rptPati.Rows(i).Record(lng����).Value Like IIf(gstrLike = "%", "*", "") & strValue & "*" Then Exit For
                Case "ҽ����"
                    If rptPati.Rows(i).Record(lngҽ����).Value = strValue Then Exit For
                Case Else
                    If rptPati.Rows(i).Record(lng����IDCol).Value = lng����ID Then Exit For
            End Select
        End If
    Next
    If i = rptPati.Rows.Count Then GoTo NotFoundPati:
    Set rptPati.FocusedRow = rptPati.Rows(i)    '����SelectionChanged�¼�
    Exit Sub
NotFoundPati:
    
    If blnNext Then
        MsgBox "�Ѿ�û�з������������Ĳ��ˣ�", vbInformation, gstrSysName
    Else
        If strErrMsg <> "" Then
            MsgBox strErrMsg, vbInformation, gstrSysName
        Else
            MsgBox "û���ҵ��������������Ĳ��ˣ�", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub chkԤ����_Click()
    txtԤ��.Enabled = chkԤ����.Value = 1
End Sub

Private Sub chkԤ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtԤ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtԤ��_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtԤ��, KeyAscii, m�����ʽ
End Sub
Private Sub zlSchemeSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '����:���˺�
    '����:2011-01-20 09:58:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    frmPatiPressMoneySet.zlShowMe Me, mlngModul, mstrPrivs
End Sub

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2011-01-31 14:22:25
    '����:35550
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    'װ������
    Call zlRptControlToVsGrid(rptPati, vsPrint)
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "�����嵥"
    objRow.Add "���˲�����" & cboUnit.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "����״̬��" & cboState.Text
    objRow.Add "��ʼ���ڣ�" & Format(dtpBegin.Value, "yyyy-mm-dd")
    objRow.Add "�������ڣ�" & Format(dtpEnd.Value, "yyyy-mm-dd")
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    bytPrn = bytFunc
    Err = 0: On Error GoTo ErrHand:
    
    Set objPrint.Body = vsPrint
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
 Private Function ExecPrePayMoney() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�н�Ԥ����
    '����:ִ�гɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2011-02-17 15:16:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim dbl�ɿ�� As Double, bln�������۲��� As Boolean
    
    On Error GoTo errHandle
    If mobjPatient Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjPatient = CreateObject("zl9Patient.clsPatient")
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   ���˹�����������,���ܽ�Ԥ��,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Err = 0
            Exit Function
        End If
    End If
    
    bln�������۲��� = ZlIsOutpatientObserve(mlng����ID, mlng��ҳID)
    strSQL = "Select Zl1_Getdef_Prepaymoney([1],[2],[3]) as �ɿ�� from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, IIf(bln�������۲���, 1, 2))
    dbl�ɿ�� = Nvl(rsTemp!�ɿ��, 0)
    
    '�������۲��˽�����Ԥ��
    'PlusDeposit(ByVal lngSys As Long, cnMain As ADODB.Connection, frmMain As Object, _
    '    ByVal strDBUser As String, Optional bytCallObject As Byte = 0, _
    '    Optional lng����ID As Long, Optional lng��ҳID As Long, Optional dblDefPrePayMoney As Double = 0) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '���ܣ� ����Ԥ�����տ��
    '    '������
    '    '   lngModul:��Ҫִ�еĹ������
    '    '   cnMain:����������ݿ�����
    '    '   frmMain:������
    '    '   strDBUser:��ǰ���ݿ��¼�û���
    '    '  bytCallObject:���˺����(0-Ԥ�������(ȱʡ��);1-���˷��ò�ѯ����)
    '    '  lng����ID-ȱʡ�Ĳ���ID
    '    '  lng��ҳID-ȱʡ����ҳID
    '    '  dblDefPrePayMoney-ȱʡ��Ԥ�����
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    ExecPrePayMoney = mobjPatient.PlusDeposit(glngSys, gcnOracle, Me, _
        gstrDBUser, 1, mlng����ID, IIf(bln�������۲���, 0, mlng��ҳID), dbl�ɿ��)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Private Sub InitMenus()
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKIND.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtFind)
    Call IDKindPati.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "��|����|0|0|0|0|0|0;��|����|0|0|0|0|0|0;ҽ|ҽ����|1|0|0|0|0|0", txt����)
    
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKIND.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKIND.Cards.��ȱʡ������
End Sub

Private Sub ModeInsurePatiDisease()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ĳ��˲���
    '����:���˺�
    '����:2013-02-20 11:41:07
    '����:31883
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String
    'ChooseDisease(ByVal lngPatiID As Long, ByVal lngPageID As Long, Optional ByVal intInsure As Integer = 0, _
    Optional ByRef strAdvance As String = "")
    If mlng����ID = 0 Or mintInsure = 0 Then Exit Sub
    Call gclsInsure.ChooseDisease(mlng����ID, mlng��ҳID, mintInsure, strAdvance)
End Sub




