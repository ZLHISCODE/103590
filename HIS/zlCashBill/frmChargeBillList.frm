VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmChargeBillList 
   Caption         =   "�տ���ϸ����"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChargeBillList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picConList 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   13065
      TabIndex        =   9
      Top             =   510
      Width           =   13065
      Begin VB.Frame fraSplit 
         Height          =   30
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   13575
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "���ݺ�:"
         Height          =   210
         Left            =   75
         TabIndex        =   12
         Top             =   150
         Width           =   735
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "ʱ�䷶Χ:"
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   150
         Width           =   945
      End
   End
   Begin VB.PictureBox picFeeList 
      BorderStyle     =   0  'None
      Height          =   7140
      Left            =   9960
      ScaleHeight     =   7140
      ScaleWidth      =   3015
      TabIndex        =   7
      Top             =   1440
      Width           =   3015
      Begin VSFlex8Ctl.VSFlexGrid vsFeeList 
         Height          =   1800
         Left            =   180
         TabIndex        =   8
         Top             =   600
         Width           =   2475
         _cx             =   4366
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillList.frx":0502
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   585
      ScaleHeight     =   3030
      ScaleWidth      =   9525
      TabIndex        =   4
      Top             =   1470
      Width           =   9525
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   1800
         Left            =   330
         TabIndex        =   5
         Top             =   645
         Width           =   8505
         _cx             =   15002
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillList.frx":057C
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.PictureBox picBalance 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   930
      ScaleHeight     =   2925
      ScaleWidth      =   4485
      TabIndex        =   2
      Top             =   3900
      Width           =   4485
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   870
         Left            =   150
         TabIndex        =   3
         Top             =   90
         Width           =   1860
         _cx             =   3281
         _cy             =   1535
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillList.frx":05F6
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.PictureBox picBillList 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   5790
      ScaleHeight     =   2685
      ScaleWidth      =   3630
      TabIndex        =   0
      Top             =   4080
      Width           =   3630
      Begin VSFlex8Ctl.VSFlexGrid vsBillList 
         Height          =   1800
         Left            =   585
         TabIndex        =   1
         Top             =   480
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillList.frx":0670
         ScrollTrack     =   -1  'True
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   7935
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   635
      SimpleText      =   $"frmChargeBillList.frx":06EA
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeBillList.frx":0731
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16272
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "���˺�"
            TextSave        =   "���˺�"
            Object.ToolTipText     =   "��ǰ����Ա:���˺�"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   360
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmChargeBillList.frx":0FC5
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeBillList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long, mstrPrivs As String
Private Enum TotalType
    EM_�շ�Ա���� = 1
    EM_С���տ� = 2
    EM_С������ = 3
    EM_�����տ� = 4
    EM_�����տ�_���շ�Ա = 5
End Enum

Private mstrCashBalance As String '�ֽ���㷽ʽ
Private mbytType As TotalType '1-�շ�Ա���ʣ�2-С���տ�;3-С������;4-�����տ�
Private mstrChargeRollingID As String '����ID���տ�ID(����mbytType)������,���ʱ,�ö��ŷָ�,��:123,23,11
Private mdtStartDate As Date, mdtendDate As Date '���ʵĿ�ʼʱ������ʱ��
Private mblnDel As Boolean '�Ƿ����ϼ�¼
Private Enum mPaneIndex
    EM_PN_ConList = 260101  '����
    EM_PN_LIST = 260102  '�տ����
    EM_PN_BALANCE = 260103    '���㷽ʽ
    EM_PN_BILL = 260104  '�˷�Ʊ��
    EM_PN_FeeLIST = 260105  '���û���
End Enum
Private mbytFontSize As Byte
Private mbytƱ�ݷ������ As Byte    'Ʊ�ݷ������:0-����ʵ�ʴ�ӡ����Ʊ��;1-����ϵͳԤ���������;2-�����û��Զ���������
Private mblnNotBrush As Boolean '��ˢ������
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mstrPersonName As String
Private mstrRollingType As String '�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�)
Private mblnFirst As Boolean

Public Sub ShowMe(ByVal frmMain As Object, _
      ByVal lngModule As Long, ByVal strPrivs As String, _
      ByVal bytType As Byte, ByVal strChargeRollingID As String, _
      Optional ByVal dtStartDate As Date, Optional ByVal dtEndDate As Date, _
      Optional ByVal blnDel As Boolean = False, Optional strPersonName As String, _
      Optional strRollingType As String)
    '-------------------------------------------------------------------------------------------------
    '����:�������,��ʾָ�����ʻ��տ��¼����ϸ����
    '���:frmMain-���õ�������
    '    lngModule-ģ���
    '    strPrivs-Ȩ�޴�
    '����bytType:1-�շ�Ա���ʣ�2-С���տ�;3-С������;4-�����տ
    '    strChargeRollingID -����ID���տ�ID(���ʱ,�ö��ŷָ�,��:123,23,11)
    '    dtStartDate-��ѡ����,��ʼ����ʱ��,strChargeRollingID=0ʱ�����봫��
    '    dtEndDate-��ѡ��������������ʱ��,strChargeRollingID=0ʱ�����봫��
    '    blnDel-�Ƿ����ϼ�¼
    '    strPersonName-ָ�����շ�Ա(Ϊ��ʱ,Ϊ��ǰ����Ա)
    '    strRollingType-�������,bytType=1ʱ��Ч�ֱ�Ϊ:
    '               0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�
    '����:���˺�
    '����:2013-09-16 10:08:39
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytType = bytType:  mstrChargeRollingID = strChargeRollingID
    mdtStartDate = dtStartDate: mdtendDate = dtEndDate: mblnDel = blnDel
    mstrPersonName = IIf(strPersonName = "", UserInfo.����, strPersonName)
    mstrRollingType = strRollingType: mblnFirst = True
    Call InitFace: Call zlDefCommandBars
    Err = 0: On Error Resume Next
    Call zlRefresh
    Me.Show 1, frmMain
End Sub
Private Function ReadListData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ص���ϸ����
    '����:���ݻ�ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-16 10:12:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Call LoadFilterRange
    If mstrChargeRollingID = "" And (mbytType = EM_�շ�Ա���� Or mbytType = EM_�����տ�_���շ�Ա) Then
         ReadListData = LoadPersonList         '�����շ�Ա������ϸ����
         Call LoadFeeData(True)   '���ط�����Ϣ
    Else
         ReadListData = LoadList          '������ص���ϸ����
    End If
    Call picBalance_Resize
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Private Function LoadPersonList() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�Ա��ǰ���ʵ���ϸ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-16 10:13:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, bytType As Byte
    Dim strWithTable As String
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim lngRow As Long, lngNo As Long
    Dim mstrCashBalance As String, str���㷽ʽ As String, strTemp As String, intƱ�� As Integer
    Dim dblToTotal As Double, strWhere As String, strRollingType As String
    
    On Error GoTo errHandle
    mblnNotBrush = True
    bytType = 1: dblToTotal = 0
    '���￨
    strWhere = " And instr([4],','||A.��������||',')>0 "
    If mstrRollingType <> "" And InStr("," & mstrRollingType & ",", ",0,") = 0 Then
        If Get���ʽ�������(mstrRollingType, strRollingType) = False Then
            Call InitGrid: LoadPersonList = False
            Exit Function
        End If
    Else
        strRollingType = ",2,3,4,5,6,"
    End If
    'Ԥ������NULL,2-����,3-�շ�,4-�Һ�,5-���￨,6-����ҽ������
    'mstrRollingType:�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�)
    '�սɶ���:1-�շ�(���Һ�),2-����,3-Ԥ��,4-���;5-���ѿ���ֵ;6--���ѿ���ֵ;7-�ݴ��(��������)�������տ���������϶��գ��������ӣ���9-���ν����¼
    strWithTable = "" & _
    "   With �������� as (" & _
    "   Select   A.��������,A.����ID,sum(A.��Ԥ��) as ��Ԥ��" & vbCrLf & _
    "   From ����Ԥ����¼ A" & vbCrLf & _
    "   Where a.����Ա���� || '' = [1]  and a.��¼����<>1 " & strWhere & vbCrLf & _
    "        And A.�տ�ʱ�� Between  [2] And [3] " & vbCrLf & _
    "   Group by  A.��������,A.����ID)" & vbCrLf

    strSQL = ""
    '��¼����:1,'�շ�',4,'�Һ�',5,'���￨',6,'����',7,'Ԥ����',8,'���ѿ���ֵ',9,'���ѿ����',10,'���',11,'���',12,'�ݴ�',13,'���ν���'
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",5,") > 0 Then  '���￨
        strSQL = "" & _
        "Select 5 AS ��¼����,A.����ID as ��¼ID,A.NO,A.��¼״̬,max(A.����) AS ��������,max(A.�Ա�) AS �Ա�,max(A.����) AS ����,max(decode(A.�����־,2,-1*NULL,A.��ʶ��)) AS �����,max(decode(A.�����־,2,��ʶ��,null)) AS סԺ��, " & vbCrLf & _
        "     Max(A.����Ա����) AS ����Ա����, NULL AS ���㷽ʽ, " & _
        "     sum(A.���ʽ��) AS ���ϼ�,Max(�Ǽ�ʱ��) AS �շ�ʱ�� " & vbCrLf & _
        "   From סԺ���ü�¼ A,�������� B " & vbCrLf & _
        "   Where A.����ID=B.����ID And B.��������=5 " & vbCrLf & _
        "       And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.����id And x.�տ�Ա = [1] And x.Id = y.�ս�ID And y.���� = 1 And x.����ʱ�� Is Null) " & vbCrLf & _
        "   GROUP BY A.NO,A.����ID,A.��¼״̬ " & vbCrLf
    End If
    

    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",1,") > 0 Or InStr("," & mstrRollingType & ",", ",4,") > 0 Then '���շ�,�Һ��￨,������
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select ��¼����, Min(��¼id) As ��¼id, NO, Max(��¼״̬) As ��¼״̬, ��������, �Ա�, ����, �����, סԺ��, ����Ա����, ���㷽ʽ, Sum(���ϼ�) As ���ϼ�, �շ�ʱ�� From ( " & _
        "   Select  MOD(A.��¼����,10) As ��¼����,A.����ID as ��¼ID,A.NO,A.��¼״̬,max(A.����) AS ��������,max(A.�Ա�) AS �Ա�,max(A.����) AS ����,max(decode(A.�����־,2,-1*NULL,A.��ʶ��)) AS �����,max(decode(A.�����־,2,��ʶ��,null)) AS סԺ��, " & vbCrLf & _
        "     Max(A.����Ա����) AS ����Ա����, NULL AS ���㷽ʽ, " & vbCrLf & _
        "     sum(A.���ʽ��) AS ���ϼ�,Max(A.�Ǽ�ʱ��) AS �շ�ʱ�� " & vbCrLf & _
        "   From ������ü�¼ A,�������� B " & vbCrLf & _
        "   Where A.����ID=B.����ID And B.�������� in (3,4) And Nvl(A.����״̬, 0)=0 " & _
        "       And Not Exists(Select y.��¼id  From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.����id And x.�տ�Ա = [1] And x.Id = y.�ս�ID And y.���� = 1 And x.����ʱ�� Is Null)     " & _
        "   GROUP BY A.��¼����,A.����ID,A.NO,A.��¼״̬ ) " & _
        "   Group By ��¼����, NO, ��������, �Ա�, ����, �����, סԺ��, ����Ա����, ���㷽ʽ, �շ�ʱ�� " & _
        "   Union ALL" & vbCrLf & _
        "   Select 13 as ��¼����, a.����id As ��¼id, a.No, a.��¼״̬, Max(c.����) As ��������, Max(c.�Ա�) As �Ա�, Max(c.����) As ����, c.����� As �����," & vbNewLine & _
        "        c.סԺ�� As סԺ��, Max(a.����Ա����) As ����Ա����, Null As ���㷽ʽ, Sum(B.��Ԥ��) As ���ϼ�, Max(a.�Ǽ�ʱ��) As �շ�ʱ��" & vbNewLine & _
        "   From ���ò����¼ A, �������� B,������Ϣ C" & vbNewLine & _
        "   Where A.����ID=B.����ID and A.����ID=C.����ID And B.��������=6 And a.��¼���� = 1 And Nvl(A.����״̬, 0)=0  " & _
        "           And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.����id  And x.Id = y.�ս�id And y.���� = 9 And x.����ʱ�� Is Null)" & vbNewLine & _
        "   Group By a.��¼����, a.����id, a.No, a.��¼״̬, c.�����, c.סԺ��"
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",3,") > 0 Then '����
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 6 AS ��¼����,A.ID as ��¼ID,A.NO,A.��¼״̬,max(C.����) AS ��������,max(C.�Ա�) AS �Ա�,max(C.����) AS ����, " & vbCrLf & _
        "     max(c.�����) AS �����,max(C.סԺ��) AS סԺ��, " & vbCrLf & _
        "     Max(A.����Ա����) AS ����Ա����,NULL AS ���㷽ʽ, " & vbCrLf & _
        "     sum(b.��Ԥ��) AS ���ϼ�,Max(A.�շ�ʱ��) AS �շ�ʱ�� " & vbCrLf & _
        "   From ���˽��ʼ�¼ A,�������� B,������Ϣ C " & vbCrLf & _
        "   Where A.ID=B.����ID  And B.��������=2 And A.����ID=C.����ID(+) And nvl(A.����״̬,0)=0 " & vbCrLf & _
        "         And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.Id And x.�տ�Ա = [1] And x.Id = y.�ս�ID And y.���� = 2 And x.����ʱ�� Is Null) " & vbCrLf & _
        "   Group by A.NO,A.ID,A.��¼״̬ " & vbCrLf
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",2,") > 0 Then 'Ԥ����(��ֵ)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 7 As ��¼����,A.ID as ��¼ID, a.No, a.��¼״̬, Max(nvl(M.����,c.����)) As ��������, Max(c.�Ա�) As �Ա�, Max(c.����) As ����, Max(c.�����) As �����, " & vbCrLf & _
        "      Max(Decode(m.סԺ��, Null, c.סԺ��, m.סԺ��)) As סԺ��, Max(a.����Ա����) As ����Ա����,Max(���㷽ʽ) AS ���㷽ʽ, Sum(a.���) As ���ϼ�, Max(a.�տ�ʱ��) As �շ�ʱ�� " & vbCrLf & _
        "   From ����Ԥ����¼ A, ������Ϣ C, ������ҳ M " & vbCrLf & _
        "   Where a.����id = c.����id(+) And a.����id = m.����id(+) And a.��ҳid = m.��ҳid(+)  " & vbCrLf & _
        "     And a.��¼���� = 1 And Nvl(a.��������,0) <> 12 And a.����Ա���� || '' = [1]  " & vbCrLf & _
        "     And a.�տ�ʱ��  between [2] And [3]  " & vbCrLf & _
        "     And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.Id And x.�տ�Ա = [1] And x.Id = y.�ս�id And y.���� = 3 And x.����ʱ�� Is Null) " & vbCrLf & _
        "   Group By a.No,A.ID, a.��¼״̬ "
    End If
    
    If InStr("," & mstrRollingType & ",", ",21,") > 0 Then '����Ԥ����(��ֵ)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 7 As ��¼����,A.ID as ��¼ID, a.No, a.��¼״̬, Max(nvl(M.����,c.����)) As ��������, Max(c.�Ա�) As �Ա�, Max(c.����) As ����, Max(c.�����) As �����, " & vbCrLf & _
        "      Max(Decode(m.סԺ��, Null, c.סԺ��, m.סԺ��)) As סԺ��, Max(a.����Ա����) As ����Ա����,Max(���㷽ʽ) AS ���㷽ʽ, Sum(a.���) As ���ϼ�, Max(a.�տ�ʱ��) As �շ�ʱ�� " & vbCrLf & _
        "   From ����Ԥ����¼ A, ������Ϣ C, ������ҳ M " & vbCrLf & _
        "   Where a.����id = c.����id(+) And a.����id = m.����id(+) And a.��ҳid = m.��ҳid(+)  " & vbCrLf & _
        "     And a.��¼���� = 1 And Nvl(a.Ԥ�����,0) = 1 And Nvl(a.��������,0) <> 12 And a.����Ա���� || '' = [1]  " & vbCrLf & _
        "     And a.�տ�ʱ��  between [2] And [3]  " & vbCrLf & _
        "     And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.Id And x.�տ�Ա = [1] And x.Id = y.�ս�id And y.���� = 3 And x.����ʱ�� Is Null) " & vbCrLf & _
        "   Group By a.No,A.ID, a.��¼״̬ "
    End If
    
    If InStr("," & mstrRollingType & ",", ",22,") > 0 Then 'סԺԤ����(��ֵ)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 7 As ��¼����,A.ID as ��¼ID, a.No, a.��¼״̬, Max(nvl(M.����,c.����)) As ��������, Max(c.�Ա�) As �Ա�, Max(c.����) As ����, Max(c.�����) As �����, " & vbCrLf & _
        "      Max(Decode(m.סԺ��, Null, c.סԺ��, m.סԺ��)) As סԺ��, Max(a.����Ա����) As ����Ա����,Max(���㷽ʽ) AS ���㷽ʽ, Sum(a.���) As ���ϼ�, Max(a.�տ�ʱ��) As �շ�ʱ�� " & vbCrLf & _
        "   From ����Ԥ����¼ A, ������Ϣ C, ������ҳ M " & vbCrLf & _
        "   Where a.����id = c.����id(+) And a.����id = m.����id(+) And a.��ҳid = m.��ҳid(+)  " & vbCrLf & _
        "     And a.��¼���� = 1 And Nvl(a.Ԥ�����,0) = 2 And Nvl(a.��������,0) <> 12 And a.����Ա���� || '' = [1]  " & vbCrLf & _
        "     And a.�տ�ʱ��  between [2] And [3]  " & vbCrLf & _
        "     And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.Id And x.�տ�Ա = [1] And x.Id = y.�ս�id And y.���� = 3 And x.����ʱ�� Is Null) " & vbCrLf & _
        "   Group By a.No,A.ID, a.��¼״̬ "
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",3,") > 0 Then '����(��ֵ)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 7 As ��¼����,A.ID as ��¼ID, a.No, a.��¼״̬, Max(nvl(M.����,c.����)) As ��������, Max(c.�Ա�) As �Ա�, Max(c.����) As ����, Max(c.�����) As �����, " & vbCrLf & _
        "      Max(Decode(m.סԺ��, Null, c.סԺ��, m.סԺ��)) As סԺ��, Max(a.����Ա����) As ����Ա����,Max(���㷽ʽ) AS ���㷽ʽ, Sum(a.���) As ���ϼ�, Max(a.�տ�ʱ��) As �շ�ʱ�� " & vbCrLf & _
        "   From ����Ԥ����¼ A, ������Ϣ C, ������ҳ M " & vbCrLf & _
        "   Where a.����id = c.����id(+) And a.����id = m.����id(+) And a.��ҳid = m.��ҳid(+)  " & vbCrLf & _
        "     And a.��¼���� = 1 And Nvl(a.��������,0) = 12  And a.����Ա���� || '' = [1]  " & vbCrLf & _
        "     And a.�տ�ʱ��  between [2] And [3]  " & vbCrLf & _
        "     And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.Id And x.�տ�Ա = [1] And x.Id = y.�ս�id And y.���� = 3 And x.����ʱ�� Is Null) " & vbCrLf & _
        "   Group By a.No,A.ID, a.��¼״̬ "
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",6,") > 0 Then '���ѿ�(��ֵ,��ֵ)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 8 as ��¼����,A.����ID as ��¼ID, a.���� As no,a.��¼״̬,'' as ��������,'' as �Ա�,'' as ����,null �����,NULL as סԺ��, " & vbCrLf & _
        "        A.����Ա���� AS ����Ա����,a.���㷽ʽ, a.ʵ�ս��  As ���ϼ�, A.�Ǽ�ʱ�� AS �շ�ʱ��  " & vbCrLf & _
        "   From ���˿������¼ A, ���˿������¼ B " & vbCrLf & _
        "   Where a.������� = b.�������(+) And a.���ѿ�id = b.���ѿ�id(+) And (a.��¼���� = 2 Or a.��¼���� = 3 And b.��¼���� = 2) And b.��¼����(+) = 2 " & vbCrLf & _
        "         And a.Id <> b.Id(+) And a.����Ա���� || '' = [1] And a.�Ǽ�ʱ�� Between [2] And [3]  " & vbCrLf & _
        "         And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.����Id And x.�տ�Ա = [1] And x.Id = y.�ս�ID And y.���� = 5 And x.����ʱ�� Is Null)" & vbCrLf & _
        "   Union All " & vbCrLf & _
        "   Select 9 As ��¼����,A.����ID as ��¼ID, a.���� As NO, a.��¼״̬, '' As ��������, '' As �Ա�, '' As ����, Null �����, " & vbCrLf & _
        "       Null As סԺ��, a.����Ա����, a.���㷽ʽ, a.ʵ�ս�� As ���ϼ�, a.�Ǽ�ʱ�� " & vbCrLf & _
        "   From ���˿������¼ A, ���˿������¼ B" & vbCrLf & _
        "   Where a.������� = b.�������(+) And a.���ѿ�id = b.���ѿ�id(+) And (a.��¼���� = 1 Or a.��¼���� = 3 And b.��¼���� = 1) And b.��¼����(+) = 1 " & vbCrLf & _
        "         And a.Id <> b.Id(+) And a.����Ա���� || '' = [1] And a.�Ǽ�ʱ�� Between [2] And [3] " & vbCrLf & _
        "         And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.����Id And x.�տ�Ա = [1] And x.Id = y.�ս�id And y.���� = 6 And x.����ʱ�� Is Null) " & vbCrLf
    End If
    
    '���ͽ��
    strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
    "   Select 10 As ��¼����,A.ID as ��¼ID, ltrIm(to_Char(a.ID)) As NO, 1 As ��¼״̬, '' As ��������, '' As �Ա�, '' As ����, Null �����, " & _
    "     Null As סԺ��, a.����� As ����Ա����, a.���㷽ʽ, a.����� As ���, a.���ʱ�� " & _
    "   From ��Ա����¼ A " & _
    "   Where a.����� || '' = [1]   And a.ȡ��ʱ�� Is Null  " & _
    "     And a.���ʱ�� Between [2] And [3] " & _
    "     And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.Id And x.�տ�Ա = [1] And x.Id = y.�ս�ID And y.���� = 4 And x.����ʱ�� Is Null)  " & _
    "        Union All " & _
    "   Select 11 As ��¼����,A.ID as ��¼ID,ltrIm( to_Char(a.ID)) As NO, 1 As ��¼״̬, '' As ��������, '' As �Ա�, '' As ����, Null �����, " & _
    "          Null As סԺ��, a.����� As ����Ա����, a.���㷽ʽ, a.����� As ���, a.���ʱ�� " & _
    "   From ��Ա����¼ A " & _
    "   Where a.����� || '' = [1]   And a.ȡ��ʱ�� Is Null  " & _
    "     And a.���ʱ�� Between [2] And [3] " & _
    "     And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.Id And x.�տ�Ա = [1] And x.Id = y.�ս�ID And y.���� = 4 And x.����ʱ�� Is Null)  "
    
    '�ݴ��
    strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
    "  Select 12 As ��¼����,A.ID as ��¼ID, a.No As NO, Decode(�ջ�ʱ��, Null, 1, 2) As ��¼״̬, '' As ��������, '' As �Ա�, '' As ����, Null �����, Null As סԺ��, " & vbCrLf & _
    "      a.�Ǽ��� As ����Ա����, '" & mstrCashBalance & "' As ���㷽ʽ, a.���, a.�Ǽ�ʱ�� " & _
    "   From ��Ա�ݴ��¼ A " & vbCrLf & _
    "   Where �տ�Ա || '' = [1] And �Ǽ�ʱ�� > [2] And  �Ǽ�ʱ�� <= [3]  " & vbCrLf & _
    "     And A.��¼����=2 And A.�ջ�ʱ�� Is Null" & vbCrLf & _
    "     And Not Exists (Select y.��¼id  From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y  Where y.��¼id = a.Id And x.�տ�Ա = [1] And x.Id = y.�ս�id And y.���� = 7 And x.����ʱ�� Is Null)" & vbCrLf
    
    If InStr(strSQL, "��������") > 0 Then
        strSQL = strWithTable & strSQL
    End If
    
    strSQL = "" & _
    "   SELECT /*+ rule */ ��¼���� As ����,��¼ID,decode(��¼����,1,'�շ�',4,'�Һ�',5,'���￨',6,'����',7,'Ԥ����',8,'���ѿ���ֵ',9,'���ѿ����',10,'���',11,'���',12,'�ݴ�',13,'������','') AS ���, " & _
    "        NO,��¼״̬,��������,�Ա�,����,�����,סԺ��,����Ա����,���㷽ʽ," & _
    "       Trim(to_char(���ϼ�,'99999999990.00')) As ���ϼ�,to_char(�շ�ʱ��,'yyyy-mm-dd hh24:mi:ss')  as �շ�ʱ�� " & _
    "   FROM ( " & strSQL & " ) " & _
    "   ORDER BY ��¼����,�շ�ʱ�� DESC,NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrPersonName, mdtStartDate, mdtendDate, strRollingType)
    
    With vsList
        .Clear 1: .Rows = 2: .FixedRows = 1
        If Not rsTemp.EOF Then
            Set .DataSource = rsTemp
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            Select Case .ColKey(i)
            Case "����", "��¼״̬", "��¼ID"
                .ColWidth(i) = 0: .ColHidden(i) = True
            Case "NO", "���", "�Ա�", "����", "�����", "סԺ��", "���㷽ʽ", "�շ�ʱ��"
                .ColAlignment(i) = flexAlignCenterCenter
            Case "���ϼ�"
                .ColAlignment(i) = flexAlignRightCenter
            Case Else
                .ColAlignment(i) = flexAlignLeftCenter
            End Select
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
        If rsTemp.RecordCount <> 0 Then
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, -1, .ColIndex("���ϼ�"), "#######0.00", &HFFC0C0, vbBlack, True, "�ϼ�"
            For i = 0 To .ColIndex("���ϼ�") - 1
                .TextMatrix(.Rows - 1, i) = "�ϼ�"
            Next
            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            .MergeRow(.Rows - 1) = True
            .MergeCells = flexMergeRestrictRows
           '�����:110535,����,2017/08/10,����ɫ�����˷Ѽ�¼�ͱ��˷Ѽ�¼
            For i = 1 To .Rows - 1
                Select Case .TextMatrix(i, .ColIndex("��¼״̬"))
                Case 1
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                Case 2
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                Case 3
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                End Select
            Next
        End If
    End With
    zl_vsGrid_Para_Restore mlngMode, vsList, Me.Name, "��ϸ��Ϣ�б�", False
    Call LoadDetailData
    mblnNotBrush = False
    LoadPersonList = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotBrush = False
End Function

Private Sub LoadDetailData()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϸ����
    '����:���˺�
    '����:2013-09-16 14:43:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, bytType As Byte, intRecordSta As Integer
    Dim lng��¼ID As Long, rsTemp As ADODB.Recordset, lngRow As Long
    Dim blnNOMoved As Boolean, bytƱ�� As Byte, strSQL As String
    
    On Error GoTo errHandle
    With vsList
        If .Row < 1 Or .Col < 0 Then GoTo GoClear:
        strNO = Trim(.TextMatrix(.Row, .ColIndex("NO")))
        bytType = Val(.TextMatrix(.Row, .ColIndex("����")))
        intRecordSta = Val(.TextMatrix(.Row, .ColIndex("��¼״̬")))
        lng��¼ID = Val(.TextMatrix(.Row, .ColIndex("��¼ID")))
        If strNO = "" Or lng��¼ID = 0 Then
            If ShowOnlyFactList Then
                vsBalance.Clear 1: vsBalance.Rows = 2
                Exit Sub
            End If
            GoTo GoClear:
        End If
    End With
    '���ؽ��㷽ʽ��Ϣ
    'decode(��¼����,1,'�շ�',4,'�Һ�',5,'���￨',6,'����',7,'Ԥ����',8,'���ѿ���ֵ',9,'���ѿ����',10,'���',11,'���',12,'�ݴ�',13,'���ν���')
    Select Case bytType
    Case 1, 4, 5, 6, 7, 13
        '�շ�,�Һ�,���￨,����
        If bytType = 1 Then
            blnNOMoved = zlDatabase.NOMoved("������ü�¼", strNO, , "1")
            bytƱ�� = 1
        ElseIf bytType = 4 Then
            blnNOMoved = zlDatabase.NOMoved("������ü�¼", strNO, , "4")
            bytƱ�� = 4
        ElseIf bytType = 5 Then
            blnNOMoved = zlDatabase.NOMoved("סԺ���ü�¼", strNO, , "5")
            bytƱ�� = 5
        ElseIf bytType = 6 Then
            blnNOMoved = zlDatabase.NOMoved("���˽��ʼ�¼", strNO)
            bytƱ�� = 3
        ElseIf bytType = 7 Then
            blnNOMoved = zlDatabase.NOMoved("����Ԥ����¼", strNO, , "1")
            bytƱ�� = 2
        ElseIf bytType = 13 Then
            blnNOMoved = zlDatabase.NOMoved("���ò����¼", strNO, , "1")
            bytƱ�� = 1
        Else
            blnNOMoved = False
        End If
        If bytƱ�� = 2 Then
            strSQL = " " & _
             " Select  A.���㷽ʽ, a.���, A.�������, A.����, A.������ˮ��, A.����˵�� " & _
             " From " & IIf(blnNOMoved, "H", "") & "����Ԥ����¼ A,���㷽ʽ B " & _
             " Where a.ID=[1] And a.���㷽ʽ=B.����(+) " & _
             " Order by  decode(nvl(B.����,0),1,1,2,2,3,10,4,11,4) ,A.���㷽ʽ"
        Else
            If bytType = 1 Then
                strSQL = "Select Decode(Mod(��¼����, 10), 1, '[��Ԥ����]', A.���㷽ʽ) As ���㷽ʽ," & vbNewLine & _
                "                   A.��Ԥ�� As ���, A.�������, A.����,A. ������ˮ��,A.����˵��" & vbNewLine & _
                "              From ����Ԥ����¼ A,���㷽ʽ B" & vbNewLine & _
                "              Where a.�������=[2] And a.���㷽ʽ=B.����(+)" & vbNewLine & _
                "              Order by decode(Mod(��¼����, 10), 0,decode(nvl(B.����,0),1,1,2,2,3,10,4,11,4)) ,���㷽ʽ"
            Else
                strSQL = " " & _
                 " Select Decode(Mod(��¼����, 10), 1, '[��Ԥ����]', A.���㷽ʽ) As ���㷽ʽ,  " & _
                 "      A.��Ԥ�� As ���, A.�������, A.����,A. ������ˮ��,A.����˵�� " & _
                 " From " & IIf(blnNOMoved, "H", "") & "����Ԥ����¼ A,���㷽ʽ B " & _
                 " Where a.����ID=[1] And a.���㷽ʽ=B.����(+) " & _
                 " Order by decode(Mod(��¼����, 10), 0,decode(nvl(B.����,0),1,1,2,2,3,10,4,11,4)) ,���㷽ʽ"
            End If
        End If
        strSQL = "" & _
        "   Select decode(nvl(b.����,0),0,0,3 ,12,4,12,  b.���� ) as ����,A.���㷽ʽ,sum(A.���) as ���, " & _
        "          max(a.�������) as �������,max(a.����) as ����,max(a.������ˮ��) as ������ˮ��,max(a.����˵��) as ����˵��" & _
        "   From (" & strSQL & ") A,���㷽ʽ B " & _
        "   Where A.���㷽ʽ=b.����(+)" & _
        "   Group by A.���㷽ʽ,decode(nvl(b.����,0),0,0,3 ,12,4,12,  b.���� )" & _
        "   Order by ����,���㷽ʽ"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, -1 * lng��¼ID)
        With vsBalance
            .Clear 1
            .Rows = rsTemp.RecordCount + IIf(rsTemp.RecordCount = 0, 2, 1)
            lngRow = 1
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = Nvl(rsTemp!���㷽ʽ)
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(rsTemp!���)), "###0.00;-###0.00; ;")
                .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(rsTemp!�������)
                .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                .TextMatrix(lngRow, .ColIndex("������ˮ��")) = Nvl(rsTemp!������ˮ��)
                .TextMatrix(lngRow, .ColIndex("����˵��")) = Nvl(rsTemp!����˵��)
                lngRow = lngRow + 1
                rsTemp.MoveNext
            Loop
            .AutoSizeMode = flexAutoSizeColWidth
            Call .AutoSize(0, .Cols - 1)
            zl_vsGrid_Para_Restore mlngMode, vsBalance, Me.Name, "������Ϣ�б�", False
        End With
        'Ʊ��ʹ�����
        'Ʊ��:1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
        Call ShowFactList(bytƱ��, strNO, blnNOMoved)
    Case 8, 9, 10, 11, 12
        ' 8,'���ѿ���ֵ',9,'���ѿ����',10,'���',11,'���',12,'�ݴ�'
        With vsBalance
            .Clear 1: .Rows = 2
            lngRow = 1
            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("���㷽ʽ")))
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ϼ�"))), "###0.00;-###0.00; ;")
            .TextMatrix(lngRow, .ColIndex("�������")) = ""
            .TextMatrix(lngRow, .ColIndex("����")) = ""
            .TextMatrix(lngRow, .ColIndex("������ˮ��")) = ""
            .TextMatrix(lngRow, .ColIndex("����˵��")) = ""
            .AutoSizeMode = flexAutoSizeColWidth
            Call .AutoSize(0, .Cols - 1)
            zl_vsGrid_Para_Restore mlngMode, vsBalance, Me.Name, "������Ϣ�б�", False
        End With
        vsBillList.Clear 1: vsBillList.Rows = 2
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
GoClear:
    With vsList
        vsBalance.Clear 1: vsBalance.Rows = 2
        vsBillList.Clear 1: vsBillList.Rows = 2
    End With
 End Sub
 
Private Sub ShowFactList(ByVal bytƱ�� As Byte, ByVal strNO As String, _
    ByVal blnNOMoved As Boolean)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ʊ��Ϣ
    '���:strNO-���ݺ�
    '       bytƱ��-Ʊ��(1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨)
    '       blnNOMoved-�Ƿ���ʷ��ռ�
    '����:���˺�
    '����:2013-09-16 15:14:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, rsTemp As ADODB.Recordset
    Dim blnIsHaveData As Boolean, lngRow As Long
    
    On Error GoTo errH
    
    blnIsHaveData = False
    If bytƱ�� = 1 And mbytƱ�ݷ������ <> 0 Then
        'Ʊ��,����Ʊ�ݸ�ʽ����
        strSQL = _
        " Select distinct B.ID,B.���� as Ʊ�ݺ�,Decode(B.ԭ��,1,'��������',2,'�����ջ�',3,'�ش򷢳�',4,'�ش��ջ�',6,'��Ʊ����') as ʹ��ԭ��," & _
        " To_Char(B.ʹ��ʱ��,'MM-DD HH24:MI') as ʹ��ʱ��,B.ʹ����" & _
        " From " & IIf(blnNOMoved, "H", "") & "Ʊ�ݴ�ӡ��ϸ A," & _
                IIf(blnNOMoved, "H", "") & "Ʊ��ʹ����ϸ B " & _
        " Where A.Ʊ��=1 And A.Ʊ��=B.���� " & _
        "             And B.Ʊ��=1 And A.NO=[1]" & _
        " Order by ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTemp.RecordCount <> 0 Then GoTo GoGridData
    End If
    
    If bytƱ�� = 3 Then
        strSQL = _
        "   Select B.ID, B.���� as Ʊ�ݺ�,Decode(B.ԭ��,1,'��������',2,'�����ջ�',3,'�ش򷢳�',4,'�ش��ջ�',6,'��Ʊ����') as ʹ��ԭ��," & _
        "           To_Char(B.ʹ��ʱ��,'MM-DD HH24:MI') as ʹ��ʱ��,B.ʹ����" & _
        "   From " & IIf(blnNOMoved, "H", "") & "Ʊ�ݴ�ӡ���� A," & _
                    IIf(blnNOMoved, "H", "") & "Ʊ��ʹ����ϸ B" & _
        " Where A.��������=[2]  And A.ID=B.��ӡID" & _
        "           And B.Ʊ�� In (1,3)  And A.NO=[1]" & _
        " Order by ID"
    Else
        If bytƱ�� = 4 Then
            strSQL = _
            "   Select B.ID, B.���� as Ʊ�ݺ�,Decode(B.ԭ��,1,'��������',2,'�����ջ�',3,'�ش򷢳�',4,'�ش��ջ�',6,'��Ʊ����') as ʹ��ԭ��," & _
            "           To_Char(B.ʹ��ʱ��,'MM-DD HH24:MI') as ʹ��ʱ��,B.ʹ����" & _
            "   From " & IIf(blnNOMoved, "H", "") & "Ʊ�ݴ�ӡ���� A," & _
                        IIf(blnNOMoved, "H", "") & "Ʊ��ʹ����ϸ B" & _
            " Where A.��������=[2]  And A.ID=B.��ӡID" & _
            "           And B.Ʊ�� In (1,4)  And A.NO=[1]" & _
            " Order by ID"
        '110414:���ϴ���2017/6/20��ҽ�ƿ�ʹ�����﷢Ʊ
        ElseIf bytƱ�� = 5 Then
            strSQL = _
            "   Select B.ID, B.���� as Ʊ�ݺ�,Decode(B.ԭ��,1,'��������',2,'�����ջ�',3,'�ش򷢳�',4,'�ش��ջ�',6,'��Ʊ����') as ʹ��ԭ��," & _
            "           To_Char(B.ʹ��ʱ��,'MM-DD HH24:MI') as ʹ��ʱ��,B.ʹ����" & _
            "   From " & IIf(blnNOMoved, "H", "") & "Ʊ�ݴ�ӡ���� A," & _
                        IIf(blnNOMoved, "H", "") & "Ʊ��ʹ����ϸ B" & _
            " Where A.��������=[2]  And A.ID=B.��ӡID" & _
            "           And B.Ʊ�� In (1,5)  And A.NO=[1]" & _
            " Order by ID"
        Else
            strSQL = _
            "   Select B.ID, B.���� as Ʊ�ݺ�,Decode(B.ԭ��,1,'��������',2,'�����ջ�',3,'�ش򷢳�',4,'�ش��ջ�',6,'��Ʊ����') as ʹ��ԭ��," & _
            "           To_Char(B.ʹ��ʱ��,'MM-DD HH24:MI') as ʹ��ʱ��,B.ʹ����" & _
            "   From " & IIf(blnNOMoved, "H", "") & "Ʊ�ݴ�ӡ���� A," & _
                        IIf(blnNOMoved, "H", "") & "Ʊ��ʹ����ϸ B" & _
            " Where A.��������=[2]  And A.ID=B.��ӡID" & _
            "           And B.Ʊ��=[2]  And A.NO=[1]" & _
            " Order by ID"
        End If
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, bytƱ��)
GoGridData:
    With vsBillList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2: lngRow = 1
        Do While Not rsTemp.EOF
            'strHead = "Ʊ�ݺ�,ʹ��ԭ��,ʹ��ʱ��,ʹ����"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݺ�")) = Nvl(rsTemp!Ʊ�ݺ�)
            .TextMatrix(lngRow, .ColIndex("ʹ��ԭ��")) = Nvl(rsTemp!ʹ��ԭ��)
            .TextMatrix(lngRow, .ColIndex("ʹ��ʱ��")) = Format(rsTemp!ʹ��ʱ��, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("ʹ����")) = Nvl(rsTemp!ʹ����)
            .Rows = .Rows + 1: lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
    End With
    '�ָ�������
    zl_vsGrid_Para_Restore mlngMode, vsBillList, Me.Name, "Ʊ����ϸ�б�", False
    vsBillList.Redraw = flexRDBuffered
    Exit Sub
errH:
    vsBillList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadFeeData(Optional ByVal blnRollingData As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�Ŀ��Ϣ
    '���:blnRollingData-�Ƿ��ȡ��������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-03-05 14:03:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String
    Dim strWithTable As String, strWhere As String
    Dim strRollingType As String, i As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If blnRollingData Then
        'Ԥ������NULL,2-����,3-�շ�,4-�Һ�,5-���￨,6-����ҽ������
        'mstrRollingType:�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�)
        '�սɶ���:1-�շ�(���Һ�),2-����,3-Ԥ��,4-���;5-���ѿ���ֵ;6--���ѿ���ֵ;7-�ݴ��(��������)�������տ���������϶��գ��������ӣ���9-���ν����¼
        If InStr(",2,6,", "," & mstrRollingType & ",") > 0 Then
            'Ԥ��������ѿ����վݷ�Ŀ
            vsFeeList.Clear 1: vsFeeList.Rows = 2
            vsFeeList.Row = 1: vsFeeList.Col = 0
            LoadFeeData = True
            Exit Function
        End If
        
        strWhere = " And instr([5],','||A.��������||',')>0 "
        If mstrRollingType <> "" And InStr("," & mstrRollingType & ",", ",0,") = 0 Then
            If Get���ʽ�������(mstrRollingType, strRollingType) = False Then
                vsFeeList.Clear 1: vsFeeList.Rows = 2
                vsFeeList.Row = 1: vsFeeList.Col = 0
                LoadFeeData = True
                Exit Function
            End If
        Else
            strRollingType = ",2,3,4,5,"
        End If
        strWithTable = "" & _
        "   Select distinct A.��������, " & vbCrLf & _
        "          Decode(nvl(A.��������,0),2,2,3,1,4,1,5,1,0) as ����,A.����ID as ��¼ID" & vbCrLf & _
        "   From ����Ԥ����¼ A " & vbCrLf & _
        "   Where a.����Ա���� || '' = [2]  and a.��¼����<>1" & strWhere & vbCrLf & _
        "       And A.�տ�ʱ�� Between  [3] And [4]  " & vbCrLf & _
        "       And Not Exists(Select 1 From ������ü�¼ B Where a.����id = b.����id And Nvl(b.����״̬, 0) = 1) " & vbNewLine & _
        "       And Not Exists(Select 1 From ���˽��ʼ�¼ B Where a.����id = b.Id And b.����״̬ Is Not Null)" & vbNewLine & _
        "       And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y " & _
        "                       Where y.��¼id = a.����id" & _
        "                               And x.Id = y.�ս�ID And Decode(nvl(A.��������,0),2,2,3,1,4,1,5,1,0)=y.����  " & _
        "                               And x.����ʱ�� Is Null) " & vbCrLf
        '���ܰ����������,����������ظ�ͳ����ͳ�ƵĽ��Ҳ����ȷ
'        If mstrRollingType = 1 Or mstrRollingType = 0 Then  '�շѺ��������,�������β���������
'            strWithTable = strWithTable & _
'            "   Union ALL " & _
'            "   Select distinct A.��������,9 as ����,B.�շѽ���ID   as ��¼ID" & vbCrLf & _
'            "   From ����Ԥ����¼ A,���ò����¼ B " & vbCrLf & _
'            "   Where a.����ID=B.����ID And nvl(B.����״̬,0)=0 And A.��������=6 " & vbCrLf & _
'            "       And a.����Ա���� || '' = [2] " & vbCrLf & _
'            "       And A.�տ�ʱ�� Between  [3] And [4]  " & vbCrLf & _
'            "       And Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y Where y.��¼id = a.����id  And x.Id = y.�ս�ID And y.���� = 1 And x.����ʱ�� Is Null) " & vbCrLf
'        End If
        strWithTable = "With c_������Ϣ As  ( " & strWithTable & " )" & vbCrLf
    Else
        strTable = ""
        If mblnDel Or mbytType = EM_�շ�Ա���� Then
            'And mblnDel = False
            If mbytType = EM_�շ�Ա���� Then
                strWhere = " And  A.ID =J.Column_Value And a.��¼����=1"
            Else
                strWhere = " And  A.ID=C.��¼ID And C.����=8 And C.�ս�ID=J.Column_Value And a.��¼����=1"
                strTable = ",��Ա�սɶ��� C"
            End If
        Else
            If mbytType = EM_С���տ� Then
                strWhere = " And  A.С���տ�ID =J.Column_Value And a.��¼����=1"
            ElseIf mbytType = EM_С������ Then
                strWhere = " And  A.С������ID =J.Column_Value And a.��¼����=1"
            Else
                strWhere = " And  A.�����տ�ID =J.Column_Value And a.��¼����=1"
            End If
        End If
         '��Ա�սɶ���.����:1-�շ�(���Һ�),2-����,3-Ԥ��,4-���;5-���ѿ���ֵ;6--���ѿ���ֵ;7-�ݴ��(��������)�������տ���������϶��գ��������ӣ���9-���ν����¼
        strWithTable = "" & _
        "   With c_������Ϣ As  ( " & _
        "           Select /*+cardinality(j,10)*/ b.�ս�id, b.����, b.��¼id" & _
        "           From ��Ա�սɼ�¼ A,��Ա�սɶ��� B,Table( f_Num2list([1])) J " & strTable & _
        "           Where a.Id = b.�ս�id And b.���� in (1,2)" & strWhere & ")" & vbCrLf
        
        '���ܰ����������,����������ظ�ͳ����ͳ�ƵĽ��Ҳ����ȷ
'        strWithTable = strWithTable & _
'        "           Union ALL " & vbCrLf & _
'        "           Select b.�ս�id, b.����,b1.�շѽ���ID" & vbCrLf & _
'        "           From ��Ա�սɼ�¼ A, ��Ա�սɶ��� B,���ò����¼ B1, " & vbCrLf & _
'        "                Table( f_Num2list([1])) J " & strTable & vbCrLf & _
'        "           Where a.Id = b.�ս�id And b.����=9 And B.��¼ID=b1.����ID " & strWhere & ") "
    End If
    
    strSQL = strWithTable & vbCrLf & _
    "   Select A.�վݷ�Ŀ,sum(A.���ʽ��) AS ���ʽ��" & _
    "   From סԺ���ü�¼ A,c_������Ϣ Q1 " & _
    "   Where A.����ID=Q1.��¼ID and Q1.���� in (1,2,9) " & _
    "   GROUP BY A.�վݷ�Ŀ " & _
    "   Union ALL  " & _
    "   Select A.�վݷ�Ŀ,sum(A.���ʽ��) AS ���ʽ��" & _
    "   From ������ü�¼ A,c_������Ϣ Q1 " & _
    "   Where  A.����ID=Q1.��¼ID  and Q1.���� in (1,2,9)" & _
    "   GROUP BY A.�վݷ�Ŀ "
    
    strSQL = "" & _
    "   SELECT A.�վݷ�Ŀ,ltrim(to_char(sum(A.���ʽ��),'99999990.00')) AS ���ʽ�� " & _
    "   FROM ( " & strSQL & " ) a " & _
    "   GROUP BY A.�վݷ�Ŀ " & _
    "   ORDER BY �վݷ�Ŀ"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID, mstrPersonName, mdtStartDate, mdtendDate, strRollingType)
    mblnNotBrush = True
    With vsFeeList
        .Clear 1
        .Rows = 2
        .FixedCols = 0: .FixedRows = 1
        If Not rsTemp.EOF Then
            Set .DataSource = rsTemp
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            Select Case .ColKey(i)
            Case "���ʽ��"
                  .ColAlignment(i) = flexAlignRightCenter
            Case Else
                  .ColAlignment(i) = flexAlignLeftCenter
            End Select
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .SubtotalPosition = flexSTBelow
        If rsTemp.RecordCount <> 0 Then
            .Subtotal flexSTSum, -1, .ColIndex("���ʽ��"), "#######0.00", &HFFC0C0, vbBlack, True, "�ϼ�"
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With
    zl_vsGrid_Para_Restore mlngMode, vsFeeList, Me.Name, "�վݷ�Ŀ�б�", False
    If mblnFirst Then Call reSetFeeListPancelWidth
    
    mblnNotBrush = False
    LoadFeeData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function ShowOnlyFactList() As Boolean
    Dim strSQL As String, i As Long, rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim strStartDate As String, strEndDate As String
    
    On Error GoTo errH
    strSQL = "Select ��ʼʱ��,��ֹʱ�� From ��Ա�սɼ�¼ Where ID= [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID)
    If rsTemp.RecordCount <> 0 Then
        strStartDate = Nvl(rsTemp!��ʼʱ��)
        strEndDate = Nvl(rsTemp!��ֹʱ��)
    Else
        Exit Function
    End If
    
    If strStartDate = "" Or strEndDate = "" Then Exit Function
    
    strSQL = "Select Distinct a.Id, a.���� As Ʊ�ݺ�, Decode(a.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�',6,'��Ʊ����') As ʹ��ԭ��," & vbNewLine & _
            "                To_Char(a.ʹ��ʱ��, 'MM-DD HH24:MI') As ʹ��ʱ��, a.ʹ����" & vbNewLine & _
            "From Ʊ��ʹ����ϸ A, (Select Ʊ��, ����, ��ʼƱ��, ��ֹƱ�� From ��Ա�ս�Ʊ�� Where �ս�id = [1]) B" & vbNewLine & _
            "Where a.Ʊ�� = b.Ʊ�� And a.���� Between b.��ʼƱ�� And b.��ֹƱ�� And a.���� = Decode(b.����, 1, 1, 2, 2, 3, 2)" & vbNewLine & _
            "      And a.ʹ��ʱ�� Between To_date([2],'yyyy-mm-dd hh24:mi:ss') And To_date([3],'yyyy-mm-dd hh24:mi:ss') " & _
            "Order By ID"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID, strStartDate, strEndDate)

    With vsBillList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2: lngRow = 1
        Do While Not rsTemp.EOF
            ShowOnlyFactList = True
            'strHead = "Ʊ�ݺ�,ʹ��ԭ��,ʹ��ʱ��,ʹ����"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݺ�")) = Nvl(rsTemp!Ʊ�ݺ�)
            .TextMatrix(lngRow, .ColIndex("ʹ��ԭ��")) = Nvl(rsTemp!ʹ��ԭ��)
            .TextMatrix(lngRow, .ColIndex("ʹ��ʱ��")) = Format(rsTemp!ʹ��ʱ��, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("ʹ����")) = Nvl(rsTemp!ʹ����)
            .Rows = .Rows + 1: lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
    End With
    '�ָ�������
    zl_vsGrid_Para_Restore mlngMode, vsBillList, Me.Name, "Ʊ����ϸ�б�", False
    vsBillList.Redraw = flexRDBuffered
    Exit Function
errH:
    vsBillList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadList() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�տƱ�ݻ�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-04 11:28:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWithTable As String, strWhere As String, i As Long
    Dim strTable As String
    strTable = ""
    If mblnDel Or mbytType = EM_�շ�Ա���� Or mbytType = EM_�����տ�_���շ�Ա Then
        If mbytType = EM_�շ�Ա���� Or mbytType = EM_�����տ�_���շ�Ա Then
            strWhere = " And  A.ID =J.Column_Value And a.��¼����=1"
        Else
            strWhere = " And  A.ID=C.��¼ID And C.����=8 And C.�ս�ID=J.Column_Value And a.��¼����=1"
            strTable = ",��Ա�սɶ��� C"
        End If
    Else
        If mbytType = EM_С���տ� Then
            strWhere = " And  A.С���տ�ID =J.Column_Value And a.��¼����=1"
        ElseIf mbytType = EM_С������ Then
            strWhere = " And  A.С������ID =J.Column_Value And a.��¼����=1"
        Else
            strWhere = " And  A.�����տ�ID =J.Column_Value And a.��¼����=1"
        End If
    End If
    
    strWithTable = "" & _
    "   With c_������Ϣ As  ( " & _
    "           Select /*+cardinality(j,10)*/ b.�ս�id, b.����, b.��¼id  " & _
    "           From ��Ա�սɼ�¼ A, ��Ա�սɶ��� B,Table( f_Num2list([1])) J " & strTable & _
    "           Where a.Id = b.�ս�id  " & strWhere & ") "
    
    strSQL = "" & _
         "   Select 5 AS ��¼����,A.����ID as ��¼ID,A.NO,A.��¼״̬,max(A.����) AS ��������,max(A.�Ա�) AS �Ա�,max(A.����) AS ����,max(decode(A.�����־,2,-1*NULL,A.��ʶ��)) AS �����,max(decode(A.�����־,2,��ʶ��,null)) AS סԺ��, " & _
         "     Max(A.����Ա����) AS ����Ա����, NULL AS ���㷽ʽ, " & _
         "     sum(���ʽ��) AS ���ϼ�,Max(�Ǽ�ʱ��) AS �շ�ʱ�� " & _
         "   From סԺ���ü�¼ A,c_������Ϣ Q1 " & _
         "   Where Nvl(A.���ʷ���, 0) = 0 And A.��¼״̬ <> 0 AND a.��¼����=5 And A.����ID=Q1.��¼ID and Q1.����=1 " & _
         "   GROUP BY A.NO,A.����ID,A.��¼״̬ " & _
         "   Union ALL  " & _
         "  Select ��¼����, Min(��¼id) As ��¼id, NO, Max(��¼״̬) As ��¼״̬, ��������, �Ա�, ����, �����, סԺ��, ����Ա����, ���㷽ʽ, Sum(���ϼ�) As ���ϼ�, �շ�ʱ�� From (" & _
         "   Select  mod(a.��¼����,10) as ��¼����,A.����ID as ��¼ID,A.NO,a.��¼״̬,max(A.����) AS ��������,max(A.�Ա�) AS �Ա�,max(A.����) AS ����,max(decode(A.�����־,2,-1*NULL,A.��ʶ��)) AS �����,max(decode(A.�����־,2,��ʶ��,null)) AS סԺ��, " & _
         "     Max(A.����Ա����) AS ����Ա����, NULL AS ���㷽ʽ, " & _
         "     sum(���ʽ��) AS ���ϼ�,Max(�Ǽ�ʱ��) AS �շ�ʱ�� " & _
         "   From ������ü�¼ A,c_������Ϣ Q1 " & _
         "   Where (MOD(A.��¼����,10)=1 OR A.��¼����=4) AND Nvl(A.���ʷ���, 0) = 0 And A.��¼״̬ <> 0 And Nvl(A.����״̬, 0) <> 1  " & _
         "        And A.����ID=Q1.��¼ID and Q1.����=1 " & _
         "   GROUP BY mod(a.��¼����,10),A.����ID,A.NO,a.��¼״̬) " & _
         "  Group By ��¼����, NO, ��������, �Ա�, ����, �����, סԺ��, ����Ա����, ���㷽ʽ, �շ�ʱ�� "
         
         strSQL = strSQL & _
         "   Union All " & _
         "   Select 6 AS ��¼����,A.ID as ��¼ID,A.NO,A.��¼״̬,max(C.����) AS ��������,max(C.�Ա�) AS �Ա�,max(C.����) AS ����, " & _
         "     max(c.�����) AS �����,max(C.סԺ��) AS סԺ��, " & _
         "     Max(A.����Ա����) AS ����Ա����,NULL AS ���㷽ʽ, " & _
         "     sum(b.��Ԥ��) AS ���ϼ�,Max(A.�շ�ʱ��) AS �շ�ʱ��  " & _
         "   From ���˽��ʼ�¼ A,����Ԥ����¼ B,������Ϣ C,c_������Ϣ Q1" & _
        "   Where A.ID=B.����ID And A.����״̬ Is Null and A.����ID=C.����ID(+)  " & _
         "         And A.ID=Q1.��¼ID and Q1.����=2  " & _
         "   group by A.NO,A.ID,A.��¼״̬ "
        
        strSQL = strSQL & _
         "   Union All " & _
         "   Select 7 As ��¼����,A.ID as ��¼ID, a.No, a.��¼״̬, Max(nvl(M.����,c.����)) As ��������, Max(c.�Ա�) As �Ա�, Max(c.����) As ����, Max(c.�����) As �����, " & _
         "      Max(Decode(m.סԺ��, Null, c.סԺ��, m.סԺ��)) As סԺ��, Max(a.����Ա����) As ����Ա����,Max(���㷽ʽ) AS ���㷽ʽ, Sum(a.���) As ���ϼ�, Max(a.�տ�ʱ��) As �շ�ʱ�� " & _
         "   From ����Ԥ����¼ A, ������Ϣ C, ������ҳ M,c_������Ϣ Q1 " & _
         "   Where a.����id = c.����id(+) And a.����id = m.����id(+) And a.��ҳid = m.��ҳid(+)  " & _
         "     And a.��¼���� = 1  And A.ID=Q1.��¼ID and Q1.����=3  " & _
         "   Group By a.No,A.ID, a.��¼״̬ " & _
         "   Union All " & _
         "   Select 8 as ��¼����,A.����ID as ��¼ID,a.���� As no,a.��¼״̬,'' as ��������,'' as �Ա�,'' as ����,null �����,NULL as סԺ��, " & _
         "        A.����Ա����,a.���㷽ʽ, a.ʵ�ս��  As ���, A.�Ǽ�ʱ�� AS �շ�ʱ��  " & _
         "   From ���˿������¼ A,c_������Ϣ Q1 " & _
         "   Where a.��¼���� In (2, 3) And A.����ID=Q1.��¼ID and Q1.����=5   " & _
         "   Union All " & _
         "   Select 9 As ��¼����,A.����ID as ��¼ID, a.���� As NO, a.��¼״̬, '' As ��������, '' As �Ա�, '' As ����, Null �����, " & _
         "       Null As סԺ��, a.����Ա����, a.���㷽ʽ, a.ʵ�ս�� As ���, a.�Ǽ�ʱ�� " & _
         "   From ���˿������¼ A,c_������Ϣ Q1 " & _
         "   Where a.��¼���� In (1, 3) And A.����ID=Q1.��¼ID and Q1.����=6    "
         
        strSQL = strSQL & _
        "   Union All" & _
        "   Select 10 As ��¼����,A.ID as ��¼ID, ltrIm(to_Char(a.ID)) As NO, 1 As ��¼״̬, '' As ��������, '' As �Ա�, '' As ����, Null �����, " & _
        "     Null As סԺ��, a.����� As ����Ա����, a.���㷽ʽ, a.����� As ���, a.���ʱ�� " & _
        "   From ��Ա����¼ A,c_������Ϣ Q1,��Ա�սɼ�¼ M " & _
        "   Where A.ID=Q1.��¼ID and Q1.����=4 and Q1.�ս�ID=M.ID and M.�տ�Ա||''=a.����� " & _
        "   Union All " & _
        "   Select 11 As ��¼����, A.ID as ��¼ID,ltrIm( to_Char(a.ID)) As NO, 1 As ��¼״̬, '' As ��������, '' As �Ա�, '' As ����, Null �����, " & _
        "     Null As סԺ��, a.����� As ����Ա����, a.���㷽ʽ, a.����� As ���, a.���ʱ�� " & _
        "   From ��Ա����¼ A,c_������Ϣ Q1,��Ա�սɼ�¼ M " & _
        "   Where   A.ID=Q1.��¼ID and Q1.����=4 and Q1.�ս�ID=M.ID and M.�տ�Ա||''=a.�����  " & _
        "   Union All " & _
        "   Select 12 As ��¼����, A.ID as ��¼ID,a.No As NO, Decode(�ջ�ʱ��, Null, 1, 2) As ��¼״̬, '' As ��������, '' As �Ա�, '' As ����, Null �����, Null As סԺ��, " & _
        "      a.�Ǽ��� As ����Ա����, '�ֽ�' As ���㷽ʽ, a.���, a.�Ǽ�ʱ�� " & _
        "   From ��Ա�ݴ��¼ A,c_������Ϣ Q1 " & _
        "   Where A.ID=Q1.��¼ID and Q1.����=7 and a.��¼����=2 And a.�ջ�ʱ�� Is Null"
        
        strSQL = strSQL & _
        "   Union ALL" & vbCrLf & _
        "   Select 13 as ��¼����, a.����id As ��¼id, a.No, a.��¼״̬, Max(c.����) As ��������, Max(c.�Ա�) As �Ա�, Max(c.����) As ����, c.����� As �����," & vbNewLine & _
        "        c.סԺ�� As סԺ��, Max(a.����Ա����) As ����Ա����, Null As ���㷽ʽ, Sum(M.��Ԥ��) As ���ϼ�, Max(a.�Ǽ�ʱ��) As �շ�ʱ��" & vbNewLine & _
        "   From (Select a.����ID,A.NO,A.��¼״̬,Max(a.����ID) as ����ID,max(a.����Ա����) as ����Ա����,max(a.�Ǽ�ʱ��) as �Ǽ�ʱ�� " & _
        "         From ���ò����¼ A, c_������Ϣ B " & _
        "         Where A.����ID=B.��¼ID and B.����=9  " & _
        "         Group by a.����ID,A.NO,A.��¼״̬) A,����Ԥ����¼ M,������Ϣ C" & vbNewLine & _
        "   Where  A.����ID=M.����ID And A.����ID=C.����ID " & _
        "   Group By  a.����id, a.No, a.��¼״̬, c.�����, c.סԺ��"
        
    
    strSQL = strWithTable & vbCrLf & strSQL
    strSQL = "" & _
    "   SELECT /*+ rule */ ��¼���� as ����,��¼ID,decode(��¼����,1,'�շ�',4,'�Һ�',5,'���￨',6,'����',7,'Ԥ����',8,'���ѿ���ֵ',9,'���ѿ����',10,'���',11,'���',12,'�ݴ�',13,'������','') AS ���, " & _
    "       NO,��¼״̬,��������,�Ա�,����,�����,סԺ��,����Ա����,���㷽ʽ," & _
    "       Trim(to_char(���ϼ�,'99999999990.00')) As ���ϼ�,to_char(�շ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as �շ�ʱ�� " & _
    "   FROM ( " & strSQL & " ) " & _
    "   ORDER BY ��¼����,�շ�ʱ�� DESC,NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID)
    mblnNotBrush = True
    With vsList
        .Clear 1
        .Rows = 2
        .FixedCols = 0: .FixedRows = 1
        If Not rsTemp.EOF Then
            Set .DataSource = rsTemp
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            Select Case .ColKey(i)
            Case "����", "��¼״̬", "��¼ID"
                  .ColWidth(i) = 0: .ColHidden(i) = True
            Case "NO", "���", "�Ա�", "����", "�����", "סԺ��", "���㷽ʽ", "�շ�ʱ��"
                  .ColAlignment(i) = flexAlignCenterCenter
            Case "���ϼ�"
                  .ColAlignment(i) = flexAlignRightCenter
            Case Else
                  .ColAlignment(i) = flexAlignLeftCenter
            End Select
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
        If rsTemp.RecordCount <> 0 Then
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, -1, .ColIndex("���ϼ�"), "#######0.00", &HFFC0C0, vbBlack, True, "�ϼ�"
            For i = 0 To .ColIndex("���ϼ�") - 1
                .TextMatrix(.Rows - 1, i) = "�ϼ�"
            Next
            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            .MergeRow(.Rows - 1) = True
            .MergeCells = flexMergeRestrictRows
            '�����:110535,����,2017/09/04,����ɫ�����˷Ѽ�¼�ͱ��˷Ѽ�¼
             For i = 1 To .Rows - 1
                Select Case .TextMatrix(i, .ColIndex("��¼״̬"))
                Case 1
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                Case 2
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                Case 3
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                End Select
            Next
        End If
        
    End With
    zl_vsGrid_Para_Restore mlngMode, vsList, Me.Name, "��ϸ��Ϣ�б�", False
    mblnNotBrush = False
    Call LoadDetailData
    Call LoadFeeData '�����վݷ�Ŀ��Ϣ
    LoadList = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotBrush = False
 End Function
Public Sub SetFontSize(ByVal bytSize As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ9����)��1-���(12��);>1: Ϊָ�����ֺ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-03 18:05:20
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������С
    '����:���˺�
    '����:2013-09-03 18:04:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Me.FontSize = mbytFontSize
    Set dkpMan.PaintManager.CaptionFont = Me.Font
    Set vsList.Font = Me.Font
    Set vsBalance.Font = Me.Font
    Set vsBillList.Font = Me.Font
 End Sub
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2013-09-03 15:28:24
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset
    mbytFontSize = 9
    mstrCashBalance = "�ֽ�"
    Set rsTemp = Get���㷽ʽ
    rsTemp.Filter = "����=1"
    If Not rsTemp.EOF Then mstrCashBalance = rsTemp!����
    strTmp = Trim(zlDatabase.GetPara("Ʊ�ݷ������", glngSys, 1121, "0||0;0;0;0;0"))
    mbytƱ�ݷ������ = Val(Split(strTmp & "||", "||")(0))
    stbThis.Panels(3).Text = UserInfo.����
    stbThis.Panels(3).ToolTipText = "��ǰ����Ա:" & UserInfo.����
    
    Call InitPanel
    Call InitGrid
    Call ReSetFontSize
 End Sub
Private Sub LoadFilterRange()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������Χ
    '����:���˺�
    '����:2013-09-16 17:18:08
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNO As String, strTime As String
    lblNO.Visible = mstrChargeRollingID <> ""
    Select Case mbytType
     Case EM_�շ�Ա����, EM_�����տ�_���շ�Ա
        If mstrChargeRollingID = "" Then
            lblRange.Caption = "ʱ�䷶Χ:" & Format(mdtStartDate, "yyyy-mm-dd HH:MM:SS") & "��" & Format(mdtendDate, "yyyy-mm-dd HH:MM:SS")
            Exit Sub
        End If
        
         strSQL = "" & _
        "   Select /*+cardinality(j,10)*/ A.NO,to_char(A.��ʼʱ��,'yyyy-mm-dd hh24:mi:ss')||'��'||to_char(A.��ֹʱ��,'yyyy-mm-dd hh24:mi:ss') as ʱ�䷶Χ " & _
        "   From ��Ա�սɼ�¼ A ,Table( f_Num2list([1])) J" & _
        "   Where A.ID= J.Column_Value" & _
        "   Order by A.NO "
     Case EM_С���տ�
        strSQL = "" & _
        "   Select /*+cardinality(j,10)*/ A.NO,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as ʱ�䷶Χ " & _
        "   From ��Ա�սɼ�¼ A ,Table( f_Num2list([1])) J" & _
        "   Where A.ID= J.Column_Value" & _
        "   Order by A.NO "
     Case EM_С������
        strSQL = "" & _
        "   Select /*+cardinality(j,10)*/ A.NO,to_char(A.��ʼʱ��,'yyyy-mm-dd hh24:mi:ss')||'��'||to_char(A.��ֹʱ��,'yyyy-mm-dd hh24:mi:ss') as ʱ�䷶Χ " & _
        "   From ��Ա�սɼ�¼ A ,Table( f_Num2list([1])) J" & _
        "   Where A.ID= J.Column_Value" & _
        "   Order by A.NO "
     Case EM_�����տ�
        strSQL = "" & _
        "   Select /*+cardinality(j,10)*/ A.NO,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as ʱ�䷶Χ " & _
        "   From ��Ա�սɼ�¼ A ,Table( f_Num2list([1])) J" & _
        "   Where A.ID= J.Column_Value" & _
        "   Order by A.NO "
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID)
    With rsTemp
        strNO = "": strTime = ""
        Do While Not .EOF
            strNO = strNO & ";" & rsTemp!NO
            strTime = strTime & ";" & rsTemp!ʱ�䷶Χ
            .MoveNext
        Loop
        lblNO.Caption = "���ݺ�:"
        If strNO <> "" Then lblNO.Caption = lblNO.Caption & Mid(strNO, 2)
        lblRange.Caption = "ʱ�䷶Χ:"
        If strTime <> "" Then lblRange.Caption = lblRange.Caption & Mid(strTime, 2)
    End With
End Sub
Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2013-09-16 16:47:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, objConPan As Pane
    Dim objFeePan As Pane
    Dim lngDetailHeight As Long 'ȱʡ��ϸ�߶�
    Dim lngTemp As Long
    lngDetailHeight = 2925 / Screen.TwipsPerPixelX
    lngTemp = picConList.Height \ Screen.TwipsPerPixelY
    With dkpMan
        Set objConPan = .CreatePane(mPaneIndex.EM_PN_ConList, 400, 400, DockBottomOf, Nothing)
        objConPan.Title = "����������Ϣ"
        objConPan.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        objConPan.Handle = picConList.hWnd
        objConPan.MaxTrackSize.Height = lngTemp
        objConPan.MinTrackSize.Height = lngTemp
        objConPan.Tag = mPaneIndex.EM_PN_ConList
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_LIST, 600, 400, DockBottomOf, objConPan)
        If mbytType = EM_�շ�Ա���� Then
            objPane.Title = "�շ�Ա������ϸ"
        ElseIf mbytType = EM_С������ Then
            objPane.Title = "С��������ϸ"
        ElseIf mbytType = EM_С���տ� Then
            objPane.Title = "С���տ���ϸ"
        Else
            objPane.Title = "�����տ���ϸ"
        End If
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
        objConPan.Tag = mPaneIndex.EM_PN_LIST
        
        Set objFeePan = .CreatePane(mPaneIndex.EM_PN_FeeLIST, 160, 400, DockRightOf, objPane)
        objFeePan.Title = "��Ŀ��Ϣ"
        objFeePan.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objFeePan.Handle = picFeeList.hWnd
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_BALANCE, 400, lngDetailHeight, DockBottomOf, objPane)
        objPane.MinTrackSize.Height = lngDetailHeight
        objPane.Title = "�տ���Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_BILL, 400, 400, DockRightOf, objPane)
        objPane.Title = "Ʊ����Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBillList.hWnd
        
        
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ���Ϣ
    '����:
    '����:���˺�
    '����:2013-09-03 11:38:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strHead As String, strArr As Variant
    '�տ������Ϣ
    strHead = "����,��¼ID,���,NO,��¼״̬,��������,�Ա�,�Ա�,����,�����,סԺ��,����Ա����,���㷽ʽ,���ϼ�,�շ�ʱ��"
    strArr = Split(strHead, ",")
    With vsList
        Set .Font = Me.Font
        .Cols = UBound(strArr) + 1: .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        For i = 0 To UBound(strArr)
            .TextMatrix(0, i) = strArr(i): .ColKey(i) = UCase(strArr(i))
          Select Case .ColKey(i)
          Case "����", "��¼״̬", "��¼ID"
                .ColWidth(i) = 0: .ColHidden(i) = True
          Case "NO", "���", "�Ա�", "����", "�����", "סԺ��", "���㷽ʽ", "�շ�ʱ��"
                .ColAlignment(i) = flexAlignCenterCenter
          Case "���ϼ�"
                .ColAlignment(i) = flexAlignRightCenter
          Case Else
                .ColAlignment(i) = flexAlignLeftCenter
          End Select
          .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        '.ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsList, Me.Name, "��ϸ��Ϣ�б�", False
    End With
    
    
    '������Ϣ��Ϣ
    strHead = "���㷽ʽ,���,�������,����,������ˮ��,����˵��"
    strArr = Split(strHead, ",")
    With vsBalance
        .Cols = UBound(strArr) + 1: .Rows = 2
        .FixedRows = 1
        For i = 0 To UBound(strArr)
            .TextMatrix(0, i) = strArr(i): .ColKey(i) = UCase(strArr(i))
          Select Case .ColKey(i)
          Case "���㷽ʽ", "����"
                .ColAlignment(i) = flexAlignCenterCenter
          Case "���"
                .ColAlignment(i) = flexAlignRightCenter
          Case Else
                .ColAlignment(i) = flexAlignLeftCenter
          End Select
          .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        '.ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsBalance, Me.Name, "������Ϣ�б�", False
    End With
   'Ʊ����Ϣ
   strHead = "Ʊ�ݺ�,ʹ��ԭ��,ʹ��ʱ��,ʹ����"
    strArr = Split(strHead, ",")
    With vsBillList
        .Cols = UBound(strArr) + 1: .Rows = 2
        .FixedRows = 1
        For i = 0 To UBound(strArr)
            .TextMatrix(0, i) = strArr(i): .ColKey(i) = UCase(strArr(i))
          Select Case .ColKey(i)
          Case "ʹ��ʱ��", "ʹ����"
                .ColAlignment(i) = flexAlignCenterCenter
          Case Else
                .ColAlignment(i) = flexAlignLeftCenter
          End Select
          .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        '.ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsBillList, Me.Name, "Ʊ����ϸ�б�", False
    End With
    
   'Ʊ����Ϣ
   strHead = "�վݷ�Ŀ,���ʽ��"
    strArr = Split(strHead, ",")
    With vsFeeList
        .Cols = UBound(strArr) + 1: .Rows = 2
        .FixedRows = 1
        For i = 0 To UBound(strArr)
            .TextMatrix(0, i) = strArr(i): .ColKey(i) = UCase(strArr(i))
          Select Case .ColKey(i)
          Case "ʵ�ս��"
                .ColAlignment(i) = flexAlignRightBottom
          Case Else
                .ColAlignment(i) = flexAlignLeftCenter
          End Select
          .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        '.ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsFeeList, Me.Name, "�վݷ�Ŀ�б�", False
    End With
 End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    End Select
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    reSetFeeListPancelWidth True
End Sub
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
    
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case EM_PN_ConList  '����
        Item.Handle = picConList.hWnd
    Case EM_PN_LIST '�տ����
        Item.Handle = picList.hWnd
    Case EM_PN_BALANCE  '���㷽ʽ
        Item.Handle = picBalance.hWnd
    Case EM_PN_BILL  '�˷�Ʊ��
        Item.Handle = picBillList.hWnd
    Case EM_PN_FeeLIST   '���û���
        Item.Handle = picFeeList.hWnd
    End Select
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
 Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Top = .ScaleTop
        vsBalance.Height = .ScaleHeight
        vsBalance.Width = .ScaleWidth
    End With
End Sub

 

Private Sub picConList_Resize()
    Err = 0: On Error Resume Next
    With picConList
        fraSplit.Top = .ScaleTop
        fraSplit.Left = .ScaleLeft
        fraSplit.Width = .ScaleWidth
        lblNO.Top = .ScaleHeight - lblNO.Height - 50
        lblRange.Top = lblNO.Top
        lblNO.Left = .ScaleLeft + 50
        lblRange.Left = IIf(mstrChargeRollingID <> "", lblNO.Left + lblNO.Width * 2 + 50, lblNO.Left)
    End With
End Sub
Private Sub picFeeList_Resize()
    Err = 0: On Error Resume Next
    With picFeeList
        vsFeeList.Left = .ScaleLeft
        vsFeeList.Top = .ScaleTop
        vsFeeList.Width = .ScaleWidth
        vsFeeList.Height = .ScaleHeight - vsFeeList.Top
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        vsList.Top = .ScaleTop
        vsList.Left = .ScaleLeft
        vsList.Height = .ScaleHeight - vsList.Top
        vsList.Width = .ScaleWidth
    End With
End Sub
 
Private Sub picBillList_Resize()
    Err = 0: On Error Resume Next
    With picBillList
        vsBillList.Left = .ScaleLeft
        vsBillList.Top = .ScaleTop
        vsBillList.Height = .ScaleHeight
        vsBillList.Width = .ScaleWidth
    End With
End Sub

Private Sub vsFeeList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsFeeList, Me.Name, "�վݷ�Ŀ�б�", False
End Sub

Private Sub vsFeeList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsFeeList, Me.Name, "�վݷ�Ŀ�б�", False
   
End Sub
Private Sub vsList_GotFocus()
    Call zl_VsGridGotFocus(vsList)
End Sub
Private Sub vsList_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsList)
End Sub
Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsList, Me.Name, "��ϸ��Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsList, OldRow, NewRow, OldCol, NewCol)
    If OldRow = NewRow Or mblnNotBrush Then Exit Sub
    Call LoadDetailData
End Sub
Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsList, Me.Name, "��ϸ��Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
 
Private Sub vsBalance_GotFocus()
    Call zl_VsGridGotFocus(vsBalance)
End Sub
Private Sub vsBalance_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBalance)
End Sub
Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsBalance, Me.Name, "������Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBalance, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsBalance, Me.Name, "������Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsBillList_GotFocus()
    Call zl_VsGridGotFocus(vsBillList)
End Sub
Private Sub vsBillList_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBillList)
End Sub
Private Sub vsBillList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsBillList, Me.Name, "Ʊ����ϸ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsBillList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBillList, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBillList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsBillList, Me.Name, "Ʊ����ϸ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б���Ϣ
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim i As Long, lngRow As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim blnFeeList As Boolean
    
    Err = 0: On Error GoTo ErrHand:
    '���������Ϣ
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    If Me.ActiveControl Is vsFeeList Then
        objPrint.Title.Text = gstr��λ���� & "��Ŀ���ܱ�"
    Else
        objPrint.Title.Text = gstr��λ���� & "�տ����"
    End If
    Set objRow = New zlTabAppRow
    If lblNO.Visible Then
        objRow.Add "" & lblNO.Caption
    End If
    If lblRange.Visible Then
        objRow.Add "" & lblRange.Caption
    End If
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = IIf(Me.ActiveControl Is vsFeeList, vsFeeList, vsList)
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Public Function zlDefCommandBars() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-09-16 16:56:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    For Each mcbrControl In mcbrToolBar.Controls
          If mcbrControl.ID <> conMenu_COMBOX_INTERFACE Then
            mcbrControl.Style = xtpButtonIconAndCaption
          End If
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
     Select Case Control.ID
        Case conMenu_File_Exit: Unload Me: '�˳�(&X)
        Case conMenu_File_PrintSet: Call zlPrintSet '��ӡ����
        Case conMenu_File_Preview: Call zlPrint(2)  'Ԥ��(&V)
        Case conMenu_File_Print: Call zlPrint(1) '��ӡ(&P)
        Case conMenu_File_Excel: Call zlPrint(3)  '�����&Excel��
        Case conMenu_View_Refresh: zlRefresh 'ˢ��(&R)
        Case conMenu_View_StatusBar '״̬��(&S)
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub zlRefresh()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '����:���˺�
    '����:2013-09-16 17:08:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ReadListData
End Sub
Private Sub reSetFeeListPancelWidth(Optional blnSetMaxWidth As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������÷��õĿ��
    '���:blnSetMaxWidth-���������
    '����:����true,���򷵻�False
    '����:���˺�
    '����:2015-03-06 12:07:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single
    Dim objPan As Pane
    Set objPan = dkpMan.FindPane(mPaneIndex.EM_PN_FeeLIST)
    If objPan Is Nothing Then Exit Sub
    
    If blnSetMaxWidth Then
        dkpMan.RecalcLayout
        sngWidth = (Me.ScaleWidth \ Screen.TwipsPerPixelY) * Round(1 / 3, 4)
        If sngWidth < 200 Then sngWidth = 200
        objPan.MaxTrackSize.Width = sngWidth
       ' dkpMan.RecalcLayout
        Exit Sub
    End If
    
    sngWidth = GetFeeListMaxWidth \ Screen.TwipsPerPixelY
    objPan.MaxTrackSize.Width = sngWidth
    dkpMan.RecalcLayout
End Sub
Private Function GetFeeListMaxWidth() As Single
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����б�������
    '����:���������
    '����:���˺�
    '����:2015-03-06 11:47:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, i As Long
    
    With vsFeeList
        sngWidth = 0
        For i = 0 To .Cols - 1
            sngWidth = sngWidth + .ColWidth(i) + 70
        Next
    End With
    GetFeeListMaxWidth = sngWidth
End Function

