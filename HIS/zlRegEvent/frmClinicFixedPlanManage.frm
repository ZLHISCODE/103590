VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicFixedPlanManage 
   BorderStyle     =   0  'None
   Caption         =   "�̶����ﰲ�Ź���"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   9720
      MaxLength       =   100
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.PictureBox picPlanDateRange 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   5040
      ScaleHeight     =   285
      ScaleWidth      =   3915
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   3915
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   285
         Left            =   900
         TabIndex        =   8
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   -2147483630
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   169869315
         CurrentDate     =   40777
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   285
         Left            =   2505
         TabIndex        =   9
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   -2147483630
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   169869315
         CurrentDate     =   40777
      End
      Begin VB.Label lblSplit 
         Caption         =   "��"
         Height          =   210
         Left            =   2250
         TabIndex        =   10
         Top             =   30
         Width           =   330
      End
      Begin VB.Label lblDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ��ʾ��"
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   45
         Width           =   900
      End
   End
   Begin VB.PictureBox picRegistPlan 
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   5280
      ScaleHeight     =   3645
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   1380
      Width           =   4395
      Begin VSFlex8Ctl.VSFlexGrid vsfRegistPlan 
         Height          =   2445
         Left            =   240
         TabIndex        =   4
         Top             =   150
         Width           =   3495
         _cx             =   6165
         _cy             =   4313
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClinicFixedPlanManage.frx":0000
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
         Begin VB.PictureBox picImgPlan 
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   75
            Picture         =   "frmClinicFixedPlanManage.frx":0075
            ScaleHeight     =   135
            ScaleWidth      =   150
            TabIndex        =   5
            Top             =   90
            Width           =   150
         End
      End
   End
   Begin VB.PictureBox picRegistRule 
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   180
      ScaleHeight     =   4065
      ScaleWidth      =   4575
      TabIndex        =   2
      Top             =   2220
      Width           =   4575
      Begin VB.Frame fraSplitRule 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   2850
         MousePointer    =   7  'Size N S
         TabIndex        =   18
         Top             =   1770
         Width           =   1005
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfRegistRule 
         Height          =   1425
         Left            =   330
         TabIndex        =   14
         Top             =   150
         Width           =   2775
         _cx             =   4895
         _cy             =   2514
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClinicFixedPlanManage.frx":016B
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
         Begin VB.PictureBox picImgRule 
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   75
            Picture         =   "frmClinicFixedPlanManage.frx":01E0
            ScaleHeight     =   135
            ScaleWidth      =   150
            TabIndex        =   15
            Top             =   90
            Width           =   150
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfRegistRuleSub 
         Height          =   1305
         Left            =   0
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
         _cx             =   4048
         _cy             =   2302
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClinicFixedPlanManage.frx":02D6
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
         Begin VB.PictureBox picImgRuleSub 
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   75
            Picture         =   "frmClinicFixedPlanManage.frx":034B
            ScaleHeight     =   135
            ScaleWidth      =   150
            TabIndex        =   17
            Top             =   90
            Width           =   150
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   1575
      Left            =   270
      TabIndex        =   0
      Top             =   570
      Width           =   1305
      _Version        =   589884
      _ExtentX        =   2302
      _ExtentY        =   2778
      _StockProps     =   64
   End
   Begin VB.Label lblValidTimeRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Чʱ�䣺2016-02-12 00:00:00��3000-01-01 00:00:00"
      Height          =   180
      Left            =   1740
      TabIndex        =   12
      Top             =   660
      Width           =   4500
   End
   Begin VB.Label lblPublishInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ˣ�Ƚ����  ����ʱ�䣺2016-01-02 12:32:12"
      Height          =   180
      Left            =   6750
      TabIndex        =   11
      Top             =   210
      Width           =   4050
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   7035
      Left            =   -600
      Top             =   -150
      Width           =   11475
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10845
      _Version        =   589884
      _ExtentX        =   19129
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "���ﰲ��>�̶�����"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmClinicFixedPlanManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mlngModule As Long
Private mstrPrivs As String

Private Enum mPgIndex '�̶�����TabPage����
    Pg_������� = 0
    Pg_���ﰲ�� = 1
End Enum

Private Enum mPanIndex
    Pan_RegistRuleMain = 0
    Pan_RegistRuleSub = 1
End Enum
Private mlng����ID As Long

Private mrsRuleRecords As ADODB.Recordset
Private mrsPlanRecords As ADODB.Recordset
Private mrsRuleRecordsSub As ADODB.Recordset
Private mlngSignalCount As Long '��Դ����
Private mintFindType As Integer

Private mlngCopyPlanID As Long, mstrCopyPlanItem As String '���ڸ���ճ��
Private mdtToday As Date
Private mstrOldSelRangePlan As String 'ѡ���������򣬸�ʽ"��ʼ��|������|��ʼ��|������"
Private mstrOldSelRangeRule As String
Private mstrOldSelRangeRuleSub As String

Private mblnShowInvalidPlan As Boolean '�Ƿ���ʾ��Ч��ʱ����

Public Sub InitCommVariable(frmParent As Form, cbsMain As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsMain
    
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom

    Err = 0: On Error GoTo errHandler

    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
'        Set cbrControl = .Add(xtpControlButton,conMenu_File_ExportToXML,"����ΪXML�ļ�(&L)��",cbrControl.Index + 1)
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "�ƶ��³����(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "�ƶ��ܳ����(&W)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PlanAdd, "�ƶ���ʱ����(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddNewSignalSource, "������Դ����(&A)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "��������(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "����ԤԼ�Һſ���(&U)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PlanVerify, "�����ʱ����(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PlanCancel, "ȡ����ʱ�������(&C)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStartNO, "ȫ��������ſ���(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStopNO, "ȫ��ȡ����ſ���(&T)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CopyPlan, "���ư���(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PastPlan, "ճ������(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearCurPlan, "�����ǰ����(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAllPlan, "�����ǰ��Դ����(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearTempPlan, "�����ǰ��ʱ����(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAll, "������к�Դ����(&A)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PublishPlan, "��������(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnPublishPlan, "ȡ������(&U)")

        '�������ŵ���
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddTempPlan, "��ʱ����(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UpdatePlan, "��������(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_LockResource, "����(&L)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnLockResource, "����(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_StopOutCall, "ͣ��(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnStopOutCall, "ȡ��ͣ��(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_OpenStopPlan, "����ͣ�ﰲ��(&O)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyDoctor, "����(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnModifyDoctor, "ȡ������(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddNumberLimit, "�Ӻ�(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ReduceNumberLimit, "����(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyDoctorOffice, "������������(&Z)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UpdateUnitRegist, "����ԤԼ�Һſ���(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintPlan, "��ӡ�����(&P)"): cbrControl.BeginGroup = True
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowDoctorStopPlan, "��ʾҽ��ͣ�ﰲ��(&P)", cbrControl.index)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_PlanChangeInfo, "��ѯ�䶯��Ϣ(&C)", cbrControl.index): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "��ʾʧЧ��ʱ����(&S)", cbrControl.index): cbrControl.BeginGroup = True
        cbrControl.Checked = mblnShowInvalidPlan
    End With

    '����������
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "�³����", cbrControl.index + 1)
        cbrControl.ToolTipText = "�ƶ��³����"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "�ܳ����", cbrControl.index + 1)
        cbrControl.ToolTipText = "�ƶ��ܳ����"

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "��������", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "ԤԼ�Һſ���", cbrControl.index + 1)
        cbrControl.ToolTipText = "����ԤԼ�Һſ���"

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PublishPlan, "��������", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnPublishPlan, "ȡ������", cbrControl.index + 1)

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_LockResource, "����", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnLockResource, "����", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_StopOutCall, "ͣ��", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyDoctor, "����", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddNumberLimit, "�Ӻ�", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ReduceNumberLimit, "����", cbrControl.index + 1)

        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowDoctorStopPlan, "ͣ�ﰲ��", cbrControl.index + 1): cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_Edit_StopOutCall
    End With

    Set objPopup = cbrToolBar.Controls.Add(xtpControlButtonPopup, conMenu_View_FindType, "��������ˡ�")
    objPopup.flags = xtpFlagRightAlign
    '���󶨵Ŀؼ����붯̬���أ���Ϊ������һ����ɾ�������󶨵Ŀؼ��ľ���ͻ���0
    Set objCustom = cbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
    If txtFind.UBound > 0 Then Unload txtFind(1)
    Load txtFind(1)
    objCustom.Handle = txtFind(1).Hwnd
    objCustom.flags = xtpFlagRightAlign

    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("Y"), conMenu_Edit_AddMonthPlan
        .Add FCONTROL, Asc("W"), conMenu_Edit_AddWeekPlan
        .Add FCONTROL, Asc("M"), conMenu_Edit_ModifyPlanItem

        .Add FCONTROL, Asc("C"), conMenu_Edit_CopyPlan
        .Add FCONTROL, Asc("V"), conMenu_Edit_PastPlan
        .Add 0, VK_DELETE, conMenu_Edit_ClearCurPlan

        .Add FCONTROL, Asc("G"), conMenu_Edit_PublishPlan
        .Add FCONTROL, Asc("I"), conMenu_Edit_UnPublishPlan
    End With

    '���ò���������
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function PlanIsValid(ByVal vsfGrid As VSFlexGrid) As Boolean
    '�жϵ�ǰѡ�����Ƿ���Ч
    Dim strCurItem As String
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        strCurItem = .Cell(flexcpData, 0, .Col)
        If IsDate(strCurItem) = False Then Exit Function
        If DateDiff("d", strCurItem, mdtToday) > 0 Then Exit Function
    End With
    PlanIsValid = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsCan�ٴ��Ű�(ByVal vsfGrid As VSFlexGrid) As Boolean
    '�жϵ�ǰѡ���Դ�Ƿ��ܹ��ٴ��ܹ��Ű�
    Dim strCurItem As String
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        'û�С����п��ҡ�Ȩ��ʱ��ֻ�ܵ����������ٴ������Űࡱ�ĺ�Դ
        If zlStr.IsHavePrivs(mstrPrivs, "���п���") Then
            IsCan�ٴ��Ű� = True
        Else
            IsCan�ٴ��Ű� = Trim(.TextMatrix(.Row, COL_�Ƿ��ٴ��Ű�)) <> ""
        End If
    End With
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim bytActiveGrid As Byte   '��ǰ������
    Dim vsfGrid As VSFlexGrid
    Dim blnEnabled As Boolean
    Dim lng����ID As Long
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    If Not tbPage.Selected Is Nothing Then
        If tbPage.Selected.index = Pg_������� Then
            If Me.ActiveControl Is vsfRegistRuleSub Then
                bytActiveGrid = 2
                Set vsfGrid = vsfRegistRuleSub
            Else
                bytActiveGrid = 1
                Set vsfGrid = vsfRegistRule
            End If
        Else
            bytActiveGrid = 3
            Set vsfGrid = vsfRegistPlan
        End If
    End If

    blnEnabled = mlng����ID <> 0
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfGrid.Rows > vsfGrid.FixedRows
    Case conMenu_EditPopup
        If mfrmMain.mFunListActived Then
            Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��;��������;ȡ������")
        Else
            Control.Visible = ((bytActiveGrid = 1 Or bytActiveGrid = 2) And HavePrivs(mstrPrivs, "���ﰲ��")) _
                Or (bytActiveGrid = 3 And (HavePrivs(mstrPrivs, "��������;��ʱ���ﰲ��;ͣ��;����;�Ӻ�;����;������������;����ԤԼ�Һ�")))
        End If
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddMonthPlan, conMenu_Edit_AddWeekPlan '�ƶ��³����,�ƶ��ܳ����
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_PublishPlan, conMenu_Edit_UnPublishPlan '��������,ȡ������
        Control.Visible = mfrmMain.mFunListActived And HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_PublishPlan, "��������", conMenu_Edit_UnPublishPlan, "ȡ������"))
        If blnEnabled Then
            If Control.ID = conMenu_Edit_PublishPlan Then
                blnEnabled = Val(lblPublishInfo.Tag) = 0
            Else
                blnEnabled = Val(lblPublishInfo.Tag) = 1
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled

    Case conMenu_Edit_PlanAdd '��ʱ����
        Control.Visible = HavePrivs(mstrPrivs, "��ʱ���ﰲ��") And mfrmMain.mFunListActived = False _
            And (bytActiveGrid = 1 Or bytActiveGrid = 2)
        lng����ID = vsfGrid.TextMatrix(vsfGrid.Row, COL_����ID)
        If blnEnabled Then blnEnabled = lng����ID <> 0
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AddNewSignalSource '������Դ����
        Control.Visible = HavePrivs(mstrPrivs, "��������") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And (bytActiveGrid = 1 Or bytActiveGrid = 2)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_PlanVerify '��ʱ�������
        Control.Visible = HavePrivs(mstrPrivs, "�����ʱ�̶�����") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And (bytActiveGrid = 1 Or bytActiveGrid = 2)
        lng����ID = vsfGrid.TextMatrix(vsfGrid.Row, COL_����ID)
        If blnEnabled Then blnEnabled = lng����ID <> 0
        If blnEnabled Then blnEnabled = IsVerified(vsfGrid) = False
        Control.Enabled = Control.Visible And blnEnabled
        If bytActiveGrid = 2 Then
            Control.Caption = "�����ʱ�̶�����"
        ElseIf bytActiveGrid = 1 Then
            Control.Caption = "���������Դ����"
        End If
    Case conMenu_Edit_PlanCancel 'ȡ����ʱ�������
        Control.Visible = HavePrivs(mstrPrivs, "ȡ����ʱ�̶��������") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And bytActiveGrid = 2
        lng����ID = vsfGrid.TextMatrix(vsfGrid.Row, COL_����ID)
        If blnEnabled Then blnEnabled = lng����ID <> 0
        If blnEnabled Then blnEnabled = IsTempPlan(vsfGrid)
        If blnEnabled Then blnEnabled = IsVerified(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ModifyPlanItem '����������
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = vsfGrid.Col >= gPlanGrid_FixedCols
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ModifyUnitRegist '����ԤԼ�Һſ���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = Is��ֹԤԼ(vsfGrid) = False
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AllStartNO, conMenu_Edit_AllStopNO 'ȫ��������ſ���,ȫ��ȡ����ſ���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And Val(lblPublishInfo.Tag) = 0 And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_CopyPlan '���ư���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_PastPlan 'ճ������
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled And mlngCopyPlanID <> 0
    Case conMenu_Edit_ClearCurPlan '�����ǰ����
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ClearAllPlan, conMenu_Edit_ClearTempPlan '�����ǰ��Դ���а���,�����ǰ��ʱ����
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived = False
        If Control.Visible Then
            If Control.ID = conMenu_Edit_ClearAllPlan Then
                Control.Visible = bytActiveGrid = 1
            Else
                Control.Visible = bytActiveGrid = 2
            End If
        End If
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If Val(lblPublishInfo.Tag) = 0 And blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ClearAll '������к�Դ����
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False And Val(lblPublishInfo.Tag) = 0
        Control.Enabled = Control.Visible And blnEnabled

    '�ѷ������ŵ���
    Case conMenu_Edit_LockResource, conMenu_Edit_UnLockResource '����,����
        Control.Visible = HavePrivs(mstrPrivs, "��������") And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then
            If Control.ID = conMenu_Edit_LockResource Then
                blnEnabled = (PlanIsSelOne(vsfGrid) = False Or PlanIsLocked(vsfGrid) = False)
            Else
                blnEnabled = (PlanIsSelOne(vsfGrid) = False Or PlanIsLocked(vsfGrid))
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_AddTempPlan, conMenu_Edit_UpdatePlan '��ʱ����,����������İ���
        Control.Visible = HavePrivs(mstrPrivs, "��������") And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If Control.ID = conMenu_Edit_UpdatePlan And blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfGrid)
        If Control.ID = conMenu_Edit_UpdatePlan And blnEnabled Then blnEnabled = PlanIsValid(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_StopOutCall, conMenu_Edit_UnStopOutCall, conMenu_Edit_OpenStopPlan 'ͣ��,ȡ��ͣ��,����ͣ�ﰲ��
        Control.Visible = HavePrivs(mstrPrivs, "ͣ��") And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfGrid)
        If blnEnabled Then
            If Control.ID = conMenu_Edit_StopOutCall Then
                blnEnabled = (PlanIsStopVisit(vsfGrid) = False)
            Else
                blnEnabled = PlanIsStopVisit(vsfGrid)
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_ModifyDoctor, conMenu_Edit_UnModifyDoctor '����,ȡ������
        Control.Visible = HavePrivs(mstrPrivs, "����") And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsStopVisit(vsfGrid) = False
        If blnEnabled Then
            If Control.ID = conMenu_Edit_ModifyDoctor Then
                blnEnabled = (PlanIsReplaceDoctor(vsfGrid) = False)
            Else
                blnEnabled = PlanIsReplaceDoctor(vsfGrid)
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_AddNumberLimit, conMenu_Edit_ReduceNumberLimit, _
        conMenu_Edit_ModifyDoctorOffice, conMenu_Edit_UpdateUnitRegist '�Ӻ�,����,������������,����ԤԼ�Һſ���
        Control.Visible = HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_AddNumberLimit, "�Ӻ�", conMenu_Edit_ReduceNumberLimit, "����", _
            conMenu_Edit_ModifyDoctorOffice, "������������", conMenu_Edit_UpdateUnitRegist, "����ԤԼ�Һ�")) _
            And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsStopVisit(vsfGrid) = False
        If Control.ID = conMenu_Edit_UpdateUnitRegist And blnEnabled Then blnEnabled = Is��ֹԤԼ(vsfGrid) = False
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_PrintPlan '    ��ӡ�����
        Control.Visible = HavePrivs(mstrPrivs, "�̶������")
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_View_FindType '���ҷ�ʽ
        Control.Caption = "��" & Decode(mintFindType, 0, "����", 1, "����", 2, "ҽ��", "����") & "���ˡ�"
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '���ҷ�ʽ
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
    Case conMenu_View_ShowDoctorStopPlan '��ʾҽ��ͣ�ﰲ��
        Control.Visible = mfrmMain.mFunListActived = False
        blnEnabled = False
        If vsfGrid.Row >= vsfGrid.FixedRows Then
            blnEnabled = Trim(vsfGrid.Cell(flexcpData, vsfGrid.Row, COL_ҽ��)) <> ""
        End If
        Control.Enabled = Control.Visible And blnEnabled
    End Select
End Sub

Public Sub InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub

    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
        With CommandBar.Controls
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "����(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "����(&2)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "ҽ��(&3)"
            End If
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frmStopVisitAndModifyDoctor As frmClinicPlanStopVisitAndModifyDoctor
    Dim frmOfficeAndUnitRegModify As frmClinicPlanOfficeAndUnitRegModify
    Dim frmNumberLimitModify As frmClinicPlanNumberLimitModify
    Dim frmEdit As frmClinicPlanEdit
    Dim lng��¼ID As Long, lng��ԴId As Long, lng����ID As Long, str���� As String, strItem As String
    Dim obj�����¼ As �����¼, obj�����Դ As �����Դ
    Dim bytActiveGrid As Byte '��ǰ������
    Dim strIDs As String, lngRowStart As Long, lngRowEnd As Long, i As Integer
    Dim lngCurCol As Long, str����IDs As String
    Dim lng����ID As Long, strTemp As String
    Dim strDoctorName As String, vsfGrid As VSFlexGrid
    Dim str��¼IDs As String

    Err = 0: On Error GoTo errHandler
    If Not tbPage.Selected Is Nothing Then
        If tbPage.Selected.index = Pg_������� Then
            If Me.ActiveControl Is vsfRegistRuleSub Then
                bytActiveGrid = 2
                Set vsfGrid = vsfRegistRuleSub
            Else
                bytActiveGrid = 1
                Set vsfGrid = vsfRegistRule
            End If
        Else
            bytActiveGrid = 3
            Set vsfGrid = vsfRegistPlan
        End If
        
        With vsfGrid
            lng��ԴId = Val(.TextMatrix(.Row, COL_��ԴID))
            str���� = Trim(.TextMatrix(.Row, COL_����))
            If bytActiveGrid = 3 Then
                '�洢�ˡ�����ID,����ID��
                strTemp = .Cell(flexcpData, .Row, GetPlanItemNameCol(.Col) + 2)
                If InStr(strTemp, ",") > 0 Then
                    lng����ID = Val(Split(strTemp, ",")(0))
                    lng����ID = Val(Split(strTemp, ",")(1))
                Else
                    lng����ID = mlng����ID
                    lng����ID = Val(.TextMatrix(.Row, COL_����ID))
                End If
            Else
                lng����ID = Val(.TextMatrix(.Row, COL_����ID))
            End If
            lng��¼ID = Val(.Cell(flexcpData, .Row, GetPlanItemNameCol(.Col)))
            strItem = .Cell(flexcpData, 0, .Col)
            strDoctorName = Trim(.Cell(flexcpData, .Row, COL_ҽ��))
        End With
    End If

    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_ModifyPlanItem '��������
        If (bytActiveGrid = 1 Or bytActiveGrid = 2) And (lng��ԴId <> 0 Or lng����ID <> 0) Then
            Set frmEdit = New frmClinicPlanEdit
            If frmEdit.ShowMe(Me, 0, IIf(bytActiveGrid = 1, Fun_Update, Fun_TempPlan), mlng����ID, lng��ԴId, lng����ID, strItem, mstrPrivs) Then
                If bytActiveGrid = 1 Then
                    Call RefreshOneData
                Else
                    Call RefreshDataSub
                End If
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ModifyUnitRegist '����ԤԼ�Һſ���
        If lng��ԴId <> 0 Or lng����ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            Call frmEdit.ShowMe(Me, 0, Fun_UpdateUnit, mlng����ID, lng��ԴId, lng����ID, strItem, mstrPrivs)
        End If
    Case conMenu_Edit_AllStartNO 'ȫ��������ſ���
        If MsgBox("��ȷ��Ҫ�Ե�ǰ�����������޺Ż���Լ�ĺ���������ſ�����", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
            Exit Sub
        End If
        Call ZlBatchSNControl(mlng����ID, True, IIf(HavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID))
    Case conMenu_Edit_AllStopNO 'ȫ��ȡ����ſ���
        If MsgBox("��ȷ��Ҫ�Ե�ǰ�����������޺Ż���Լ�ĺ���ȡ����ſ�����", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
            Exit Sub
        End If
        Call ZlBatchSNControl(mlng����ID, False, IIf(HavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID))
    Case conMenu_Edit_CopyPlan '���ư���
        If lng����ID <> 0 And strItem <> "" Then
            mlngCopyPlanID = lng����ID
            mstrCopyPlanItem = strItem
        End If
    Case conMenu_Edit_PastPlan 'ճ������
        If PastPlan(mlng����ID, mlngCopyPlanID, mstrCopyPlanItem) Then
            If bytActiveGrid = 1 Then
                Call RefreshOneData
            Else
                Call RefreshDataSub
            End If
        End If
    Case conMenu_Edit_ClearCurPlan '�����ǰ����
        If strItem = "" Then Exit Sub
        If bytActiveGrid = 2 Then '��ʱ����
            Dim rsFixedRecord As ADODB.Recordset
            With vsfRegistRuleSub
                Set rsFixedRecord = GetԤԼ�Һż�¼(lng��ԴId, _
                    CDate(Format(.TextMatrix(.Row, COL_��ʼʱ��), "yyyy-mm-dd")), CDate(Format(.TextMatrix(.Row, COL_��ֹʱ��), "yyyy-mm-dd")))
            End With
            If Not rsFixedRecord Is Nothing Then
                Do While Not rsFixedRecord.EOF
                    If Nvl(rsFixedRecord!������Ŀ) = strItem Then
                        MsgBox "��ǰ��Դ�ڸ���ʱ������Чʱ�䷶Χ�ڵġ�" & strItem & "������ԤԼ�Һż�¼��" & _
                            "��" & strItem & "���İ����ڸ���ʱ�����б���̶�������ճ����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    rsFixedRecord.MoveNext
                Loop
            End If
        End If
        If MsgBox("��ȷ��Ҫ�������Ϊ��" & str���� & "����" & strItem & "���İ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlan(lng����ID, strItem, False) Then
            If bytActiveGrid = 1 Then
                Call RefreshOneData
            Else
                Call RefreshDataSub
            End If
            If mlngCopyPlanID = lng����ID And mstrCopyPlanItem = strItem Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ClearAllPlan '�����ǰ��Դ���а���
        If MsgBox("��ȷ��Ҫ�������Ϊ��" & str���� & "�������а�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlanBatch(mlng����ID, lng��ԴId, , , _
            Val(lblPublishInfo.Tag) = 1 And Val(vsfRegistRule.TextMatrix(vsfRegistRule.Row, COL_�Ƿ����)) = 0) Then
            Call RefreshOneData
            If mlngCopyPlanID = lng����ID Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ClearTempPlan '�����ǰ��ʱ����
        If MsgBox("��ȷ��Ҫ�������Ϊ��" & str���� & "���ĵ�ǰ��ʱ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlanBatch(mlng����ID, lng��ԴId, , lng����ID, True) Then
            If mlngCopyPlanID = lng����ID Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
            Call RefreshDataSub
        End If
    Case conMenu_Edit_ClearAll '������к�Դ����
        If MsgBox("��ȷ��Ҫ������к�Դ�İ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If ZlClearPlanBatch(mlng����ID, 0, IIf(HavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID)) Then
            Call RefreshData(mlng����ID)
            mlngCopyPlanID = 0: mstrCopyPlanItem = ""
        End If
    Case conMenu_Edit_PublishPlan '��������
        If PublishPlan(mlng����ID, True) Then
            Call PrintPlan(mlng����ID)
            'ˢ������
            Call mfrmMain.NodeChanged("K1_" & mlng����ID) '�̶������ڵ㣺K1_����ID
        End If
    Case conMenu_Edit_UnPublishPlan 'ȡ������
       If PublishPlan(mlng����ID, False) Then
            'ˢ������
            Call mfrmMain.NodeChanged("K1_" & mlng����ID) '�̶������ڵ㣺K1_����ID
        End If

    Case conMenu_Edit_PlanAdd '��ʱ����
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 0, Fun_TempPlan, mlng����ID, lng��ԴId, , strItem, mstrPrivs) Then
            Call RefreshDataSub 'ˢ������
        End If
    '�ѷ������ŵ���
    Case conMenu_Edit_AddNewSignalSource '������Դ����
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 0, Fun_AddSignalSourcePlan, mlng����ID, , , strItem, mstrPrivs) Then
            Call RefreshData(mlng����ID)
            With vsfRegistRule
                If .Rows > .FixedRows And .Cols > .FixedCols Then
                    .ShowCell .Rows - 1, .Col '������ʾ��ָ����Ԫ
                End If
            End With
        End If
    Case conMenu_Edit_PlanVerify '��ʱ�������
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 0, Fun_TempPlanVerify, mlng����ID, lng��ԴId, lng����ID, strItem, mstrPrivs) Then
            'ˢ�³������
            If bytActiveGrid = 1 Then
                Call RefreshOneData
            End If
            Call RefreshDataSub
            '�л�ҳǩʱ����ˢ�³����¼����ʹ��RefreshOneData()����ˢ������Ϊ���ܵ������շ���Ŀ�����¶�һ������
            tbPage(Pg_���ﰲ��).Tag = "0"
        End If
    Case conMenu_Edit_PlanCancel 'ȡ����ʱ�������
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 0, Fun_TempPlanCancel, mlng����ID, lng��ԴId, lng����ID, strItem, mstrPrivs) Then
            'ˢ�³������
            Call RefreshDataSub
            '�л�ҳǩʱ����ˢ�³����¼����ʹ��RefreshOneData()����ˢ������Ϊ���ܵ������շ���Ŀ������һ������
            tbPage(Pg_���ﰲ��).Tag = "0"
        End If
    Case conMenu_Edit_LockResource '����
        Call LockPlan(False)
    Case conMenu_Edit_UnLockResource '����
        Call LockPlan(True)
    Case conMenu_Edit_AddTempPlan '��ʱ����
        Set frmEdit = New frmClinicPlanEdit

        If CheckCanTempVisit(lng����ID, strItem) = False Then Exit Sub
        If frmEdit.ShowMe(Me, 1, Fun_TempPlanRecord, lng����ID, lng��ԴId, lng����ID, strItem, mstrPrivs) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_UpdatePlan '����������İ���
        Call LockPlanByDay(False, str��¼IDs)
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 1, Fun_UpdatePlan, lng����ID, lng��ԴId, lng����ID, strItem, mstrPrivs) Then
            Call LockPlanByDay(True, str��¼IDs)
            Call RefreshOneData(True)
        End If
        Call LockPlanByDay(True, str��¼IDs)
    Case conMenu_Edit_StopOutCall 'ͣ��
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        If lng��¼ID = 0 Then Exit Sub
        'bytFun ���ܣ�1-ͣ��,2-ȡ��ͣ��,3-����,4-ȡ������
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 1, lng��¼ID) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_UnStopOutCall 'ȡ��ͣ��
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        If lng��¼ID = 0 Then Exit Sub
        'bytFun ���ܣ�1-ͣ��,2-ȡ��ͣ��,3-����,4-ȡ������
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 2, lng��¼ID) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_OpenStopPlan '����ͣ�ﰲ��
        'zlOpenStopedPlanBySN(ByVal frmMain As Object, ByVal lngModule As Long, _
            Optional ByVal lng��¼ID As Long, _
            Optional ByVal lngDeptID As Long, Optional ByVal lngDoctorID As Long) As Boolean
        '���ܣ�����������ſ��Ʒ�ʱ�ε���ͣ�ﰲ�Ű���ſ��ŹҺ�
        '��Σ�
        '   frmMain ���õ�������
        '   lngModule ����ģ���
        '   lng��¼ID ��¼ID,1114ģ�����ʱ����
        '   lngDeptID ����ID
        '   lngDoctorID ҽ��ID
        '���أ��ɹ�����True��ʧ�ܷ���False
        If lng��¼ID <> 0 And Not gobjRegist Is Nothing Then
            gobjRegist.zlOpenStopedPlanBySN Me, mlngModule, lng��¼ID
        End If
    Case conMenu_Edit_ModifyDoctor '����
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        If lng��¼ID = 0 Then Exit Sub
        'bytFun ���ܣ�1-ͣ��,2-ȡ��ͣ��,3-����,4-ȡ������
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 3, lng��¼ID) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_UnModifyDoctor 'ȡ������
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        If lng��¼ID = 0 Then Exit Sub
        'bytFun ���ܣ�1-ͣ��,2-ȡ��ͣ��,3-����,4-ȡ������
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 4, lng��¼ID) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_AddNumberLimit '�Ӻ�
        If Get�����¼(lng��ԴId, lng��¼ID, True, obj�����Դ, obj�����¼) Then
            Set frmNumberLimitModify = New frmClinicPlanNumberLimitModify
            If frmNumberLimitModify.ShowMe(Me, 1, obj�����Դ, obj�����¼) Then
                Call RefreshOneData(True)
            End If
        End If
    Case conMenu_Edit_ReduceNumberLimit '����
        If Get�����¼(lng��ԴId, lng��¼ID, True, obj�����Դ, obj�����¼) Then
            Set frmNumberLimitModify = New frmClinicPlanNumberLimitModify
            If frmNumberLimitModify.ShowMe(Me, 2, obj�����Դ, obj�����¼) Then
                Call RefreshOneData(True)
            End If
        End If
    Case conMenu_Edit_ModifyDoctorOffice '������������
        If Get�����¼(lng��ԴId, lng��¼ID, True, obj�����Դ, obj�����¼) Then
            Set frmOfficeAndUnitRegModify = New frmClinicPlanOfficeAndUnitRegModify
            Call frmOfficeAndUnitRegModify.ShowMe(Me, 1, obj�����Դ, obj�����¼, True)
        End If
    Case conMenu_Edit_UpdateUnitRegist '����ԤԼ�Һſ���
        If Get�����¼(lng��ԴId, lng��¼ID, True, obj�����Դ, obj�����¼) Then
            Set frmOfficeAndUnitRegModify = New frmClinicPlanOfficeAndUnitRegModify
            Call frmOfficeAndUnitRegModify.ShowMe(Me, 2, obj�����Դ, obj�����¼, True)
        End If
    Case conMenu_Edit_PrintPlan '    ��ӡ�����
        Call PrintPlan(mlng����ID, 1)
    Case conMenu_View_PlanChangeInfo '��ѯ��Ϣ
        Dim frmPlanChangeHistory As New frmClinicPlanChangeHistory
        frmPlanChangeHistory.ShowMe Me, mlngModule
    Case conMenu_View_Refresh
        '104266����ԤԼ�������ܱ��ˣ�ˢ��ʱҪ���³�ʼ����
        If tbPage.Selected.index = Pg_���ﰲ�� Then
            Call RefreshData(mlng����ID, True, True)
        Else
            Call RefreshData(mlng����ID)
        End If
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 3 '���ҷ�ʽ
        mintFindType = Val(Right(Control.ID, 2)) - 1
        mcbsMain.RecalcLayout
        txtFind(1).Text = ""
        If txtFind(1).Visible And txtFind(1).Enabled Then txtFind(1).SetFocus
    Case conMenu_View_ShowStoped '�Ƿ���ʾ��Ч��ʱ����
        Control.Checked = Not Control.Checked
        mblnShowInvalidPlan = Control.Checked
        Call zlDatabase.SetPara("��ʾ��Ч��ʱ����", IIf(mblnShowInvalidPlan, "1", "0"), glngSys, mlngModule)
        Call RefreshDataSub 'ˢ������
    Case conMenu_View_ShowDoctorStopPlan '��ʾҽ��ͣ�ﰲ��
        If strDoctorName <> "" Then
            Dim frmDoctorStopVisit As New frmClinicPlanStopVisitManage
            frmDoctorStopVisit.ShowDoctorStopVisit Me, strDoctorName
        End If
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub PrintPlan(ByVal lng����ID As Long, Optional ByVal bytMode As Byte)
    '��ӡ�����
    '��Σ�
    '   bytMode 0-�������ӡ,1-�˵�ѡ���ӡ
    Err = 0: On Error GoTo errHandler
    If bytMode = 1 Then '��ֹ�����
        If MsgBox("Ҫ��ӡ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If

    If gVisitPlan_ModulePara.byt������ӡ��ʽ = 1 Or bytMode = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_1", Me, "����ID=" & mlng����ID, 2)
    ElseIf gVisitPlan_ModulePara.byt������ӡ��ʽ = 2 Then
        If MsgBox("Ҫ��ӡ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_1", Me, "����ID=" & mlng����ID, 2)
        End If
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecuteFilter(Optional ByVal blnOnlyRefrashRule As Boolean, _
    Optional ByVal blnOnlyRefrashRecord As Boolean)
    '��������
    Dim strKey As String

    Err = 0: On Error GoTo errHandler
    Call zlControl.TxtSelAll(txtFind(1))

    Screen.MousePointer = vbHourglass
    If blnOnlyRefrashRecord = False Then
        If Not mrsRuleRecords Is Nothing Then
            With mrsRuleRecords
                If Trim(txtFind(1).Text) = "" Then
                    .Filter = ""
                Else
                    strKey = Replace(gstrLike, "%", "*") & UCase(txtFind(1).Text) & "*"
                    Select Case mintFindType
                    Case 0   '����
                        .Filter = "���� Like '" & strKey & "'"
                    Case 1   '����(����)
                        .Filter = "���� Like '" & strKey & "' Or ���Ҽ��� Like '" & strKey & "'"
                    Case 2   'ҽ��(����)
                        .Filter = "ҽ������ Like '" & strKey & "' Or ҽ������ Like '" & strKey & "'"
                    Case Else
                        .Filter = ""
                    End Select
                End If
            End With
        End If
        Call LoadPlanDataByRecordset(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecords, 0, mlngSignalCount)
    End If
    
    If blnOnlyRefrashRule = False Then
        If Not mrsPlanRecords Is Nothing Then
            With mrsPlanRecords
                If Trim(txtFind(1).Text) = "" Then
                    .Filter = ""
                Else
                    strKey = Replace(gstrLike, "%", "*") & UCase(txtFind(1).Text) & "*"
                    Select Case mintFindType
                    Case 0   '����
                        .Filter = "���� Like '" & strKey & "'"
                    Case 1   '����(����)
                        .Filter = "���� Like '" & strKey & "' Or ���Ҽ��� Like '" & strKey & "'"
                    Case 2   'ҽ��(����)
                        .Filter = "ҽ������ Like '" & strKey & "' Or ҽ������ Like '" & strKey & "'"
                    Case Else
                        .Filter = ""
                    End Select
                End If
            End With
        End If
        Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, 0, , , Val(lblPublishInfo.Tag) = 1)
    End If
    If mintFindType = 8 Then mintFindType = 0 '���
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LockPlan(ByVal blnUnlock As Boolean) As Boolean
    '����/���������¼
    '��Σ�
    '   blnUnlock �Ƿ����,True-����,False-����
    Dim str��¼ID As String
    Dim lngRowStart As Long, lngRowEnd As Long '��ʼ�к���ֹ��
    Dim lngColStart As Long, lngColEnd As Long '��ʼ�к���ֹ��
    Dim i As Long, j As Long, lngTemp As Long
    Dim cll��¼ID As Collection

    Err = 0: On Error GoTo errHandler
    Set cll��¼ID = New Collection
    With vsfRegistPlan
        'ѡ���з�Χ
        lngRowStart = .Row: lngRowEnd = .RowSel
        If lngRowStart > lngRowEnd Then lngTemp = lngRowStart: lngRowStart = lngRowEnd: lngRowEnd = lngTemp

        'ѡ���з�Χ
        lngColStart = GetPlanItemNameCol(.Col) 'ȷ��"ʱ���"��
        lngColEnd = GetPlanItemNameCol(.ColSel)
        If lngColStart > lngColEnd Then lngTemp = lngColStart: lngColStart = lngColEnd: lngColEnd = lngTemp

        For i = lngRowStart To lngRowEnd
            For j = lngColStart To lngColEnd Step 3
                If PlanIsLocked(vsfRegistPlan, i, j) = blnUnlock Then
                    If Val(.Cell(flexcpData, i, j)) <> 0 Then
                        If zlStr.ActualLen(str��¼ID & "," & Val(.Cell(flexcpData, i, j))) >= 4000 Then
                            cll��¼ID.Add Mid(str��¼ID, 2)
                            str��¼ID = ""
                        End If
                        str��¼ID = str��¼ID & "," & Val(.Cell(flexcpData, i, j))
                    End If
                End If
            Next
        Next
        If str��¼ID <> "" Then
            cll��¼ID.Add Mid(str��¼ID, 2)
        End If
        If cll��¼ID.Count = 0 Then
            MsgBox "��ǰû��ѡ����Ҫ" & IIf(blnUnlock, "����", "����") & "�ĳ��ﰲ�ţ�", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    gcnOracle.BeginTrans
    For i = 1 To cll��¼ID.Count
        If ZlBatchLockPlan(cll��¼ID(i), blnUnlock) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    Next
    gcnOracle.CommitTrans
    LockPlan = True

    'ˢ�½���
    If LockPlan Then
        For i = lngRowStart To lngRowEnd
            Call RefreshOneData(True, i, i = lngRowStart)
        Next
    End If
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LockPlanByDay(ByVal blnUnlock As Boolean, ByRef str��¼IDs As String) As Boolean
    '����/����ĳһ��ĳ����¼
    '��Σ�
    '   blnUnlock �Ƿ����,True-����,False-����
    '˵����
    '   ������ʱ�������ʱ�ļ�¼ID
    Dim lngRowStart As Long, lngRowEnd As Long '��ʼ�к���ֹ��
    Dim lngCol  As Long, i As Long

    Err = 0: On Error GoTo errHandler
    If blnUnlock Then
        LockPlanByDay = ZlBatchLockPlan(str��¼IDs, True)
        Exit Function
    End If

    With vsfRegistPlan
        'ѡ���з�Χ
        GetPlanGroupRange vsfRegistPlan, .Row, lngRowStart, lngRowEnd
        lngCol = GetPlanItemNameCol(.Col)  'ȷ��"ʱ���"��

        For i = lngRowStart To lngRowEnd
            If PlanIsLocked(vsfRegistPlan, i, lngCol) = False Then
                If Val(.Cell(flexcpData, i, lngCol)) <> 0 Then
                    str��¼IDs = str��¼IDs & "," & Val(.Cell(flexcpData, i, lngCol))
                End If
            End If
        Next
    End With
    If str��¼IDs = "" Then Exit Function
    str��¼IDs = Mid(str��¼IDs, 2)
    
    LockPlanByDay = ZlBatchLockPlan(str��¼IDs, False)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckCanTempVisit(ByVal lng����ID As Long, ByVal strCurDate As String) As Boolean
    '��鵱ǰ��Դ�Ƿ�ɽ�����ʱ����
    Dim strSQL As String, rsTemp As ADODB.Recordset

    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Or IsDate(strCurDate) = False Then Exit Function
    strSQL = "Select 1" & vbNewLine & _
            " From �ٴ����ﰲ�� A, �ٴ������Դ B" & vbNewLine & _
            " Where a.��Դid = b.Id And a.ID = [1]" & vbNewLine & _
            "       And ([2] Between a.��ʼʱ�� And a.��ֹʱ��" & vbNewLine & _
            "       Or ([2] > a.��ֹʱ�� Or [2] < a.��ʼʱ��) And Nvl(b.�Ƿ�ɾ��, 0) = 0" & vbNewLine & _
            "           And (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And b.�Ű෽ʽ = 0)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Դ�Ƿ����ʱ����", lng����ID, CDate(strCurDate))
    If rsTemp.EOF Then
        MsgBox "�ú�Դ�����ѱ�ͣ�û�ǰ���ڰ�����������������У�����ͨ����ǰ����������ʱ���", vbInformation, gstrSysName
        Exit Function
    End If
    CheckCanTempVisit = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function PublishPlan(ByVal lng����ID As Long, ByVal blnPublish As Boolean) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim obj�ϰ�ʱ�� As �ϰ�ʱ��
    
    Err = 0: On Error GoTo errHandler
    If MsgBox("��ȷ��Ҫ" & IIf(blnPublish, "", "ȡ��") & "������ǰ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    If blnPublish Then
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ��������� B, �ٴ������ C" & vbNewLine & _
                " Where a.Id = b.����id And a.����id = c.Id And c.�Ű෽ʽ = 0 And c.Id = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If rsTemp.EOF Then
            MsgBox "��ǰ���������Ч�İ��ţ����ܷ�����", vbInformation, gstrSysName: Exit Function
        End If
        
        strSQL = "Select Distinct d.����, d.����, e.վ��, b.������Ŀ, b.�ϰ�ʱ��, To_Char(c.��ʼʱ��, 'hh24:mi:ss') As ��ʼʱ��, " & vbNewLine & _
                "       To_Char(c.��ֹʱ��, 'hh24:mi:ss') As ��ֹʱ��" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ��������� B, �ٴ�����ʱ�� C, �ٴ������Դ D, ���ű� E" & vbNewLine & _
                " Where a.Id = b.����id And b.Id = c.����id And a.��Դid = d.Id And d.����id = e.Id" & vbNewLine & _
                "       And c.��� = 1 And ����id = [1]" & vbNewLine & _
                " Order By " & IIf(gVisitPlan_ModulePara.byt����ȽϷ�ʽ = 0, "d.����", "Lpad(d.����,5,'0')")
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ϰ�ʱ������ŷ�ʱ��ʱ���Ƿ�һ��", lng����ID)
        Do While Not rsTemp.EOF
            Set obj�ϰ�ʱ�� = GetWorkTimeRange(Nvl(rsTemp!�ϰ�ʱ��), Nvl(rsTemp!վ��), Nvl(rsTemp!����))
            If Format(obj�ϰ�ʱ��.��ʼʱ��, "hh:mm:00") <> Format(Nvl(rsTemp!��ʼʱ��), "hh:mm:00") Then
                If MsgBox("��ǰ������в��ַ�ʱ�εİ��Ų��Ǹ����ϰ�ʱ�ε�ʱ����зֶεģ��磺" & vbCrLf & _
                    "����Ϊ " & Nvl(rsTemp!����) & " ��" & Nvl(rsTemp!������Ŀ) & " " & Nvl(rsTemp!�ϰ�ʱ��) & _
                    "[" & Format(obj�ϰ�ʱ��.��ʼʱ��, "hh:mm") & "-" & Format(obj�ϰ�ʱ��.����ʱ��, "hh:mm") & "]��" & _
                    "��һ�����ʱ��Ϊ[" & Format(Nvl(rsTemp!��ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsTemp!��ֹʱ��), "hh:mm") & "])" & vbCrLf & vbCrLf & _
                    "�Ƿ���Ҫ����������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Exit Do
            End If
            rsTemp.MoveNext
        Loop

        'Zl_�ٴ����ﰲ��_Publish
        strSQL = "Zl_�ٴ����ﰲ��_Publish("
        '  Id_In       �ٴ������.Id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ������_In   �ٴ������.������%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ����ʱ��_In �ٴ������.����ʱ��%Type := Null,
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
        '  ȡ������_In Number:=0
        Screen.MousePointer = vbHourglass
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        PublishPlan = True

        '�Զ������ٴ������¼
        'Zl1_Auto_Buildingregisterplan
        '  --����˵�����Զ������ٴ������¼
        '  --          1�����ݺ�Դ�Զ�����ԤԼ���ڵ��ٴ������¼;
        '  --          2��ԤԼ������ȷ��:��ԴԤԼ����-->ԤԼ��ʽ��������ȡ���)-->ϵͳԤԼ����
        '  --���:�Һ�ʱ��_IN:NULLʱ���Զ�����;����ֻ���ָ�������Ƿ������˳����¼û��
        strSQL = "Zl1_Auto_Buildingregisterplan("
        '    �Һ�ʱ��_In In Date := Null
        strSQL = strSQL & "" & "NULL" & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Screen.MousePointer = vbDefault
    Else
        strSQL = "Select 1" & vbNewLine & _
                " From ���˹Һż�¼ C, �ٴ������¼ A, �ٴ����ﰲ�� B" & vbNewLine & _
                " Where c.�����¼id = a.Id And a.����id = b.Id And b.����id = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ�����İ����ѱ�ʹ�ã�������ȡ��������", vbInformation, gstrSysName: Exit Function
        End If

        'Zl_�ٴ����ﰲ��_Publish
        strSQL = "Zl_�ٴ����ﰲ��_Publish("
        '  Id_In       �ٴ������.Id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ������_In   �ٴ������.������%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ����ʱ��_In �ٴ������.����ʱ��%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ȡ������_In Number:=0
        strSQL = strSQL & "" & 1 & ")"
        Screen.MousePointer = vbHourglass
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Screen.MousePointer = vbDefault
    End If
    PublishPlan = True
    Exit Function
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function PastPlan(ByVal lng����ID As Long, ByVal lngԭ����ID As Long, ByVal strԭ��Ŀ As String) As Long
    '���ܣ�ճ������
    '������
    Dim strSQL As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim rsPlan As ADODB.Recordset, rsSignalSource As ADODB.Recordset
    Dim blnTran As Boolean, blnNoPlan As Boolean
    Dim lng����ID  As Long, lng��ԴId As Long, strApplyItem  As String
    Dim dtStart  As Date, dtEnd As Date, str���� As String
    Dim rsFixedRecord As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Then Exit Function
    If lngԭ����ID = 0 Then Exit Function
    If strԭ��Ŀ = "" Then Exit Function

    If tbPage.Selected Is Nothing Then Exit Function
    If tbPage.Selected.index <> Pg_������� Then Exit Function

    If Me.ActiveControl Is vsfRegistRule Then
        With vsfRegistRule
            lng��ԴId = Val(.TextMatrix(.Row, COL_��ԴID))
            lng����ID = Val(.TextMatrix(.Row, COL_����ID))
            strApplyItem = .Cell(flexcpData, 0, .Col)
            str���� = Trim(.TextMatrix(.Row, COL_����))
        End With
    Else
        With vsfRegistRuleSub
            lng��ԴId = Val(.TextMatrix(.Row, COL_��ԴID))
            lng����ID = Val(.TextMatrix(.Row, COL_����ID))
            strApplyItem = .Cell(flexcpData, 0, .Col)
            str���� = Trim(.TextMatrix(.Row, COL_����))
            Set rsFixedRecord = GetԤԼ�Һż�¼(lng��ԴId, _
                CDate(Format(.TextMatrix(.Row, COL_��ʼʱ��), "yyyy-mm-dd")), CDate(Format(.TextMatrix(.Row, COL_��ֹʱ��), "yyyy-mm-dd")))
        End With
    End If

    If lng��ԴId = 0 Then Exit Function
    If strApplyItem = "" Then Exit Function
    If lngԭ����ID = lng����ID And strԭ��Ŀ = strApplyItem Then
        MsgBox "��ǰ�����븴�ư�����ͬ������ճ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not rsFixedRecord Is Nothing Then
        Do While Not rsFixedRecord.EOF
            If Nvl(rsFixedRecord!������Ŀ) = strApplyItem Then
                MsgBox "��ǰ��Դ�ڸ���ʱ������Чʱ�䷶Χ�ڵġ�" & strApplyItem & "������ԤԼ�Һż�¼��" & _
                    "��" & strApplyItem & "���İ����ڸ���ʱ�����б���̶�������ճ����", vbInformation, gstrSysName
                Exit Function
            End If
            rsFixedRecord.MoveNext
        Loop
    End If
    
    '���ĳ���ϰ�ʱ���Ƿ������ڵ�ǰ��Դ
    strSQL = "Select c.����, a.�ϰ�ʱ��" & vbNewLine & _
            " From �ٴ��������� A,�ٴ����ﰲ�� B, �ٴ������Դ C" & vbNewLine & _
            " Where a.����ID =b.ID And b.��ԴID = c.ID And b.ID = [1] And a.������Ŀ = [2] And a.�ϰ�ʱ�� Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngԭ����ID, strԭ��Ŀ)
    Do While Not rsTemp.EOF
        If GetWorkTimeRange(Nvl(rsTemp!�ϰ�ʱ��), gstrNodeNo, str����) Is Nothing Then
            MsgBox "�ϰ�ʱ�Ρ�" & Nvl(rsTemp!�ϰ�ʱ��) & "����������" & str���� & "�ţ�����ճ����", vbInformation, gstrSysName
            Exit Function
        End If
        rsTemp.MoveNext
    Loop
    
    If Me.ActiveControl Is vsfRegistRule Then
        If Trim(vsfRegistRule.TextMatrix(vsfRegistRule.Row, vsfRegistRule.Col)) <> "" Then
            If MsgBox("��ճ�������ڵ�ǰ�Ѵ��ڳ��ﰲ�ţ�ճ�����ⲿ�ְ��Ž��ᱻ���ǣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    Else
        If Trim(vsfRegistRuleSub.TextMatrix(vsfRegistRuleSub.Row, vsfRegistRuleSub.Col)) <> "" Then
            If MsgBox("��ճ�������ڵ�ǰ�Ѵ��ڳ��ﰲ�ţ�ճ�����ⲿ�ְ��Ž��ᱻ���ǣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If

    If lng����ID = 0 Then
        '��ȡ���ŵ�ʱ�䷶Χ����ԭ����һ��
        strSQL = "Select ��ʼʱ��,��ֹʱ�� From �ٴ����ﰲ�� Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ�䷶Χ", lngԭ����ID)
        If rsTemp.EOF Then
            dtStart = Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd")
            dtEnd = CDate("3000-01-01")
        Else
            dtStart = Format(Nvl(rsTemp!��ʼʱ��, DateAdd("d", 1, zlDatabase.Currentdate)), "yyyy-mm-dd hh:mm:ss")
            dtEnd = Format(Nvl(rsTemp!��ֹʱ��, CDate("3000-01-01")), "yyyy-mm-dd hh:mm:ss")
        End If
        blnNoPlan = True

        Set rsSignalSource = GetSignalSource("", lng��ԴId)
        If rsSignalSource.EOF Then
            MsgBox "��Դ��Ϣδ�ҵ���", vbInformation, gstrSysName
            Exit Function
        End If

        lng����ID = zlDatabase.GetNextId("�ٴ����ﰲ��")
        'Zl_�ٴ����ﰲ��_Insert(
        strSQL = "Zl_�ٴ����ﰲ��_Insert("
        'Id_In           �ٴ����ﰲ��.Id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '����id_In       �ٴ����ﰲ��.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '��Դid_In       �ٴ����ﰲ��.��Դid%Type,
        strSQL = strSQL & "" & lng��ԴId & ","
        '��Ŀid_In       �ٴ����ﰲ��.��Ŀid%Type,
        strSQL = strSQL & "" & ZVal(Nvl(rsSignalSource!��ĿID)) & ","
        'ҽ��id_In       �ٴ����ﰲ��.ҽ��id%Type,
        strSQL = strSQL & "" & ZVal(Nvl(rsSignalSource!ҽ��ID)) & ","
        'ҽ������_In     �ٴ����ﰲ��.ҽ������%Type,
        strTemp = Nvl(rsSignalSource!ҽ������)
        strSQL = strSQL & "" & IIf(strTemp = "", "NULL", "'" & strTemp & "'") & ","
        '�Ű����_In     �ٴ����ﰲ��.�Ű����%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '�Ƿ���������_In �ٴ����ﰲ��.�Ƿ���������%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '�Ƿ����ճ���_In �ٴ����ﰲ��.�Ƿ����ճ���%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '��ʼʱ��_In     �ٴ����ﰲ��.��ʼʱ��%Type,
        strSQL = strSQL & "" & ZDate(dtStart) & ","
        '��ֹʱ��_In     �ٴ����ﰲ��.��ֹʱ��%Type,
        strSQL = strSQL & "" & ZDate(dtEnd) & ","
        '����Ա����_In   �ٴ����ﰲ��.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '�Ǽ�ʱ��_In     �ٴ����ﰲ��.�Ǽ�ʱ��%Type
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    End If

    blnTran = True
    gcnOracle.BeginTrans
        If blnNoPlan And strSQL <> "" Then
            zlDatabase.ExecuteProcedure strSQL, "��������"
        End If
        If ZlPlanApplyTo(0, lngԭ����ID, strԭ��Ŀ, lng����ID, strApplyItem) = False Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
    gcnOracle.CommitTrans
    blnTran = False
    PastPlan = True
    Exit Function
errHandler:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshData(Optional ByVal lng����ID As Long, Optional ByVal blnClear As Boolean, _
    Optional ByVal blnLoadRecord As Boolean)
    '���ܣ�ˢ�°�����������
    '������
    '   lng����ID - ����ID
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dtStart As Date, dtEnd As Date
    Dim lngOldRow As Long, lngoldCol As Long

    Err = 0: On Error GoTo errHandler
    If blnLoadRecord Then
        lngOldRow = vsfRegistPlan.Row: lngoldCol = vsfRegistPlan.Col
    End If
    mdtToday = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
    tbPage(Pg_���ﰲ��).Tag = ""
    If blnClear Then
        mlng����ID = lng����ID '�洢����ID
        mlngSignalCount = 0
        Set mrsRuleRecords = Nothing
        Set mrsPlanRecords = Nothing
        Set mrsRuleRecordsSub = Nothing
        mlngCopyPlanID = 0: mstrCopyPlanItem = ""

        sccTitle.Caption = "���ﰲ��>�̶������" & IIf(lng����ID = 0, "(�޳����)", "")
        lblDateRange.Visible = (tbPage.Selected.index = Pg_�������) And lng����ID <> 0
        lblDateRange.Caption = "��Чʱ�䣺" & Format(mdtToday, "yyyy-mm-dd hh:mm:ss") & "��" & "3000-01-01 00:00:00"
        lblPublishInfo.Tag = ""

        '��ʾʱ�䷶Χ
'        picPlanDateRange.Visible = (tbPage.Selected.index = Pg_���ﰲ��) And lng����ID <> 0
'        dtpStartDate.MaxDate = CDate("3000-01-01")
'        dtpStartDate.MinDate = CDate(Format(mdtToday, "yyyy-mm-dd")): dtpStartDate.MaxDate = CDate(Format(mdtToday, "yyyy-mm-dd")) + 6
'        dtpEndDate.MaxDate = CDate("3000-01-01")
'        dtpEndDate.MinDate = CDate(Format(mdtToday, "yyyy-mm-dd")): dtpEndDate.MaxDate = CDate(Format(mdtToday, "yyyy-mm-dd")) + 6
        dtpStartDate.Value = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
        dtpEndDate.Value = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd")) + GetԤԼ����(lng����ID)

'        dtStart = mdtToday
'        dtEnd = DateAdd("d", mdtToday, 7)

        strSQL = "Select b.�������, b.������, b.����ʱ�� From �ٴ������ B Where b.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������Ϣ", mlng����ID)
        If Not rsTemp.EOF Then
            sccTitle.Caption = "���ﰲ��>" & Nvl(rsTemp!�������)
            lblPublishInfo.Tag = IIf(Nvl(rsTemp!����ʱ��) = "", "", "1") '����Ƿ񷢲�
            lblPublishInfo.Caption = "�����ˣ�" & IIf(Nvl(rsTemp!������) = "", "      ", Nvl(rsTemp!������)) & _
                "  ����ʱ�䣺" & IIf(Nvl(rsTemp!����ʱ��) = "", "                   ", Format(Nvl(rsTemp!����ʱ��), "yyyy-mm-dd hh:mm:ss"))
        End If

        '�����¼
'        If txtPublisher.Caption <> "" Then
'            'ȱʡ��ʾʱ�䷶Χ
'            strSql = "Select Min(a.��������) As ��С����, Max(a.��������) As �������" & vbNewLine & _
'                    " From �ٴ������¼ A, �ٴ����ﰲ�� B" & vbNewLine & _
'                    " Where a.����id = b.Id And b.����Id = [1]"
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�����¼ʱ�䷶Χ", lng����ID)
'            If rsTemp.EOF Then Exit Sub
'
'            dtpStartDate.MaxDate = CDate("3000-01-01")
'            dtpStartDate.MinDate = CDate(Nvl(rsTemp!��С����, CDate(Format(mdtToday, "yyyy-mm-dd"))))
'            dtpStartDate.MaxDate = CDate(Nvl(rsTemp!�������, CDate(Format(mdtToday, "yyyy-mm-dd")) + 6))
'            dtpEndDate.MaxDate = CDate("3000-01-01")
'            dtpEndDate.MinDate = CDate(Nvl(rsTemp!��С����, CDate(Format(mdtToday, "yyyy-mm-dd"))))
'            dtpEndDate.MaxDate = CDate(Nvl(rsTemp!�������, CDate(Format(mdtToday, "yyyy-mm-dd")) + 6))
'
'            dtpStartDate.Value = CDate(Nvl(rsTemp!��С����, CDate(Format(mdtToday, "yyyy-mm-dd"))))
'            dtpEndDate.Value = CDate(Nvl(rsTemp!�������, CDate(Format(mdtToday, "yyyy-mm-dd")) + 6))

'            If DateDiff("d", dtpStartDate.Value, mdtToday) > 0 Then dtpStartDate.Value = CDate(Format(mdtToday, "yyyy-mm-dd"))
'        End If
        
        Call InitPlanGrid(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRule, Me.Name, "����")
        Call InitPlanGrid(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRuleSub, Me.Name, "����")
        Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, dtpStartDate.Value, dtpEndDate.Value, Val(lblPublishInfo.Tag) = 1)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "����")
        Call ShowHolidayToPlan(vsfRegistPlan, Format(dtpStartDate.Value, "yyyy-mm-dd hh:mm:ss"), Format(dtpEndDate.Value, "yyyy-mm-dd hh:mm:ss"))
    End If
    
    If lng����ID <> 0 Then
        '��������
        Screen.MousePointer = vbHourglass
        If blnLoadRecord = False Then
            tbPage.Enabled = False
            tbPage(Pg_�������).Selected = True
            tbPage.Enabled = True
        End If
        '����
        Set mrsRuleRecords = GetPlanRuleData(lng����ID, 0, Val(lblPublishInfo.Tag) = 1)
        Call ExecuteFilter(True)
        
        If blnLoadRecord Then
            tbPage(Pg_���ﰲ��).Tag = "1"
            '�����¼
            If Val(lblPublishInfo.Tag) = 1 Then
                Set mrsPlanRecords = GetPlanRecords(lng����ID, Format(dtpStartDate.Value, "yyyy-mm-dd"), Format(dtpEndDate.Value, "yyyy-mm-dd"))
                Call ExecuteFilter(False, True)
            End If
            '��λ��һ����
            With vsfRegistPlan
                If .Rows > .FixedRows And .Cols > .FixedCols Then     'ȱʡ��λ��
                    .Row = -1 '��֤��ѡ���в���������Ҳ����RowColChange�¼�
                    .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
                    .Col = IIf(lngoldCol = 0 Or lngoldCol > .Cols - 1, .FixedCols, lngoldCol)
                    .ShowCell .Row, .Col  '������ʾ��ָ����Ԫ
                End If
            End With
        End If
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshOneData(Optional ByVal blnRecord As Boolean, Optional ByVal lngCurRow As Long = -1, _
    Optional ByVal blnReLoadData As Boolean = True)
    'ˢ��ָ���к�Դ����
    '��Σ�
    '   blnRecord �Ƿ�ˢ�³����¼
    Dim lng����ID As Long, lng��ԴId As Long
    Dim strSQL  As String, rsData As ADODB.Recordset
    Dim str�շ���Ŀ As String

    Err = 0: On Error GoTo errHandle
    If blnRecord Then
        '1.��¼ԭ���ݣ�����ȡ������
        With vsfRegistPlan
            lng��ԴId = Val(.TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_��ԴID))
            str�շ���Ŀ = .TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_��Ŀ)
        End With
        
        If blnReLoadData Then
            '���±��ؼ�¼��,��Ҫ�Ǹ��±����޸ĵ�
            Set mrsPlanRecords = GetPlanRecords(mlng����ID, Format(dtpStartDate.Value, "yyyy-mm-dd"), Format(dtpEndDate.Value, "yyyy-mm-dd"))
        End If
        
        '2.���½���
        mrsPlanRecords.Filter = "��ԴID=" & lng��ԴId & " And �շ���Ŀ='" & str�շ���Ŀ & "'"
        Call RefreshOnePlanData(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, lngCurRow, _
            Val(lblPublishInfo.Tag) = 1, 0)
        mrsPlanRecords.Filter = ""
    Else
        '1.��¼ԭ���ݣ�����ȡ������
        With vsfRegistRule
            lng����ID = Val(.TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_����ID))
            lng��ԴId = Val(.TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_��ԴID))
            str�շ���Ŀ = .TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_��Ŀ)
        End With
        '�������IDΪ0����ʾ�����ɳ����������ĺ�Դ����Ҫ���»�ȡ����ID
        '��Ҫ���»�ȡ����ID����Ϊ�ڵ�������ʱ������ID�����Ѿ�����
    '    If lng����ID = 0 Then
            strSQL = "Select a.Id, a.��ʼʱ��, a.��ֹʱ�� From �ٴ����ﰲ�� A Where a.����id = [1] And a.��Դid = [2]"
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ID", mlng����ID, lng��ԴId)
            If Not rsData.EOF Then
                lng����ID = Val(Nvl(rsData!ID))
            Else
                '������ID��Ϊ�㣬���˳�
                '�����˳���Ҫ��յ�ǰ��
                'Exit Sub
            End If
    '    End If
        
        If blnReLoadData Then
            '���±��ؼ�¼��
            Set mrsRuleRecords = GetPlanRuleData(mlng����ID, 0, Val(lblPublishInfo.Tag) = 1)
        End If
        
        '2.���½���
        mrsRuleRecords.Filter = "����ID=" & lng����ID & " And �շ���Ŀ='" & str�շ���Ŀ & "'"
        Call RefreshOnePlanData(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecords, , _
            Val(lblPublishInfo.Tag) = 1, 0)
        mrsRuleRecords.Filter = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim strSQL As String

    Err = 0: On Error GoTo errHandler
'    dtpStartDate.MaxDate = CDate("3000-01-01")
'    dtpStartDate.MinDate = CDate(Format(Now, "yyyy-mm-dd")): dtpStartDate.MaxDate = CDate(Format(Now, "yyyy-mm-dd")) + 6
'    dtpEndDate.MaxDate = CDate("3000-01-01")
'    dtpEndDate.MinDate = CDate(Format(Now, "yyyy-mm-dd")): dtpEndDate.MaxDate = CDate(Format(Now, "yyyy-mm-dd")) + 6
'    dtpStartDate.Value = CDate(Format(Now, "yyyy-mm-dd"))
'    dtpEndDate.Value = CDate(Format(Now, "yyyy-mm-dd")) + 6
    
    
    mblnShowInvalidPlan = Val(zlDatabase.GetPara("��ʾ��Ч��ʱ����", glngSys, mlngModule, "0")) = 1
    
    Call InitPage

    Dim strFindType As String
    Call GetRegInFor(g˽��ģ��, Me.Name, "FindType", strFindType)
    mintFindType = Val(strFindType)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    lblPublishInfo.Move sccTitle.Width - lblPublishInfo.Width - 100, sccTitle.Top + sccTitle.Height - lblPublishInfo.Height - 50

    With tbPage
        .Left = sccTitle.Left
        .Top = sccTitle.Top + sccTitle.Height
        .Width = sccTitle.Width
        .Height = shpBorder.Height - .Top - 10
    End With
    lblDateRange.Move 2500, tbPage.Top + 40, Me.ScaleWidth - lblDateRange.Left - 16
End Sub

Private Sub InitPage()
    '����:��ʼ��ҳ��ؼ�
    Dim i As Long, ObjItem As TabControlItem

    Err = 0: On Error GoTo errHandler
    tbPage.RemoveAll
    tbPage.InsertItem mPgIndex.Pg_�������, "�������", picRegistRule.Hwnd, 0
    tbPage.InsertItem mPgIndex.Pg_���ﰲ��, "���ﰲ��", picRegistPlan.Hwnd, 0

     With tbPage.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003 '��ʾ���
        .BoldSelected = True '��ʾҳ��������Ӵ�
        .ClientFrame = xtpTabFrameSingleLine 'ҳ��߿�
        .Layout = xtpTabLayoutAutoSize
    End With
    tbPage.Enabled = False
    tbPage.Item(Pg_���ﰲ��).Selected = True
    tbPage.Item(Pg_�������).Selected = True
    tbPage.Enabled = True
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsRuleRecords = Nothing
    Set mrsPlanRecords = Nothing
    Set mrsRuleRecordsSub = Nothing

    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRule, Me.Name, "����")
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "����")
    Call SaveRegInFor(g˽��ģ��, Me.Name, "FindType", mintFindType)
End Sub

Private Sub picImgRule_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT

    vRect = zlControl.GetControlRect(picImgRule.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgRule.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfRegistRule, lngLeft, lngTop, picImgRule.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRule, Me.Name, "����")
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRuleSub, Me.Name, "����")
End Sub

Private Sub picImgRuleSub_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT

    vRect = zlControl.GetControlRect(picImgRuleSub.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgRuleSub.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfRegistRuleSub, lngLeft, lngTop, picImgRuleSub.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRuleSub, Me.Name, "����")
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRule, Me.Name, "����")
End Sub

Private Sub picRegistPlan_GotFocus()
    On Error Resume Next
    If vsfRegistPlan.Visible And vsfRegistPlan.Enabled Then vsfRegistPlan.SetFocus
End Sub

Private Sub picRegistPlan_Resize()
    On Error Resume Next
    vsfRegistPlan.Move -10, 0, picRegistPlan.ScaleWidth + 20, picRegistPlan.ScaleHeight
End Sub

Private Sub picRegistRule_GotFocus()
    On Error Resume Next
    If vsfRegistRule.Visible And vsfRegistRule.Enabled Then vsfRegistRule.SetFocus
End Sub

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If tbPage.Selected Is Nothing Then
        If vsfRegistRule.Visible And vsfRegistRule.Enabled Then vsfRegistRule.SetFocus
    Else
        If tbPage.Selected.index = Pg_������� Then
            If vsfRegistRule.Visible And vsfRegistRule.Enabled Then vsfRegistRule.SetFocus
        Else
            If vsfRegistPlan.Visible And vsfRegistPlan.Enabled Then vsfRegistPlan.SetFocus
        End If
    End If
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Err = 0: On Error GoTo errHandler
    If tbPage.ItemCount < 2 Then Exit Sub
    If tbPage.Tag = Item.Caption Then Exit Sub
    tbPage.Tag = Item.Caption
    lblDateRange.Visible = (Item.index = Pg_�������)

    If Item.index = Pg_���ﰲ�� And Val(tbPage(Pg_���ﰲ��).Tag) = 0 Then
        '�����¼
        If Val(lblPublishInfo.Tag) = 1 Then
            Screen.MousePointer = vbHourglass
            Set mrsPlanRecords = GetPlanRecords(mlng����ID, Format(dtpStartDate.Value, "yyyy-mm-dd"), Format(dtpEndDate.Value, "yyyy-mm-dd"))
            Call ExecuteFilter(False, True)
            Screen.MousePointer = vbDefault
        End If
        tbPage(Pg_���ﰲ��).Tag = "1"
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlDataPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrSysName = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    Dim vsfTemp As VSFlexGrid

    Err = 0: On Error GoTo errHandler
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objOut)
        If bytMode = 0 Then Exit Sub
    End If
    
    objOut.Title.Text = Mid(sccTitle.Caption, InStr(sccTitle.Caption, ">") + 1) & IIf(tbPage.Selected.index = Pg_�������, "����", "��¼") & "�嵥"
    If VSFlexGridCopyTo(IIf(tbPage.Selected.index = Pg_�������, vsfRegistRule, vsfRegistPlan), _
        vsfTemp, bytMode) = False Then Exit Sub
    If vsfTemp Is Nothing Then Exit Sub
    vsfTemp.ColWidth(COL_ͼ��) = 0 '����ͼ����
    Set objOut.Body = vsfTemp

    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow

    zlPrintOrView1Grd objOut, bytMode
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng��¼ID As Long
    Dim strTemp As String

    On Error Resume Next
    '��ʾͣ����Ϣ
    lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, NewRow, GetPlanItemNameCol(NewCol)))
    If lng��¼ID = 0 Or mrsPlanRecords Is Nothing Then
        strTemp = ""
    Else
        Dim strFilter As Variant
        strFilter = mrsPlanRecords.Filter
        mrsPlanRecords.Filter = "��¼ID=" & lng��¼ID
        If mrsPlanRecords.EOF Then
            strTemp = ""
        Else
            If Nvl(mrsPlanRecords!ͣ�￪ʼʱ��) = "" Then
                strTemp = ""
            Else
                strTemp = Nvl(mrsPlanRecords!�ϰ�ʱ��) & _
                " ͣ��ʱ�䣺" & Format(Nvl(mrsPlanRecords!ͣ�￪ʼʱ��), "mm-dd hh:mm") & _
                "��" & Format(Nvl(mrsPlanRecords!ͣ����ֹʱ��), "mm-dd hh:mm") & "��ͣ��ԭ��" & Nvl(mrsPlanRecords!ͣ��ԭ��)
            End If
        End If
        mrsPlanRecords.Filter = strFilter
    End If
    Call mfrmMain.StatusShowInfoChanged(2, strTemp)
End Sub

Private Sub vsfRegistPlan_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    On Error Resume Next
    Call SetPlanGridRangeColor(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mstrOldSelRangePlan)
    mstrOldSelRangePlan = vsfRegistPlan.Row & "|" & vsfRegistPlan.RowSel & "|" & vsfRegistPlan.Col & "|" & vsfRegistPlan.ColSel
End Sub

Private Sub vsfRegistPlan_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "����")
End Sub

Private Sub vsfRegistPlan_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    If Val(vsfRegistPlan.RowData(NewRow)) = -1 Then Cancel = True
End Sub

Private Sub vsfRegistPlan_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = gPlanGrid_ColIndex.COL_ͼ�� Then Cancel = True: Exit Sub
End Sub

Private Sub vsfRegistPlan_DblClick()
    Dim lng��ԴId As Long, lng����ID As Long, lng����ID As Long
    Dim frmEdit As New frmClinicPlanEdit
    Dim strCurItem As String, strTemp As String
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String

    Err = 0: On Error GoTo errHandler
    lngCol = vsfRegistPlan.MouseCol
    lngRow = vsfRegistPlan.MouseRow
    If lngRow = 0 Or lngRow = 1 Then
        '����
        If mrsPlanRecords Is Nothing Then Exit Sub
        If mrsPlanRecords.RecordCount = 0 Then Exit Sub
        strSort = GetPlanSortCircleStr(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, lngRow, lngCol)
        If strSort <> "" Then
            mrsPlanRecords.Sort = strSort
            Screen.MousePointer = vbHourglass
            Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, 0, , True, Val(lblPublishInfo.Tag) = 1)
            Screen.MousePointer = vbDefault
        End If
    Else
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        lngCol = GetPlanItemNameCol(vsfRegistPlan.Col)
        '�洢�ˡ�����ID,����ID��
        strTemp = vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, lngCol + 2)
        If InStr(strTemp, ",") > 0 Then
            lng����ID = Val(Split(strTemp, ",")(0))
            lng����ID = Val(Split(strTemp, ",")(1))
        Else
            lng����ID = mlng����ID
            lng����ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_����ID))
        End If
        strCurItem = vsfRegistPlan.Cell(flexcpData, 0, lngCol)
        If lng��ԴId = 0 And lng����ID = 0 Then Exit Sub
        If IsDate(strCurItem) = False Then Exit Sub
        If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) = "" Then Exit Sub

        Call frmEdit.ShowMe(Me, 1, Fun_View, lng����ID, lng��ԴId, lng����ID, strCurItem, mstrPrivs)
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistPlan_GotFocus()
    Call SetSelectedBackColor(vsfRegistPlan, True)
End Sub

Private Sub vsfRegistPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RegistPlan_KeyDown(vsfRegistPlan, KeyCode, Shift)
End Sub

Private Sub vsfRegistPlan_LostFocus()
    Call SetSelectedBackColor(vsfRegistPlan, False)
End Sub

Private Sub vsfRegistPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbCommandBar As CommandBar

    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)

    Set cbCommandBar = GetPopupCommandBar(Me, mcbsMain)
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub
    
    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRule_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If NewRow < 2 Then
        Call mfrmMain.StatusShowInfoChanged(2, "")
    Else
        Call mfrmMain.StatusShowInfoChanged(2, "��ǰ��" & mlngSignalCount & "����Դ����ǰ��Դ��" & vsfRegistRule.TextMatrix(NewRow, COL_����) & _
            "����ʼʱ�䣺" & vsfRegistRule.TextMatrix(NewRow, COL_��ʼʱ��) & "����ֹʱ�䣺" & vsfRegistRule.TextMatrix(NewRow, COL_��ֹʱ��) & "")
    End If
End Sub

Private Sub vsfRegistRule_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    On Error Resume Next
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    vsfRegistRuleSub.LeftCol = NewLeftCol
End Sub

Private Sub vsfRegistRule_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    On Error Resume Next
    Call SetPlanGridRangeColor(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, mstrOldSelRangeRule)
    mstrOldSelRangeRule = vsfRegistRule.Row & "|" & vsfRegistRule.RowSel & "|" & vsfRegistRule.Col & "|" & vsfRegistRule.ColSel
End Sub

Private Sub vsfRegistRule_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRule, Me.Name, "����")
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRuleSub, Me.Name, "����")
End Sub

Private Sub vsfRegistRule_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    If Val(vsfRegistRule.RowData(NewRow)) = -1 Then Cancel = True
End Sub

Private Sub vsfRegistRule_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = gPlanGrid_ColIndex.COL_ͼ�� Then Cancel = True: Exit Sub
End Sub

Private Sub vsfRegistRule_DblClick()
    Dim lng��ԴId As Long, lng����ID As Long
    Dim frmEdit As New frmClinicPlanEdit
    Dim strCurItem As String, blnUpdate As Boolean
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String

    Err = 0: On Error GoTo errHandler
    lngCol = vsfRegistRule.MouseCol
    lngRow = vsfRegistRule.MouseRow
    If lngRow = 0 Or lngRow = 1 Then
        '����
        If Not mrsRuleRecords Is Nothing Then
            strSort = GetPlanSortCircleStr(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, lngRow, lngCol)
            If strSort <> "" Then
                mrsRuleRecords.Sort = strSort
                Screen.MousePointer = vbHourglass
                Call LoadPlanDataByRecordset(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecords, 0, , True)
                Screen.MousePointer = vbDefault
            End If
        End If
    Else
        With vsfRegistRule
            lng��ԴId = Val(.TextMatrix(.Row, COL_��ԴID))
            lng����ID = Val(.TextMatrix(.Row, COL_����ID))
            strCurItem = .Cell(flexcpData, 0, .Col)
            If lng��ԴId = 0 And lng����ID = 0 Then Exit Sub
            If strCurItem = "" Then Exit Sub
            
            blnUpdate = zlStr.IsHavePrivs(mstrPrivs, "���ﰲ��") _
                And (Val(lblPublishInfo.Tag) = 0 Or Val(.TextMatrix(.Row, COL_�Ƿ����)) = 0)
            If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
                'û�С����п��ҡ�Ȩ��ʱ��ֻ�ܵ����������ٴ������Űࡱ�ĺ�Դ
                If Trim(.TextMatrix(.Row, COL_�Ƿ��ٴ��Ű�)) = "" Then blnUpdate = False
            End If
    
            If frmEdit.ShowMe(Me, 0, IIf(blnUpdate, Fun_Update, Fun_View), mlng����ID, lng��ԴId, lng����ID, strCurItem, mstrPrivs) Then
                If blnUpdate Then Call RefreshOneData
            End If
        End With
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRule_EnterCell()
    Dim lng����ID As Long

    Err = 0: On Error GoTo errHandler
    If Val(vsfRegistRule.Tag) = vsfRegistRule.Row Then Exit Sub
    vsfRegistRule.Tag = vsfRegistRule.Row
    lng����ID = Val(vsfRegistRule.TextMatrix(vsfRegistRule.Row, COL_����ID))
    LoadPlanDataSub mlng����ID, lng����ID
    If vsfRegistRule.Visible And vsfRegistRule.Enabled Then vsfRegistRule.SetFocus
    Call SetSelectedBackColor(vsfRegistRuleSub, False)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRule_GotFocus()
    Call SetSelectedBackColor(vsfRegistRule, True)
End Sub

Private Sub vsfRegistRule_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RegistPlan_KeyDown(vsfRegistRule, KeyCode, Shift)
End Sub

Private Sub vsfRegistRule_LostFocus()
    Call SetSelectedBackColor(vsfRegistRule, False)
End Sub

Private Sub vsfRegistRule_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbCommandBar As CommandBar

    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)

    Set cbCommandBar = GetPopupCommandBar(Me, mcbsMain)
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub

    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistPlan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim varData As Variant, strTemp As String
    Dim lngRow As Long, lngCol As Long

    On Error GoTo errHandler
    With vsfRegistPlan
        If Not .Visible Then Call mfrmMain.StatusShowInfoChanged(3, ""): Exit Sub

        lngRow = .MouseRow: lngCol = .MouseCol
        If .Tag = lngRow & "," & lngCol Then Exit Sub
        .Tag = lngRow & "," & lngCol

        If lngRow < .FixedRows Or lngCol < gPlanGrid_FixedCols Then Call mfrmMain.StatusShowInfoChanged(3, ""): Exit Sub
        If (lngCol - gPlanGrid_FixedCols) Mod 3 = 0 Then Call mfrmMain.StatusShowInfoChanged(3, ""): Exit Sub '"ʱ��"���˳�
        strTemp = Trim(.TextMatrix(lngRow, lngCol))
        If (strTemp = "" Or InStr(strTemp, "/") = 0) And strTemp <> "-" Then Call mfrmMain.StatusShowInfoChanged(3, ""): Exit Sub

        '2.��ʾ����
        If strTemp = "-" Then
            strTemp = "��ֹԤԼ��"
        Else
            varData = Split(strTemp, "/")
            If (lngCol - gPlanGrid_FixedCols) Mod 3 = 1 Then
                strTemp = "�޺���:" & IIf(Trim(varData(1)) = "��", "������", Trim(varData(1))) & ", �����ѹ���:" & Trim(varData(0))
            ElseIf (lngCol - gPlanGrid_FixedCols) Mod 3 = 2 Then
                strTemp = "��Լ��:" & IIf(Trim(varData(1)) = "��", "������", Trim(varData(1))) & ", ������Լ��:" & Trim(varData(0))
            End If
        End If
        Call mfrmMain.StatusShowInfoChanged(3, strTemp)
    End With
    Exit Sub
errHandler:
    Err.Clear
    Call mfrmMain.StatusShowInfoChanged(3, "")
End Sub

Private Sub picImgPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT

    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfRegistPlan, lngLeft, lngTop, picImgPlan.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "����")
End Sub

Private Function GetPlanRuleData(ByVal lng����ID As Long, Optional ByVal lng����ID As Long, _
    Optional ByVal blnPublished As Boolean) As ADODB.Recordset
    '���ܣ���ȡ���Ź���
    '�������ܳ����
    '   lng����ID   - ����ID
    '   blnPublished- �Ƿ��ѷ���
    Dim strSQL As String, strColSub As String
    Dim strWhere As String, str�Ƿ���Ч As String
    Dim str������� As String

    Err = 0: On Error GoTo errHandler
    str������� = IIf(gVisitPlan_ModulePara.byt����ȽϷ�ʽ = 0, "a.����", "Lpad(a.����,5,'0')")
    strColSub = "       " & str������� & " As �������, a.Id As ��Դid, a.����, a.����, Nvl(a.�Ƿ񽨲���, 0) As �Ƿ񽨲���,a.ԤԼ����, a.����Ƶ��," & vbNewLine & _
                "       Decode(a.���տ���״̬, 1, '����ԤԼ', 2, '��ֹԤԼ', 3, '�ܽڼ������ÿ���', '���ϰ�') As ���տ���״̬," & vbNewLine & _
                "       Nvl(a.�Ƿ��ٴ��Ű�, 0) As �Ƿ��ٴ��Ű�, Decode(a.�Ű෽ʽ, 1, '�����Ű�', 2, '�����Ű�', '�̶��Ű�') As �Ű෽ʽ," & vbNewLine & _
                "       f.���� As ����, f.���� As ���Ҽ���, Nvl(a.�Ƿ���ջ���, 0) As �Ƿ���ջ���," & vbNewLine

    'û��"���п���"Ȩ�޵Ĳ���Աֻ�ܲ����Լ��������ҵĺ�Դ
    If HavePrivs(mstrPrivs, "���п���") = False Then
        strWhere = "      And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = [3])"
    End If
    
    '��Ч���ţ����¹������ͬʱ���㣺
    '    --1.�����
    '    --2.��ֹʱ����ڵ�ǰʱ��
    '    --3.������������ʱ���Ż������������ŵ�ʱ�䷶Χû�б�����
    '    --4.û�е���Ϊ�����Ű෽ʽ�����ߵ���Ϊ�����Ű෽ʽ����û�г��ﰲ��
    '˵�����ɶ����ʱ����һ�𸲸ǵİ��ţ��жϲ���
    str�Ƿ���Ч = "Nvl((Select 1" & vbNewLine & _
                " From Dual" & vbNewLine & _
                " Where b.���ʱ�� Is Not Null And b.��ֹʱ�� > Sysdate" & vbNewLine & _
                "       And Not Exists(Select 1" & vbNewLine & _
                "        From �ٴ����ﰲ��" & vbNewLine & _
                "        Where ���ʱ�� Is Not Null And ��Դid = b.��Դid And �Ǽ�ʱ�� > b.�Ǽ�ʱ��" & vbNewLine & _
                "              And (Nvl(b.�Ƿ���ʱ����, 0) = 0 And Nvl(�Ƿ���ʱ����, 0) = 0 Or Nvl(b.�Ƿ���ʱ����, 0) = 1)" & vbNewLine & _
                "              And Decode(Sign(Sysdate - ��ʼʱ��), 1, Sysdate, ��ʼʱ��) <= Decode(Sign(Sysdate - b.��ʼʱ��), 1, Sysdate, b.��ʼʱ��)" & vbNewLine & _
                "              And ��ֹʱ�� >= b.��ֹʱ��)" & vbNewLine & _
                "       And Not Exists(Select 1" & vbNewLine & _
                "           From �ٴ����ﰲ�� P, �ٴ������ Q" & vbNewLine & _
                "           Where p.����id = q.Id And p.��Դid = b.��Դid And Nvl(q.�Ű෽ʽ, 0) In (1, 2)And p.��ʼʱ�� < Sysdate)" & vbNewLine & _
                "   ), 0) As �Ƿ���Ч,"
    
    If lng����ID = 0 Then
        strSQL = "Select m.Id As ����ID, m.����ID, m.��Դid, m.��ĿID, m.ҽ��ID, m.ҽ������, m.�Ű����," & _
                "        m.��ʼʱ��, m.��ֹʱ��, m.�Ǽ�ʱ��,m.���ʱ��,m.�Ƿ���ʱ����" & vbNewLine & _
                " From �ٴ����ﰲ�� M" & vbNewLine & _
                " Where m.����id = [1] And Nvl(m.�Ƿ���ʱ����, 0) = 0"
        If blnPublished Then
            strSQL = "Select " & str�Ƿ���Ч & _
                    "       b.����ID, b.����id, e.���� As �շ���Ŀ, b.ҽ������, g.���� As ҽ������, g.רҵ����ְ�� as ҽ��ְ��,i.��ʶ��," & vbNewLine & _
                    "       b.�Ű����, b.��ʼʱ��, b.��ֹʱ��, b.�Ǽ�ʱ��, b.�Ƿ���ʱ���� As ��ʱ����, Decode(b.���ʱ��,Null,0,1) As �Ƿ����," & vbNewLine & strColSub & _
                    "       c.Id As ��¼id, c.������Ŀ, c.�ϰ�ʱ��, c.�޺���, c.��Լ��, c.ԤԼ���� As ԤԼ���Ʒ�ʽ" & vbNewLine & _
                    "From �ٴ������Դ A, (" & strSQL & ") B, �ٴ��������� C, �շ���ĿĿ¼ E, ���ű� F, ��Ա�� G, �ٴ������ H,רҵ����ְ�� I" & vbNewLine & _
                    "Where a.Id = b.��Դid And b.����ID = c.����id(+) And a.����id = f.Id And b.��Ŀid = e.Id And b.ҽ��ID = g.ID(+) And b.����ID = h.ID" & strWhere & vbNewLine & _
                    "      And g.רҵ����ְ��=i.����(+) And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
                    "      And (e.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.����ʱ�� Is Null)" & vbNewLine & _
                    "      And (f.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.����ʱ�� Is Null)" & vbNewLine & _
                    "      And (g.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.����ʱ�� Is Null)" & vbNewLine & _
                    "Order By " & str������� & ", b.�Ǽ�ʱ�� Desc, c.������Ŀ, c.�ϰ�ʱ��"
        Else
            strSQL = "Select " & str�Ƿ���Ч & _
                    "       b.����ID, b.����id," & _
                    "       Decode(b.����ID,Null,e.����,m.����) As �շ���Ŀ, Decode(b.����ID,Null,a.ҽ������,b.ҽ������) As ҽ������," & vbNewLine & _
                    "       Decode(b.����ID,Null,g.����,n.����) As ҽ������, Decode(b.����ID,Null,g.רҵ����ְ��,n.רҵ����ְ��) as ҽ��ְ��, 0 As ��ʱ����," & vbNewLine & _
                    "       Decode(b.����ID,Null,i.��ʶ��,j.��ʶ��) as ��ʶ��,b.�Ű����, b.��ʼʱ��, b.��ֹʱ��, b.�Ǽ�ʱ��, 0 As �Ƿ����," & vbNewLine & strColSub & _
                    "       c.Id As ��¼id, c.������Ŀ, c.�ϰ�ʱ��, c.�޺���, c.��Լ��, c.ԤԼ���� As ԤԼ���Ʒ�ʽ" & vbNewLine & _
                    "From �ٴ������Դ A, (" & strSQL & ") B, �ٴ��������� C, �շ���ĿĿ¼ E, ���ű� F, ��Ա�� G, �շ���ĿĿ¼ M, ��Ա�� N, �ٴ������ H,רҵ����ְ�� I,רҵ����ְ�� J" & vbNewLine & _
                    "Where a.Id = b.��Դid(+) And b.����ID = c.����id(+) And a.����id = f.Id(+)" & vbNewLine & _
                    "      And a.��Ŀid = e.Id And a.ҽ��ID= g.ID(+) And b.��Ŀid = m.Id(+) And b.ҽ��ID= n.ID(+) And b.����ID = h.ID(+)" & strWhere & vbNewLine & _
                    "      And g.רҵ����ְ��=i.����(+) And n.רҵ����ְ��=j.����(+)" & vbNewLine & _
                    "      And (b.����ID Is Not Null Or (b.����ID Is Null And a.�Ű෽ʽ = 0))" & vbNewLine & _
                    "      And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
                    "      And (e.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.����ʱ�� Is Null)" & vbNewLine & _
                    "      And (f.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.����ʱ�� Is Null)" & vbNewLine & _
                    "      And (g.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.����ʱ�� Is Null)" & vbNewLine & _
                    "      And Nvl(Nvl(f.վ��,[5]),Nvl([4],'-')) = Nvl([4],'-')" & vbNewLine & _
                    "      And Not Exists(Select 1 From �ٴ����ﰲ�� P,�ٴ������ Q" & vbNewLine & _
                    "                     Where p.����ID = q.ID And p.��ԴID = a.ID And q.�Ű෽ʽ = 0 And q.ID <> Nvl(b.����id,0))" & vbNewLine & _
                    "Order By " & str������� & ", b.�Ǽ�ʱ�� Desc, c.������Ŀ, c.�ϰ�ʱ��"
        End If
    Else
        strSQL = "Select m.Id As ����id, m.����ID, m.��Դid, m.��ĿID, m.ҽ��ID, m.ҽ������, m.�Ű����, m.��ʼʱ��, m.��ֹʱ��, m.�Ǽ�ʱ��, m.���ʱ��, m.�Ƿ���ʱ����" & vbNewLine & _
                " From �ٴ����ﰲ�� M, �ٴ����ﰲ�� J" & vbNewLine & _
                " Where m.����id = j.����id And m.��Դid = j.��Դid And j.id = [2] And Nvl(m.�Ƿ���ʱ����, 0) = 1"
        strSQL = "Select " & str�Ƿ���Ч & _
                "       b.����ID, b.����id, e.���� As �շ���Ŀ, b.ҽ������, g.���� As ҽ������, g.רҵ����ְ�� as ҽ��ְ��,i.��ʶ��," & _
                "       b.�Ű����, b.��ʼʱ��, b.��ֹʱ��, b.�Ǽ�ʱ��,Decode(b.���ʱ��,Null,0,1) As �Ƿ����," & vbNewLine & _
                "       Case When b.�Ǽ�ʱ�� > Nvl(h.����ʱ��, To_date('3000-01-01','yyyy-mm-dd')) Then 1 Else 0 End As ��ʱ����," & vbNewLine & strColSub & _
                "       c.Id As ��¼id, c.������Ŀ, c.�ϰ�ʱ��, c.�޺���, c.��Լ��, c.ԤԼ���� As ԤԼ���Ʒ�ʽ" & vbNewLine & _
                "From �ٴ������Դ A, (" & strSQL & ") B, �ٴ��������� C, �շ���ĿĿ¼ E, ���ű� F, ��Ա�� G, �ٴ������ H,רҵ����ְ�� I" & vbNewLine & _
                "Where a.Id = b.��Դid And b.����ID = c.����id(+) And a.����id = f.Id And b.��Ŀid = e.Id And b.ҽ��ID = g.ID(+) And b.����ID = h.ID and g.רҵ����ְ��=i.����(+)" & vbNewLine & _
                "      And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
                "      And (e.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.����ʱ�� Is Null)" & vbNewLine & _
                "      And (f.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.����ʱ�� Is Null)" & vbNewLine & _
                "      And (g.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.����ʱ�� Is Null)" & vbNewLine & _
                "Order By " & str������� & ", b.�Ǽ�ʱ�� Desc, c.������Ŀ, c.�ϰ�ʱ��"
        If mblnShowInvalidPlan = False Then
            strSQL = "Select �Ƿ���Ч, ����ID, ����id, �շ���Ŀ, ҽ������, ҽ������, ҽ��ְ��,��ʶ��," & vbNewLine & _
                    "        �Ű����, ��ʼʱ��, ��ֹʱ��, �Ǽ�ʱ��, �Ƿ����, ��ʱ����," & vbNewLine & _
                    "        ��Դid, ����, ����, �Ƿ񽨲���, ԤԼ����, ����Ƶ��," & vbNewLine & _
                    "        ���տ���״̬, �Ƿ��ٴ��Ű�, �Ű෽ʽ, ����, ���Ҽ���, �Ƿ���ջ���," & vbNewLine & _
                    "        ��¼id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, ԤԼ���Ʒ�ʽ" & vbNewLine & _
                    " From (" & strSQL & ")" & vbNewLine & _
                    " Where �Ƿ���Ч=1 Or �Ƿ����=0"
        End If
    End If
    Set GetPlanRuleData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ű���Ϣ", lng����ID, lng����ID, UserInfo.ID, _
        gstrNodeNo, gVisitPlan_ModulePara.str��Դά��վ��)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPlanRecords(ByVal lng����ID As Long, _
    Optional ByVal dtStart As Date, Optional ByVal dtEnd As Date) As ADODB.Recordset
    '���ܣ���ȡ���ż�¼
    Dim strSQL As String, str�Ƿ���Ч As String
    Dim strPrivsWhere As String
    Dim str������� As String

    Err = 0: On Error GoTo errHandler
    str������� = IIf(gVisitPlan_ModulePara.byt����ȽϷ�ʽ = 0, "c.����", "Lpad(c.����,5,'0')")
    'û��"���п���"Ȩ�޵Ĳ���Աֻ�ܲ����Լ��������ҵĺ�Դ
    If HavePrivs(mstrPrivs, "���п���") = False Then
        strPrivsWhere = "      And Exists (Select 1 From ������Ա Where ����id = c.����id And ��Աid = [2])"
    End If
    
    '��Ч���ţ����¹������ͬʱ���㣺
    '    --1.���������
    '    --2.������ֹʱ����ڵ�ǰʱ��
    '    --3.������������ʱ���Ż������������ŵ�ʱ�䷶Χû�б�����
    '    --4.û�е���Ϊ�����Ű෽ʽ�����ߵ���Ϊ�����Ű෽ʽ����û�г��ﰲ��
    '    --5.����Ч�ĳ����¼
    '˵�����ɶ����ʱ����һ�𸲸ǵİ��ţ��жϲ���
    str�Ƿ���Ч = "Nvl((Select 1" & vbNewLine & _
                " From Dual" & vbNewLine & _
                " Where a.���ʱ�� Is Not Null And a.��ֹʱ�� > Sysdate" & vbNewLine & _
                "       And Not Exists(Select 1" & vbNewLine & _
                "           From �ٴ����ﰲ��" & vbNewLine & _
                "           Where ���ʱ�� Is Not Null And ��Դid = a.��Դid And �Ǽ�ʱ�� > a.�Ǽ�ʱ��" & vbNewLine & _
                "                 And (Nvl(a.�Ƿ���ʱ����, 0) = 0 And Nvl(�Ƿ���ʱ����, 0) = 0 Or Nvl(a.�Ƿ���ʱ����, 0) = 1)" & vbNewLine & _
                "                 And Decode(Sign(Sysdate - ��ʼʱ��), 1, Sysdate, ��ʼʱ��) <= Decode(Sign(Sysdate - a.��ʼʱ��), 1, Sysdate, a.��ʼʱ��)" & vbNewLine & _
                "                 And ��ֹʱ�� >= b.��ֹʱ��)" & vbNewLine & _
                "       And Not Exists(Select 1" & vbNewLine & _
                "           From �ٴ����ﰲ�� P, �ٴ������ Q" & vbNewLine & _
                "           Where p.����id = q.Id And p.��Դid = a.��Դid And Nvl(q.�Ű෽ʽ, 0) In (1, 2) And p.��ʼʱ�� < Sysdate)" & vbNewLine & _
                "       And Exists(Select 1" & vbNewLine & _
                "           From �ٴ������¼ P, �ٴ����ﰲ�� Q" & vbNewLine & _
                "           Where p.����id = q.Id And q.����id = [1] And q.��Դid = a.��ԴID And p.�������� + 1 > Sysdate)" & vbNewLine & _
                "   ), 0) As �Ƿ���Ч,"
                
    strSQL = "Select " & str�Ƿ���Ч & vbNewLine & _
            "        " & str������� & " As �������, c.Id As ��Դid, c.����, c.����, Nvl(c.�Ƿ񽨲���, 0) As �Ƿ񽨲���, c.ԤԼ����, c.����Ƶ��," & vbNewLine & _
            "        Decode(c.���տ���״̬, 1, '����ԤԼ', 2, '��ֹԤԼ', 3, '�ܽڼ������ÿ���', '���ϰ�') As ���տ���״̬," & vbNewLine & _
            "        Decode(c.�Ű෽ʽ, 1, '�����Ű�', 2, '�����Ű�', '�̶��Ű�') As �Ű෽ʽ," & vbNewLine & _
            "        Nvl(c.�Ƿ���ջ���, 0) As �Ƿ���ջ���, Nvl(c.�Ƿ��ٴ��Ű�, 0) As �Ƿ��ٴ��Ű�," & vbNewLine & _
            "        f.���� As ����, f.���� As ���Ҽ���, e.���� As �շ���Ŀ, a.ҽ������, g.רҵ����ְ�� As ҽ��ְ��,h.��ʶ��,g.���� As ҽ������," & vbNewLine & _
            "        a.����id, a.Id As ����id, a.��ʼʱ��, a.��ֹʱ��," & vbNewLine & _
            "        b.Id As ��¼id, b.��������, b.�ϰ�ʱ��, b.�޺���, b.��Լ��, b.�ѹ���, b.��Լ��, b.ԤԼ���� As ԤԼ���Ʒ�ʽ," & vbNewLine & _
            "        b.ͣ�￪ʼʱ��, b.ͣ����ֹʱ��, b.ͣ��ԭ��, b.����ҽ������, b.�Ƿ���ʱ����, b.�Ƿ�����" & vbNewLine & _
            " From �ٴ����ﰲ�� A, �ٴ������¼ B, �ٴ������Դ C, �շ���ĿĿ¼ E, ���ű� F, ��Ա�� G,רҵ����ְ�� H" & vbNewLine & _
            " Where a.Id = b.����id(+) And a.��Դid = c.Id And c.����id = f.Id And a.��Ŀid = e.Id And a.ҽ��id = g.Id(+) And a.����id = [1]" & vbNewLine & _
            "       And a.���ʱ�� Is Not Null" & vbNewLine & _
            "       And g.רҵ����ְ��=h.����(+) " & vbNewLine & _
            "       And Nvl(c.�Ƿ�ɾ��, 0) = 0 And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
            "       And (e.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.����ʱ�� Is Null)" & vbNewLine & _
            "       And (f.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.����ʱ�� Is Null)" & vbNewLine & _
            "       And (g.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.����ʱ�� Is Null)" & vbNewLine & _
                    strPrivsWhere & vbNewLine & _
            "       And Nvl(b.��������,[3]) Between [3] And [4]" & vbNewLine & _
            "       And Exists(Select 1 From �ٴ������¼ Where ����ID = a.ID And b.�������� Between [3] And [4])" & vbNewLine & _
            " Order By " & str������� & ", ����, �շ���Ŀ, ҽ������, ��������, �ϰ�ʱ��"
    Set GetPlanRecords = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ű���Ϣ", lng����ID, UserInfo.ID, dtStart, dtEnd, gstrNodeNo)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadPlanDataSub(Optional ByVal lng����ID As Long, Optional ByVal lng����ID As Long)
    '���ܣ�����ͬһ����Դ������ŵİ�������
    '������
    '   lng����ID - ����ID
    Err = 0: On Error GoTo errHandler
    Set mrsRuleRecordsSub = Nothing
    If lng����ID <> 0 And lng����ID <> 0 Then
        '��������
        Screen.MousePointer = vbHourglass
        '����
        Set mrsRuleRecordsSub = GetPlanRuleData(lng����ID, lng����ID, Val(lblPublishInfo.Tag) = 1)
        Call LoadPlanDataByRecordset(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecordsSub, 0)
        Screen.MousePointer = vbDefault
    Else
        Call LoadPlanDataByRecordset(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecordsSub, 0)
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRuleSub_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If NewRow < 2 Then
        Call mfrmMain.StatusShowInfoChanged(2, "")
    Else
        Call mfrmMain.StatusShowInfoChanged(2, "��ǰ��" & mlngSignalCount & "����Դ����ǰ��Դ��" & vsfRegistRuleSub.TextMatrix(NewRow, COL_����) & _
            "����ʼʱ�䣺" & vsfRegistRuleSub.TextMatrix(NewRow, COL_��ʼʱ��) & "����ֹʱ�䣺" & vsfRegistRuleSub.TextMatrix(NewRow, COL_��ֹʱ��) & "")
    End If
End Sub

Private Sub vsfRegistRuleSub_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    On Error Resume Next
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    vsfRegistRule.LeftCol = NewLeftCol
End Sub

Private Sub vsfRegistRuleSub_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    On Error Resume Next
    Call SetPlanGridRangeColor(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, mstrOldSelRangeRuleSub)
    mstrOldSelRangeRuleSub = vsfRegistRuleSub.Row & "|" & vsfRegistRuleSub.RowSel & "|" & vsfRegistRuleSub.Col & "|" & vsfRegistRuleSub.ColSel
End Sub

Private Sub vsfRegistRuleSub_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRuleSub, Me.Name, "����")
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRule, Me.Name, "����")
End Sub

Private Sub vsfRegistRuleSub_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    If Val(vsfRegistRuleSub.RowData(NewRow)) = -1 Then Cancel = True
End Sub

Private Sub vsfRegistRuleSub_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = gPlanGrid_ColIndex.COL_ͼ�� Then Cancel = True: Exit Sub
End Sub

Private Sub vsfRegistRuleSub_DblClick()
    Dim lng��ԴId As Long, lng����ID As Long
    Dim frmEdit As New frmClinicPlanEdit
    Dim strCurItem As String, blnUpdate As Boolean
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String

    Err = 0: On Error GoTo errHandler
    lngCol = vsfRegistRuleSub.MouseCol
    lngRow = vsfRegistRuleSub.MouseRow
    If lngRow = 0 Or lngRow = 1 Then
        '����
        If mrsRuleRecordsSub Is Nothing Then Exit Sub
        If mrsRuleRecordsSub.RecordCount = 0 Then Exit Sub
        strSort = GetPlanSortCircleStr(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, lngRow, lngCol)
        If strSort <> "" Then
            mrsRuleRecordsSub.Sort = strSort
            Screen.MousePointer = vbHourglass
            Call LoadPlanDataByRecordset(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecordsSub, 0, , True)
            Screen.MousePointer = vbDefault
        End If
    Else
        With vsfRegistRuleSub
            lng��ԴId = Val(.TextMatrix(.Row, COL_��ԴID))
            lng����ID = Val(.TextMatrix(.Row, COL_����ID))
            strCurItem = .Cell(flexcpData, 0, .Col)
            If lng��ԴId = 0 And lng����ID = 0 Then Exit Sub
            If strCurItem = "" Then Exit Sub
            
            blnUpdate = zlStr.IsHavePrivs(mstrPrivs, "���ﰲ��") _
                And (Val(.TextMatrix(.Row, COL_��ʱ����)) = 1 And Val(.TextMatrix(.Row, COL_�Ƿ����)) = 0 Or Val(lblPublishInfo.Tag) = 0)
            If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
                'û�С����п��ҡ�Ȩ��ʱ��ֻ�ܵ����������ٴ������Űࡱ�ĺ�Դ
                If Trim(.TextMatrix(.Row, COL_�Ƿ��ٴ��Ű�)) = "" Then blnUpdate = False
            End If
    
            If frmEdit.ShowMe(Me, 0, IIf(blnUpdate, Fun_TempPlan, Fun_View), mlng����ID, lng��ԴId, lng����ID, strCurItem, mstrPrivs) Then
                If blnUpdate Then Call RefreshDataSub
            End If
        End With
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshDataSub()
    'ˢ����ʱ�����б�
    Dim lngOldRow As Long, lngoldCol As Long

    Err = 0: On Error GoTo errHandler
    With vsfRegistRuleSub
        lngOldRow = .Row
        lngoldCol = .Col

        vsfRegistRule.Tag = "": Call vsfRegistRule_EnterCell

        If .Rows > .FixedRows And .Cols > .FixedCols Then
            .Row = IIf(lngOldRow = 0 Or lngOldRow > .Rows - 1, .FixedRows, lngOldRow)
            .Col = IIf(lngoldCol = 0 Or lngoldCol > .Cols - 1, .FixedCols, lngoldCol)
            .ShowCell .Row, .Col  '������ʾ��ָ����Ԫ
            If .Visible And .Enabled Then .SetFocus
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRuleSub_GotFocus()
    Call SetSelectedBackColor(vsfRegistRuleSub, True)
End Sub

Private Sub vsfRegistRuleSub_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RegistPlan_KeyDown(vsfRegistRuleSub, KeyCode, Shift)
End Sub

Private Sub vsfRegistRuleSub_LostFocus()
    Call SetSelectedBackColor(vsfRegistRuleSub, False)
End Sub

Private Sub vsfRegistRuleSub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbCommandBar As CommandBar

    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)

    Set cbCommandBar = GetPopupCommandBar(Me, mcbsMain)
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub

    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txtFind_KeyPress(index As Integer, KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call ExecuteFilter(Val(tbPage(Pg_���ﰲ��).Tag) = 0)
    End If
End Sub

Private Sub SetSelectedBackColor(vsfGrid As VSFlexGrid, ByVal blnFocus As Boolean)
    '�������񼤻�״̬����ѡ���б�����ɫ
    Dim lngRowStart As Long, lngColStart As Long, lngRowEnd As Long, lngColEnd As Long
    Dim strOldSelRange As String, dataType As gPlanGrid_DataStyle

    Err = 0: On Error GoTo errHandler
    If vsfGrid Is vsfRegistPlan Then
        strOldSelRange = mstrOldSelRangePlan
        dataType = gPlanGrid_DataStyle.Data_Plan
    ElseIf vsfGrid Is vsfRegistRule Then
        strOldSelRange = mstrOldSelRangeRule
        dataType = gPlanGrid_DataStyle.Data_FixedRule
    ElseIf vsfGrid Is vsfRegistRuleSub Then
        strOldSelRange = mstrOldSelRangeRuleSub
        dataType = gPlanGrid_DataStyle.Data_FixedRule
    Else
        Exit Sub
    End If
    If blnFocus Then
        Call SetPlanGridRangeColor(vsfGrid, dataType, strOldSelRange)
    Else
        If GetSelectRange(vsfGrid, strOldSelRange, lngRowStart, lngRowEnd, lngColStart, lngColEnd) Then
            vsfGrid.Cell(flexcpBackColor, lngRowStart, lngColStart, lngRowEnd, lngColEnd) = G_LostFocusColor
        End If
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub fraSplitRule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Err = 0: On Error Resume Next
    If Button <> vbLeftButton Then Exit Sub
    If vsfRegistRule.Height + Y < 1200 Or vsfRegistRuleSub.Height - Y < 1200 Then Exit Sub

    fraSplitRule.Top = fraSplitRule.Top + Y
    
    vsfRegistRule.Height = vsfRegistRule.Height + Y
    vsfRegistRuleSub.Top = vsfRegistRuleSub.Top + Y
    vsfRegistRuleSub.Height = vsfRegistRuleSub.Height - Y
    Me.Refresh
End Sub

Private Sub picRegistRule_Resize()
    On Error Resume Next
    vsfRegistRule.Move -10, 0, picRegistRule.ScaleWidth + 20, picRegistRule.ScaleHeight * 2 / 3
    fraSplitRule.Move 0, vsfRegistRule.Top + vsfRegistRule.Height, picRegistRule.ScaleWidth + 20
    With vsfRegistRuleSub
        .Left = -10
        .Top = fraSplitRule.Top + fraSplitRule.Height
        .Width = picRegistRule.ScaleWidth + 20
        .Height = picRegistRule.ScaleHeight - .Top + 10
    End With
End Sub

Private Sub txtFind_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        '�����Ҽ��˵���ݼ������ճ��������
        If Clipboard.GetText <> "" Then Clipboard.Clear
    End If
End Sub

Private Sub txtFind_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtFind(index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtFind(index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtFind_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtFind(index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


