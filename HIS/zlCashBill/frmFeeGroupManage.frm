VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmFeeGroupManage 
   Caption         =   "�������տ����"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmFeeGroupManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11790
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picHistory 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   3240
      ScaleHeight     =   6855
      ScaleWidth      =   8055
      TabIndex        =   7
      Top             =   720
      Width           =   8055
      Begin VB.PictureBox picImgPlanHistory 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   14
         Top             =   510
         Width           =   210
         Begin VB.Image imgColPlanHistory 
            Height          =   195
            Left            =   0
            Picture         =   "frmFeeGroupManage.frx":058A
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.PictureBox picSendInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   3720
         ScaleHeight     =   3135
         ScaleWidth      =   4935
         TabIndex        =   12
         Top             =   2280
         Width           =   4935
         Begin VB.PictureBox picImgPlan 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   60
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   15
            Top             =   30
            Width           =   210
            Begin VB.Image imgColPlan 
               Height          =   195
               Left            =   0
               Picture         =   "frmFeeGroupManage.frx":0AD8
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsCollectorDetail 
            Height          =   1095
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   2535
            _cx             =   4471
            _cy             =   1931
            Appearance      =   0
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmFeeGroupManage.frx":1026
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
      End
      Begin VB.PictureBox picGroupCollectInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   6840
         ScaleHeight     =   2895
         ScaleWidth      =   3735
         TabIndex        =   10
         Top             =   960
         Width           =   3735
         Begin VB.PictureBox picImgPlanGroup 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   60
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   16
            Top             =   30
            Width           =   210
            Begin VB.Image imgColPlanGroup 
               Height          =   195
               Left            =   0
               Picture         =   "frmFeeGroupManage.frx":1282
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsGroupCollectInfo 
            Height          =   1935
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   2535
            _cx             =   4471
            _cy             =   3413
            Appearance      =   0
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmFeeGroupManage.frx":17D0
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
      End
      Begin XtremeSuiteControls.TabControl tabHistory 
         Height          =   1815
         Left            =   15
         TabIndex        =   9
         Top             =   2880
         Width           =   2295
         _Version        =   589884
         _ExtentX        =   4048
         _ExtentY        =   3201
         _StockProps     =   64
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRollingCurtainHistory 
         Height          =   1335
         Left            =   15
         TabIndex        =   5
         Top             =   480
         Width           =   4215
         _cx             =   7435
         _cy             =   2355
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeGroupManage.frx":1967
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
      Begin VB.CommandButton cmdReFilter 
         Caption         =   "���¹�������(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8640
         TabIndex        =   4
         Top             =   75
         Width           =   1935
      End
      Begin VB.TextBox txtSendFeeNO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8000
         MaxLength       =   20
         TabIndex        =   3
         Top             =   75
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker dtpTerminateTime 
         Height          =   300
         Left            =   4560
         TabIndex        =   2
         Top             =   75
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   181469185
         CurrentDate     =   41521
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   75
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   181469185
         CurrentDate     =   41521.0020833333
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��                         ��ֹʱ��                         ���ʵ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   7770
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8025
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   635
      SimpleText      =   $"frmFeeGroupManage.frx":1B8E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFeeGroupManage.frx":1BD5
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13150
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
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
   Begin XtremeSuiteControls.TabControl tabMain 
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   870
      Width           =   2175
      _Version        =   589884
      _ExtentX        =   3836
      _ExtentY        =   2355
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1080
      Top             =   3720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFeeGroupManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mcbrPopupMain As CommandBar, mcbrMenuView As CommandBarPopup, mcbrListView As CommandBar
Private mcbrCmb As CommandBarComboBox, mstrPrivs As String, mlngModule As Long
Private mlngGroupID As Long '�ɿ���ID
Private mobjChargeBillHistory As New clsChargeBill, mfrmChargeBillTotalHistory As Form  '�տ���Ϣ��Ʊ�ݶ���
Private mfrmFeeGroupRollingCurtain As Form    '�տ��������ҳ��
Private WithEvents mfrmFeeGroupCollectFee As frmFeeGroupCollectFee
Attribute mfrmFeeGroupCollectFee.VB_VarHelpID = -1
Private mblnCancel As Boolean   '�ⲿж�ش����ʶ
Private mintReceiptPrint As Integer, mintChargeBookPrint As Integer

Private Enum EM_Pan
    EM_Pan_��Ա�� = 1
    EM_Pan_�շ�������Ϣ = 2
    EM_Pan_�տƱ����Ϣ = 3
End Enum

Private Enum EM_Tab
    EM_Tab_�տ� = 1
    EM_Tab_���� = 2
    EM_Tab_��ʷ������Ϣ = 3
    EM_Tab_�տƱ�ݻ��� = 4
    EM_Tab_�շ�Ա������ϸ = 5
    EM_Tab_���տ���Ϣ = 6
    EM_Tab_�շ�Ա������Ϣ = 7
End Enum
Private mstrTitle As String '���ڴ�����Ի�����Ĵ�����

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    Select Case Control.ID
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            Control.Checked = stbThis.Visible
            Form_Resize
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            Control.Checked = Not Control.Checked
            cbsThis(2).Visible = Not cbsThis(2).Visible
            Form_Resize
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            Control.Checked = Not Control.Checked
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            Control.Checked = Not Control.Checked
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            Form_Resize
            cbsThis.RecalcLayout
        Case conMenu_View_LargeICO
            Call ChangeListViewType(1)
        Case conMenu_View_MinICO
            Call ChangeListViewType(2)
        Case conMenu_View_ListICO
            Call ChangeListViewType(3)
        Case conMenu_View_DetailsICO
            Call ChangeListViewType(4)
        Case conMenu_Edit_ChargeBook_Reprint
            Call PrintReport
        Case conMenu_Edit_ReprintReceipt
            Call PrintReport
        Case conMenu_View_ChargeAndBilllTotal
            Call ChargeAndBilllTotal
        Case conMenu_View_Detail
            Call ViewDetail
        Case conMenu_File_PrintSet
            zlPrintSet
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
        Case conMenu_File_Parameter
            Call SetPara
        Case conMenu_Edit_CollectFees
            Call ButtonCollectFees
        Case conMenu_Edit_RollingCurtain
            Call mfrmFeeGroupRollingCurtain.RollingCurtain
        Case conMenu_Edit_CollectFees_Cancel
            Call mfrmFeeGroupRollingCurtain.ButtonCancelCollect
        Case conMenu_Edit_RollingCurtain_Cancel
            Call ButtonRollingCurtainCancel
        Case conMenu_File_Print
            Call zlRptPrint(1)
        Case conMenu_File_Preview
            Call zlRptPrint(2)
        Case conMenu_File_Excel
            Call zlRptPrint(3)
        Case conMenu_Edit_CheckCash
            frmMoneyEnum.Show vbModal, Me
        Case conMenu_Help_Help
            ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
        Case conMenu_Help_Web_Home
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_Web_Forum
            Call zlWebForum(Me.hWnd)
        Case conMenu_Help_About
            ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
        Case conMenu_Edit_SelAll
            Call mfrmFeeGroupCollectFee.SetVSFCheckBat(0)
        Case conMenu_Edit_ClsAll
            Call mfrmFeeGroupCollectFee.SetVSFCheckBat(1)
        Case Else
            If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
            End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetPara()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:������
    '����:2013-10-19
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Call frmFeeGroupSetting.ParaSetting(Me, mlngModule, mstrPrivs)
    mintChargeBookPrint = Val(zlDatabase.GetPara("�ɿ����ӡ��ʽ", glngSys, mlngModule, "0"))
    mintReceiptPrint = Val(zlDatabase.GetPara("�տ��վݴ�ӡ��ʽ", glngSys, mlngModule, "0"))
End Sub

Public Function GetListViewMenu() As CommandBarPopup
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����˵�
    '����:�����˵�
    '����:������
    '����:2013-10-09
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Set GetListViewMenu = mcbrMenuView
End Function

Private Sub ChangeListViewType(intTYPE As Integer)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:������Ա�б���ʾ��ʽ
    '���:intType-�б���ʾ��ʽ: 1-��ͼ��;2-Сͼ��;3-�б�;4-��ϸ�б�
    '����:������
    '����:2013-10-09
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim cbrListView As CommandBar
    Set cbrListView = mcbrListView
    Select Case intTYPE
        Case 1
            cbrListView.Controls.Find(, conMenu_View_LargeICO).Checked = True
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_LargeICO).Checked = True
            cbrListView.Controls.Find(, conMenu_View_MinICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_MinICO).Checked = False
            cbrListView.Controls.Find(, conMenu_View_ListICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_ListICO).Checked = False
            cbrListView.Controls.Find(, conMenu_View_DetailsICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_DetailsICO).Checked = False
        Case 2
            cbrListView.Controls.Find(, conMenu_View_MinICO).Checked = True
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_MinICO).Checked = True
            cbrListView.Controls.Find(, conMenu_View_LargeICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_LargeICO).Checked = False
            cbrListView.Controls.Find(, conMenu_View_ListICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_ListICO).Checked = False
            cbrListView.Controls.Find(, conMenu_View_DetailsICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_DetailsICO).Checked = False
        Case 3
            cbrListView.Controls.Find(, conMenu_View_ListICO).Checked = True
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_ListICO).Checked = True
            cbrListView.Controls.Find(, conMenu_View_MinICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_MinICO).Checked = False
            cbrListView.Controls.Find(, conMenu_View_LargeICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_LargeICO).Checked = False
            cbrListView.Controls.Find(, conMenu_View_DetailsICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_DetailsICO).Checked = False
        Case 4
            cbrListView.Controls.Find(, conMenu_View_DetailsICO).Checked = True
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_DetailsICO).Checked = True
            cbrListView.Controls.Find(, conMenu_View_MinICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_MinICO).Checked = False
            cbrListView.Controls.Find(, conMenu_View_ListICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_ListICO).Checked = False
            cbrListView.Controls.Find(, conMenu_View_LargeICO).Checked = False
            mcbrMenuView.CommandBar.Controls.Find(, conMenu_View_LargeICO).Checked = False
    End Select
    Call mfrmFeeGroupCollectFee.ChangeListViewType(intTYPE)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub PrintReport()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�ش��վ�/�ɿ��鰴ť����
    '����:������
    '����:2013-09-22
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Select Case tabMain.Selected.Index
        Case 1
            Call zl_OpenReport(glngSys, "zl" & glngSys \ 100 & "_BILL_1507", 2)
        Case 2
            Call zl_OpenReport(glngSys, "zl" & glngSys \ 100 & "_INSIDE_1507", 2)
    End Select
End Sub

Private Sub ChargeAndBilllTotal()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�տƱ�ݻ��ܰ�ť����
    '����:������
    '����:2013-09-27
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    If ActiveControl = vsCollectorDetail Then
        With vsCollectorDetail
            Call frmFeeGroupChargeAndBillTotal.ShowMe(Me, BT_�շ�Ա����, mlngModule, mstrPrivs, Val(.TextMatrix(.RowSel, .ColIndex("ID"))), _
                                                      .TextMatrix(.RowSel, .ColIndex("���ʵ���")), .TextMatrix(.RowSel, .ColIndex("����ʱ��")), _
                                                      .TextMatrix(.RowSel, .ColIndex("�շ�Ա")))
        End With
    End If
    If ActiveControl = vsGroupCollectInfo Then
        With vsGroupCollectInfo
            Call frmFeeGroupChargeAndBillTotal.ShowMe(Me, BT_С���տ�, mlngModule, mstrPrivs, Val(.TextMatrix(.RowSel, .ColIndex("ID"))), _
                                                      .TextMatrix(.RowSel, .ColIndex("�տ��")), .TextMatrix(.RowSel, .ColIndex("�տ�ʱ��")), _
                                                      .TextMatrix(.RowSel, .ColIndex("�տ���")))
        End With
    End If
End Sub

Private Sub ViewDetail()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�鿴��ϸ��ť����
    '����:������
    '����:2013-09-22
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strIDs As String, i As Integer
    
    Select Case tabMain.Selected.Index
        Case 0
            With mfrmFeeGroupCollectFee.vsCollectorInfo
                For i = .Row To .RowSel
                    strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
                Next i
                strIDs = Mid(strIDs, 2)
                Call mfrmFeeGroupCollectFee.ChargeRollingListShow(Me, EM_�շ�Ա����, strIDs)
            End With
        Case 1
            Call mfrmFeeGroupRollingCurtain.ViewDetail
        Case 2
            If ActiveControl = vsCollectorDetail Then
                With vsCollectorDetail
                    For i = .Row To .RowSel
                        strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
                    Next i
                    strIDs = Mid(strIDs, 2)
                    Call mfrmFeeGroupCollectFee.ChargeRollingListShow(Me, EM_�շ�Ա����, strIDs)
                End With
                Exit Sub
            End If
            If ActiveControl = vsGroupCollectInfo Then
                With vsGroupCollectInfo
                    For i = .Row To .RowSel
                        strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
                    Next i
                    strIDs = Mid(strIDs, 2)
                    Call mfrmFeeGroupCollectFee.ChargeRollingListShow(Me, EM_С���տ�, strIDs)
                End With
                Exit Sub
            End If
            With vsRollingCurtainHistory
                For i = .Row To .RowSel
                    strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
                Next i
                strIDs = Mid(strIDs, 2)
                Call mfrmFeeGroupCollectFee.ChargeRollingListShow(Me, EM_С������, strIDs)
            End With
    End Select
End Sub

Private Sub ButtonCollectFees()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�շѰ�ť����
    '����:������
    '����:2013-09-22
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim i As Integer, strOut As String, j As Integer, blnRefresh As Boolean
    With mfrmFeeGroupCollectFee.vsCollectorInfo
        If .Rows = 1 Then Exit Sub
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) = -1 Then
                strOut = strOut & "," & Val(.TextMatrix(i, .ColIndex("ID")))
            End If
        Next i
        strOut = Mid(strOut, 2)
        If strOut = "" Then strOut = Val(.TextMatrix(.RowSel, .ColIndex("ID")))
    End With
    With mfrmFeeGroupCollectFee.lvwSubWorker_S
        blnRefresh = frmFeeGroupCollectEdit.ShowMe(Me, mlngModule, Val(Right(.SelectedItem.Key, Len(.SelectedItem.Key) - 1)), _
                                                   .SelectedItem.Text, mlngGroupID, strOut)
    End With
    If blnRefresh = True Then
        Call mfrmFeeGroupCollectFee.AfterCollectEdit
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ButtonRollingCurtainCancel()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�������ϰ�ť����
    '����:������
    '����:2013-09-22
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    With vsRollingCurtainHistory
        If MsgBox("�����ʼ�¼[" & .TextMatrix(.RowSel, .ColIndex("���ʵ���")) & "]���ϣ�ȷ�����ϣ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End With
    
    Call CancelRollingCurtain
    Call SetDefaultHistory(True)
    Call mfrmFeeGroupRollingCurtain.SetDefaultRollingCurtain(False)
    
    With vsCollectorDetail
        .Clear 1
        .Rows = 2
    End With
    
    With vsGroupCollectInfo
        .Clear 1
        .Rows = 2
    End With
End Sub

Public Sub AutoPrint(ByVal lngID As Long, ByVal strNO As String, ByVal intTYPE As Integer)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�տ�/���ʺ��ӡ����
    '���:lngID-��¼ID  strNO-��¼NO    intType-��������:1-�տ�;2-�ɿ�
    '����:������
    '����:2013-10-22
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strRNO As String, strID As String
    
    strRNO = "NO=" & strNO
    strID = "ID=" & lngID
    '��ӡ�տ��վ�
    If intTYPE = 1 Then
        If zlStr.IsHavePrivs(mstrPrivs, "��Ա�ɿ��վ�") = False Then Exit Sub
        Select Case mintReceiptPrint
            Case 1
                Call ReportOpen(gcnOracle, glngSys, "zl" & glngSys \ 100 & "_BILL_1507", Me, strRNO, strID, 2)
            Case 2
                If MsgBox("���Ƿ�Ҫ��ӡ�տ��վݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
                Call ReportOpen(gcnOracle, glngSys, "zl" & glngSys \ 100 & "_BILL_1507", Me, strRNO, strID, 2)
        End Select
    End If
    
    '��ӡ�ɿ���
    If intTYPE = 2 Then
        If zlStr.IsHavePrivs(mstrPrivs, "�ɿ����ӡ") = False Then Exit Sub
        Select Case mintChargeBookPrint
            Case 1
                Call ReportOpen(gcnOracle, glngSys, "zl" & glngSys \ 100 & "_INSIDE_1507", Me, strRNO, strID, 2)
            Case 2
                If MsgBox("���Ƿ�Ҫ��ӡ�ɿ��飿", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
                Call ReportOpen(gcnOracle, glngSys, "zl" & glngSys \ 100 & "_INSIDE_1507", Me, strRNO, strID, 2)
        End Select
    End If
End Sub

Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String, Optional ByVal intTYPE As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ������
    '���:lngSys-ϵͳ��
    '     strReportCode-������
    '     intType-�����������:0-Ĭ��,1-ֱ��Ԥ��,2-ֱ�Ӵ�ӡ,3-�����EXCEL
    '����:������
    '����:2013-09-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNO As String, strID As String
    Select Case tabMain.Selected.Index
        '��ʷ���ʽ���
        Case 2
            With vsRollingCurtainHistory
                If .RowSel < 1 Then Exit Sub
                strNO = "NO=" & .TextMatrix(.RowSel, .ColIndex("���ʵ���"))
                strID = "ID=" & Val(.TextMatrix(.RowSel, .ColIndex("ID")))
            End With
        '���ʽ���
        Case 1
            With mfrmFeeGroupRollingCurtain.vsCollectHistory
                If .RowSel < 1 Then Exit Sub
                strNO = "NO=" & .TextMatrix(.RowSel, .ColIndex("�տ��"))
                strID = "ID=" & Val(.TextMatrix(.RowSel, .ColIndex("ID")))
            End With
        '�տ����
        Case 0
            With mfrmFeeGroupCollectFee.vsCollectorInfo
                If .RowSel < 1 Then Exit Sub
                strNO = "NO=" & .TextMatrix(.RowSel, .ColIndex("���ʵ���"))
                strID = "ID=" & Val(.TextMatrix(.RowSel, .ColIndex("ID")))
            End With
    End Select
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, strNO, strID, intTYPE)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub CancelRollingCurtain()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�������ϲ���
    '����:������
    '����:2013-09-10
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strSQL As String
    With vsRollingCurtainHistory
        strSQL = "Zl_С�����ʼ�¼_Cancel(" & Val(.TextMatrix(.RowSel, .ColIndex("ID"))) & ",'" & UserInfo.���� & _
                 "'," & UserInfo.ID & ",to_date('" & zlDatabase.Currentdate & "','yyyy-MM-dd HH24:mi:ss')," & mlngGroupID & ")"
    End With
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CancelCollect()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�տ����ϲ���
    '����:������
    '����:2013-09-10
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNO As String, strSQL As String
    With mfrmFeeGroupRollingCurtain.vsCollectHistory
        strSQL = "Zl_С���տ��¼_Cancel(" & Val(.TextMatrix(.RowSel, .ColIndex("ID"))) & ",'" & UserInfo.���� & _
                 "',to_date('" & zlDatabase.Currentdate & "','yyyy-MM-dd HH24:mi:ss'))"
    End With
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, blnCollect As Boolean
    Select Case Control.ID
        Case conMenu_Edit_CollectFees '�տť
            If zlStr.IsHavePrivs(mstrPrivs, "�տ�") = False Then Exit Sub
            If tabMain.Selected.Index = 0 Then
                Control.Visible = True
                With mfrmFeeGroupCollectFee.vsCollectorInfo
                    For i = 1 To .Rows - 1
                        If Val(.TextMatrix(i, .ColIndex("ѡ��"))) = -1 Then blnCollect = True
                    Next i
                    Control.Enabled = blnCollect Or (.RowSel >= 1 And .TextMatrix(.RowSel, .ColIndex("ID")) <> "")
                End With
            Else
                Control.Visible = False
                Control.Enabled = False
            End If
        Case conMenu_Edit_RollingCurtain '���ʰ�ť
            If zlStr.IsHavePrivs(mstrPrivs, "����") = False Then Exit Sub
            If tabMain.Selected.Index = 1 Then
                Control.Visible = True
                With mfrmFeeGroupRollingCurtain.vsCollectHistory
                    Control.Enabled = .Rows > 1 And .TextMatrix(1, .ColIndex("ID")) <> ""
                    mfrmFeeGroupRollingCurtain.cmdSendFees.Enabled = Control.Enabled
                End With
            Else
                Control.Visible = False
                Control.Enabled = False
                mfrmFeeGroupRollingCurtain.cmdSendFees.Enabled = False
            End If
        Case conMenu_Edit_CollectFees_Cancel '�տ����ϰ�ť
            If zlStr.IsHavePrivs(mstrPrivs, "�տ�����") = False Then Exit Sub
            If tabMain.Selected.Index = 1 Then
                Control.Visible = True
                With mfrmFeeGroupRollingCurtain.vsCollectHistory
                    Control.Enabled = .RowSel >= 1 And .TextMatrix(1, .ColIndex("ID")) <> ""
                End With
            Else
                Control.Visible = False
                Control.Enabled = False
            End If
        Case conMenu_View_Detail '�鿴��ϸ��ť
            Select Case tabMain.Selected.Index
                Case 1
                    With mfrmFeeGroupRollingCurtain.vsCollectHistory
                        Control.Enabled = .RowSel >= 1 And .TextMatrix(1, .ColIndex("ID")) <> ""
                    End With
                Case 2
                    With vsRollingCurtainHistory
                        Control.Enabled = .RowSel >= 1 And .TextMatrix(1, .ColIndex("ID")) <> ""
                    End With
                Case 0
                    With mfrmFeeGroupCollectFee.vsCollectorInfo
                        Control.Enabled = .RowSel >= 1 And .TextMatrix(1, .ColIndex("ID")) <> ""
                    End With
            End Select
        Case conMenu_Edit_RollingCurtain_Cancel '�������ϰ�ť
            If zlStr.IsHavePrivs(mstrPrivs, "��������") = False Then Exit Sub
            If tabMain.Selected.Index = 2 Then
                Control.Visible = True
                With vsRollingCurtainHistory
                    If .TextMatrix(.RowSel, .ColIndex("�����տ���")) = "" And _
                       .TextMatrix(.RowSel, .ColIndex("������")) = "" _
                       And .RowSel >= 1 And .TextMatrix(1, .ColIndex("ID")) <> "" Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                End With
            Else
                Control.Visible = False
                Control.Enabled = False
            End If
            
        Case conMenu_Edit_ReprintReceipt    '�ش��վݰ�ť
            If zlStr.IsHavePrivs(mstrPrivs, "��Ա�ɿ��վ�") = False Then Exit Sub
            If zlStr.IsHavePrivs(mstrPrivs, "�ش��վ�") = False Then Exit Sub
            Select Case tabMain.Selected.Index
                Case 0
                    Control.Visible = False
                    Control.Enabled = False
                Case 1
                    Control.Visible = True
                    With mfrmFeeGroupRollingCurtain.vsCollectHistory
                        Control.Enabled = .RowSel >= 1 And .TextMatrix(1, .ColIndex("ID")) <> ""
                    End With
                Case 2
                    Control.Visible = False
                    Control.Enabled = False
            End Select
            
        Case conMenu_Edit_ChargeBook_Reprint '�ش�ɿ��鰴ť
            If zlStr.IsHavePrivs(mstrPrivs, "�ɿ����ӡ") = False Then Exit Sub
            If zlStr.IsHavePrivs(mstrPrivs, "�ش�ɿ���") = False Then Exit Sub
            If tabMain.Selected.Index = 2 Then
                Control.Visible = True
                With vsRollingCurtainHistory
                    Control.Enabled = .RowSel >= 1 And .TextMatrix(1, .ColIndex("ID")) <> "" And .TextMatrix(.RowSel, .ColIndex("����ʱ��")) = ""
                End With
            Else
                Control.Visible = False
                Control.Enabled = False
            End If
        Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll
            With mfrmFeeGroupCollectFee.vsCollectorInfo
                If .TextMatrix(1, .ColIndex("ID")) <> "" And tabMain(0).Selected = True Then
                    Control.Visible = True
                    Control.Enabled = True
                Else
                    Control.Visible = False
                    Control.Enabled = False
                End If
            End With
    End Select
End Sub

Private Sub cmdReFilter_Click()
    If (dtpTerminateTime.Value - dtpStartTime.Value) > 178 Then
        MsgBox "��ѯ��ʱ�䷶Χ���ܳ������꣬������ѡ��ʱ�䷶Χ��", vbInformation, gstrSysName
        If dtpStartTime.Visible And dtpStartTime.Enabled Then dtpStartTime.SetFocus
        Exit Sub
    End If
    Call SetDefaultHistory(True)
End Sub

Public Sub FailInit()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�ⲿ�������ж�ش���
    '����:������
    '����:2013-10-11
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    mblnCancel = True
End Sub

Public Sub SetGroupID(ByVal lngGroupID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡѡ��Ľɿ���ID
    '���:lngGroupID-�ɿ���ID
    '����:������
    '����:2013-11-07
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    mlngGroupID = lngGroupID
End Sub

Private Sub SetDefaultHistory(ByVal blnReload As Boolean)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:����Ĭ����ʷ���ʽ�����Ϣ
    '����:������
    '����:2013-09-11
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strSQL As String, rsTmp As New ADODB.Recordset, i As Integer
    
    With vsRollingCurtainHistory
        .Rows = 1
        strSQL = "" & _
        "Select NO, ��ʼʱ��, ��ֹʱ��, �Ǽ���, �Ǽ�ʱ��, С���տ���, С���տ�ʱ��, Trim(to_char(��Ԥ����,'99999999990.00')) As ��Ԥ����, " & _
        "Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�, Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�, �����տ���, �����տ�ʱ��, ������, ����ʱ��, ժҪ, ID" & vbNewLine & _
        "From ��Ա�սɼ�¼" & vbNewLine & _
        "Where ��¼���� = 3 And �Ǽ��� = [1] And �ɿ���ID = [4]" & _
        "      And �Ǽ�ʱ�� Between [2] And [3] Order By �Ǽ�ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, CDate(dtpStartTime.Value), CDate(dtpTerminateTime.Value), mlngGroupID)
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("���ʵ���")) = NVL(rsTmp!NO)
            .TextMatrix(.Rows - 1, .ColIndex("��ʼʱ��")) = NVL(rsTmp!��ʼʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ֹʱ��")) = NVL(rsTmp!��ֹʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTmp!�Ǽ���)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = NVL(rsTmp!�Ǽ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = NVL(rsTmp!��Ԥ����)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTmp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTmp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ���")) = NVL(rsTmp!�����տ���)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ�ʱ��")) = NVL(rsTmp!�����տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTmp!������)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = NVL(rsTmp!����ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ע")) = NVL(rsTmp!ժҪ)
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTmp!ID)
            rsTmp.MoveNext
        Loop
        'Set .DataSource = rsTmp
        .AutoSize 1, .Cols - 1
        Call zl_vsGrid_Para_Restore(mlngModule, vsRollingCurtainHistory, Me.Caption, "��ʷ������Ϣ", False)
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�����տ�ʱ��")) = "" Then
                .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &HC0FFFF
            End If
            If .TextMatrix(i, .ColIndex("������")) <> "" Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &HFF
            End If
        Next i
        .Select 0, 0
        If .Rows = 1 Then .Rows = 2
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetTabControl()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:����TAB�ؼ�
    '����:������
    '����:2013-09-04
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With tabMain
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        Set .PaintManager.Font = txtSendFeeNO.Font
        .InsertItem EM_Tab_�տ�, " �տ�  ", mfrmFeeGroupCollectFee.hWnd, 0
        .InsertItem EM_Tab_����, " ����  ", mfrmFeeGroupRollingCurtain.hWnd, 0
        .InsertItem EM_Tab_��ʷ������Ϣ, " ��ʷ������Ϣ  ", picHistory.hWnd, 0
        .Item(0).Selected = True
        .PaintManager.BoldSelected = True
    End With
    
    With tabHistory
        Set .PaintManager.Font = txtSendFeeNO.Font
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        .InsertItem EM_Tab_�տƱ�ݻ���, " �տƱ�ݻ���  ", mfrmChargeBillTotalHistory.hWnd, 0
        .InsertItem EM_Tab_���տ���Ϣ, " ���տ���Ϣ  ", picGroupCollectInfo.hWnd, 0
        .InsertItem EM_Tab_�շ�Ա������Ϣ, " �շ�Ա������Ϣ  ", picSendInfo.hWnd, 0
        .Item(0).Selected = True
        .PaintManager.BoldSelected = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtpStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtpTerminateTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Function LoadGroup() As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ������Ա���еĽɿ���
    '����:�ɹ�����True,ʧ�ܷ���False
    '����:������
    '����:2013-11-06
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select Id,������,������ID From ����ɿ���� Where (ɾ������ Is Null or ɾ������ Between Sysdate And to_date('3000-01-01','YYYY-MM-DD')) And ������ID=[1]"
    strSQL = strSQL & " Union Select A.��ID,B.������,A.�鳤ID From �������鳤���� A,����ɿ���� B Where A.��ID=B.ID And A.�鳤ID=[1] And (B.ɾ������ Is Null or B.ɾ������ Between Sysdate And to_date('3000-01-01','YYYY-MM-DD'))"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    If rsTmp.RecordCount = 0 Then
        LoadGroup = False
        Exit Function
    End If
    
    If rsTmp.RecordCount = 1 Then
        mlngGroupID = Val(rsTmp!ID)
    Else
        frmFeeGroupSelDept.Show vbModal, Me
    End If
    
    LoadGroup = True
    Exit Function
errHandle:
    LoadGroup = False
    If ErrCenter = 1 Then Resume
End Function

Private Sub Form_Load()
    mlngModule = glngModul
    mstrPrivs = gstrPrivs
    mblnCancel = False
    
    If LoadGroup = False Then
        MsgBox "�޷���ȡ��������Ϣ,��ȷ�����ǲ���ɿ����鳤!", vbCritical, gstrSysName
    End If
    
    mintChargeBookPrint = Val(zlDatabase.GetPara("�ɿ����ӡ��ʽ", glngSys, mlngModule, "0"))
    mintReceiptPrint = Val(zlDatabase.GetPara("�տ��վݴ�ӡ��ʽ", glngSys, mlngModule, "0"))
    '��ʼ����TAB�������
    If mfrmFeeGroupCollectFee Is Nothing Then Set mfrmFeeGroupCollectFee = New frmFeeGroupCollectFee
    Call mfrmFeeGroupCollectFee.InitMe(Me, mlngModule, mstrPrivs, mlngGroupID)
    Load mfrmFeeGroupCollectFee
    
    If mfrmFeeGroupRollingCurtain Is Nothing Then Set mfrmFeeGroupRollingCurtain = New frmFeeGroupRollingCurtain
    Call mfrmFeeGroupRollingCurtain.InitMe(mlngModule, mstrPrivs, mlngGroupID)
    Load mfrmFeeGroupRollingCurtain
    
    mobjChargeBillHistory.SetFontSize lblStartTime.Font.Size
    '��ͬ�����Ʊ����Ϣģ������
    
    Set mfrmChargeBillTotalHistory = mobjChargeBillHistory.GetChargeAndBillTotalForm
    
    Call zlDefCommandBars
    '����TAB��Ϣ
    Call SetTabControl
    stbThis.Panels(3).Text = UserInfo.����
    Call SetDateUnit
    Call SetGrid
    mfrmFeeGroupCollectFee.lblCurrentMoney(0).Caption = " ��ǰ�ݴ��:"
    '��ʷ������Ϣ����Ĭ����Ϣ
    Call SetDefaultHistory(False)
    mstrTitle = "�������տ����"
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
End Sub

Private Sub SetGrid()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSF�ؼ�
    '����:������
    '����:2013-10-13
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    With vsRollingCurtainHistory
        For i = 0 To .Cols - 1
            If .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "������" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "����" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "���ʵ���" Or .ColKey(i) = "��ʼʱ��" Or .ColKey(i) = "��ֹʱ��" Or _
            .ColKey(i) = "������" Or .ColKey(i) = "����ʱ��" Then .ColData(i) = "1|0"
            If i = .ColIndex("��Ԥ����") Or i = .ColIndex("����ϼ�") Or i = .ColIndex("����ϼ�") Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
    End With
    
    With vsGroupCollectInfo
        For i = 0 To .Cols - 1
            If .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "�տ���" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "����" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "�տ��" Or .ColKey(i) = "�տ�ʱ��" Then .ColData(i) = "1|0"
        Next
    End With
    
    With vsCollectorDetail
        For i = 0 To .Cols - 1
            If .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "�շ�Ա" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "����" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "���ʵ���" Or .ColKey(i) = "��ʼʱ��" Or .ColKey(i) = "��ֹʱ��" Then .ColData(i) = "1|0"
        Next
    End With
    
    zl_vsGrid_Para_Restore mlngModule, vsGroupCollectInfo, Me.Caption, "С���տ���ϸ", False
    zl_vsGrid_Para_Restore mlngModule, vsCollectorDetail, Me.Caption, "��ʷ�շ�Ա������ϸ", False
    zl_vsGrid_Para_Restore mlngModule, vsRollingCurtainHistory, Me.Caption, "��ʷ������Ϣ", False
End Sub

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:������
    '����:2013-09-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long, intActive As Integer
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String
    
    If mfrmFeeGroupCollectFee.ActiveControl Is mfrmFeeGroupCollectFee.vsCollectorInfo Then
        With mfrmFeeGroupCollectFee.vsCollectorInfo
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("ID"))) = 0 Then Exit Sub
        End With
        Set vsBill = mfrmFeeGroupCollectFee.vsCollectorInfo: strTittle = GetUnitName & "�շ�Ա������Ϣ"
    End If
    If mfrmFeeGroupRollingCurtain.ActiveControl Is mfrmFeeGroupRollingCurtain.vsCollectHistory Then
        With mfrmFeeGroupRollingCurtain.vsCollectHistory
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("ID"))) = 0 Then Exit Sub
        End With
        Set vsBill = mfrmFeeGroupRollingCurtain.vsCollectHistory: strTittle = GetUnitName & "�������տ��¼"
    End If
    If mfrmFeeGroupRollingCurtain.ActiveControl Is mfrmFeeGroupRollingCurtain.vsSubCollectorInfo Then
        With mfrmFeeGroupRollingCurtain.vsSubCollectorInfo
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("ID"))) = 0 Then Exit Sub
        End With
        Set vsBill = mfrmFeeGroupRollingCurtain.vsSubCollectorInfo: strTittle = GetUnitName & "�շ�Ա������ϸ"
    End If
    If Me.ActiveControl Is vsRollingCurtainHistory Then
        With vsRollingCurtainHistory
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("ID"))) = 0 Then Exit Sub
        End With
        Set vsBill = vsRollingCurtainHistory: strTittle = GetUnitName & "�����������ʼ�¼"
    End If
    If Me.ActiveControl Is vsGroupCollectInfo Then
        With vsGroupCollectInfo
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("ID"))) = 0 Then Exit Sub
        End With
        Set vsBill = vsGroupCollectInfo: strTittle = GetUnitName & "�������տ��¼"
    End If
    If Me.ActiveControl Is vsCollectorDetail Then
        With vsCollectorDetail
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("ID"))) = 0 Then Exit Sub
        End With
        Set vsBill = vsCollectorDetail: strTittle = GetUnitName & "�շ�Ա������Ϣ"
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTittle
    
    If mfrmFeeGroupCollectFee.ActiveControl Is mfrmFeeGroupCollectFee.vsCollectorInfo Then
        objRow.Add "�ɿ���:" & mfrmFeeGroupCollectFee.dkpCollectFees.FindPane(EM_Pan_��Ա��).Title
        objRow.Add "�տ�Ա:" & mfrmFeeGroupCollectFee.lvwSubWorker_S.SelectedItem.Text
    End If
    If mfrmFeeGroupRollingCurtain.ActiveControl Is mfrmFeeGroupRollingCurtain.vsCollectHistory Then
        objRow.Add "�ɿ���:" & mfrmFeeGroupCollectFee.dkpCollectFees.FindPane(EM_Pan_��Ա��).Title
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add "�ϴ�����ʱ��:" & mfrmFeeGroupRollingCurtain.dtpLastTime.Value
        objRow.Add "��ֹʱ��:" & mfrmFeeGroupRollingCurtain.dtpEndTime.Value
    End If
    If mfrmFeeGroupRollingCurtain.ActiveControl Is mfrmFeeGroupRollingCurtain.vsSubCollectorInfo Then
        objRow.Add "�ɿ���:" & mfrmFeeGroupCollectFee.dkpCollectFees.FindPane(EM_Pan_��Ա��).Title
        objRow.Add "�տ�Ա:" & mfrmFeeGroupRollingCurtain.vsSubCollectorInfo.TextMatrix(1, mfrmFeeGroupRollingCurtain.vsSubCollectorInfo.ColIndex("�տ�Ա"))
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add "С���տ��:" & mfrmFeeGroupRollingCurtain.vsCollectHistory.TextMatrix(mfrmFeeGroupRollingCurtain.vsCollectHistory.RowSel, mfrmFeeGroupRollingCurtain.vsCollectHistory.ColIndex("�տ��"))
    End If
    If Me.ActiveControl Is vsRollingCurtainHistory Then
        objRow.Add "�ɿ���:" & mfrmFeeGroupCollectFee.dkpCollectFees.FindPane(EM_Pan_��Ա��).Title
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add "��ʼʱ��:" & dtpStartTime.Value
        objRow.Add "��ֹʱ��:" & dtpTerminateTime.Value
    End If
    If Me.ActiveControl Is vsGroupCollectInfo Then
        objRow.Add "�ɿ���:" & mfrmFeeGroupCollectFee.dkpCollectFees.FindPane(EM_Pan_��Ա��).Title
        objRow.Add "���ʵ���:" & vsRollingCurtainHistory.TextMatrix(vsRollingCurtainHistory.RowSel, vsRollingCurtainHistory.ColIndex("���ʵ���"))
    End If
    
    If Me.ActiveControl Is vsCollectorDetail Then
        objRow.Add "�ɿ���:" & mfrmFeeGroupCollectFee.dkpCollectFees.FindPane(EM_Pan_��Ա��).Title
        objRow.Add "���ʵ���:" & vsRollingCurtainHistory.TextMatrix(vsRollingCurtainHistory.RowSel, vsRollingCurtainHistory.ColIndex("���ʵ���"))
    End If
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    If vsBill Is Nothing Then Exit Sub
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsBill
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsBill
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
    '�ָ�
    With vsBill
        For i = 0 To .Cols - 1
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetDateUnit()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�������ڿؼ���ʽ����
    '����:������
    '����:2013-09-09
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    dtpStartTime.Format = dtpCustom
    dtpStartTime.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpStartTime.Value = zlDatabase.Currentdate
    dtpStartTime.Value = dtpStartTime.Value - 7
    dtpTerminateTime.Format = dtpCustom
    dtpTerminateTime.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpTerminateTime.Value = zlDatabase.Currentdate
End Sub

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:������
    '����:2013-09-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
        
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    '��ʼ������
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
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CollectFees, "�տ�(&S)")
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�տ�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CollectFees_Cancel, "�տ�����(&C)")
        mcbrControl.IconId = 3589
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�տ�����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "����(&Z)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 227
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain_Cancel, "��������(&D)")
        mcbrControl.IconId = 229
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "�ֽ�㳮(&E)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3590
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "�鿴��ϸ����(&V)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 2322
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintReceipt, "�ش��վ�(&R)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�ش��վ�") And zlStr.IsHavePrivs(mstrPrivs, "��Ա�ɿ��վ�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeBook_Reprint, "�ش�ɿ���(&R)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�ش�ɿ���") And zlStr.IsHavePrivs(mstrPrivs, "�ɿ����ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��")
    End With
    
    Set mcbrMenuView = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&E)", -1, False)
    mcbrMenuView.ID = conMenu_ViewPopup
    With mcbrMenuView.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        cbrControl.Checked = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "��ͼ��"): mcbrControl.BeginGroup = True
        mcbrControl.Checked = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "Сͼ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "�б�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "��ϸ�б�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��"): mcbrControl.BeginGroup = True
        mcbrControl.Checked = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&K)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '���������˵�
    Set mcbrPopupMain = cbsThis.Add("�����˵�1", xtpBarPopup)
    With mcbrPopupMain.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ChargeAndBilllTotal, "�տ���ܼ�Ʊ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "��ʾ��ϸ�տ�")
    End With
    
    Set mcbrListView = cbsThis.Add("��Ա�����˵�", xtpBarPopup)
    With mcbrListView.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "��ͼ��")
        mcbrControl.Checked = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "Сͼ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "�б�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "��ϸ�б�")
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("S"), conMenu_Edit_RollingCurtain
        .Add FCONTROL, Asc("D"), conMenu_Edit_RollingCurtain_Cancel
        .Add FCONTROL, Asc("R"), conMenu_Edit_ChargeBook_Reprint
        .Add FCONTROL, Asc("C"), conMenu_Edit_CollectFees_Cancel
        .Add FCONTROL, Asc("A"), conMenu_Edit_SelAll
        .Add FCONTROL, Asc("C"), conMenu_Edit_ClsAll
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F6, conMenu_Edit_CheckCash
        .Add 0, VK_F2, conMenu_Edit_CollectFees
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_File_Exit
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ModifyStyle &H400000, 0
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CollectFees, "�տ�"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�տ�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CollectFees_Cancel, "�տ�����")
        mcbrControl.IconId = 3589
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�տ�����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "����"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 227
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain_Cancel, "��������")
        mcbrControl.IconId = 229
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "�ֽ�㳮"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3590
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "�鿴��ϸ"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 2322
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintReceipt, "�ش��վ�")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�ش��վ�") And zlStr.IsHavePrivs(mstrPrivs, "��Ա�ɿ��վ�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeBook_Reprint, "�ش�ɿ���")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�ش�ɿ���") And zlStr.IsHavePrivs(mstrPrivs, "�ɿ����ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): mcbrControl.BeginGroup = True
    End With
    
    For Each mcbrControl In mcbrToolBar.Controls
        If mcbrControl.ID <> conMenu_Edit_UserType Then
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

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tabMain.Width = Me.Width - 225
    If cbsThis(2).Visible Then
        If cbsThis.Options.LargeIcons Then
            tabMain.Top = 870
        Else
            tabMain.Top = 750
        End If
    Else
        tabMain.Top = 370
    End If
    '����״̬����������
    If stbThis.Visible Then
        tabMain.Height = Me.Height - 910 - tabMain.Top
    Else
        tabMain.Height = Me.Height - 910 - tabMain.Top + stbThis.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    '�洢�б�ĸ��Ի�����(����)
    zl_vsGrid_Para_Save mlngModule, mfrmFeeGroupRollingCurtain.vsCollectHistory, Me.Caption, "С���տ���Ϣ", False
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtainHistory, Me.Caption, "��ʷ������Ϣ", False
    zl_vsGrid_Para_Save mlngModule, mfrmFeeGroupCollectFee.vsCollectorInfo, Me.Caption, "�շ�Ա������Ϣ", False
    zl_vsGrid_Para_Save mlngModule, mfrmFeeGroupRollingCurtain.vsSubCollectorInfo, Me.Caption, "�շ�Ա������ϸ", False
    zl_vsGrid_Para_Save mlngModule, vsGroupCollectInfo, Me.Caption, "С���տ���ϸ", False
    zl_vsGrid_Para_Save mlngModule, vsCollectorDetail, Me.Caption, "��ʷ�շ�Ա������ϸ", False
    
    SaveWinState Me, App.ProductName, mstrTitle
    'ж�ؼ��ش������
    If Not mfrmFeeGroupCollectFee Is Nothing Then Unload mfrmFeeGroupCollectFee
    Set mfrmFeeGroupCollectFee = Nothing
    If Not mfrmFeeGroupRollingCurtain Is Nothing Then Unload mfrmFeeGroupRollingCurtain
    Set mfrmFeeGroupRollingCurtain = Nothing
    If Not frmMoneyEnum Is Nothing Then Unload frmMoneyEnum
    Set frmMoneyEnum = Nothing
    If Not frmFeeGroupSetting Is Nothing Then Unload frmFeeGroupSetting
    Set frmFeeGroupSetting = Nothing
    If Not frmFeeGroupCollectEdit Is Nothing Then Unload frmFeeGroupCollectEdit
    Set frmFeeGroupCollectEdit = Nothing
    If Not frmFeeGroupChargeAndBillTotal Is Nothing Then Unload frmFeeGroupChargeAndBillTotal
    Set frmFeeGroupChargeAndBillTotal = Nothing
    If Not mfrmChargeBillTotalHistory Is Nothing Then Unload mfrmChargeBillTotalHistory
    Set mfrmChargeBillTotalHistory = Nothing
    If Not mobjChargeBillHistory Is Nothing Then Set mobjChargeBillHistory = Nothing
    If Not frmFeeGroupSelDept Is Nothing Then Unload frmFeeGroupSelDept
    Set frmFeeGroupSelDept = Nothing
End Sub


Private Sub mfrmFeeGroupCollectFee_ShowPopupMenu(ByVal bytType As Byte)
    Dim cbrPopup As CommandBarPopup
    
    If bytType = 1 Then
        Set cbrPopup = cbsThis.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_EditPopup, , 1)
    ElseIf bytType = 2 Then
        If Not mcbrListView Is Nothing Then mcbrListView.ShowPopup
        Exit Sub
    End If
    
    If cbrPopup Is Nothing Then Exit Sub
    cbrPopup.CommandBar.ShowPopup
End Sub

Private Sub picGroupCollectInfo_Resize()
    With vsGroupCollectInfo
        .Width = picGroupCollectInfo.Width
        .Height = picGroupCollectInfo.Height
    End With
End Sub

Private Sub picHistory_Resize()
    On Error Resume Next
    cmdReFilter.Left = picHistory.Width - 2200
    If cmdReFilter.Left < txtSendFeeNO.Left + txtSendFeeNO.Width + 200 Then
        cmdReFilter.Left = txtSendFeeNO.Left + txtSendFeeNO.Width + 200
    End If
    With vsRollingCurtainHistory
        .Width = picHistory.Width
        .Height = picHistory.Height * 0.35
    End With
    tabHistory.Top = vsRollingCurtainHistory.Top + vsRollingCurtainHistory.Height + 50
    tabHistory.Width = picHistory.Width
    tabHistory.Height = picHistory.Height - tabHistory.Top - 15
End Sub

Private Sub picSendInfo_Resize()
    With vsCollectorDetail
        .Width = picSendInfo.Width
        .Height = picSendInfo.Height
    End With
End Sub

Private Sub tabMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    mfrmFeeGroupCollectFee.ClearChargeAndBillTotalForm
    mobjChargeBillHistory.ClearChargeAndBillTotalForm
    mfrmFeeGroupRollingCurtain.ClearChargeAndBillTotalForm
    
    If mfrmFeeGroupRollingCurtain.vsCollectHistory.RowSel = -1 Then Exit Sub
    mfrmFeeGroupCollectFee.lblCurrentMoney(0).Caption = " ��ǰ�ݴ��:"
    '������б�
    mfrmFeeGroupRollingCurtain.vsSubCollectorInfo.Clear 1
    mfrmFeeGroupRollingCurtain.vsSubCollectorInfo.Rows = 2
    vsCollectorDetail.Clear 1
    vsCollectorDetail.Rows = 2
    vsGroupCollectInfo.Clear 1
    vsGroupCollectInfo.Rows = 2
    Select Case Item.Index
        Case 0
            With mfrmFeeGroupCollectFee.vsCollectorInfo
                .Clear 1
                .Rows = 2
            End With
            Call zl_vsGrid_Para_Restore(mlngModule, mfrmFeeGroupCollectFee.vsCollectorInfo, Me.Caption, "�շ�Ա������Ϣ", False)
            mcbrToolBar.FindControl(, conMenu_Edit_CollectFees_Cancel).BeginGroup = False
            mcbrToolBar.FindControl(, conMenu_Edit_RollingCurtain_Cancel).BeginGroup = False
        Case 1
            Call mfrmFeeGroupRollingCurtain.RefreshPage
            mcbrToolBar.FindControl(, conMenu_Edit_CollectFees_Cancel).BeginGroup = True
            mcbrToolBar.FindControl(, conMenu_Edit_RollingCurtain_Cancel).BeginGroup = False
        Case 2
            Call zl_vsGrid_Para_Restore(mlngModule, vsRollingCurtainHistory, Me.Caption, "��ʷ������Ϣ", False)
            Call zl_vsGrid_Para_Restore(mlngModule, vsGroupCollectInfo, Me.Caption, "С���տ���ϸ", False)
            Call zl_vsGrid_Para_Restore(mlngModule, vsCollectorDetail, Me.Caption, "��ʷ�շ�Ա������ϸ", False)
            dtpStartTime.Value = zlDatabase.Currentdate
            dtpStartTime.Value = dtpStartTime.Value - 7
            dtpTerminateTime.Value = zlDatabase.Currentdate
            Call SetDefaultHistory(True)
            mcbrToolBar.FindControl(, conMenu_Edit_CollectFees_Cancel).BeginGroup = False
            mcbrToolBar.FindControl(, conMenu_Edit_RollingCurtain_Cancel).BeginGroup = True
    End Select
End Sub

Private Sub txtSendFeeNO_GotFocus()
    Call zlControl.TxtSelAll(txtSendFeeNO)
End Sub

Private Sub txtSendFeeNO_KeyPress(KeyAscii As Integer)
    '���Ƶ�������(��ĸ������)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
       (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        If txtSendFeeNO.Text = "" Then
            KeyAscii = 0
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        Dim i As Integer
        '��ȫƥ�����뵥��
        With vsRollingCurtainHistory
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ʵ���")) = txtSendFeeNO.Text Then
                    If .Visible And .Enabled Then .SetFocus
                    .Select i, 0
                    .TopRow = i
                    Exit Sub
                End If
            Next i
        End With
        
        '�Զ��������뵥��,�ٴν��в���
        txtSendFeeNO.Text = GetFullNO(txtSendFeeNO.Text, 139)
        With vsRollingCurtainHistory
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ʵ���")) = txtSendFeeNO.Text Then
                    If .Visible And .Enabled Then .SetFocus
                    .Select i, 0
                    .TopRow = i
                    Exit Sub
                End If
            Next i
        End With
        MsgBox "û���ҵ����ʵ���[" & txtSendFeeNO.Text & "]�ļ�¼��", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(txtSendFeeNO)
    End If
End Sub

Private Sub vsCollectorDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call zl_VsGridRowChange(vsCollectorDetail, OldRow, NewRow, OldCol, NewCol)
    With vsCollectorDetail
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsCollectorDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsCollectorDetail, Me.Caption, "��ʷ�շ�Ա������ϸ", False)
End Sub

Private Sub vsCollectorDetail_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsCollectorDetail_DblClick()
    With vsCollectorDetail
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        Call mfrmFeeGroupCollectFee.ChargeRollingListShow(Me, EM_�շ�Ա����, Val(.TextMatrix(.RowSel, .ColIndex("ID"))))
    End With
End Sub

Private Sub vsCollectorDetail_GotFocus()
    Call zl_VsGridGotFocus(vsCollectorDetail)
    With vsCollectorDetail
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsCollectorDetail_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsCollectorDetail)
End Sub

Private Sub vsGroupCollectInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call zl_VsGridRowChange(vsGroupCollectInfo, OldRow, NewRow, OldCol, NewCol)
    With vsGroupCollectInfo
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsGroupCollectInfo_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsGroupCollectInfo, Me.Caption, "С���տ���ϸ", False)
End Sub

Private Sub vsGroupCollectInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsGroupCollectInfo_DblClick()
    With vsGroupCollectInfo
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        Call mfrmFeeGroupCollectFee.ChargeRollingListShow(Me, EM_С���տ�, Val(.TextMatrix(.RowSel, .ColIndex("ID"))))
    End With
End Sub

Private Sub vsGroupCollectInfo_GotFocus()
    Call zl_VsGridGotFocus(vsGroupCollectInfo)
    With vsGroupCollectInfo
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsGroupCollectInfo_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsGroupCollectInfo)
End Sub

Private Sub vsGroupCollectInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intRow As Integer
    With vsGroupCollectInfo
        If .TextMatrix(1, .ColIndex("ID")) = "" Then Exit Sub
        If Button = 2 Then
            If .MouseRow < 1 Then Exit Sub
            If .MouseRow > .Rows - 1 Then Exit Sub
            If .Enabled And .Visible Then .SetFocus
            .Select .MouseRow, 0
            mcbrPopupMain.ShowPopup
        End If
    End With
End Sub

Private Sub vsRollingCurtainHistory_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error GoTo errHandle
    Dim i As Integer
    If OldRow = NewRow Then Exit Sub
    
    With vsRollingCurtainHistory
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then
            Exit Sub
        End If
    End With
    
    With vsGroupCollectInfo
        .Rows = 1
        Dim strSQL As String, rsTmp As New ADODB.Recordset
        strSQL = "" & _
        "Select No, �տ�Ա, �Ǽ�ʱ�� As �տ�ʱ��, Trim(to_char(��Ԥ����,'99999999990.00')) As ��Ԥ����, " & _
        "       Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�, Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�, �����տ���, �����տ�ʱ��, ժҪ, Id" & vbNewLine & _
        "From ��Ա�սɼ�¼ " & vbNewLine & _
        "Where ��¼���� = 2 And С������id = [1]"
        With vsRollingCurtainHistory
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.RowSel, .ColIndex("ID"))))
        End With
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("�տ��")) = NVL(rsTmp!NO)
            .TextMatrix(.Rows - 1, .ColIndex("�տ�ʱ��")) = NVL(rsTmp!�տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = NVL(rsTmp!��Ԥ����)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTmp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTmp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("�տ���")) = NVL(rsTmp!�տ�Ա)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ���")) = NVL(rsTmp!�����տ���)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ�ʱ��")) = NVL(rsTmp!�����տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ע")) = NVL(rsTmp!ժҪ)
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTmp!ID)
            rsTmp.MoveNext
        Loop
        'Set .DataSource = rsTmp
        .AutoSize 1, .Cols - 1
        zl_vsGrid_Para_Restore mlngModule, vsGroupCollectInfo, Me.Caption, "С���տ���ϸ", False
        If .Rows = 1 Then .Rows = 2
    End With
    
    With vsCollectorDetail
        .Rows = 1
        strSQL = "" & _
        "Select No, Substr(Decode(�Ƿ�Һ�,1,',�Һ�','') || Decode(�Ƿ���￨,1,',���￨','') || Decode(�Ƿ����ѿ�,1,',���ѿ�','') || Decode(�Ƿ��շ�,1,',�շ�','') || Decode(�Ƿ����,1,',����','') || Decode(Ԥ�����,1,',Ԥ��',2,',����Ԥ��',3,',סԺԤ��',''),2) As �������," & _
        "       �Ǽ�ʱ�� As ����ʱ��, �տ�Ա, ��ʼʱ��, ��ֹʱ��, Trim(to_char(��Ԥ����,'99999999990.00')) As ��Ԥ����, " & _
        "       Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�, Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�, С���տ���, С���տ�ʱ��, " & _
        "       �����տ���, �����տ�ʱ��, ժҪ, Id " & vbNewLine & _
        "From ��Ա�սɼ�¼ " & vbNewLine & _
        "Where ��¼���� = 1 And С������id = [1]"
        With vsRollingCurtainHistory
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.RowSel, .ColIndex("ID"))))
        End With
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("���ʵ���")) = NVL(rsTmp!NO)
            .TextMatrix(.Rows - 1, .ColIndex("�������")) = NVL(rsTmp!�������)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = NVL(rsTmp!����ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("�շ�Ա")) = NVL(rsTmp!�տ�Ա)
            .TextMatrix(.Rows - 1, .ColIndex("��ʼʱ��")) = NVL(rsTmp!��ʼʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ֹʱ��")) = NVL(rsTmp!��ֹʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = NVL(rsTmp!��Ԥ����)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTmp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTmp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ���")) = NVL(rsTmp!С���տ���)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ�ʱ��")) = NVL(rsTmp!С���տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ���")) = NVL(rsTmp!�����տ���)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ�ʱ��")) = NVL(rsTmp!�����տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ע")) = NVL(rsTmp!ժҪ)
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTmp!ID)
            rsTmp.MoveNext
        Loop
        'Set .DataSource = rsTmp
        .AutoSize 1, .Cols - 1
        zl_vsGrid_Para_Restore mlngModule, vsCollectorDetail, Me.Caption, "��ʷ�շ�Ա������ϸ", False
        If .Rows = 1 Then .Rows = 2
    End With
    With vsRollingCurtainHistory
        mobjChargeBillHistory.LoadChargeAndBillTotalData Me, mlngModule, mstrPrivs, EM_С������, .TextMatrix(.RowSel, .ColIndex("ID"))
        Call zl_VsGridRowChange(vsRollingCurtainHistory, OldRow, NewRow, OldCol, NewCol)
        If .TextMatrix(OldRow, .ColIndex("�����տ�ʱ��")) = "" Then
            .Cell(flexcpBackColor, OldRow, 1, OldRow, .Cols - 1) = &HC0FFFF
        End If
            .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub vsCollectorDetail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intRow As Integer
    With vsCollectorDetail
        If .TextMatrix(1, .ColIndex("ID")) = "" Then Exit Sub
        If Button = 2 Then
            If .MouseRow < 1 Then Exit Sub
            If .MouseRow > .Rows - 1 Then Exit Sub
            If .Enabled And .Visible Then .SetFocus
            .Select .MouseRow, 0
            mcbrPopupMain.ShowPopup
        End If
    End With
End Sub

Private Sub vsRollingCurtainHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsRollingCurtainHistory, Me.Caption, "��ʷ������Ϣ", False)
End Sub

Private Sub vsRollingCurtainHistory_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsRollingCurtainHistory_DblClick()
    With vsRollingCurtainHistory
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        Call mfrmFeeGroupCollectFee.ChargeRollingListShow(Me, EM_С������, Val(.TextMatrix(.RowSel, .ColIndex("ID"))))
    End With
End Sub

Private Sub vsRollingCurtainHistory_GotFocus()
    Call zl_VsGridGotFocus(vsRollingCurtainHistory)
    With vsRollingCurtainHistory
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsRollingCurtainHistory_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsRollingCurtainHistory)
End Sub

Private Sub imgColPlanHistory_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanHistory.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlanHistory.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRollingCurtainHistory, lngLeft, lngTop, imgColPlanHistory.Height)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtainHistory, Me.Caption, "��ʷ������Ϣ", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlanHistory_Click()
    Call imgColPlanHistory_Click
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCollectorDetail, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsCollectorDetail, Me.Caption, "��ʷ�շ�Ա������ϸ", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub

Private Sub imgColPlanGroup_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanGroup.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlanGroup.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGroupCollectInfo, lngLeft, lngTop, imgColPlanGroup.Height)
    zl_vsGrid_Para_Save mlngModule, vsGroupCollectInfo, Me.Caption, "С���տ���ϸ", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlanGroup_Click()
    Call imgColPlanGroup_Click
End Sub
